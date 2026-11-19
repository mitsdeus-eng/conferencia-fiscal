import csv
import os
import re
import tempfile
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from pypdf import PdfReader
from odf.opendocument import load as load_ods
from odf.table import Table, TableRow, TableCell
from odf.teletype import extractText
from services.file_reader import ler_arquivo_tabular, FileReadError

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

CNPJ_FORMATADO_RE = re.compile(r'\b(\d{2}\.?\d{3}\.?\d{3}\s*/\s*\d{4}\s*-\s*\d{2})\b')
CNPJ_ROTULADO_RE = re.compile(r'(?:cnpj|cpf/cnpj|inscri[çc][ãa]o)[^\d]{0,18}(\d{14})\b', re.IGNORECASE)
RAZAO_ROTULADA_RE = re.compile(
    r'(?:empresa|raz[ãa]o\s*social|contribuinte|nome\s*empresarial|estabelecimento)\s*[:\-]\s*([^\n|;]{3,90})',
    re.IGNORECASE
)
SUFIXO_EMPRESA_RE = re.compile(
    r'\b([A-Z0-9À-Ú][A-Z0-9À-Ú\s&.,\-]{2,70}\b(?:LTDA|EIRELI|S/A|S\.A\.|ME|EPP|COMERCIO|COMÉRCIO|INDUSTRIA|INDÚSTRIA|SERVICOS|SERVIÇOS|DISTRIBUIDORA))\b'
)


def validar_cnpj(cnpj_digits: str) -> bool:
    if len(cnpj_digits) != 14 or len(set(cnpj_digits)) == 1:
        return False

    pesos_1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    soma_1 = sum(int(d) * p for d, p in zip(cnpj_digits[:12], pesos_1))
    resto_1 = soma_1 % 11
    dig_1 = '0' if resto_1 < 2 else str(11 - resto_1)

    pesos_2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    soma_2 = sum(int(d) * p for d, p in zip(cnpj_digits[:13], pesos_2))
    resto_2 = soma_2 % 11
    dig_2 = '0' if resto_2 < 2 else str(11 - resto_2)

    return cnpj_digits[-2:] == dig_1 + dig_2


def formatar_cnpj(cnpj_digits: str) -> str:
    if len(cnpj_digits) != 14:
        return cnpj_digits
    return f"{cnpj_digits[:2]}.{cnpj_digits[2:5]}.{cnpj_digits[5:8]}/{cnpj_digits[8:12]}-{cnpj_digits[12:]}"


def extrair_cabecalho_texto(caminho_temp: str, extensao: str, max_linhas: int = 40) -> str:
    linhas = []
    try:
        if extensao == '.pdf':
            reader = PdfReader(caminho_temp)
            if reader.pages:
                texto_p1 = reader.pages[0].extract_text() or ''
                linhas = [l.strip() for l in texto_p1.splitlines() if l.strip()][:max_linhas]

        elif extensao == '.xlsx':
            wb = load_workbook(caminho_temp, read_only=True, data_only=True)
            try:
                ws = wb.active
                for row in ws.iter_rows(max_row=max_linhas, values_only=True):
                    vals = [str(c).strip() for c in row if c is not None and str(c).strip()]
                    if vals:
                        linhas.append(' | '.join(vals))
            finally:
                wb.close()

        elif extensao == '.ods':
            doc = load_ods(caminho_temp)
            tabelas = doc.getElementsByType(Table)
            if tabelas:
                rows = tabelas[0].getElementsByType(TableRow)[:max_linhas]
                for row in rows:
                    vals = []
                    for cell in row.getElementsByType(TableCell):
                        txt = extractText(cell).strip()
                        if txt:
                            vals.append(txt)
                    if vals:
                        linhas.append(' | '.join(vals))

        elif extensao == '.csv':
            for enc in ('utf-8-sig', 'latin-1'):
                try:
                    with open(caminho_temp, 'r', encoding=enc, errors='strict') as f:
                        reader = csv.reader(f, delimiter=';')
                        for idx, row in enumerate(reader):
                            if idx >= max_linhas:
                                break
                            vals = [c.strip() for c in row if c and c.strip()]
                            if vals:
                                linhas.append(' | '.join(vals))
                    break
                except UnicodeDecodeError:
                    continue
    except Exception:
        pass

    return '\n'.join(linhas)


def extrair_identificacao(cabecalho_texto: str) -> dict:
    if not cabecalho_texto:
        return {'cnpj': '', 'cnpj_formatado': '', 'empresa_detectada': ''}

    cnpj_valido = ''

    for match in CNPJ_FORMATADO_RE.findall(cabecalho_texto):
        digits = re.sub(r'\D', '', match)
        if validar_cnpj(digits):
            cnpj_valido = digits
            break

    if not cnpj_valido:
        for match in CNPJ_ROTULADO_RE.findall(cabecalho_texto):
            if validar_cnpj(match):
                cnpj_valido = match
                break

    empresa_detectada = ''
    m_rot = RAZAO_ROTULADA_RE.search(cabecalho_texto)
    if m_rot:
        cand = CNPJ_FORMATADO_RE.sub('', m_rot.group(1))
        cand = re.sub(r'\b(?:cnpj|ie|inscri[çc][ãa]o|per[íi]odo|folha|p[áa]gina).*$', '', cand, flags=re.IGNORECASE)
        cand = re.sub(r'\s+', ' ', cand).strip(' -:|')
        if len(cand) >= 3 and not cand.isdigit():
            empresa_detectada = cand.upper()

    if not empresa_detectada:
        for linha in cabecalho_texto.splitlines()[:20]:
            linha_up = linha.upper()
            m_suf = SUFIXO_EMPRESA_RE.search(linha_up)
            if m_suf:
                cand = CNPJ_FORMATADO_RE.sub('', m_suf.group(1))
                cand = re.sub(r'\s+', ' ', cand).strip(' -:|')
                if len(cand) >= 4:
                    empresa_detectada = cand
                    break

    return {
        'cnpj': cnpj_valido,
        'cnpj_formatado': formatar_cnpj(cnpj_valido) if cnpj_valido else '',
        'empresa_detectada': empresa_detectada
    }


def parse_valor_br(valor_raw) -> float:
    if isinstance(valor_raw, (int, float)):
        return float(valor_raw)
    if valor_raw is None:
        return 0.0
    texto = str(valor_raw).replace('R$', '').strip()
    if not texto:
        return 0.0
    texto = re.sub(r'[^\d,.-]', '', texto)
    if not texto:
        return 0.0
    if ',' in texto and '.' in texto:
        if texto.rfind(',') > texto.rfind('.'):
            texto = texto.replace('.', '').replace(',', '.')
        else:
            texto = texto.replace(',', '')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return 0.0


def limpar_cfop(valor_raw) -> str:
    if valor_raw is None:
        return ''
    digits = re.sub(r'\D', '', str(valor_raw))
    return digits if len(digits) == 4 else ''


def limpar_nota(valor_raw) -> str:
    if valor_raw is None:
        return ''
    digits = re.sub(r'\D', '', str(valor_raw))
    if not digits:
        return ''
    return digits.lstrip('0') or '0'


def detectar_mapeamento_colunas(colunas: list[str]) -> dict:
    cols_lower = {c: c.strip().lower() for c in colunas}

    termos_imposto = [
        'icms', 'ipi', 'pis', 'cofins', 'ibs', 'cbs', 'iss', 'is ',
        'st', 'subst', 'alíquota', 'aliquota', 'aliq', 'base', 'bc ',
        'isentas', 'outras', 'retido', 'diferido', 'crédito', 'credito'
    ]

    nomes_prioritarios_valor = [
        'valor contábil', 'valor contabil', 'vlr. contábil', 'vlr contábil', 'vlr contabil',
        'valor total', 'vlr. total', 'vlr total', 'total da nota', 'total nota',
        'valor da nota', 'valor nota', 'valor líquido', 'valor liquido',
        'valor dos produtos', 'valor produtos', 'valor mercadoria', 'valor merc.', 'valor'
    ]

    col_cfop = None
    for orig, low in cols_lower.items():
        if 'cfop' in low or low in ('código', 'codigo', 'cod. fiscal', 'cod fiscal'):
            col_cfop = orig
            break

    col_nota = None
    for orig, low in cols_lower.items():
        if any(t in low for t in ['valor', 'total', 'vlr', 'data', 'serie', 'série', 'chave']):
            continue
        if any(k in low for k in ['nota', 'nf', 'nº', 'numero', 'número', 'doc', 'documento']):
            col_nota = orig
            break

    col_data = None
    for orig, low in cols_lower.items():
        if any(k in low for k in ['data', 'emissão', 'emissao', 'entrada', 'dt.']):
            col_data = orig
            break

    col_valor = None
    for prio in nomes_prioritarios_valor:
        for orig, low in cols_lower.items():
            if any(imp in low for imp in termos_imposto):
                continue
            if prio in low:
                col_valor = orig
                break
        if col_valor:
            break

    return {
        'cfop': col_cfop,
        'nota': col_nota,
        'data': col_data,
        'valor': col_valor
    }


def normalizar_registros(registros_brutos: list[dict], extensao: str, modo_operacao: str) -> list[dict]:
    if not registros_brutos:
        return []

    if extensao == '.pdf':
        saida = []
        for item in registros_brutos:
            cfop = limpar_cfop(item.get('cfop', ''))
            if not cfop:
                continue
            nota = 'RESUMO' if modo_operacao == 'resumo_cfop' else limpar_nota(item.get('nota', ''))
            if modo_operacao == 'notas' and not nota:
                continue
            saida.append({
                'data': str(item.get('data', '-')).strip() or '-',
                'nota': nota,
                'cfop': cfop,
                'valor': round(parse_valor_br(item.get('valor', 0)), 2)
            })
        return saida

    colunas = list(registros_brutos[0].keys())
    mapa = detectar_mapeamento_colunas(colunas)

    if not mapa['cfop']:
        raise FileReadError(f"Coluna de CFOP não encontrada. Colunas: {', '.join(colunas)}")
    if not mapa['valor']:
        raise FileReadError(f"Coluna de Valor Contábil/Total não encontrada. Colunas: {', '.join(colunas)}")
    if modo_operacao == 'notas' and not mapa['nota']:
        raise FileReadError(f"Coluna de Número da Nota não encontrada. Colunas: {', '.join(colunas)}")

    saida = []
    for row in registros_brutos:
        cfop = limpar_cfop(row.get(mapa['cfop'], ''))
        if not cfop:
            continue

        if modo_operacao == 'resumo_cfop':
            nota = 'RESUMO'
        else:
            nota = limpar_nota(row.get(mapa['nota'], ''))
            if not nota:
                continue

        data_str = str(row.get(mapa['data'], '-')).strip() if mapa['data'] else '-'
        if not data_str:
            data_str = '-'

        valor_num = round(parse_valor_br(row.get(mapa['valor'], 0)), 2)
        saida.append({
            'data': data_str,
            'nota': nota,
            'cfop': cfop,
            'valor': valor_num
        })

    return saida


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/api/extrair', methods=['POST'])
def api_extrair():
    if 'arquivo' not in request.files:
        return jsonify({'sucesso': False, 'erro': 'Nenhum arquivo enviado.'}), 400

    arquivo = request.files['arquivo']
    if not arquivo or not arquivo.filename:
        return jsonify({'sucesso': False, 'erro': 'Arquivo inválido.'}), 400

    modo_operacao = request.form.get('modo_operacao', 'notas')
    nome_original = arquivo.filename
    extensao = os.path.splitext(nome_original)[1].lower()
    nome_seguro = secure_filename(nome_original) or f"upload{extensao}"

    fd, caminho_temp = tempfile.mkstemp(suffix=f"_{nome_seguro}")
    os.close(fd)

    try:
        arquivo.save(caminho_temp)

        cabecalho_texto = extrair_cabecalho_texto(caminho_temp, extensao, max_linhas=40)
        identificacao = extrair_identificacao(cabecalho_texto)

        resultado_leitura = ler_arquivo_tabular(caminho_temp, modo_operacao=modo_operacao)

        if isinstance(resultado_leitura, dict):
            registros_brutos = resultado_leitura.get('registros', [])
            cab_extra = resultado_leitura.get('cabecalho_texto', '')
            if cab_extra and not identificacao['cnpj']:
                cabecalho_texto = cab_extra
                identificacao = extrair_identificacao(cabecalho_texto)
        else:
            registros_brutos = resultado_leitura

        registros_limpos = normalizar_registros(registros_brutos, extensao, modo_operacao)

        return jsonify({
            'sucesso': True,
            'arquivo': nome_original,
            'cnpj': identificacao['cnpj'],
            'cnpj_formatado': identificacao['cnpj_formatado'],
            'empresa_detectada': identificacao['empresa_detectada'],
            'cabecalho_texto': cabecalho_texto[:1500],
            'registros': registros_limpos
        })
    except FileReadError as e:
        return jsonify({
            'sucesso': False,
            'arquivo': nome_original,
            'erro': str(e)
        }), 400
    except Exception as e:
        return jsonify({
            'sucesso': False,
            'arquivo': nome_original,
            'erro': f"Erro ao processar {nome_original}: {str(e)}"
        }), 500
    finally:
        if os.path.exists(caminho_temp):
            try:
                os.remove(caminho_temp)
            except OSError:
                pass


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
