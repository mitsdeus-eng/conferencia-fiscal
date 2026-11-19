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

# Captura CNPJ com 14 ou 15 dígitos (ex: 012.345.678/0001-99), com ou sem pontos e com espaços do pypdf
CNPJ_FLEX_RE = re.compile(
    r'(?<!\d)(0?\d{2})\s*\.?\s*(\d{3})\s*\.?\s*(\d{3})\s*/\s*(\d{4})\s*[-\s]?\s*(\d{2})(?!\d)'
)

# Captura CNPJ sem barra/traço próximo a rótulos (CNPJ, C.N.P.J., CGC, Inscrição)
CNPJ_ROTULADO_RE = re.compile(
    r'(?:c\.?\s*n\.?\s*p\.?\s*j\.?|c\.?\s*g\.?\s*c\.?|cpf\s*/\s*cnpj|inscri[çc][ãa]o)[^\d]{0,25}(0?\d{14})(?!\d)',
    re.IGNORECASE
)

RAZAO_ROTULADA_RE = re.compile(
    r'(?:empresa|raz[ãa]o\s*social|firma|contribuinte|nome\s*empresarial|estabelecimento)\s*[:\-]\s*([^\n|;]{3,90})',
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


def normalizar_14_digitos(digits: str) -> str:
    if len(digits) == 15 and digits.startswith('0'):
        return digits[1:]
    return digits


def formatar_cnpj(cnpj_digits: str) -> str:
    if len(cnpj_digits) != 14:
        return cnpj_digits
    return f"{cnpj_digits[:2]}.{cnpj_digits[2:5]}.{cnpj_digits[5:8]}/{cnpj_digits[8:12]}-{cnpj_digits[12:]}"


def extrair_texto_pdf_paginas(caminho_temp: str, max_paginas: int = 2) -> str:
    partes = []
    try:
        reader = PdfReader(caminho_temp)
        for page in reader.pages[:max_paginas]:
            txt = ''
            try:
                txt = page.extract_text(extraction_mode="layout") or ''
            except TypeError:
                txt = page.extract_text() or ''
            except Exception:
                txt = page.extract_text() or ''

            if txt.strip():
                linhas_limpas = [re.sub(r'[ \t]{2,}', ' | ', l.strip()) for l in txt.splitlines() if l.strip()]
                partes.append('\n'.join(linhas_limpas))
    except Exception:
        pass
    return '\n'.join(partes)


def extrair_cabecalho_texto(caminho_temp: str, extensao: str, max_linhas: int = 60) -> str:
    linhas = []
    try:
        if extensao == '.pdf':
            return extrair_texto_pdf_paginas(caminho_temp, max_paginas=2)

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
                        for idx, raw_line in enumerate(f):
                            if idx >= max_linhas:
                                break
                            parts = [p.strip() for p in re.split(r'[;,\t]', raw_line) if p.strip()]
                            if parts:
                                linhas.append(' | '.join(parts))
                    break
                except UnicodeDecodeError:
                    continue
    except Exception:
        pass

    return '\n'.join(linhas)


def extrair_identificacao(texto_documento: str) -> dict:
    if not texto_documento:
        return {'cnpj': '', 'cnpj_formatado': '', 'empresa_detectada': ''}

    cnpj_valido = ''
    linha_do_cnpj = ''

    # 1. Busca CNPJs com barra/traço (14 ou 15 dígitos com zero à esquerda, com ou sem pontos/espaços)
    for m in CNPJ_FLEX_RE.finditer(texto_documento):
        bruto = ''.join(m.groups())
        digits = normalizar_14_digitos(bruto)
        if validar_cnpj(digits):
            cnpj_valido = digits
            inicio = texto_documento.rfind('\n', 0, m.start()) + 1
            fim = texto_documento.find('\n', m.end())
            linha_do_cnpj = texto_documento[inicio: fim if fim != -1 else len(texto_documento)]
            break

    # 2. Fallback: CNPJ sem pontuação próximo a rótulo (CNPJ, CGC, Inscrição)
    if not cnpj_valido:
        for m in CNPJ_ROTULADO_RE.finditer(texto_documento):
            digits = normalizar_14_digitos(m.group(1))
            if validar_cnpj(digits):
                cnpj_valido = digits
                inicio = texto_documento.rfind('\n', 0, m.start()) + 1
                fim = texto_documento.find('\n', m.end())
                linha_do_cnpj = texto_documento[inicio: fim if fim != -1 else len(texto_documento)]
                break

    # 3. Extrai Razão Social / Nome da Empresa
    empresa_detectada = ''
    m_rot = RAZAO_ROTULADA_RE.search(texto_documento[:3000])
    if m_rot:
        cand = CNPJ_FLEX_RE.sub('', m_rot.group(1))
        cand = re.sub(r'\b(?:c\.?n\.?p\.?j\.?|c\.?g\.?c\.?|ie|inscri[çc][ãa]o|per[íi]odo|folha|p[áa]gina).*$', '', cand, flags=re.IGNORECASE)
        cand = re.sub(r'\s+', ' ', cand).strip(' -:|')
        if len(cand) >= 3 and not cand.isdigit():
            empresa_detectada = cand.upper()

    if not empresa_detectada:
        for linha in texto_documento.splitlines()[:35]:
            linha_up = linha.upper()
            m_suf = SUFIXO_EMPRESA_RE.search(linha_up)
            if m_suf:
                cand = CNPJ_FLEX_RE.sub('', m_suf.group(1))
                cand = re.sub(r'\s+', ' ', cand).strip(' -:|')
                if len(cand) >= 4:
                    empresa_detectada = cand
                    break

    # 4. Fallback: se achou CNPJ na mesma linha do nome da empresa (ex: "EMPRESA XYZ | CNPJ: 12.345...")
    if not empresa_detectada and linha_do_cnpj:
        cand = CNPJ_FLEX_RE.sub('', linha_do_cnpj)
        cand = re.sub(r'\b(?:empresa|raz[ãa]o\s*social|firma|c\.?n\.?p\.?j\.?|c\.?g\.?c\.?|ie|inscri[çc][ãa]o\s*estadual)\b\s*[:\-]?', '', cand, flags=re.IGNORECASE)
        cand = re.sub(r'\s+', ' ', cand).strip(' -:|')
        if len(cand) >= 4 and not cand.isdigit():
            empresa_detectada = cand.upper()[:70]

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

        cabecalho_texto = extrair_cabecalho_texto(caminho_temp, extensao, max_linhas=60)
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
        cab_preview = '\n'.join(cabecalho_texto.splitlines()[:25])[:1500]

        return jsonify({
            'sucesso': True,
            'arquivo': nome_original,
            'cnpj': identificacao['cnpj'],
            'cnpj_formatado': identificacao['cnpj_formatado'],
            'empresa_detectada': identificacao['empresa_detectada'],
            'cabecalho_texto': cab_preview,
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
