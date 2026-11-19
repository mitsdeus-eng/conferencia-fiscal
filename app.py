import os
import re
import tempfile
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from services.file_reader import ler_arquivo_tabular, FileReadError

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024


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
        registros_brutos = ler_arquivo_tabular(caminho_temp, modo_operacao=modo_operacao)
        registros_limpos = normalizar_registros(registros_brutos, extensao, modo_operacao)

        return jsonify({
            'sucesso': True,
            'arquivo': nome_original,
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
