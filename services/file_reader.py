import os
import re
import csv
import openpyxl
from pypdf import PdfReader

try:
    from odf.opendocument import load as load_ods
    from odf.table import Table, TableRow, TableCell
    from odf.text import P
except ImportError:
    load_ods = None


class FileReadError(Exception):
    pass


def _extrair_linhas_ods(caminho_arquivo: str) -> list[list[str]]:
    if load_ods is None:
        raise FileReadError("Biblioteca odfpy não instalada no servidor para ler arquivos .ods.")
    doc = load_ods(caminho_arquivo)
    tabelas = doc.spreadsheet.getElementsByType(Table)
    if not tabelas:
        return []
    primeira_aba = tabelas[0]
    linhas_resultado = []

    for row in primeira_aba.getElementsByType(TableRow):
        valores_linha = []
        for cell in row.getElementsByType(TableCell):
            repeticao = int(cell.getAttribute("numbercolumnsrepeated") or 1)
            paragrafos = cell.getElementsByType(P)
            texto_celula = " ".join(
                "".join(n.data for n in p.childNodes if hasattr(n, "data"))
                for p in paragrafos
            ).strip()
            if not texto_celula:
                val_attr = cell.getAttribute("value")
                if val_attr is not None:
                    texto_celula = str(val_attr).strip()
            if repeticao > 50 and not texto_celula:
                continue
            valores_linha.extend([texto_celula] * min(repeticao, 50))
        if any(valores_linha):
            linhas_resultado.append(valores_linha)
    return linhas_resultado


def _processar_matriz_planilha(linhas: list) -> list[dict]:
    if len(linhas) < 2:
        raise FileReadError("A planilha está vazia ou sem dados.")

    idx_cabecalho = 0
    palavras_chave_cab = [
        'cfop', 'código', 'codigo', 'nota', 'nf', 'valor', 'total',
        'contabil', 'contábil', 'data', 'emissao', 'emissão'
    ]
    for idx, row in enumerate(linhas[:25]):
        texto_linha = " ".join([str(c).lower() for c in row if c is not None])
        if sum(1 for p in palavras_chave_cab if p in texto_linha) >= 2 or any(
            k in texto_linha for k in ['cfop', 'código', 'codigo']
        ):
            idx_cabecalho = idx
            break

    cabecalho = [
        str(c).strip() if c is not None and str(c).strip() != '' else f"col_{i}"
        for i, c in enumerate(linhas[idx_cabecalho])
    ]
    registros = []
    for row in linhas[idx_cabecalho + 1:]:
        if not any(row):
            continue
        reg = {}
        for i in range(len(cabecalho)):
            val = row[i] if i < len(row) and row[i] is not None else ''
            reg[cabecalho[i]] = str(val).strip()
        registros.append(reg)
    return registros


def ler_arquivo_tabular(caminho_arquivo: str, modo_operacao: str = 'resumo_cfop') -> list[dict]:
    if not os.path.exists(caminho_arquivo):
        raise FileReadError(f"Arquivo não encontrado: {caminho_arquivo}")

    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    registros = []

    try:
        if extensao == '.xlsx':
            wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
            ws = wb.active
            linhas = list(ws.iter_rows(values_only=True))
            registros = _processar_matriz_planilha(linhas)

        elif extensao == '.ods':
            linhas = _extrair_linhas_ods(caminho_arquivo)
            registros = _processar_matriz_planilha(linhas)

        elif extensao == '.csv':
            conteudo = None
            for encoding in ['utf-8-sig', 'utf-8', 'latin-1']:
                try:
                    with open(caminho_arquivo, 'r', encoding=encoding) as f:
                        conteudo = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            if not conteudo:
                raise FileReadError("Não foi possível decodificar o arquivo CSV.")

            sep = ';' if conteudo.count(';') >= conteudo.count(',') else ','
            linhas_csv = [l for l in conteudo.splitlines() if l.strip()]

            idx_cab = 0
            for idx, l in enumerate(linhas_csv[:25]):
                if any(p in l.lower() for p in ['cfop', 'código', 'codigo', 'nota', 'valor', 'total', 'contabil', 'contábil']):
                    idx_cab = idx
                    break

            leitor = csv.DictReader(linhas_csv[idx_cab:], delimiter=sep)
            for row in leitor:
                if any(row.values()):
                    registros.append({str(k).strip(): (str(v).strip() if v is not None else '') for k, v in row.items() if k})

        elif extensao == '.pdf':
            reader = PdfReader(caminho_arquivo)

            if modo_operacao == 'resumo_cfop':
                paginas_alvo = reader.pages[-4:] if len(reader.pages) > 4 else reader.pages
                texto_final = "\n".join([p.extract_text() or '' for p in paginas_alvo])

                pos_inicio = -1
                for marcador in ['DEMONSTRATIVO POR CFOP', 'DEMONSTRATIVO', 'COMPARATIVO POR CFOP']:
                    idx_m = texto_final.upper().rfind(marcador)
                    if idx_m > pos_inicio:
                        pos_inicio = idx_m

                if pos_inicio == -1:
                    linhas_ultimas = [l.strip() for l in texto_final.splitlines() if l.strip()]
                    raise FileReadError(f"Quadro 'Demonstrativo por CFOP' não encontrado. Últimas linhas: {' | '.join(linhas_ultimas[-30:])}")

                trecho_quadro = texto_final[pos_inicio:]
                linhas_quadro = trecho_quadro.splitlines()

                padrao_linha_cfop = re.compile(
                    r'(\d{1,3}(?:\.\d{3})*,\d{2})\s+([123567][\.\-]\d{3})\b'
                )

                cfops_encontrados = {}
                for linha in linhas_quadro:
                    linha_limpa = linha.strip()
                    upper = linha_limpa.upper()
                    if any(t in upper for t in ['TOTAL', 'SUBTOTAL', 'INSCRI', 'CNPJ', 'FOLHA', 'PÁGINA', 'PAGINA']):
                        continue

                    m = padrao_linha_cfop.search(linha_limpa)
                    if m:
                        valor_str, cfop_raw = m.groups()
                        cfop_limpo = re.sub(r'\D', '', cfop_raw)
                        if cfop_limpo.startswith('202'):
                            continue
                        cfops_encontrados[cfop_limpo] = {
                            "cfop": cfop_limpo,
                            "valor": valor_str
                        }

                if not cfops_encontrados:
                    amostra = " | ".join([l.strip() for l in linhas_quadro if l.strip()][:30])
                    raise FileReadError(f"Diagnóstico Demonstrativo por CFOP: [{amostra[:600]}]")

                registros = list(cfops_encontrados.values())

            else:
                texto_bruto = " ".join([p.extract_text() or '' for p in reader.pages])
                texto_limpo = re.sub(r'\s+', ' ', texto_bruto)

                # Captura exata:
                # 01/09/2026Nota 1 08601/09/2026 MT 56.690,00 2-918 2 40.000,0034408
                padrao_modelo_p1 = re.compile(
                    r'(\d{2}/\d{2}/\d{4})\s*'                     # 1. Data de entrada
                    r'(?:(?!\d{2}/\d{2}/\d{4}).){0,120}?'         # Espécie / Série / Código emitente
                    r'(\d{2}/\d{2}/\d{4})\s*'                     # 2. Data de emissão
                    r'[A-Z]{2}\s*'                                # UF
                    r'(\d{1,3}(?:\.\d{3})*,\d{2})\s+'             # 3. Valor Contábil total da nota
                    r'([123567][\.\-]\d{3})\s+'                   # 4. CFOP
                    r'[123]\s+'                                   # Código tributação (1, 2 ou 3)
                    r'\d{1,3}(?:\.\d{3})*,\d{2}\s*'               # 1º valor tributado (logo antes da nota)
                    r'(\d{1,9})'                                  # 5. Número da Nota grudado logo após os centavos
                    r'(?=\s|$)'
                )

                for match in padrao_modelo_p1.finditer(texto_limpo):
                    data_ent, data_emi, valor_str, cfop_raw, nota_raw = match.groups()
                    nota_limpa = nota_raw.lstrip('0') or '0'
                    cfop_limpo = re.sub(r'\D', '', cfop_raw)

                    registros.append({
                        "data": data_ent,
                        "nota": nota_limpa,
                        "cfop": cfop_limpo,
                        "valor": valor_str
                    })

                if not registros:
                    amostra = texto_limpo[:500]
                    raise FileReadError(f"Nenhuma nota capturada no PDF. Amostra: {amostra}")
        else:
            raise FileReadError(f"Formato não suportado ({extensao}). Envie .xlsx, .ods, .csv ou .pdf.")

        if not registros:
            raise FileReadError("O arquivo enviado está vazio ou não possui linhas válidas.")

        return registros

    except FileReadError:
        raise
    except Exception as e:
        raise FileReadError(f"Falha ao ler o arquivo '{os.path.basename(caminho_arquivo)}': {str(e)}")
