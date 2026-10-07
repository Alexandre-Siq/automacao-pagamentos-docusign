import json
import os
import re
import unicodedata


CAMPOS = (
    "cnpj_emissor",
    "valor_total",
    "data_vencimento",
    "data_emissao",
    "fornecedor_nome",
    "numero_nf",
    "linha_digitavel",
)

# Nomes oficiais já reconhecidos pelas regras da planilha.
FORNECEDORES = (
    (r"\bINFOCOPPY\b", "INFOCOPPY"),
    (r"SOUZA\s*(?:&|E)\s*SANTANA", "Souza & Santana Suprimentos e Soluções Técnicas Ltda – ME"),
    (r"\bPCTEC\b", "PCTEC OUTSOURCING LTDA"),
    (r"\bWTT\b", "WTT TECNOLOGIA E CONSULTORIA"),
    (r"\bTASY\b|\bTUCANO\b", "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA"),
    (r"\bTELEFONICA\b|\bVIVO\b", "Telefônica Brasil S.A."),
    (r"\bCLARO\b", "Claro S/A"),
)

# Raiz do CNPJ (8 primeiros dígitos) quando o nome vem só como logotipo.
RAIZES_CNPJ = {
    "40432544": "Claro S/A",
    "02558157": "Telefônica Brasil S.A.",
}

ROTULOS_VALOR = (
    r"VALOR\s+DO\s+DOCUMENTO",
    r"VALOR\s+COBRADO",
    r"TOTAL\s+A\s+PAGAR",
    r"VALOR\s+TOTAL\s+DA\s+NOTA",
    r"VALOR\s+TOTAL\s+DO\s+SERVICO",
    r"VALOR\s+TOTAL\s+DOS\s+SERVICOS",
    r"VALOR\s+TOTAL\s+DA\s+NFS-?E",
    r"VALOR\s+TOTAL\s+DA\s+FATURA",
    r"VALOR\s+LIQUIDO\s+DA\s+NOTA",
    r"VALOR\s+LIQUIDO",
    r"VALOR\s+DOS\s+SERVICOS",
    r"VALOR\s+TOTAL",
)

ROTULOS_NUMERO = (
    r"NUMERO\s+DA\s+NOTA(?:\s+FISCAL)?",
    r"NUMERO\s+DA\s+NFS-?E",
    r"NUMERO\s+DA\s+NF-?E",
    r"NUMERO\s+DA\s+FATURA",
    r"N(?:O|\.)?\s+DA\s+NOTA(?:\s+FISCAL)?",
    r"N(?:O|\.)?\s+DA\s+FATURA",
    r"N(?:O|\.)?\s+FATURA",
    r"NFS-?E\s+N(?:O|\.)?",
    r"NF-?E\s+N(?:O|\.)?",
    r"NOTA\s+FISCAL\s+N(?:O|\.)?",
    r"FATURA\s+N(?:O|\.)?",
)

CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
DATA_RE = re.compile(r"\b(\d{2})[/-](\d{2})[/-](\d{4})\b")
DINHEIRO_RE = re.compile(
    r"R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})"
    r"|"
    r"(?<!\d)(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?!\d)"
)
LINHA_FORMATADA_RE = re.compile(
    r"(\d{5})\.(\d{5})\s+(\d{5})\.(\d{6})\s+(\d{5})\.(\d{6})\s+(\d)\s+(\d{14})"
)
LINHA_47_RE = re.compile(r"(?<!\d)(\d{47})(?!\d)")
LINHA_ARRECADACAO_RE = re.compile(
    r"(\d{11})-?(\d)\s+(\d{11})-?(\d)\s+(\d{11})-?(\d)\s+(\d{11})-?(\d)"
)


def extrair_texto_pdf_local(caminho_pdf):
    """Lê a camada de texto do PDF. Sem texto, o arquivo é imagem e não entra na fila de API."""
    textos = []
    for leitor in (_texto_pypdf, _texto_pymupdf):
        try:
            texto = leitor(caminho_pdf)
        except Exception:
            texto = None
        if texto and texto.strip():
            textos.append(texto.strip())
    if not textos:
        return None
    return max(textos, key=len)


def extrair_dados_pdf(caminho_pdf):
    """Lê a fatura no próprio Python e devolve o JSON usado pela planilha."""
    if not os.path.exists(caminho_pdf):
        raise FileNotFoundError(f"Ficheiro PDF não encontrado em: {caminho_pdf}")

    print("-> Lendo a fatura localmente, sem enviar ao Gemini...")
    candidatos = []
    for leitor in (_texto_pypdf, _texto_pymupdf):
        try:
            texto = leitor(caminho_pdf)
        except Exception:
            texto = None
        if not texto or not texto.strip():
            continue
        dados = analisar_texto_fatura(texto)
        candidatos.append(dados)

    if not candidatos:
        raise ValueError(
            "O PDF não tem texto selecionável. A leitura local não envia o arquivo ao Gemini."
        )

    dados = max(candidatos, key=_pontuacao)
    if not dados.get("fornecedor_nome") or not dados.get("valor_total"):
        raise ValueError(
            "Não foi possível identificar o fornecedor e o valor no texto do PDF. "
            f"Encontrado: fornecedor={dados.get('fornecedor_nome') or '—'}, "
            f"valor={dados.get('valor_total') or '—'}, "
            f"nf={dados.get('numero_nf') or '—'}."
        )
    return json.dumps(dados, ensure_ascii=False)


def analisar_texto_fatura(texto):
    """Interpreta o texto de uma NFS-e, fatura ou boleto e devolve os campos da planilha."""
    busca = _normalizar(texto)
    plano = re.sub(r"\s+", " ", busca)
    cnpj = _extrair_cnpj(busca)
    linha = _extrair_linha_digitavel(plano)
    dados = {
        "cnpj_emissor": cnpj,
        "valor_total": _extrair_valor(plano) or _valor_da_linha(linha),
        "data_vencimento": _extrair_data(plano, (r"DATA\s+DE\s+VENCIMENTO", r"\bVENCIMENTO\b", r"\bVENCTO\b")),
        "data_emissao": _extrair_data(plano, (r"DATA\s+(?:E\s+HORA\s+)?DE\s+EMISSAO", r"\bEMISSAO\b")),
        "fornecedor_nome": _extrair_fornecedor(busca, cnpj),
        "numero_nf": _extrair_numero(plano),
        "linha_digitavel": linha,
    }
    return dados


def _normalizar(texto):
    texto = texto.replace("\xa0", " ").replace("\u200b", "")
    texto = "".join(
        caractere for caractere in unicodedata.normalize("NFD", texto)
        if unicodedata.category(caractere) != "Mn"
    )
    texto = texto.replace("º", " ").replace("°", " ").replace("ª", " ")
    texto = texto.upper()
    texto = texto.replace("\r", "\n")
    texto = re.sub(r"[ \t]+", " ", texto)
    return texto


def _pontuacao(dados):
    pesos = {
        "fornecedor_nome": 4,
        "valor_total": 4,
        "numero_nf": 2,
        "data_vencimento": 1,
        "data_emissao": 1,
        "cnpj_emissor": 1,
        "linha_digitavel": 1,
    }
    return sum(pesos[campo] for campo in CAMPOS if dados.get(campo))


def _texto_pypdf(caminho_pdf):
    import pypdf

    reader = pypdf.PdfReader(caminho_pdf)
    partes = []
    for pagina in reader.pages:
        extraido = pagina.extract_text()
        if extraido:
            partes.append(extraido)
    return "\n".join(partes)


def _texto_pymupdf(caminho_pdf):
    import pymupdf

    documento = pymupdf.open(caminho_pdf)
    try:
        return "\n".join(pagina.get_text("text") for pagina in documento)
    finally:
        documento.close()


def _formatar_cnpj(bruto):
    digitos = re.sub(r"\D", "", bruto)
    if len(digitos) != 14:
        return None
    return f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:]}"


def _primeiro_cnpj(trecho):
    for bruto in CNPJ_RE.findall(trecho or ""):
        formatado = _formatar_cnpj(bruto)
        if formatado:
            return formatado
    return None


def _extrair_cnpj(busca):
    bloco = re.search(r"\b(?:PRESTADOR|EMITENTE|EMISSOR)\b(.{0,500})", busca, re.S)
    if bloco:
        encontrado = _primeiro_cnpj(bloco.group(1))
        if encontrado:
            return encontrado
    corte = re.search(r"\b(?:DESTINATARIO|TOMADOR|SACADO)\b", busca)
    if corte:
        encontrado = _primeiro_cnpj(busca[:corte.start()])
        if encontrado:
            return encontrado
    return _primeiro_cnpj(busca)


def _extrair_fornecedor(busca, cnpj):
    cabecalho = busca[:800]
    universo = cabecalho if any(re.search(padrao, cabecalho) for padrao, _ in FORNECEDORES) else busca
    achados = []
    for padrao, nome in FORNECEDORES:
        encontrado = re.search(padrao, universo)
        if encontrado:
            achados.append((encontrado.start(), nome))
    if achados:
        achados.sort(key=lambda item: item[0])
        return achados[0][1]

    if cnpj:
        raiz = re.sub(r"\D", "", cnpj)[:8]
        if raiz in RAIZES_CNPJ:
            return RAIZES_CNPJ[raiz]

    bloco = re.search(r"\b(?:PRESTADOR|EMITENTE|EMISSOR)\b(.{0,400})", busca, re.S)
    if not bloco:
        return None
    for linha in bloco.group(1).splitlines():
        linha = linha.strip(" -:")
        if re.search(r"\b(?:LTDA|S\.?A\.?|S/A|EIRELI|ME)\b", linha) and "CNPJ" not in linha:
            return linha
    return None


def _formatar_valor(bruto):
    limpo = bruto.replace(" ", "").replace(".", "")
    if "," not in limpo:
        return None
    inteiro, decimal = limpo.split(",", 1)
    if not inteiro.isdigit() or not decimal.isdigit():
        return None
    if int(inteiro) == 0 and int(decimal) == 0:
        return None
    inteiro = str(int(inteiro))
    grupos = []
    while inteiro:
        grupos.append(inteiro[-3:])
        inteiro = inteiro[:-3]
    return ".".join(reversed(grupos)) + "," + decimal


def _valor_na_janela(trecho):
    for encontrado in DINHEIRO_RE.finditer(trecho):
        bruto = encontrado.group(1) or encontrado.group(2)
        formatado = _formatar_valor(bruto)
        if formatado and formatado != "0,00":
            return formatado
    return None


def _extrair_valor(plano):
    for rotulo in ROTULOS_VALOR:
        ultimo = None
        for encontrado in re.finditer(rotulo, plano):
            valor = _valor_na_janela(plano[encontrado.end():encontrado.end() + 120])
            if valor:
                ultimo = valor
        if ultimo:
            return ultimo
    return None


def _extrair_linha_digitavel(plano):
    encontrada = LINHA_FORMATADA_RE.search(plano)
    if encontrada:
        grupos = encontrada.groups()
        return (
            f"{grupos[0]}.{grupos[1]} {grupos[2]}.{grupos[3]} "
            f"{grupos[4]}.{grupos[5]} {grupos[6]} {grupos[7]}"
        )
    continua = LINHA_47_RE.search(plano)
    if continua:
        digitos = continua.group(1)
        return (
            f"{digitos[0:5]}.{digitos[5:10]} {digitos[10:15]}.{digitos[15:21]} "
            f"{digitos[21:26]}.{digitos[26:32]} {digitos[32]} {digitos[33:47]}"
        )
    arrecadacao = LINHA_ARRECADACAO_RE.search(plano)
    if arrecadacao:
        grupos = arrecadacao.groups()
        return " ".join(f"{grupos[i]}-{grupos[i + 1]}" for i in range(0, 8, 2))
    return None


def _valor_da_linha(linha):
    if not linha:
        return None
    digitos = re.sub(r"\D", "", linha)
    if len(digitos) != 47:
        return None
    centavos = int(digitos[-10:])
    if centavos <= 0:
        return None
    inteiro, decimal = divmod(centavos, 100)
    return _formatar_valor(f"{inteiro},{decimal:02d}")


def _data_valida(dia, mes, ano):
    if not (1 <= mes <= 12 and 1 <= dia <= 31):
        return False
    if mes in (4, 6, 9, 11) and dia > 30:
        return False
    if mes == 2 and dia > 29:
        return False
    return 1990 <= ano <= 2100


def _extrair_data(plano, rotulos):
    for rotulo in rotulos:
        for encontrado in re.finditer(rotulo, plano):
            janela_depois = plano[encontrado.end():encontrado.end() + 80]
            data = _primeira_data(janela_depois)
            if data:
                return data
            janela_antes = plano[max(0, encontrado.start() - 40):encontrado.start()]
            data = _primeira_data(janela_antes)
            if data:
                return data
    return None


def _primeira_data(trecho):
    for encontrado in DATA_RE.finditer(trecho):
        dia, mes, ano = (int(encontrado.group(i)) for i in (1, 2, 3))
        if _data_valida(dia, mes, ano):
            return f"{dia:02d}/{mes:02d}/{ano}"
    return None


def _extrair_numero(plano):
    for rotulo in ROTULOS_NUMERO:
        encontrado = re.search(rotulo + r"\s*[:\-]?\s*(\d{1,3}(?:\.\d{3})+|\d{1,12})\b", plano)
        if not encontrado:
            continue
        numero = encontrado.group(1).replace(".", "")
        if numero.strip("0"):
            return numero
    return None


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Uso: python -m src.modules.extrator_pdf <arquivo.pdf>")
        raise SystemExit(1)
    print(extrair_dados_pdf(sys.argv[1]))
