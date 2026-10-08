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
    (r"\bINFOCOPP?Y\b", "INFOCOPPY"),
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
    r"VALOR\s+A\s+PAGAR",
    r"VALOR\s+TOTAL\s+DA\s+NOTA",
    r"VALOR\s+TOTAL\s+DO\s+SERVICO",
    r"VALOR\s+TOTAL\s+DOS\s+SERVICOS",
    r"VALOR\s+TOTAL\s+DA\s+NFS[-\s]?E",
    r"VALOR\s+TOTAL\s+DA\s+FATURA",
    r"VALOR\s+LIQUIDO\s+DA\s+NOTA",
    r"VALOR\s+LIQUIDO\s+A\s+PAGAR",
    r"VALOR\s+LIQUIDO",
    r"VALOR\s+DOS\s+SERVICOS",
    r"VALOR\s+DO\s+SERVICO",
    r"TOTAL\s+DA\s+NFS[-\s]?E",
    r"TOTAL\s+DA\s+NOTA",
    r"VALOR\s+TOTAL",
)

ROTULOS_NUMERO = (
    r"NUMERO\s+DA\s+NOTA(?:\s+FISCAL)?",
    r"NUMERO\s+DA\s+NFS[-\s]?E",
    r"NUMERO\s+DA\s+NF[-\s]?E",
    r"NUMERO\s+DA\s+FATURA",
    r"N(?:O|\.)?\s+DA\s+NOTA(?:\s+FISCAL)?",
    r"N(?:O|\.)?\s+DA\s+FATURA",
    r"N(?:O|\.)?\s+FATURA",
    r"NFS[-\s]?E\s+N(?:O|\.|UMERO)?",
    r"NF[-\s]?E\s+N(?:O|\.|UMERO)?",
    r"NOTA\s+FISCAL\s+N(?:O|\.)?",
    r"FATURA\s+N(?:O|\.)?",
)

_ocr_engine = None

CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
DATA_RE = re.compile(r"\b(\d{2})[/-](\d{2})[/-](\d{4})\b")
DINHEIRO_RE = re.compile(
    r"(?:R\$|RS)\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})"
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
    """Lê o PDF no computador: texto selecionável primeiro, imagem em seguida."""
    textos = list(_textos_digitais(caminho_pdf))
    if not textos:
        ocr = _texto_ocr(caminho_pdf)
        if ocr:
            textos.append(ocr)
    if not textos:
        return None
    return max(textos, key=len)


def extrair_dados_pdf(caminho_pdf):
    """Lê a fatura no próprio Python e devolve o JSON usado pela planilha."""
    if not os.path.exists(caminho_pdf):
        raise FileNotFoundError(f"Ficheiro PDF não encontrado em: {caminho_pdf}")

    print("-> Lendo a fatura localmente, sem enviar ao Gemini...")
    candidatos = []
    for texto in _textos_digitais(caminho_pdf):
        candidatos.append(analisar_texto_fatura(texto))

    melhor = max(candidatos, key=_pontuacao) if candidatos else {}
    if not _completo(melhor):
        print("-> Sem texto selecionável útil. Lendo a imagem do PDF no computador...")
        texto_ocr = _texto_ocr(caminho_pdf)
        if texto_ocr:
            candidatos.append(analisar_texto_fatura(texto_ocr))

    if not candidatos:
        raise ValueError(
            "O PDF não tem texto selecionável e a leitura da imagem também não "
            "encontrou conteúdo. Use um PDF mais nítido, gerado pelo sistema."
        )

    dicas = _dicas_arquivo(caminho_pdf)
    for dados in candidatos:
        _completar_com_dicas(dados, dicas)

    dados = max(candidatos, key=_pontuacao)
    if not dados.get("fornecedor_nome") or not dados.get("valor_total"):
        raise ValueError(
            "Não foi possível identificar o fornecedor e o valor no PDF. "
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


def _completo(dados):
    return bool(dados.get("fornecedor_nome") and dados.get("valor_total"))


def _textos_digitais(caminho_pdf):
    textos = []
    for leitor in (_texto_pypdf, _texto_pymupdf):
        try:
            texto = leitor(caminho_pdf)
        except Exception:
            texto = None
        if texto and texto.strip():
            textos.append(texto.strip())
    return textos


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
        partes = []
        for pagina in documento:
            texto = pagina.get_text("text") or ""
            if not texto.strip():
                palavras = pagina.get_text("words") or []
                texto = " ".join(item[4] for item in palavras)
            partes.append(texto)
        return "\n".join(partes)
    finally:
        documento.close()


def _obter_ocr():
    global _ocr_engine
    if _ocr_engine is False:
        return None
    if _ocr_engine is not None:
        return _ocr_engine
    try:
        from rapidocr import RapidOCR

        _ocr_engine = RapidOCR()
        return _ocr_engine
    except Exception as erro:
        print(f"   [Aviso] OCR indisponível: {erro}")
        _ocr_engine = False
        return None


def _texto_ocr(caminho_pdf):
    """Rasteriza as páginas e lê a imagem no computador, sem enviar o arquivo."""
    engine = _obter_ocr()
    if engine is None:
        return None
    import pymupdf

    documento = pymupdf.open(caminho_pdf)
    partes = []
    try:
        for indice, pagina in enumerate(documento):
            if indice >= 4:
                break
            maior = max(pagina.rect.width, pagina.rect.height) or 1
            zoom = 2.0 if maior <= 1600 else 1600 / maior
            pix = pagina.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            try:
                resultado = engine(pix.tobytes("png"))
            except Exception as erro:
                print(f"   [Aviso] Falha no OCR da página {indice + 1}: {erro}")
                continue
            texto = _texto_do_ocr(resultado)
            if texto:
                partes.append(texto)
    finally:
        documento.close()
    texto = "\n".join(partes).strip()
    return texto or None


def _texto_do_ocr(resultado):
    if resultado is None:
        return ""
    txts = getattr(resultado, "txts", None)
    boxes = getattr(resultado, "boxes", None)
    if txts:
        linhas = []
        for indice, txt in enumerate(txts):
            box = boxes[indice] if boxes is not None else None
            y, x = _origem_caixa(box, indice)
            linhas.append((y, x, str(txt)))
        return _juntar_linhas(ocr_linhas=linhas)
    if isinstance(resultado, (tuple, list)) and resultado and resultado[0]:
        linhas = []
        for indice, item in enumerate(resultado[0]):
            box = item[0] if isinstance(item, (list, tuple)) else None
            txt = item[1] if isinstance(item, (list, tuple)) and len(item) > 1 else ""
            y, x = _origem_caixa(box, indice)
            linhas.append((y, x, str(txt)))
        return _juntar_linhas(ocr_linhas=linhas)
    return ""


def _origem_caixa(box, indice):
    try:
        ponto = box[0]
        return float(ponto[1]), float(ponto[0])
    except (TypeError, IndexError, ValueError):
        return float(indice), 0.0


def _juntar_linhas(ocr_linhas):
    if not ocr_linhas:
        return ""
    ocr_linhas = sorted(ocr_linhas, key=lambda item: (item[0], item[1]))
    grupos = []
    for y, x, txt in ocr_linhas:
        if not txt or not str(txt).strip():
            continue
        if not grupos or abs(y - grupos[-1][0][0]) > 14:
            grupos.append([(y, x, str(txt).strip())])
        else:
            grupos[-1].append((y, x, str(txt).strip()))
    saidas = []
    for grupo in grupos:
        grupo.sort(key=lambda item: item[1])
        saidas.append(" ".join(item[2] for item in grupo))
    return "\n".join(saidas)


def _dicas_arquivo(caminho_pdf):
    """Nome do arquivo ajuda quando a imagem só entrega parte dos dados."""
    nome = os.path.splitext(os.path.basename(caminho_pdf or ""))[0]
    busca = _normalizar(re.sub(r"[_\-]+", " ", nome))
    dicas = {}
    for padrao, fornecedor in FORNECEDORES:
        if re.search(padrao, busca):
            dicas["fornecedor_nome"] = fornecedor
            break
    if re.search(r"\bNFS?\s*E\b|\bNF\b", busca):
        numeros = re.findall(r"\d{3,12}", nome)
        if numeros:
            escolhido = max(numeros, key=len).lstrip("0") or numeros[0]
            dicas["numero_nf"] = escolhido
    return dicas


def _completar_com_dicas(dados, dicas):
    for campo, valor in (dicas or {}).items():
        if valor and not dados.get(campo):
            dados[campo] = valor
    return dados


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
