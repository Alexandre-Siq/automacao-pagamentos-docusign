import os
import re
import json
from copy import copy
from datetime import datetime
import openpyxl
from openpyxl.styles import Font

def sanitizar_nome(texto):
    """Remove caracteres inválidos para nomes de pastas e arquivos."""
    if not texto:
        return "OUTROS"
    texto_limpo = re.sub(r'[\\/*?:"<>|]', "", str(texto)).strip()
    return texto_limpo if texto_limpo else "OUTROS"

def formatar_moeda(valor_float):
    """Garante que a Capa tem o visual bonito da moeda brasileira"""
    return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def obter_regras_fornecedor(nome_fornecedor, valor_float):
    nome_upper = str(nome_fornecedor).upper()
    
    regra = {
        "nome_oficial": nome_fornecedor,
        "descricao": "Serviços prestados conforme nota fiscal.",
        "contrato": "",
        "boleto": "sim",
        "valor_igual": "sim",
        "contrato_vigente": "sim",
        "tem_rateio": False,
        "template_file": "Template_PADRAO.xlsx",
        "rateio_coords": {},
        "pasta": sanitizar_nome(nome_fornecedor)
    }
    
    if "CLARO" in nome_upper:
        regra["nome_oficial"] = "Claro S/A"
        regra["pasta"] = "CLARO"
        regra["tem_rateio"] = True
        regra["rateio_coords"] = {
            "fornecedor": "B6", "nf": "B7", "emissao": "F7", 
            "valor": "H7", "periodo": "B8"
        }
        if valor_float <= 500:
            regra["descricao"] = "Conta referente ao pacote Claro Internet Empresa."
            regra["template_file"] = "Template_CLARO_MODENS.xlsx"
        else:
            regra["descricao"] = "Conta telefônica de 29 celulares corporativos."
            regra["template_file"] = "Template_CLARO_CELULARES.xlsx"
            
    elif "VIVO" in nome_upper or "TELEFÔNICA" in nome_upper or "TELEFONICA" in nome_upper:
        regra["nome_oficial"] = "Telefônica Brasil S.A."
        regra["pasta"] = "VIVO"
        regra["descricao"] = "Conta telefônica do terminal 13 3228-3000."
        regra["valor_igual"] = ""
        regra["contrato_vigente"] = ""
        regra["tem_rateio"] = True
        regra["template_file"] = "Template_VIVO.xlsx"
        regra["rateio_coords"] = {
            "fornecedor": "B6", "nf": "B7", "emissao": "F7", 
            "valor": "H7", "periodo": "B8"
        }
        
    elif "INFOCOPPY" in nome_upper or "SOUZA" in nome_upper or "SANTANA" in nome_upper:
        regra["nome_oficial"] = "Souza & Santana Suprimentos e Soluções Técnicas Ltda – ME"
        regra["pasta"] = "INFOCOPPY"
        regra["descricao"] = "48 Impressoras/ Multifuncionais Brother, 5 Impressoras Coloridas Epson, 31 Impressoras Zebras, 10 Multifuncionais Lexmark, 1 Multifuncional Elgin, 2 Scanners Espon/Avision."
        regra["contrato"] = "S052/2022"
        regra["contrato_vigente"] = "nao"
        regra["tem_rateio"] = True
        regra["template_file"] = "Template_INFOCOPPY.xlsx"
        regra["rateio_coords"] = {
            "fornecedor": "B4", "nf": "B5", "emissao": "F5", 
            "valor": "H5", "periodo": "B6"
        }
        
    elif "PCTEC" in nome_upper:
        regra["nome_oficial"] = "PCTEC OUTSOURCING LTDA"
        regra["pasta"] = "PCTEC"
        regra["descricao"] = "Locação de notebooks"
        regra["boleto"] = "nao"
        regra["valor_igual"] = ""
        regra["contrato_vigente"] = ""
        regra["tem_rateio"] = True
        regra["template_file"] = "Template_PCTEC.xlsx"
        regra["rateio_coords"] = {
            "fornecedor": "B6", "nf": "B7", "emissao": "F7", 
            "valor": "H7", "periodo": "F28"
        }

    elif "WTT" in nome_upper:
        regra["nome_oficial"] = "WTT TECNOLOGIA E CONSULTORIA"
        regra["pasta"] = "WTT"
        regra["descricao"] = "Licença de uso, suporte e manutenção do software D-Server, WTT Print e CAP to PACS."
        regra["contrato"] = "S059/2022"
        regra["contrato_vigente"] = "nao"
        
    elif "TASY" in nome_upper or "TUCANO" in nome_upper:
        regra["nome_oficial"] = "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA"
        regra["pasta"] = "TASY"
        regra["descricao"] = "Licença de uso, suporte e manutenção do software Tasy."
        
    return regra

def formatar_checkbox(status):
    if status == "sim": return "(X) Sim", "( ) Não"
    elif status == "nao": return "( ) Sim", "(X) Não"
    else: return "( ) Sim", "( ) Não"

def limpar_valor(valor_str):
    if not valor_str: return 0.0
    v = str(valor_str).upper().replace('R$', '').strip()
    v = v.replace('.', '').replace(',', '.') 
    try: return float(v)
    except ValueError: return 0.0

def escrever_celula(ws, coordenada, valor, aplicar_estilo=True):
    fonte_destaque = Font(name='Verdana', size=13, bold=True)
    for intervalo in ws.merged_cells.ranges:
        if coordenada in intervalo:
            celula_principal = ws.cell(row=intervalo.min_row, column=intervalo.min_col)
            celula_principal.value = valor
            
            if aplicar_estilo:
                for linha in range(intervalo.min_row, intervalo.max_row + 1):
                    for coluna in range(intervalo.min_col, intervalo.max_col + 1):
                        ws.cell(row=linha, column=coluna).font = fonte_destaque
            return
            
    ws[coordenada].value = valor
    if aplicar_estilo:
        ws[coordenada].font = fonte_destaque

def encontrar_linha_assinatura(ws):
    for linha in range(1, 150):
        valor_celula = str(ws[f'A{linha}'].value).strip()
        if "Elaborado por:" in valor_celula:
            return linha
    return 69

def salvar_na_planilha(dados_json_string, pasta_templates, caminho_saida):
    dados = json.loads(dados_json_string)
    valor_str = dados.get("valor_total", "")
    
    valor_float = limpar_valor(valor_str)
    valor_formatado = formatar_moeda(valor_float)
    
    fornecedor_extraido = dados.get("fornecedor_nome", "")
    regras = obter_regras_fornecedor(fornecedor_extraido, valor_float)
    caminho_template = os.path.join(pasta_templates, regras["template_file"])
    
    if not os.path.exists(caminho_template):
        raise FileNotFoundError(f"Ficheiro de template não encontrado: {caminho_template}")

    wb = openpyxl.load_workbook(caminho_template)
    ws_capa = wb.worksheets[0]
    
    agora = datetime.now()
    data_hoje = agora.strftime("%d/%m/%Y")
    meses = {1:"jan", 2:"fev", 3:"mar", 4:"abr", 5:"mai", 6:"jun", 7:"jul", 8:"ago", 9:"set", 10:"out", 11:"nov", 12:"dez"}
    competencia = f"{meses[agora.month]}/{str(agora.year)[-2:]}"

    # Preenchimento da Capa (mantém o estilo padrão da Capa)
    escrever_celula(ws_capa, 'B7', regras["nome_oficial"], True)
    escrever_celula(ws_capa, 'G7', regras["contrato"], True)
    escrever_celula(ws_capa, 'B9', dados.get("numero_nf", ""), True)
    escrever_celula(ws_capa, 'D9', dados.get("data_vencimento", ""), True)
    escrever_celula(ws_capa, 'E9', valor_formatado, True) 
    escrever_celula(ws_capa, 'H9', valor_formatado, True)
    escrever_celula(ws_capa, 'B18', regras["descricao"], True)
    escrever_celula(ws_capa, 'J4', data_hoje, True)
    escrever_celula(ws_capa, 'D31', competencia, True)
    
    bol_sim, bol_nao = formatar_checkbox(regras["boleto"])
    escrever_celula(ws_capa, 'B22', f"Boleto  {bol_sim}")
    escrever_celula(ws_capa, 'D22', f"{bol_nao}  - Preencher os dados bancários")

    val_sim, val_nao = formatar_checkbox(regras["valor_igual"])
    escrever_celula(ws_capa, 'B33', f"Valor Faturado é Igual ao Contratado?      {val_sim}          {val_nao}")

    vig_sim, vig_nao = formatar_checkbox(regras["contrato_vigente"])
    escrever_celula(ws_capa, 'B34', f"Contrato e/ou Termo Aditivo Vigente?      {vig_sim}          {vig_nao}")

    aba_rateio_existe = any(aba.lower() == "rateio" for aba in wb.sheetnames)

    if regras["tem_rateio"] and aba_rateio_existe:
        for aba in wb.sheetnames:
            if aba.lower() == "rateio":
                ws_rateio = wb[aba]
                break
                
        coords = regras["rateio_coords"]
        
        # Preenchimento do cabeçalho do Rateio (preservando o padrão nativo do template)
        escrever_celula(ws_rateio, coords["fornecedor"], regras["nome_oficial"], False)
        escrever_celula(ws_rateio, coords["nf"], dados.get("numero_nf", ""), False)
        
        data_emissao = dados.get("data_emissao") or dados.get("data_vencimento", "")
        escrever_celula(ws_rateio, coords["emissao"], data_emissao, False)
        
        escrever_celula(ws_rateio, coords["valor"], valor_float, False)
        escrever_celula(ws_rateio, coords["periodo"], competencia, False)
        
        # --- PRESERVAÇÃO NATIVA DA LINHA DE ASSINATURA ---
        linha_assinatura = encontrar_linha_assinatura(ws_rateio)
        
        # 1. Escreve o Nome no campo B e HERDA a fonte original do rótulo A ("Elaborado por:")
        celula_nome = ws_rateio[f'B{linha_assinatura}']
        celula_nome.value = "Alexandre Siqueira Souza Costa"
        if ws_rateio[f'A{linha_assinatura}'].font:
            celula_nome.font = copy(ws_rateio[f'A{linha_assinatura}'].font)

        # 2. Localiza a coluna da "Data:" e herda a fonte nativa para o valor da Data
        coluna_data_rotulo = 'G'
        for col in ['E', 'F', 'G']:
            if "Data:" in str(ws_rateio[f'{col}{linha_assinatura}'].value):
                coluna_data_rotulo = col
                break
                
        celula_data = ws_rateio[f'H{linha_assinatura}']
        celula_data.value = data_hoje
        if ws_rateio[f'{coluna_data_rotulo}{linha_assinatura}'].font:
            celula_data.font = copy(ws_rateio[f'{coluna_data_rotulo}{linha_assinatura}'].font)
        
    elif not regras["tem_rateio"] and aba_rateio_existe:
        for aba in wb.sheetnames:
            if aba.lower() == "rateio":
                del wb[aba]

    wb.save(caminho_saida)
    return caminho_saida