import os
import sys
import json

# Ajusta sys.path para reconhecer o pacote 'src'
diretorio_atual = os.path.dirname(os.path.abspath(__file__))
raiz_projeto = os.path.abspath(os.path.join(diretorio_atual, "..", ".."))
if raiz_projeto not in sys.path:
    sys.path.insert(0, raiz_projeto)

from src.modules.gemini_extractor import extrair_dados_pdf
from src.modules.pdf_exporter import converter_excel_para_pdf
from src.modules.sheet_manager import (
    limpar_valor,
    obter_regras_fornecedor,
    salvar_na_planilha,
    sanitizar_nome,
)


def carregar_json_seguro(texto_json):
    """Converte a resposta textual em dicionário Python de forma segura."""
    if not texto_json:
        return None

    texto_limpo = texto_json.strip()
    # Remove marcações do markdown se o modelo retornar ```json ... ```
    if texto_limpo.startswith("```json"):
        texto_limpo = texto_limpo.replace("```json", "", 1)
    if texto_limpo.startswith("```"):
        texto_limpo = texto_limpo.replace("```", "", 1)
    if texto_limpo.endswith("```"):
        texto_limpo = texto_limpo[:-3]

    return json.loads(texto_limpo.strip())


def _resolver_caminho(diretorio_raiz, caminho):
    if os.path.isabs(caminho):
        return caminho
    return os.path.join(diretorio_raiz, caminho)


def _nome_saida(pasta_fornecedor, numero_nf, nome_arquivo, reservados):
    """Separa duas faturas do mesmo lote que compartilham o número da NF.

    Uma nova execução do mesmo arquivo reutiliza o nome e substitui a saída anterior.
    """
    nome_base = f"fatura_NF_{numero_nf}"
    caminho_excel = os.path.join(pasta_fornecedor, f"{nome_base}.xlsx")
    caminho_pdf = os.path.join(pasta_fornecedor, f"{nome_base}.pdf")
    if caminho_excel in reservados or caminho_pdf in reservados:
        stem = sanitizar_nome(os.path.splitext(nome_arquivo)[0])
        nome_base = f"fatura_NF_{numero_nf}_{stem}"
        caminho_excel = os.path.join(pasta_fornecedor, f"{nome_base}.xlsx")
        caminho_pdf = os.path.join(pasta_fornecedor, f"{nome_base}.pdf")
    reservados.add(caminho_excel)
    reservados.add(caminho_pdf)
    return caminho_excel, caminho_pdf


def processar_fatura(caminho_pdf, pasta_templates, pasta_output, reservados=None):
    """Extrai uma fatura e grava Excel e PDF em data/output/<FORNECEDOR>/."""
    json_resposta = extrair_dados_pdf(caminho_pdf)
    dados_extraidos = carregar_json_seguro(json_resposta)
    if not isinstance(dados_extraidos, dict):
        raise ValueError("A extração não retornou um objeto JSON.")

    valor_float = limpar_valor(dados_extraidos.get("valor_total"))
    regras = obter_regras_fornecedor(dados_extraidos.get("fornecedor_nome") or "", valor_float)
    pasta_nome = regras.get("pasta") or "OUTROS"
    numero_nf = sanitizar_nome(dados_extraidos.get("numero_nf") or "SEM_NUMERO")

    pasta_fornecedor = os.path.join(pasta_output, pasta_nome)
    os.makedirs(pasta_fornecedor, exist_ok=True)

    nome_arquivo = os.path.basename(caminho_pdf)
    if reservados is None:
        reservados = set()
    caminho_excel, caminho_pdf_saida = _nome_saida(pasta_fornecedor, numero_nf, nome_arquivo, reservados)

    payload = json.dumps(dados_extraidos, ensure_ascii=False)
    salvar_na_planilha(payload, pasta_templates, caminho_excel)

    erro_pdf = None
    pdf_gerado = None
    try:
        converter_excel_para_pdf(caminho_excel, caminho_pdf_saida)
        pdf_gerado = caminho_pdf_saida
    except Exception as e:
        erro_pdf = f"Falha ao exportar PDF: {e}"

    return {
        "arquivo_origem": nome_arquivo,
        "status": "PARCIAL" if erro_pdf else "SUCESSO",
        "dados": dados_extraidos,
        "excel": caminho_excel,
        "pdf": pdf_gerado,
        "erro": erro_pdf,
    }


def processar_lote_faturas(pasta_input="data/input", pasta_output="data/output", pasta_templates="data/templates"):
    # Obtém caminhos absolutos baseados no diretório do projeto
    diretorio_raiz = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    caminho_input = _resolver_caminho(diretorio_raiz, pasta_input)
    caminho_output = _resolver_caminho(diretorio_raiz, pasta_output)
    caminho_templates = _resolver_caminho(diretorio_raiz, pasta_templates)

    os.makedirs(caminho_input, exist_ok=True)
    os.makedirs(caminho_output, exist_ok=True)

    # Busca todos os arquivos .pdf (independente de maiúsculas/minúsculas no nome)
    arquivos_pdf = sorted(
        f for f in os.listdir(caminho_input)
        if f.lower().endswith(".pdf")
    )

    if not arquivos_pdf:
        print(f"⚠️ Nenhuma fatura PDF encontrada na pasta '{caminho_input}'.")
        return []

    print("==================================================")
    quantidade = len(arquivos_pdf)
    rotulo = "fatura encontrada" if quantidade == 1 else "faturas encontradas"
    print(f"🚀 INICIANDO PROCESSAMENTO EM LOTE ({quantidade} {rotulo})")
    print("==================================================\n")

    resultados = []
    reservados = set()
    sucessos = 0
    parciais = 0
    falhas = 0

    for index, nome_arquivo in enumerate(arquivos_pdf, start=1):
        caminho_completo = os.path.join(caminho_input, nome_arquivo)
        print(f"[{index}/{len(arquivos_pdf)}] Processando: '{nome_arquivo}'...")

        try:
            resultado_item = processar_fatura(
                caminho_completo, caminho_templates, caminho_output, reservados
            )
            resultados.append(resultado_item)

            fornecedor = (resultado_item.get("dados") or {}).get("fornecedor_nome") or "Não identificado"
            valor = (resultado_item.get("dados") or {}).get("valor_total") or "Não identificado"

            if resultado_item["status"] == "SUCESSO":
                sucessos += 1
                print(f"  ✅ Concluído! Fornecedor: {fornecedor} | Valor: {valor}")
                print(f"     Excel: {resultado_item['excel']}")
                print(f"     PDF:   {resultado_item['pdf']}\n")
            else:
                parciais += 1
                print(f"  ⚠️ Excel gerado, PDF pendente. Fornecedor: {fornecedor} | Valor: {valor}")
                print(f"     Excel: {resultado_item['excel']}")
                print(f"     {resultado_item['erro']}\n")

        except Exception as e:
            falhas += 1
            print(f"  ❌ Erro ao processar '{nome_arquivo}': {e}\n")
            resultados.append({
                "arquivo_origem": nome_arquivo,
                "status": "ERRO",
                "dados": None,
                "excel": None,
                "pdf": None,
                "erro": str(e)
            })

    # Salva relatório em JSON na pasta output
    caminho_relatorio = os.path.join(caminho_output, "relatorio_lote.json")
    with open(caminho_relatorio, "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=2)

    print("==================================================")
    print("📊 RESUMO DO PROCESSAMENTO EM LOTE")
    print("==================================================")
    print(f" Total de arquivos: {len(arquivos_pdf)}")
    print(f" Processados com sucesso: {sucessos}")
    print(f" Excel sem PDF: {parciais}")
    print(f" Falhas: {falhas}")
    print(f" Relatório salvo em: {caminho_relatorio}")
    print("==================================================\n")

    return resultados

if __name__ == "__main__":
    processar_lote_faturas()
