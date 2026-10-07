import os
import sys
import json

# Ajusta sys.path para reconhecer o pacote 'src'
diretorio_atual = os.path.dirname(os.path.abspath(__file__))
raiz_projeto = os.path.abspath(os.path.join(diretorio_atual, "..", ".."))
if raiz_projeto not in sys.path:
    sys.path.insert(0, raiz_projeto)

from src.modules.gemini_extractor import extrair_dados_pdf

# ... MANTENHA O RESTANTE DO BATCH_PROCESSOR COMO ESTAVA ...

try:
    from src.modules.gemini_extractor import extrair_dados_pdf
except ModuleNotFoundError:
    from gemini_extractor import extrair_dados_pdf

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

def processar_lote_faturas(pasta_input="data/input", pasta_output="data/output"):
    # Obtém caminhos absolutos baseados no diretório do projeto
    diretorio_raiz = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    caminho_input = os.path.join(diretorio_raiz, pasta_input)
    caminho_output = os.path.join(diretorio_raiz, pasta_output)

    os.makedirs(caminho_input, exist_ok=True)
    os.makedirs(caminho_output, exist_ok=True)

    # Busca todos os arquivos .pdf (independente de maiúsculas/minúsculas no nome)
    arquivos_pdf = [
        f for f in os.listdir(caminho_input)
        if f.lower().endswith(".pdf")
    ]

    if not arquivos_pdf:
        print(f"⚠️ Nenhuma fatura PDF encontrada na pasta '{caminho_input}'.")
        return []

    print("==================================================")
    print(f"🚀 INICIANDO PROCESSAMENTO EM LOTE ({len(arquivos_pdf)} faturas encontradas)")
    print("==================================================\n")

    resultados = []
    sucessos = 0
    falhas = 0

    for index, nome_arquivo in enumerate(arquivos_pdf, start=1):
        caminho_completo = os.path.join(caminho_input, nome_arquivo)
        print(f"[{index}/{len(arquivos_pdf)}] Processando: '{nome_arquivo}'...")

        try:
            # Executa a extração usando o módulo estável
            json_resposta = extrair_dados_pdf(caminho_completo)
            dados_extraidos = carregar_json_seguro(json_resposta)

            fornecedor = dados_extraidos.get("fornecedor_nome") or "Não identificado"
            valor = dados_extraidos.get("valor_total") or "Não identificado"
            
            resultado_item = {
                "arquivo_origem": nome_arquivo,
                "status": "SUCESSO",
                "dados": dados_extraidos,
                "erro": None
            }
            resultados.append(resultado_item)
            sucessos += 1
            
            print(f"  ✅ Concluído! Fornecedor: {fornecedor} | Valor: {valor}\n")

        except Exception as e:
            falhas += 1
            print(f"  ❌ Erro ao processar '{nome_arquivo}': {e}\n")
            resultados.append({
                "arquivo_origem": nome_arquivo,
                "status": "ERRO",
                "dados": None,
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
    print(f" Falhas: {falhas}")
    print(f" Relatório salvo em: {caminho_relatorio}")
    print("==================================================\n")

    return resultados

if __name__ == "__main__":
    processar_lote_faturas()