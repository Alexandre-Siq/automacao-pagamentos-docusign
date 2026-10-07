import os

import pymupdf

ANCORA_DOCUSIGN = "\\assinatura_dir_fin\\"


def preparar_pdf_docusign(lista_caminhos_pdf, caminho_saida, texto_ancora=ANCORA_DOCUSIGN):
    """Junta os PDFs e carimba a âncora branca do DocuSign na última página."""
    pdf_final = pymupdf.open()
    try:
        print("Mesclando documentos...")
        for caminho in lista_caminhos_pdf:
            if not caminho or not os.path.exists(caminho):
                print(f"Aviso: Arquivo não encontrado e será ignorado: {caminho}")
                continue
            try:
                pdf_atual = pymupdf.open(caminho)
            except Exception as e:
                print(f"Aviso: PDF ilegível e será ignorado: {caminho} ({e})")
                continue
            if pdf_atual.page_count == 0:
                pdf_atual.close()
                continue
            pdf_final.insert_pdf(pdf_atual)
            pdf_atual.close()

        if pdf_final.page_count == 0:
            raise Exception("Nenhum PDF válido foi encontrado para mesclar.")

        # A solicitação fica por último, então a âncora cai na página que será assinada.
        ultima_pagina = pdf_final[-1]
        retangulo_pagina = ultima_pagina.rect
        x = retangulo_pagina.width - 180
        y = retangulo_pagina.height - 50

        print(f"Carimbando âncora do DocuSign ({texto_ancora}) na última página...")
        ultima_pagina.insert_text(
            pymupdf.Point(x, y),
            texto_ancora,
            fontsize=8,
            color=(1, 1, 1),
        )

        pasta_saida = os.path.dirname(os.path.abspath(caminho_saida))
        if pasta_saida:
            os.makedirs(pasta_saida, exist_ok=True)
        pdf_final.save(caminho_saida)
    finally:
        pdf_final.close()

    print(f"PDF pronto para o DocuSign gerado em: {caminho_saida}")
    return caminho_saida


if __name__ == "__main__":
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_teste = os.path.abspath(os.path.join(diretorio_atual, "..", "..", "data", "input", "teste_pagamento.pdf"))
    caminho_saida = os.path.abspath(os.path.join(diretorio_atual, "..", "..", "data", "output", "documento_docusign_pronto.pdf"))

    try:
        preparar_pdf_docusign([caminho_teste, caminho_teste], caminho_saida)
    except Exception as e:
        print(f"Erro: {e}")
