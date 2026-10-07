import os
import sys
import tempfile
import unittest

import pymupdf

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.modules.pdf_processor import ANCORA_DOCUSIGN, preparar_pdf_docusign


def _gravar(caminho, texto):
    documento = pymupdf.open()
    pagina = documento.new_page()
    pagina.insert_text((72, 72), texto)
    documento.save(caminho)
    documento.close()


class TestPacoteDocuSign(unittest.TestCase):
    def test_junta_fatura_e_solicitacao_e_carimba_a_ultima_pagina(self):
        with tempfile.TemporaryDirectory() as pasta:
            fatura = os.path.join(pasta, "fatura.pdf")
            solicitacao = os.path.join(pasta, "solicitacao.pdf")
            saida = os.path.join(pasta, "pacote.pdf")
            _gravar(fatura, "Nota fiscal")
            documento = pymupdf.open()
            documento.new_page()
            documento.new_page()
            documento.save(solicitacao)
            documento.close()

            preparar_pdf_docusign([fatura, solicitacao], saida)

            pacote = pymupdf.open(saida)
            try:
                self.assertEqual(pacote.page_count, 3)
                self.assertIn("Nota fiscal", pacote[0].get_text())
                self.assertNotIn(ANCORA_DOCUSIGN, pacote[0].get_text())
                self.assertIn(ANCORA_DOCUSIGN, pacote[-1].get_text())
            finally:
                pacote.close()

    def test_ignora_arquivo_ausente_e_falha_se_nada_restar(self):
        with tempfile.TemporaryDirectory() as pasta:
            with self.assertRaises(Exception) as ctx:
                preparar_pdf_docusign(
                    [os.path.join(pasta, "nao-existe.pdf")],
                    os.path.join(pasta, "vazio.pdf"),
                )
            self.assertIn("Nenhum PDF válido", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
