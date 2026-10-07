"""Abre a tela de pagamentos no navegador."""

import os
import traceback

from src.interface_web import avisar_erro, iniciar


def main():
    try:
        iniciar(abrir_navegador=True)
    except Exception:
        detalhe = traceback.format_exc()
        pasta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "output")
        try:
            os.makedirs(pasta, exist_ok=True)
            with open(os.path.join(pasta, "interface.log"), "w", encoding="utf-8") as arquivo:
                arquivo.write(detalhe)
        except OSError:
            pass
        avisar_erro("Não foi possível abrir a tela de pagamentos.\n\n" + detalhe[-800:])


if __name__ == "__main__":
    main()
