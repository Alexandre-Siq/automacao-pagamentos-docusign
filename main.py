from src.modules.batch_processor import processar_lote_faturas


def main():
    """Varre data/input e gera Excel e PDF em data/output/<FORNECEDOR>/."""
    return processar_lote_faturas()


if __name__ == "__main__":
    main()
