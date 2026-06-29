from .fnb_importer import parse_fnb_statement


def import_bank_statement(file, bank_name="FNB"):
    if bank_name == "FNB":
        return parse_fnb_statement(file)

    raise ValueError(f"Unsupported bank format: {bank_name}")