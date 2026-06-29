import re
from datetime import datetime, date
import pdfplumber


def parse_fnb_statement(file):
    transactions = []

    with pdfplumber.open(file) as pdf:
        text = ""

        for page in pdf.pages:
            text += "\n" + (page.extract_text() or "")

    statement_year = date.today().year

    year_match = re.search(
        r"Statement Period\s*:\s*\d{1,2}\s+\w+\s+(\d{4})",
        text
    )

    if year_match:
        statement_year = int(year_match.group(1))

    pattern = re.compile(
        r"^(\d{2}\s(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec))\s+(.+?)\s+([\d,]+\.\d{2})(Cr)?\s+([\d,]+\.\d{2})Cr"
    )

    for raw_line in text.splitlines():
        line = raw_line.strip()
        match = pattern.match(line)

        if not match:
            continue

        date_text = match.group(1)
        description = match.group(2).strip()
        amount = match.group(3).replace(",", "")
        is_credit = bool(match.group(4))

        statement_date = datetime.strptime(
            f"{date_text} {statement_year}",
            "%d %b %Y"
        ).date()

        transactions.append({
            "statement_date": statement_date,
            "description": description,
            "reference": "",
            "amount": amount,
            "entry_type": "In" if is_credit else "Out",
        })

    return transactions