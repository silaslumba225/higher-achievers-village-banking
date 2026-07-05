from decimal import Decimal


class FinancialIntelligenceService:

    def __init__(
        self,
        cash_total=Decimal("0.00"),
        income_total=Decimal("0.00"),
        expense_total=Decimal("0.00"),
        asset_total=Decimal("0.00"),
        liability_total=Decimal("0.00"),
        equity_total=Decimal("0.00"),
        total_debits=Decimal("0.00"),
        total_credits=Decimal("0.00")
    ):
        self.cash_total = cash_total
        self.income_total = income_total
        self.expense_total = expense_total
        self.asset_total = asset_total
        self.liability_total = liability_total
        self.equity_total = equity_total
        self.total_debits = total_debits
        self.total_credits = total_credits

    def build_health(self):
        score = 100

        if self.cash_total <= Decimal("0.00"):
            score -= 35

        if self.expense_total > self.income_total:
            score -= 25

        if self.asset_total < self.liability_total:
            score -= 20

        if self.total_debits != self.total_credits:
            score -= 20

        score = max(score, 0)

        if score >= 90:
            status = "Excellent"
            colour = "success"
        elif score >= 75:
            status = "Healthy"
            colour = "primary"
        elif score >= 60:
            status = "Needs Attention"
            colour = "warning"
        else:
            status = "High Risk"
            colour = "danger"

        return {
            "financial_health_score": score,
            "financial_health_status": status,
            "financial_health_colour": colour
        }

    def build_position(self):
        net_worth = self.asset_total - self.liability_total

        if self.liability_total == Decimal("0.00"):
            asset_cover = "Excellent"
            asset_cover_colour = "success"
        elif self.asset_total >= self.liability_total:
            asset_cover = "Healthy"
            asset_cover_colour = "primary"
        else:
            asset_cover = "High Risk"
            asset_cover_colour = "danger"

        return {
            "net_worth": net_worth,
            "asset_cover": asset_cover,
            "asset_cover_colour": asset_cover_colour
        }

    def build_integrity(self):
        checks = []

        if self.total_debits == self.total_credits:
            checks.append("General Ledger Balanced")

        if self.cash_total >= Decimal("0.00"):
            checks.append("Cash Accounts Verified")

        if self.asset_total >= Decimal("0.00"):
            checks.append("Assets Verified")

        if self.liability_total >= Decimal("0.00"):
            checks.append("Liabilities Verified")

        integrity_score = len(checks) * 25

        return {
            "integrity_checks": checks,
            "integrity_score": integrity_score
        }

    def build_liquidity(self):
        if self.expense_total > Decimal("0.00"):
            months = self.cash_total / self.expense_total
        else:
            months = Decimal("99")

        if months >= Decimal("3"):
            status = "Excellent"
            colour = "success"
        elif months >= Decimal("1"):
            status = "Healthy"
            colour = "primary"
        else:
            status = "Low"
            colour = "danger"

        return {
            "cash_months": round(float(months), 1),
            "liquidity_status": status,
            "liquidity_colour": colour
        }

    def build_ratios(self):
        if self.income_total > Decimal("0.00"):
            surplus_margin = (
                (self.income_total - self.expense_total)
                / self.income_total
            ) * Decimal("100")
        else:
            surplus_margin = Decimal("0.00")

        return {
            "surplus_margin": round(float(surplus_margin), 1)
        }

    def build_assistant(self):
        assistant = []

        if self.cash_total <= Decimal("0.00"):
            assistant.append({
                "level": "danger",
                "icon": "fa-wallet",
                "title": "Cash reserves exhausted",
                "message": "The organisation currently has no available cash."
            })
        else:
            assistant.append({
                "level": "success",
                "icon": "fa-coins",
                "title": "Cash position healthy",
                "message": "Cash reserves are available for normal operations."
            })

        if self.expense_total > self.income_total:
            assistant.append({
                "level": "warning",
                "icon": "fa-chart-line",
                "title": "Expenses exceed income",
                "message": "Operating costs are higher than income. Monitor spending carefully."
            })
        else:
            assistant.append({
                "level": "success",
                "icon": "fa-arrow-trend-up",
                "title": "Income exceeds expenditure",
                "message": "The organisation is operating with a positive surplus."
            })

        if self.total_debits != self.total_credits:
            assistant.append({
                "level": "danger",
                "icon": "fa-scale-balanced",
                "title": "Ledger imbalance detected",
                "message": "The General Ledger is out of balance and should be investigated immediately."
            })
        else:
            assistant.append({
                "level": "success",
                "icon": "fa-scale-balanced",
                "title": "Ledger balanced",
                "message": "Debits and credits agree."
            })

        return {
            "financial_assistant": assistant
        }

    def build(self):
        data = {}

        data.update(self.build_health())
        data.update(self.build_position())
        data.update(self.build_integrity())
        data.update(self.build_liquidity())
        data.update(self.build_ratios())
        data.update(self.build_assistant())

        return data