from datetime import date


class LoanIntelligenceService:
    def __init__(
        self,
        total_loans=0,
        applied_loans=0,
        approved_loans=0,
        disbursed_loans=0,
        paid_loans=0,
        overdue_loans=0,
        portfolio_balance=0,
        active_loans=None
    ):
        self.total_loans = total_loans
        self.applied_loans = applied_loans
        self.approved_loans = approved_loans
        self.disbursed_loans = disbursed_loans
        self.paid_loans = paid_loans
        self.overdue_loans = overdue_loans
        self.portfolio_balance = portfolio_balance
        self.active_loans = active_loans or []

    # ----------------------------------------------------
    # Portfolio Health
    # ----------------------------------------------------

    def build_portfolio_health(self):

        score = 100

        if self.overdue_loans > 0:
            score -= min(self.overdue_loans * 5, 30)

        if self.disbursed_loans > 0:

            repayment_ratio = self.paid_loans / (
                self.disbursed_loans + self.paid_loans
            )

            if repayment_ratio < 0.30:
                score -= 20

            elif repayment_ratio < 0.60:
                score -= 10

        score = max(score, 0)

        if score >= 90:
            status = "Excellent"
            colour = "success"

        elif score >= 75:
            status = "Good"
            colour = "primary"

        elif score >= 60:
            status = "Needs Attention"
            colour = "warning"

        else:
            status = "High Risk"
            colour = "danger"

        return {
            "portfolio_health_score": score,
            "portfolio_health_status": status,
            "portfolio_health_colour": colour,
        }

    # ----------------------------------------------------
    # Loan Assistant
    # ----------------------------------------------------

    def build_assistant(self):

        assistant = []

        if self.overdue_loans > 0:
            assistant.append({
                "level": "danger",
                "icon": "fa-clock",
                "title": "Follow up overdue loans",
                "message":
                    f"{self.overdue_loans} overdue loan(s) require immediate attention."
            })

        if self.approved_loans > 0:
            assistant.append({
                "level": "watch",
                "icon": "fa-circle-check",
                "title": "Approved loans awaiting disbursement",
                "message":
                    f"{self.approved_loans} approved loan(s) are ready for disbursement."
            })

        if self.applied_loans > 0:
            assistant.append({
                "level": "primary",
                "icon": "fa-file-signature",
                "title": "Loan applications awaiting review",
                "message":
                    f"{self.applied_loans} application(s) need committee review."
            })

        if self.portfolio_balance > 0:
            assistant.append({
                "level": "good",
                "icon": "fa-chart-line",
                "title": "Outstanding portfolio",
                "message":
                    "Continue monitoring repayments to maintain a healthy portfolio."
            })

        if not assistant:
            assistant.append({
                "level": "good",
                "icon": "fa-circle-check",
                "title": "Portfolio looks healthy",
                "message":
                    "There are no urgent loan issues requiring attention today."
            })

        return {
            "loan_assistant": assistant
        }

    # ----------------------------------------------------
    # Risk Monitor
    # ----------------------------------------------------

    def build_risk_monitor(self):

        risk = []

        for loan in self.active_loans:

            if getattr(loan, "overdue", False):

                level = "High"
                colour = "danger"

            elif loan.balance > 0:

                level = "Medium"
                colour = "warning"

            else:

                level = "Low"
                colour = "success"

            risk.append({

                "loan_no":
                    loan.loan_no or f"LN{loan.id:04d}",

                "member":
                    loan.member.full_name,

                "level":
                    level,

                "colour":
                    colour

            })

        return {

            "loan_risk_monitor":
                risk[:10]

        }

    # ----------------------------------------------------
    # Collection Calendar
    # ----------------------------------------------------

    def build_collection_calendar(self):

        today = date.today()

        calendar = []

        for loan in self.active_loans:

            if (
                loan.balance > 0
                and loan.due_on
                and loan.due_on >= today
            ):

                days = (loan.due_on - today).days

                if days <= 7:

                    calendar.append({

                        "member":
                            loan.member.full_name,

                        "loan_no":
                            loan.loan_no or f"LN{loan.id:04d}",

                        "due_on":
                            loan.due_on,

                        "days":
                            days

                    })

        calendar.sort(key=lambda x: x["due_on"])

        return {

            "collection_calendar":
                calendar[:10]

        }

    # ----------------------------------------------------
    # Build Everything
    # ----------------------------------------------------

    def build(self):

        data = {}

        data.update(self.build_portfolio_health())

        data.update(self.build_assistant())

        data.update(self.build_risk_monitor())

        data.update(self.build_collection_calendar())

        return data