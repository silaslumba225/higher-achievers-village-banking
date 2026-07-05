class WelfareIntelligenceService:

    def __init__(
        self,
        balance=0,
        total_contributions=0,
        total_paid=0,
        pending_claims=0,
        approved_claims=0,
        paid_claims=0,
        this_month_contributions=0,
        this_month_claims_paid=0
    ):

        self.balance = balance
        self.total_contributions = total_contributions
        self.total_paid = total_paid
        self.pending_claims = pending_claims
        self.approved_claims = approved_claims
        self.paid_claims = paid_claims
        self.this_month_contributions = this_month_contributions
        self.this_month_claims_paid = this_month_claims_paid

    # -----------------------------------------

    def build_health(self):

        score = 100

        if self.pending_claims > 0:
            score -= min(self.pending_claims * 5, 20)

        if self.balance <= 0:
            score -= 40

        elif self.balance < self.total_paid * 0.20:
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
            "welfare_health_score": score,
            "welfare_health_status": status,
            "welfare_health_colour": colour
        }

    # -----------------------------------------

    def build_assistant(self):

        messages = []

        if self.pending_claims > 0:

            messages.append({

                "level": "warning",

                "icon": "fa-heart",

                "title": "Pending welfare claims",

                "message":
                    f"{self.pending_claims} claim(s) require committee review."

            })

        if self.balance <= 0:

            messages.append({

                "level": "danger",

                "icon": "fa-wallet",

                "title": "Fund depleted",

                "message":
                    "No welfare funds are currently available."

            })

        if self.this_month_contributions > self.this_month_claims_paid:

            messages.append({

                "level": "success",

                "icon": "fa-chart-line",

                "title": "Fund growing",

                "message":
                    "Monthly contributions currently exceed welfare payments."

            })

        if not messages:

            messages.append({

                "level": "success",

                "icon": "fa-circle-check",

                "title": "Everything looks good",

                "message":
                    "The Welfare Fund is operating normally."

            })

        return {

            "welfare_assistant": messages

        }

    # -----------------------------------------

    def build(self):

        data = {}

        data.update(self.build_health())

        data.update(self.build_assistant())

        return data