class DashboardService:
    def __init__(self, total_cash, overdue_loans, pending_welfare_claims):
        self.total_cash = total_cash
        self.overdue_loans = overdue_loans
        self.pending_welfare_claims = pending_welfare_claims

    def build_group_health(self):
        group_health = "Healthy"
        group_health_colour = "good"
        group_health_message = "Your group is doing well today."

        if self.overdue_loans > 0:
            group_health = "Needs Attention"
            group_health_colour = "watch"
            group_health_message = f"{self.overdue_loans} loan(s) need follow-up."

        if self.pending_welfare_claims > 0:
            group_health = "Needs Attention"
            group_health_colour = "watch"
            group_health_message = f"{self.pending_welfare_claims} welfare claim(s) need review."

        if self.total_cash <= 0:
            group_health = "Action Needed"
            group_health_colour = "danger"
            group_health_message = "The group has no available cash."

        return {
            "group_health": group_health,
            "group_health_colour": group_health_colour,
            "group_health_message": group_health_message
        }

    def build(self):
        dashboard_data = {}

        dashboard_data.update(self.build_group_health())

        return dashboard_data