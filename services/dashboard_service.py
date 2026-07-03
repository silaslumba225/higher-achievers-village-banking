def build_group_health(total_cash, overdue_loans, pending_welfare_claims):
    group_health = "Healthy"
    group_health_colour = "good"
    group_health_message = "Your group is doing well today."

    if overdue_loans > 0:
        group_health = "Needs Attention"
        group_health_colour = "watch"
        group_health_message = f"{overdue_loans} loan(s) need follow-up."

    if pending_welfare_claims > 0:
        group_health = "Needs Attention"
        group_health_colour = "watch"
        group_health_message = f"{pending_welfare_claims} welfare claim(s) need review."

    if total_cash <= 0:
        group_health = "Action Needed"
        group_health_colour = "danger"
        group_health_message = "The group has no available cash."

    return {
        "group_health": group_health,
        "group_health_colour": group_health_colour,
        "group_health_message": group_health_message
    }

def build_action_centre(self):
        action_items = []

        if self.overdue_loans > 0:
            action_items.append({
                "level": "danger",
                "icon": "fa-clock",
                "title": f"{self.overdue_loans} loan(s) need follow-up",
                "message": "Some loan repayments may be overdue.",
                "link": "loans",
                "button": "Review Loans"
            })

        if self.pending_welfare_claims > 0:
            action_items.append({
                "level": "watch",
                "icon": "fa-heart",
                "title": f"{self.pending_welfare_claims} welfare claim(s) waiting",
                "message": "Review pending emergency fund requests.",
                "link": "welfare",
                "button": "Review Claims"
            })

        if self.total_cash <= 0:
            action_items.append({
                "level": "danger",
                "icon": "fa-wallet",
                "title": "No money available",
                "message": "The group currently has no available cash.",
                "link": "contributions",
                "button": "Record Savings"
            })

        if not action_items:
            action_items.append({
                "level": "good",
                "icon": "fa-circle-check",
                "title": "Nothing urgent today",
                "message": "Your group has no urgent items requiring attention.",
                "link": "executive_dashboard",
                "button": "View Overview"
            })

        return {
            "action_items": action_items
        }    

def build(self):
        dashboard_data = {}
        dashboard_data.update(self.build_daily_briefing())
        dashboard_data.update(self.build_group_health())
        dashboard_data.update(self.build_action_centre())
        dashboard_data.update(self.build_today_assistant())
        dashboard_data.update(self.build_group_pulse())

        return dashboard_data

def build_today_assistant(self):
        assistant_messages = []

        if self.total_cash <= 0:
            assistant_messages.append({
                "level": "danger",
                "icon": "fa-wallet",
                "title": "Record savings first",
                "message": "The group has no money available. Start by recording today's savings."
            })

        if self.overdue_loans > 0:
            assistant_messages.append({
                "level": "watch",
                "icon": "fa-hand-holding-dollar",
                "title": "Follow up loan repayments",
                "message": f"{self.overdue_loans} loan(s) may need follow-up today."
            })

        if self.pending_welfare_claims > 0:
            assistant_messages.append({
                "level": "watch",
                "icon": "fa-heart",
                "title": "Review emergency fund requests",
                "message": f"{self.pending_welfare_claims} request(s) are waiting for review."
            })

        if not assistant_messages:
            assistant_messages.append({
                "level": "good",
                "icon": "fa-circle-check",
                "title": "Your group looks fine today",
                "message": "There are no urgent issues. You may continue with savings, loans, or reports."
            })

        return {
            "assistant_messages": assistant_messages
        }
def build_group_pulse(self):
    score = 100

    if self.total_cash <= 0:
        score -= 20

    if self.overdue_loans > 0:
        score -= 20

    if self.pending_welfare_claims > 0:
        score -= 20

    if score >= 90:
        status = "Excellent"
        colour = "success"

    elif score >= 70:
        status = "Good"
        colour = "primary"

    elif score >= 50:
        status = "Needs Attention"
        colour = "warning"

    else:
        status = "Action Required"
        colour = "danger"

    return {
        "group_pulse": score,
        "group_pulse_status": status,
        "group_pulse_colour": colour
    }        

from datetime import datetime

def build_daily_briefing(self):

    hour = datetime.now().hour

    if hour < 12:
        greeting = "Good Morning"

    elif hour < 17:
        greeting = "Good Afternoon"

    else:
        greeting = "Good Evening"

    recommendation = "Your group is progressing well today."

    if self.overdue_loans > 0:
        recommendation = (
            f"I recommend following up "
            f"{self.overdue_loans} overdue loan(s) today."
        )

    elif self.pending_welfare_claims > 0:
        recommendation = (
            f"There are "
            f"{self.pending_welfare_claims} welfare request(s) waiting."
        )

    elif self.total_cash <= 0:
        recommendation = (
            "Record today's savings before issuing any new loans."
        )

    return {

        "greeting": greeting,

        "briefing_title":
            f"{greeting}, Treasurer.",

        "briefing_message":
            "Welcome back. Here's today's summary of your village banking group.",

        "recommendation":
            recommendation

    }