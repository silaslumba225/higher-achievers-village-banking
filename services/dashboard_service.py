from datetime import datetime


class DashboardService:
    """
    Builds the intelligent data used by the Group Overview workspace.

    This service should remain independent of Flask routes and templates.
    It receives values from app.py and returns plain dictionaries for rendering.
    """

    def __init__(self, total_cash, overdue_loans, pending_welfare_claims, next_meeting=None):
        self.total_cash = total_cash
        self.overdue_loans = overdue_loans
        self.pending_welfare_claims = pending_welfare_claims
        self.next_meeting = next_meeting

    def build_daily_briefing(self):
        hour = datetime.now().hour

        if hour < 12:
            greeting = "Good Morning"
        elif hour < 17:
            greeting = "Good Afternoon"
        else:
            greeting = "Good Evening"

        recommendation = "Your group is progressing well today."

        if self.total_cash <= 0:
            recommendation = "Record today's savings before issuing any new loans."
        elif self.overdue_loans > 0:
            recommendation = f"I recommend following up {self.overdue_loans} overdue loan(s) today."
        elif self.pending_welfare_claims > 0:
            recommendation = f"There are {self.pending_welfare_claims} welfare request(s) waiting."

        return {
            "greeting": greeting,
            "briefing_title": f"{greeting}, Treasurer.",
            "briefing_message": "Welcome back. Here's today's summary of your village banking group.",
            "recommendation": recommendation,
        }

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
            "group_health_message": group_health_message,
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
            message = "Your group is strong and running well."
        elif score >= 70:
            status = "Good"
            colour = "primary"
            message = "Your group is doing well, with a few things to watch."
        elif score >= 50:
            status = "Needs Attention"
            colour = "warning"
            message = "Some matters need committee attention."
        else:
            status = "Action Required"
            colour = "danger"
            message = "The group needs urgent follow-up."

        return {
            "group_pulse": score,
            "group_pulse_status": status,
            "group_pulse_colour": colour,
            "group_pulse_message": message,
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
                "button": "Review Loans",
            })

        if self.pending_welfare_claims > 0:
            action_items.append({
                "level": "watch",
                "icon": "fa-heart",
                "title": f"{self.pending_welfare_claims} welfare claim(s) waiting",
                "message": "Review pending emergency fund requests.",
                "link": "welfare",
                "button": "Review Claims",
            })

        if self.total_cash <= 0:
            action_items.append({
                "level": "danger",
                "icon": "fa-wallet",
                "title": "No money available",
                "message": "The group currently has no available cash.",
                "link": "contributions",
                "button": "Record Savings",
            })

        if not action_items:
            action_items.append({
                "level": "good",
                "icon": "fa-circle-check",
                "title": "Nothing urgent today",
                "message": "Your group has no urgent items requiring attention.",
                "link": "executive_dashboard",
                "button": "View Overview",
            })

        return {
            "action_items": action_items,
        }

    def build_today_assistant(self):
        assistant_messages = []

        if self.total_cash <= 0:
            assistant_messages.append({
                "level": "danger",
                "icon": "fa-wallet",
                "title": "Record savings first",
                "message": "The group has no money available. Start by recording today's savings.",
            })

        if self.overdue_loans > 0:
            assistant_messages.append({
                "level": "watch",
                "icon": "fa-hand-holding-dollar",
                "title": "Follow up loan repayments",
                "message": f"{self.overdue_loans} loan(s) may need follow-up today.",
            })

        if self.pending_welfare_claims > 0:
            assistant_messages.append({
                "level": "watch",
                "icon": "fa-heart",
                "title": "Review emergency fund requests",
                "message": f"{self.pending_welfare_claims} request(s) are waiting for review.",
            })

        if not assistant_messages:
            assistant_messages.append({
                "level": "good",
                "icon": "fa-circle-check",
                "title": "Your group looks fine today",
                "message": "There are no urgent issues. You may continue with savings, loans, or reports.",
            })

        return {
            "assistant_messages": assistant_messages,
        }
    
    def build_group_health_check(self):
        health_items = []

        if self.total_cash > 0:
            health_items.append({
                "label": "Cash Position",
                "status": "Strong",
                "level": "good",
                "icon": "fa-wallet",
                "message": "The group has money available."
            })
        else:
            health_items.append({
                "label": "Cash Position",
                "status": "Action Needed",
                "level": "danger",
                "icon": "fa-wallet",
                "message": "The group has no money available."
            })

        if self.overdue_loans == 0:
            health_items.append({
                "label": "Loan Recovery",
                "status": "Healthy",
                "level": "good",
                "icon": "fa-hand-holding-dollar",
                "message": "No overdue loans need follow-up."
            })
        else:
            health_items.append({
                "label": "Loan Recovery",
                "status": "Needs Attention",
                "level": "watch",
                "icon": "fa-hand-holding-dollar",
                "message": f"{self.overdue_loans} loan(s) need follow-up."
            })

        if self.pending_welfare_claims == 0:
            health_items.append({
                "label": "Emergency Fund",
                "status": "Healthy",
                "level": "good",
                "icon": "fa-heart",
                "message": "No emergency fund requests are waiting."
            })
        else:
            health_items.append({
                "label": "Emergency Fund",
                "status": "Needs Review",
                "level": "watch",
                "icon": "fa-heart",
                "message": f"{self.pending_welfare_claims} request(s) are waiting."
            })

        health_items.append({
            "label": "Membership",
            "status": "Active",
            "level": "good",
            "icon": "fa-users",
            "message": "Member records are being tracked."
        })

        health_items.append({
            "label": "Savings",
            "status": "Open",
            "level": "good",
            "icon": "fa-coins",
            "message": "Savings can be recorded today."
        })

        return {
            "health_items": health_items
        }

    def build_success_celebrations(self):

        celebrations = []

        if self.overdue_loans == 0:
            celebrations.append({
            "icon": "fa-trophy",
            "title": "Excellent!",
            "message": "Your group has no overdue loans."
        })

        if self.total_cash > 0:
         celebrations.append({
            "icon": "fa-wallet",
            "title": "Money Available",
            "message": "The group has funds available for lending."
        })

        if self.pending_welfare_claims == 0:
            celebrations.append({
            "icon": "fa-heart",
            "title": "Well Done!",
            "message": "No welfare requests are waiting."
        })

        return {
        "celebrations": celebrations
     } 
    
    def build_meeting_countdown(self):
        if not self.next_meeting:
            return {
                "meeting_countdown": {
                    "has_meeting": False,
                    "title": "No meeting scheduled",
                    "message": "No upcoming meeting has been recorded yet.",
                    "days_remaining": None,
                    "meeting_date": None,
                    "meeting_type": None,
                    "agenda": None
                }
            }

        today = datetime.now().date()
        days_remaining = (self.next_meeting.meeting_date - today).days

        if days_remaining == 0:
            message = "The meeting is today. Remember to record attendance and resolutions."
        elif days_remaining == 1:
            message = "The meeting is tomorrow. Prepare the agenda and reports."
        elif days_remaining <= 7:
            message = "The meeting is coming soon. Start preparing committee reports."
        else:
            message = "There is enough time to prepare for this meeting."

        return {
            "meeting_countdown": {
                "has_meeting": True,
                "title": "Next Meeting",
                "message": message,
                "days_remaining": days_remaining,
                "meeting_date": self.next_meeting.meeting_date,
                "meeting_type": self.next_meeting.meeting_type,
                "agenda": self.next_meeting.agenda
                }
            }
        

    def build(self):
        dashboard_data = {}

        dashboard_data.update(self.build_daily_briefing())
        dashboard_data.update(self.build_group_health())
        dashboard_data.update(self.build_group_pulse())
        dashboard_data.update(self.build_action_centre())
        dashboard_data.update(self.build_today_assistant())
        dashboard_data.update(self.build_group_health_check())
        dashboard_data.update(self.build_success_celebrations())
        dashboard_data.update(self.build_meeting_countdown())
        
        return dashboard_data