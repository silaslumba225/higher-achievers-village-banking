class MeetingIntelligenceService:

    def __init__(
        self,
        next_meeting=None,
        days_to_next=None,
        average_attendance=0,
        pending_resolutions=0,
        meetings_this_year=0
    ):

        self.next_meeting = next_meeting
        self.days_to_next = days_to_next
        self.average_attendance = average_attendance
        self.pending_resolutions = pending_resolutions
        self.meetings_this_year = meetings_this_year

    # --------------------------------

    def build_health(self):

        score = 100

        if self.average_attendance < 10:
            score -= 20

        if self.days_to_next is None:
            score -= 10

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
            status = "Poor"
            colour = "danger"

        return {
            "meeting_health_score": score,
            "meeting_health_status": status,
            "meeting_health_colour": colour
        }

    # --------------------------------

    def build_assistant(self):

        assistant = []

        if self.next_meeting:

            assistant.append({

                "level": "info",

                "icon": "fa-calendar",

                "title": "Next meeting",

                "message":
                    f"{self.next_meeting.meeting_type} in {self.days_to_next} day(s)."
            })

        if self.pending_resolutions > 0:

            assistant.append({

                "level": "info",

                "icon": "fa-list-check",

                "title": "Recorded resolutions",

                "message":
                    f"{self.pending_resolutions} meeting(s) contain recorded resolutions or minutes."

            })

        if not assistant:

            assistant.append({

                "level": "success",

                "icon": "fa-circle-check",

                "title": "Everything looks good",

                "message":
                    "Meeting governance is running smoothly."

            })

        return {
            "meeting_assistant": assistant
        }

    # --------------------------------

    def build(self):

        data = {}

        data.update(self.build_health())
        data.update(self.build_assistant())

        return data
