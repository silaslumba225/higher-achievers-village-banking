class MemberIntelligenceService:
    def __init__(
        self,
        member,
        total_contributions,
        loan_balance,
        fine_balance,
        recent_contributions=None,
        recent_loans=None,
        recent_fines=None
    ):
        self.member = member
        self.total_contributions = total_contributions
        self.loan_balance = loan_balance
        self.fine_balance = fine_balance
        self.recent_contributions = recent_contributions or []
        self.recent_loans = recent_loans or []
        self.recent_fines = recent_fines or []

    def calculate_trust_score(self):
        score = 0

        if self.member.status == "Active":
            score += 20

        if self.total_contributions > 0:
            score += 25

        if self.loan_balance <= 0:
            score += 25
        else:
            score += 10

        if self.fine_balance <= 0:
            score += 20

        if len(self.recent_contributions) > 0:
            score += 10

        return min(score, 100)

    def calculate_rating(self, score):
        if score >= 95:
            return "Outstanding Member", "★★★★★", "good"
        elif score >= 80:
            return "Trusted Member", "★★★★☆", "good"
        elif score >= 65:
            return "Good Standing", "★★★☆☆", "watch"
        elif score >= 50:
            return "Needs Improvement", "★★☆☆☆", "watch"
        else:
            return "High Attention", "★☆☆☆☆", "danger"

    def determine_loan_eligibility(self, score):
        if self.member.status != "Active":
            return "Not Recommended Yet", "Member is currently inactive.", "danger"

        if self.fine_balance > 0:
            return "Not Recommended Yet", "Outstanding fines should be cleared first.", "watch"

        if self.loan_balance > 0:
            return "Review Carefully", "Member already has an outstanding loan balance.", "watch"

        if score >= 75:
            return "Eligible for Loan Consideration", "Member appears suitable for committee review.", "good"

        return "Review Carefully", "Committee should review savings and repayment history first.", "watch"

    def build_badges(self):
        badges = []

        if self.total_contributions > 0:
            badges.append({
                "icon": "fa-coins",
                "title": "Saver",
                "message": "This member has recorded savings."
            })

        if self.loan_balance <= 0:
            badges.append({
                "icon": "fa-circle-check",
                "title": "No Outstanding Loan",
                "message": "No current loan balance is showing."
            })

        if self.fine_balance <= 0:
            badges.append({
                "icon": "fa-thumbs-up",
                "title": "Good Standing",
                "message": "No unpaid fines are currently showing."
            })

        if self.member.committee_position:
            badges.append({
                "icon": "fa-star",
                "title": "Committee Member",
                "message": self.member.committee_position
            })

        return badges

    def build(self):
        score = self.calculate_trust_score()
        rating, stars, level = self.calculate_rating(score)
        eligibility, eligibility_message, eligibility_level = self.determine_loan_eligibility(score)

        return {
            "member_trust_score": score,
            "member_rating": rating,
            "member_stars": stars,
            "member_rating_level": level,
            "loan_eligibility": eligibility,
            "loan_eligibility_message": eligibility_message,
            "loan_eligibility_level": eligibility_level,
            "member_badges": self.build_badges()
        }