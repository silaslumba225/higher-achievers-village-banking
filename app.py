from datetime import date, datetime, timedelta
from openpyxl import load_workbook
import os
from decimal import Decimal
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, Response, session, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import csv
import io
import os
import shutil
import json
from pathlib import Path
from werkzeug.utils import secure_filename
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'change-this-secret-key')

# Local development uses SQLite. Render/production uses PostgreSQL through DATABASE_URL.
database_url = os.environ.get('DATABASE_URL')
if database_url:
    # Render provides postgresql:// URLs. Use the psycopg v3 SQLAlchemy driver.
    if database_url.startswith('postgresql://'):
        database_url = database_url.replace('postgresql://', 'postgresql+psycopg://', 1)
    elif database_url.startswith('postgres://'):
        database_url = database_url.replace('postgres://', 'postgresql+psycopg://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = database_url or 'sqlite:///village_banking.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

INTEREST_RATE = Decimal('0.15')
PAYMENT_METHODS = ['Bank Transfer', 'Mobile Money', 'Cash']
ABSENCE_FINE_AMOUNT = Decimal('50.00')
LATE_ATTENDANCE_FINE_AMOUNT = Decimal('20.00')
LATE_CONTRIBUTION_FINE_AMOUNT = Decimal('10.00')
CLIENT_NAME = 'Higher Achievers'
PRODUCER_NAME = 'Excelling Foundation'
BACKUP_RETENTION = 30


FINE_CATEGORIES = [
    'Late Meeting Attendance',
    'Absence from Meeting',
    'Late Contribution',
    'Late Loan Repayment',
    'Rule Violation',
    'Other Penalty',
]
FINE_STATUSES = ['Unpaid', 'Partially Paid', 'Paid', 'Waived']

WELFARE_CATEGORIES = [
    'Funeral Support',
    'Hospitalization',
    'Birth Support',
    'School Emergency',
    'House Fire',
    'Other Emergency',
]
WELFARE_STATUSES = ['Requested', 'Reviewed', 'Approved', 'Paid', 'Rejected']
ATTENDANCE_STATUSES = ['Present', 'Absent', 'Late', 'Excused']
NOTIFICATION_CHANNELS = ['SMS', 'WhatsApp']
NOTIFICATION_TYPES = ['Contribution Reminder', 'Loan Repayment Reminder', 'Meeting Reminder', 'Welfare Notification', 'Share-Out Notification', 'General Notice']
ACCOUNT_TYPES = ['Asset', 'Liability', 'Equity', 'Income', 'Expense']

ROLES = ['Administrator', 'Chairperson', 'Treasurer', 'Secretary', 'Auditor', 'Data Clerk']
ROLE_PERMISSIONS = {
    'Administrator': [
    'dashboard',
    'members',
    'contributions',
    'loans',
    'repayments',
    'distributions',
    'meetings',
    'attendance',
    'reports',
    'statements',
    'shareout',
    'fines',
    'welfare',
    'users',
    'audit',
    'backups',
    'notifications',
    'accounting',
    'exports',
    'settings'
],
    'Chairperson': ['dashboard', 'loans', 'distributions', 'meetings', 'attendance', 'reports', 'statements', 'shareout', 'fines', 'welfare', 'audit', 'notifications', 'accounting', 'exports'],
    'Treasurer': ['dashboard', 'contributions', 'loans', 'repayments', 'distributions', 'reports', 'statements', 'shareout', 'fines', 'welfare', 'notifications', 'accounting', 'exports'],
    'Secretary': ['dashboard', 'members', 'meetings', 'attendance', 'reports', 'statements', 'welfare', 'notifications', 'accounting', 'exports'],
    'Auditor': ['dashboard', 'attendance', 'reports', 'statements', 'shareout', 'fines', 'welfare', 'audit', 'notifications', 'accounting', 'exports'],
    'Data Clerk': ['dashboard', 'members', 'contributions', 'notifications'],
}

def user_can(permission):
    user = session.get('user') or {}
    return permission in ROLE_PERMISSIONS.get(user.get('role'), [])

def ensure_month_end_columns():
    columns = {
        'members_processed': 'INTEGER DEFAULT 0',
        'loans_processed': 'INTEGER DEFAULT 0',
        'reversed': 'BOOLEAN DEFAULT FALSE',
        'reversed_by': 'VARCHAR(120)',
        'reversed_on': 'DATE',
        'reversal_reason': 'VARCHAR(250)',
    }

    for column, definition in columns.items():
        try:
            db.session.execute(
                db.text(f'ALTER TABLE month_end_process ADD COLUMN IF NOT EXISTS {column} {definition}')
            )
            db.session.commit()
        except Exception:
            db.session.rollback()

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    full_name = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(30), default='Administrator')
    password_hash = db.Column(db.String(255), nullable=False)
    active = db.Column(db.Boolean, default=True)


class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer)
    username = db.Column(db.String(50))
    full_name = db.Column(db.String(120))
    role = db.Column(db.String(30))
    action = db.Column(db.String(80), nullable=False)
    entity = db.Column(db.String(80))
    entity_id = db.Column(db.String(40))
    details = db.Column(db.Text)
    ip_address = db.Column(db.String(60))
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class BackupRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_size = db.Column(db.Integer, default=0)
    backup_type = db.Column(db.String(30), default='Manual')
    created_by = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    notes = db.Column(db.Text)


class NotificationLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    channel = db.Column(db.String(30), nullable=False)
    notification_type = db.Column(db.String(80), nullable=False)
    recipient_type = db.Column(db.String(80))
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'))
    phone = db.Column(db.String(40))
    subject = db.Column(db.String(160))
    message = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(30), default='Prepared')
    created_by = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    sent_at = db.Column(db.DateTime)
    provider_reference = db.Column(db.String(120))
    member = db.relationship('Member', backref='notification_logs')

class Member(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_no = db.Column(db.String(20), unique=True, nullable=False)
    full_name = db.Column(db.String(120), nullable=False)
    phone = db.Column(db.String(30))
    national_id = db.Column(db.String(50))
    group_name = db.Column(db.String(80))
    status = db.Column(db.String(20), default='Active')
    created_at = db.Column(db.Date, default=date.today)

class Contribution(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    month = db.Column(db.String(7), nullable=False)
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    method = db.Column(db.String(30), nullable=False)
    reference = db.Column(db.String(80))
    paid_on = db.Column(db.Date, default=date.today)
    member = db.relationship('Member', backref='contributions')

class Loan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    principal = db.Column(db.Numeric(12, 2), nullable=False)
    interest_rate = db.Column(db.Numeric(5, 2), default=INTEREST_RATE)
    issued_on = db.Column(db.Date, default=date.today)
    due_on = db.Column(db.Date, nullable=False)
    purpose = db.Column(db.String(200))
    status = db.Column(db.String(20), default='Applied')  # Applied, Reviewed, Approved, Disbursed, Open, Rejected, Closed
    approved_by = db.Column(db.String(120), default='')
    reviewed_by = db.Column(db.String(120))
    disbursed_by = db.Column(db.String(120))
    reviewed_on = db.Column(db.Date)
    approved_on = db.Column(db.Date)
    disbursed_on = db.Column(db.Date)
    rejected_on = db.Column(db.Date)
    rejection_reason = db.Column(db.String(250))
    member = db.relationship('Member', backref='loans')

    @property
    def interest_amount(self):
        return money(self.principal * self.interest_rate)

    @property
    def total_due(self):
        return money(self.principal + self.interest_amount)

    @property
    def total_paid(self):
        return money(sum((r.amount for r in self.repayments), Decimal('0.00')))

    @property
    def balance(self):
        return money(self.total_due - self.total_paid)

    @property
    def overdue(self):
        return self.status in ['Open', 'Disbursed'] and self.due_on < date.today() and self.balance > 0

class Repayment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    loan_id = db.Column(db.Integer, db.ForeignKey('loan.id'), nullable=False)
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    method = db.Column(db.String(30), nullable=False)
    reference = db.Column(db.String(80))
    paid_on = db.Column(db.Date, default=date.today)
    loan = db.relationship('Loan', backref='repayments')

class Distribution(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    method = db.Column(db.String(30), nullable=False)
    reference = db.Column(db.String(80))
    authorized_by = db.Column(db.String(120))
    paid_on = db.Column(db.Date, default=date.today)
    member = db.relationship('Member', backref='distributions')


class WelfareContribution(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    month = db.Column(db.String(7), nullable=False)
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    method = db.Column(db.String(30), nullable=False)
    reference = db.Column(db.String(80))
    paid_on = db.Column(db.Date, default=date.today)
    member = db.relationship('Member', backref='welfare_contributions')


class WelfareClaim(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    category = db.Column(db.String(60), nullable=False)
    amount_requested = db.Column(db.Numeric(12, 2), nullable=False)
    amount_approved = db.Column(db.Numeric(12, 2), default=0)
    status = db.Column(db.String(20), default='Requested')
    reason = db.Column(db.String(250))
    requested_on = db.Column(db.Date, default=date.today)
    reviewed_by = db.Column(db.String(120))
    reviewed_on = db.Column(db.Date)
    approved_by = db.Column(db.String(120))
    approved_on = db.Column(db.Date)
    paid_by = db.Column(db.String(120))
    paid_on = db.Column(db.Date)
    payment_method = db.Column(db.String(30))
    reference = db.Column(db.String(80))
    rejection_reason = db.Column(db.String(250))
    member = db.relationship('Member', backref='welfare_claims')

    @property
    def payable_amount(self):
        return money(self.amount_approved or 0)



class FinePenalty(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    category = db.Column(db.String(60), nullable=False)
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    reason = db.Column(db.String(250))
    fine_date = db.Column(db.Date, default=date.today)
    status = db.Column(db.String(20), default='Unpaid')
    recorded_by = db.Column(db.String(120))
    waived_by = db.Column(db.String(120))
    waived_on = db.Column(db.Date)
    member = db.relationship('Member', backref='fines')

    @property
    def total_paid(self):
        return money(sum((p.amount for p in self.payments), Decimal('0.00')))

    @property
    def balance(self):
        return money(self.amount - self.total_paid)
    
class FinePayment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fine_id = db.Column(db.Integer, db.ForeignKey('fine_penalty.id'), nullable=False)
    amount = db.Column(db.Numeric(12, 2), nullable=False)
    method = db.Column(db.String(30), nullable=False)
    reference = db.Column(db.String(80))
    paid_on = db.Column(db.Date, default=date.today)
    fine = db.relationship('FinePenalty', backref='payments')

class Meeting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    meeting_type = db.Column(db.String(40), nullable=False)
    meeting_date = db.Column(db.Date, nullable=False)
    agenda = db.Column(db.Text)
    resolutions = db.Column(db.Text)
    attendance_count = db.Column(db.Integer, default=0)


class MeetingAttendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    meeting_id = db.Column(db.Integer, db.ForeignKey('meeting.id'), nullable=False)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    status = db.Column(db.String(20), default='Present')
    remarks = db.Column(db.String(250))
    recorded_by = db.Column(db.String(120))
    recorded_at = db.Column(db.DateTime, default=datetime.utcnow)
    fine_generated = db.Column(db.Boolean, default=False)
    meeting = db.relationship('Meeting', backref='attendance_records')
    member = db.relationship('Member', backref='attendance_records')


class Account(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(20), unique=True, nullable=False)
    name = db.Column(db.String(120), nullable=False)
    account_type = db.Column(db.String(30), nullable=False)
    normal_balance = db.Column(db.String(10), default='Debit')
    active = db.Column(db.Boolean, default=True)


class JournalEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    entry_date = db.Column(db.Date, default=date.today, nullable=False)
    description = db.Column(db.String(250), nullable=False)
    reference = db.Column(db.String(100))
    source_type = db.Column(db.String(60))
    source_id = db.Column(db.String(60))
    posted_by = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    @property
    def total_debit(self):
        return money(sum((line.debit for line in self.lines), Decimal('0.00')))

    @property
    def total_credit(self):
        return money(sum((line.credit for line in self.lines), Decimal('0.00')))

    @property
    def balanced(self):
        return self.total_debit == self.total_credit


class JournalLine(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    journal_entry_id = db.Column(db.Integer, db.ForeignKey('journal_entry.id'), nullable=False)
    account_id = db.Column(db.Integer, db.ForeignKey('account.id'), nullable=False)
    debit = db.Column(db.Numeric(12, 2), default=0)
    credit = db.Column(db.Numeric(12, 2), default=0)
    memo = db.Column(db.String(250))
    entry = db.relationship('JournalEntry', backref='lines')
    account = db.relationship('Account', backref='journal_lines')

class SavingsInterest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    month = db.Column(db.String(7), nullable=False)  # YYYY-MM
    opening_balance = db.Column(db.Numeric(12, 2), nullable=False)
    interest_rate = db.Column(db.Numeric(5, 2), default=15.00)
    interest_amount = db.Column(db.Numeric(12, 2), nullable=False)
    closing_balance = db.Column(db.Numeric(12, 2), nullable=False)
    processed_on = db.Column(db.Date, default=date.today)

    member = db.relationship('Member', backref='savings_interest')


class LoanInterest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    loan_id = db.Column(db.Integer, db.ForeignKey('loan.id'), nullable=False)
    member_id = db.Column(db.Integer, db.ForeignKey('member.id'), nullable=False)
    month = db.Column(db.String(7), nullable=False)  # YYYY-MM
    opening_balance = db.Column(db.Numeric(12, 2), nullable=False)
    interest_rate = db.Column(db.Numeric(5, 2), default=15.00)
    interest_amount = db.Column(db.Numeric(12, 2), nullable=False)
    closing_balance = db.Column(db.Numeric(12, 2), nullable=False)
    processed_on = db.Column(db.Date, default=date.today)

    loan = db.relationship('Loan', backref='loan_interest_entries')
    member = db.relationship('Member', backref='loan_interest_entries')


class MonthEndProcess(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    month = db.Column(db.String(7), nullable=False)  # YYYY-MM

    savings_interest_total = db.Column(db.Numeric(12, 2), default=0)
    loan_interest_total = db.Column(db.Numeric(12, 2), default=0)

    members_processed = db.Column(db.Integer, default=0)
    loans_processed = db.Column(db.Integer, default=0)

    processed_by = db.Column(db.String(120))
    processed_on = db.Column(db.Date, default=date.today)

    reversed = db.Column(db.Boolean, default=False)
    reversed_by = db.Column(db.String(120))
    reversed_on = db.Column(db.Date)
    reversal_reason = db.Column(db.String(250))

class SystemSetting(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    organization_name = db.Column(db.String(200), default='Higher Achievers Village Banking')
    contribution_amount = db.Column(db.Numeric(12, 2), default=100)
    savings_interest_rate = db.Column(db.Numeric(5, 2), default=15)
    loan_interest_rate = db.Column(db.Numeric(5, 2), default=15)
    welfare_contribution_amount = db.Column(db.Numeric(12, 2), default=0)

    updated_on = db.Column(db.DateTime, default=datetime.utcnow)
    organization_address = db.Column(db.String(250))
    organization_phone = db.Column(db.String(50))
    organization_email = db.Column(db.String(120))
    registration_number = db.Column(db.String(100))

def ensure_settings_columns():
    columns = {
        'logo_url': 'VARCHAR(500)',
    }

    for column, definition in columns.items():
        try:
            db.session.execute(
                db.text(f'ALTER TABLE system_setting ADD COLUMN IF NOT EXISTS {column} {definition}')
            )
            db.session.commit()
        except Exception:
            db.session.rollback() 
def ensure_settings_columns():
    columns = {
        'organization_address': 'VARCHAR(250)',
        'organization_phone': 'VARCHAR(50)',
        'organization_email': 'VARCHAR(120)',
        'registration_number': 'VARCHAR(100)',
    }

    for column, definition in columns.items():
        try:
            db.session.execute(
                db.text(f'ALTER TABLE system_setting ADD COLUMN IF NOT EXISTS {column} {definition}')
            )
            db.session.commit()
        except Exception:
            db.session.rollback()   

def money(value):
    return Decimal(value or 0).quantize(Decimal('0.01'))

@app.template_filter('kwacha')
def kwacha(value):
    return f"K {money(value):,.2f}"

@app.context_processor
def inject_globals():

    setting = SystemSetting.query.first()

    client_name = (
        setting.organization_name
        if setting and setting.organization_name
        else CLIENT_NAME
    )

    return dict(
        payment_methods=PAYMENT_METHODS,
        interest_percent=int(INTEREST_RATE * 100),
        client_name=client_name,
        setting=setting,
        producer_name=PRODUCER_NAME,
        current_year=date.today().year,
        current_user=session.get('user'),
        user_can=user_can,
        roles=ROLES,
        fine_categories=FINE_CATEGORIES,
        fine_statuses=FINE_STATUSES,
        welfare_categories=WELFARE_CATEGORIES,
        welfare_statuses=WELFARE_STATUSES,
        attendance_statuses=ATTENDANCE_STATUSES,
        notification_channels=NOTIFICATION_CHANNELS,
        notification_types=NOTIFICATION_TYPES,
        account_types=ACCOUNT_TYPES
    )
def get_settings():
    setting = SystemSetting.query.first()

    if not setting:
        setting = SystemSetting()
        db.session.add(setting)
        db.session.commit()

    return setting

def log_audit(action, entity=None, entity_id=None, details=None):
    user = session.get('user') or {}
    try:
        entry = AuditLog(
            user_id=user.get('id'),
            username=user.get('username', 'system'),
            full_name=user.get('full_name', 'System'),
            role=user.get('role', 'System'),
            action=action,
            entity=entity,
            entity_id=str(entity_id) if entity_id is not None else None,
            details=details,
            ip_address=request.headers.get('X-Forwarded-For', request.remote_addr) if request else None,
        )
        db.session.add(entry)
        db.session.commit()
    except Exception:
        db.session.rollback()


def seed_chart_of_accounts():
    default_accounts = [
        ('1000', 'Cash on Hand', 'Asset', 'Debit'),
        ('1010', 'Bank Account', 'Asset', 'Debit'),
        ('1020', 'Mobile Money', 'Asset', 'Debit'),
        ('1100', 'Loans Receivable', 'Asset', 'Debit'),
        ('2000', 'Member Savings / Contributions', 'Liability', 'Credit'),
        ('2010', 'Welfare Fund Payable', 'Liability', 'Credit'),
        ('3000', 'Accumulated Surplus', 'Equity', 'Credit'),
        ('4000', 'Loan Interest Income', 'Income', 'Credit'),
        ('4010', 'Fines and Penalties Income', 'Income', 'Credit'),
        ('4020', 'Other Income', 'Income', 'Credit'),
        ('5000', 'Administrative Expenses', 'Expense', 'Debit'),
        ('5010', 'Bank Charges', 'Expense', 'Debit'),
        ('5020', 'Transport and Meetings', 'Expense', 'Debit'),
        ('5030', 'Welfare Support Expense', 'Expense', 'Debit'),
        ('5040', 'Share-Out / Dividend Distribution', 'Expense', 'Debit'),
    ]
    for code, name, account_type, normal_balance in default_accounts:
        if not Account.query.filter_by(code=code).first():
            db.session.add(Account(code=code, name=name, account_type=account_type, normal_balance=normal_balance))
    db.session.commit()


def account_by_code(code):
    return Account.query.filter_by(code=code).first()


def cash_account_for_method(method):
    if method == 'Bank Transfer':
        return account_by_code('1010')
    if method == 'Mobile Money':
        return account_by_code('1020')
    return account_by_code('1000')


def post_journal(entry_date, description, reference, source_type, source_id, lines):
    if source_type and source_id and JournalEntry.query.filter_by(source_type=source_type, source_id=str(source_id)).first():
        return None
    debit_total = money(sum((money(line.get('debit', 0)) for line in lines), Decimal('0.00')))
    credit_total = money(sum((money(line.get('credit', 0)) for line in lines), Decimal('0.00')))
    if debit_total != credit_total:
        raise ValueError('Journal entry is not balanced.')
    entry = JournalEntry(entry_date=entry_date or date.today(), description=description, reference=reference, source_type=source_type, source_id=str(source_id) if source_id is not None else None, posted_by=(session.get('user') or {}).get('username'))
    db.session.add(entry)
    db.session.flush()
    for line in lines:
        db.session.add(JournalLine(journal_entry_id=entry.id, account_id=line['account'].id, debit=money(line.get('debit', 0)), credit=money(line.get('credit', 0)), memo=line.get('memo')))
    db.session.commit()
    return entry


def sync_operational_transactions_to_gl():
    seed_chart_of_accounts()
    created = 0
    member_savings = account_by_code('2000')
    loans_receivable = account_by_code('1100')
    interest_income = account_by_code('4000')
    fines_income = account_by_code('4010')
    welfare_payable = account_by_code('2010')
    welfare_expense = account_by_code('5030')
    # Contributions: Dr payment account / Cr member savings
    for c in Contribution.query.all():
        if not JournalEntry.query.filter_by(source_type='Contribution', source_id=str(c.id)).first():
            post_journal(c.paid_on, f'Contribution from {c.member.full_name} for {c.month}', c.reference, 'Contribution', c.id, [
                {'account': cash_account_for_method(c.method), 'debit': c.amount},
                {'account': member_savings, 'credit': c.amount},
            ])
            created += 1
    # Disbursed loans: Dr loans receivable / Cr cash on hand
    for l in Loan.query.filter(Loan.status.in_(['Open','Disbursed','Closed'])).all():
        if not JournalEntry.query.filter_by(source_type='Loan', source_id=str(l.id)).first():
            post_journal(l.issued_on, f'Loan issued to {l.member.full_name}', f'Loan-{l.id}', 'Loan', l.id, [
                {'account': loans_receivable, 'debit': l.principal},
                {'account': account_by_code('1000'), 'credit': l.principal},
            ])
            created += 1
    # Repayments: Dr payment account / Cr loans receivable and interest income estimate
    for r in Repayment.query.all():
        if not JournalEntry.query.filter_by(source_type='Repayment', source_id=str(r.id)).first():
            interest_part = money(min(money(r.amount), money(r.loan.interest_amount)))
            principal_part = money(money(r.amount) - interest_part)
            lines = [{'account': cash_account_for_method(r.method), 'debit': r.amount}]
            if principal_part > 0:
                lines.append({'account': loans_receivable, 'credit': principal_part})
            if interest_part > 0:
                lines.append({'account': interest_income, 'credit': interest_part})
            post_journal(r.paid_on, f'Loan repayment from {r.loan.member.full_name}', r.reference, 'Repayment', r.id, lines)
            created += 1
    # Fine payments: Dr payment account / Cr fines income
    for fp in FinePayment.query.all():
        if not JournalEntry.query.filter_by(source_type='FinePayment', source_id=str(fp.id)).first():
            post_journal(fp.paid_on, f'Fine payment from {fp.fine.member.full_name}', fp.reference, 'FinePayment', fp.id, [
                {'account': cash_account_for_method(fp.method), 'debit': fp.amount},
                {'account': fines_income, 'credit': fp.amount},
            ])
            created += 1
    # Welfare contributions: Dr payment account / Cr welfare fund payable
    if 'WelfareContribution' in globals():
        for wc in WelfareContribution.query.all():
            if not JournalEntry.query.filter_by(source_type='WelfareContribution', source_id=str(wc.id)).first():
                post_journal(wc.paid_on, f'Welfare contribution from {wc.member.full_name}', wc.reference, 'WelfareContribution', wc.id, [
                    {'account': cash_account_for_method(wc.method), 'debit': wc.amount},
                    {'account': welfare_payable, 'credit': wc.amount},
                ])
                created += 1
    # Welfare paid claims: Dr welfare support expense / Cr cash on hand
    if 'WelfareClaim' in globals():
        for claim in WelfareClaim.query.filter_by(status='Paid').all():
            if not JournalEntry.query.filter_by(source_type='WelfareClaim', source_id=str(claim.id)).first():
                post_journal(claim.paid_on or date.today(), f'Welfare payment to {claim.member.full_name}', f'Welfare-{claim.id}', 'WelfareClaim', claim.id, [
                    {'account': welfare_expense, 'debit': claim.amount},
                    {'account': account_by_code('1000'), 'credit': claim.amount},
                ])
                created += 1
    return created


def ledger_balances(start_date=None, end_date=None):
    query = JournalLine.query.join(JournalEntry).join(Account)
    if start_date:
        query = query.filter(JournalEntry.entry_date >= start_date)
    if end_date:
        query = query.filter(JournalEntry.entry_date <= end_date)
    balances = []
    for account in Account.query.order_by(Account.code).all():
        lines = JournalLine.query.join(JournalEntry).filter(JournalLine.account_id == account.id)
        if start_date:
            lines = lines.filter(JournalEntry.entry_date >= start_date)
        if end_date:
            lines = lines.filter(JournalEntry.entry_date <= end_date)
        debit = money(sum((line.debit for line in lines.all()), Decimal('0.00')))
        credit = money(sum((line.credit for line in lines.all()), Decimal('0.00')))
        balance = money(debit - credit) if account.normal_balance == 'Debit' else money(credit - debit)
        balances.append({'account': account, 'debit': debit, 'credit': credit, 'balance': balance})
    return balances

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get('user'):
            return redirect(url_for('login', next=request.path))
        return view(*args, **kwargs)
    return wrapped

def role_required(permission):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not session.get('user'):
                return redirect(url_for('login', next=request.path))
            if not user_can(permission):
                flash('You do not have permission to access that section.', 'error')
                return redirect(url_for('dashboard'))
            return view(*args, **kwargs)
        return wrapped
    return decorator


def database_file_path():
    # Flask stores sqlite:///village_banking.db inside the instance folder.
    return os.path.join(app.instance_path, 'village_banking.db')

def backups_folder():
    folder = os.path.join(app.instance_path, 'backups')
    os.makedirs(folder, exist_ok=True)
    return folder

def make_backup_filename(prefix='backup'):
    return f"{prefix}_{datetime.now().strftime('%Y_%m_%d_%H%M%S')}.db"

def prune_old_backups(keep=BACKUP_RETENTION):
    folder = backups_folder()
    files = sorted(Path(folder).glob('*.db'), key=lambda x: x.stat().st_mtime, reverse=True)
    for old in files[keep:]:
        try:
            old.unlink()
        except OSError:
            pass

def create_database_backup(prefix='backup', notes=None):
    os.makedirs(app.instance_path, exist_ok=True)
    src = database_file_path()
    if not os.path.exists(src):
        raise FileNotFoundError('Database file was not found. Start the application once before creating backups.')
    filename = make_backup_filename(prefix)
    dst = os.path.join(backups_folder(), filename)
    db.session.remove()
    db.engine.dispose()
    shutil.copy2(src, dst)
    size = os.path.getsize(dst)
    user = session.get('user') or {}
    record = BackupRecord(filename=filename, file_size=size, backup_type='Manual', created_by=user.get('full_name') or user.get('username') or 'System', notes=notes)
    db.session.add(record)
    db.session.commit()
    prune_old_backups()
    return record

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = User.query.filter_by(
            username=request.form.get('username', '').strip(),
            active=True
        ).first()

        if user and check_password_hash(user.password_hash, request.form.get('password', '')):
            session['user'] = {
                'id': user.id,
                'username': user.username,
                'full_name': user.full_name,
                'role': user.role
            }

            log_audit('LOGIN_SUCCESS', 'User', user.id, f'{user.full_name} logged in')
            flash('Welcome back. You are logged in securely.')

            return redirect(request.args.get('next') or url_for('home'))

        flash('Invalid username or password.', 'error')

    return render_template('login.html')

@app.route('/logout')
def logout():
    log_audit('LOGOUT', 'User', session.get('user', {}).get('id'), 'User logged out')
    session.clear()
    flash('You have been logged out.')
    return redirect(url_for('login'))

@app.route('/')
@login_required
def home():
    return render_template('home.html')

@app.route('/settings', methods=['GET', 'POST'])
@login_required
@role_required('settings')
def settings():
    setting = SystemSetting.query.first()

    if not setting:
        setting = SystemSetting()
        db.session.add(setting)
        db.session.commit()

    if request.method == 'POST':
        setting.organization_name = request.form['organization_name']
        setting.contribution_amount = money(request.form['contribution_amount'])
        setting.savings_interest_rate = money(request.form['savings_interest_rate'])
        setting.loan_interest_rate = money(request.form['loan_interest_rate'])
        setting.welfare_contribution_amount = money(request.form['welfare_contribution_amount'])
        setting.organization_address = request.form.get('organization_address')
        setting.organization_phone = request.form.get('organization_phone')
        setting.organization_email = request.form.get('organization_email')
        setting.registration_number = request.form.get('registration_number')

        db.session.commit()

        log_audit(
            'UPDATE_SETTINGS',
            'SystemSetting',
            setting.id,
            'System settings updated'
        )

        flash('Settings updated successfully.')
        return redirect(url_for('settings'))

    return render_template('settings.html', setting=setting)

@app.route('/dashboard')
@login_required
@role_required('dashboard')
def dashboard():
    members = Member.query.count()
    active_members = Member.query.filter_by(status='Active').count()
    contribution_total = money(db.session.query(db.func.coalesce(db.func.sum(Contribution.amount), 0)).scalar())
    month = date.today().strftime('%Y-%m')
    month_contribution_total = money(db.session.query(db.func.coalesce(db.func.sum(Contribution.amount), 0)).filter(Contribution.month == month).scalar())
    loans_total = money(db.session.query(db.func.coalesce(db.func.sum(Loan.principal), 0)).scalar())
    repayment_total = money(db.session.query(db.func.coalesce(db.func.sum(Repayment.amount), 0)).scalar())
    distribution_total = money(db.session.query(db.func.coalesce(db.func.sum(Distribution.amount), 0)).scalar())
    fine_total = money(db.session.query(db.func.coalesce(db.func.sum(FinePenalty.amount), 0)).scalar())
    fine_paid_total = money(db.session.query(db.func.coalesce(db.func.sum(FinePayment.amount), 0)).scalar())
    welfare_contribution_total = money(db.session.query(db.func.coalesce(db.func.sum(WelfareContribution.amount), 0)).scalar())
    welfare_paid_total = money(db.session.query(db.func.coalesce(db.func.sum(WelfareClaim.amount_approved), 0)).filter(WelfareClaim.status == 'Paid').scalar())
    welfare_balance = money(welfare_contribution_total - welfare_paid_total)
    welfare_pending = WelfareClaim.query.filter(WelfareClaim.status.in_(['Requested', 'Reviewed', 'Approved'])).count()
    open_loans = Loan.query.filter(Loan.status.in_(['Open', 'Disbursed'])).all()
    loan_balance = money(sum((l.balance for l in open_loans), Decimal('0.00')))
    interest_earned = money(sum((l.interest_amount for l in Loan.query.all()), Decimal('0.00')))
    available_fund = money(contribution_total + repayment_total + fine_paid_total - loans_total - distribution_total)
    next_committee = next_month_end(date.today())
    next_full = next_quarter_meeting(date.today())
    recent_contribs = Contribution.query.order_by(Contribution.paid_on.desc()).limit(5).all()
    recent_loans = Loan.query.order_by(Loan.issued_on.desc()).limit(5).all()
    overdue_loans = [l for l in open_loans if l.overdue]
    savings_interest_total = money(
    db.session.query(db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0)).scalar()
    )

    loan_interest_charged = money(
        db.session.query(db.func.coalesce(db.func.sum(LoanInterest.interest_amount), 0)).scalar()
    )

    outstanding_fines = money(
        sum(
            (f.balance for f in FinePenalty.query.all() if f.status != 'Waived'),
            Decimal('0.00')
        )
    )

    net_surplus = money(
        contribution_total
        + savings_interest_total
        + repayment_total
        + fine_paid_total
        + welfare_contribution_total
        - loans_total
        - distribution_total
        - welfare_paid_total
    )
    return render_template('dashboard.html', **locals())

@app.route('/members')
@login_required
@role_required('members')
def members():
    q = request.args.get('q', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 25

    query = Member.query

    if q:
        query = query.filter(
            Member.full_name.contains(q) |
            Member.member_no.contains(q) |
            Member.phone.contains(q)
        )

    pagination = query.order_by(Member.member_no).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    return render_template(
        'members.html',
        members=pagination.items,
        pagination=pagination,
        q=q
    )

@app.route('/members/new', methods=['GET','POST'])
@login_required
@role_required('members')
def member_new():
    if request.method == 'POST':
        m = Member(member_no=request.form['member_no'], full_name=request.form['full_name'], phone=request.form.get('phone'), national_id=request.form.get('national_id'), group_name=request.form.get('group_name'))
        db.session.add(m); db.session.commit(); log_audit('CREATE_MEMBER', 'Member', m.id, f'{m.member_no} - {m.full_name}'); flash('Member added successfully.'); return redirect(url_for('members'))
    return render_template('member_form.html')

@app.route('/members/<int:member_id>/edit', methods=['GET', 'POST'])
@login_required
@role_required('members')
def member_edit(member_id):
    member = Member.query.get_or_404(member_id)

    if request.method == 'POST':
        member.member_no = request.form['member_no'].strip()
        member.full_name = request.form['full_name'].strip()
        member.phone = request.form.get('phone')
        member.national_id = request.form.get('national_id')
        member.group_name = request.form.get('group_name')
        member.status = request.form.get('status') or 'Active'

        db.session.commit()
        log_audit('UPDATE_MEMBER', 'Member', member.id, f'{member.member_no} - {member.full_name}')
        flash('Member updated successfully.')
        return redirect(url_for('members'))

    return render_template('member_form.html', member=member)


@app.route('/members/<int:member_id>/toggle', methods=['POST'])
@login_required
@role_required('members')
def member_toggle(member_id):
    member = Member.query.get_or_404(member_id)

    member.status = 'Inactive' if member.status == 'Active' else 'Active'

    db.session.commit()
    log_audit('TOGGLE_MEMBER_STATUS', 'Member', member.id, f'{member.member_no} status changed to {member.status}')
    flash(f'Member status changed to {member.status}.')
    return redirect(url_for('members'))
@app.route('/members/import', methods=['GET', 'POST'])
@login_required
@role_required('members')
def members_import():
    results = {
        'created': 0,
        'updated': 0,
        'skipped': 0,
        'errors': []
    }

    if request.method == 'POST':
        file = request.files.get('file')

        if not file or file.filename == '':
            flash('Please select a CSV or Excel file to import.', 'error')
            return render_template('members_import.html', results=results)

        filename = file.filename.lower()

        try:
            rows = []

            if filename.endswith('.xlsx'):
                workbook = load_workbook(file)
                sheet = workbook.active

                headers = [
                    str(cell.value or '').strip().lower()
                    for cell in sheet[1]
                ]

                required = ['member_no', 'full_name']
                missing = [col for col in required if col not in headers]

                if missing:
                    flash(f'Missing required columns: {", ".join(missing)}', 'error')
                    return render_template('members_import.html', results=results)

                for line_no, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                    row = dict(zip(headers, values))
                    rows.append((line_no, row))

            elif filename.endswith('.csv'):
                raw = file.stream.read()

                try:
                    decoded = raw.decode('utf-8-sig')
                except UnicodeDecodeError:
                    decoded = raw.decode('latin-1')

                stream = io.StringIO(decoded, newline=None)
                reader = csv.DictReader(stream)

                required = ['member_no', 'full_name']
                missing = [col for col in required if col not in reader.fieldnames]

                if missing:
                    flash(f'Missing required columns: {", ".join(missing)}', 'error')
                    return render_template('members_import.html', results=results)

                for line_no, row in enumerate(reader, start=2):
                    rows.append((line_no, row))

            else:
                flash('Unsupported file type. Please upload .xlsx or .csv file.', 'error')
                return render_template('members_import.html', results=results)

            for line_no, row in rows:
                member_no = str(row.get('member_no') or '').strip()
                full_name = str(row.get('full_name') or '').strip()

                if not member_no or not full_name:
                    results['skipped'] += 1
                    results['errors'].append(f'Line {line_no}: member_no and full_name are required.')
                    continue

               
                phone = str(row.get('phone') or '').strip()
                national_id = str(row.get('national_id') or '').strip()
                group_name = str(row.get('group_name') or '').strip()
                status = str(row.get('status') or 'Active').strip() or 'Active'

                member = Member.query.filter_by(member_no=member_no).first()

                if member:
                    member.full_name = full_name
                    member.phone = phone
                    member.national_id = national_id
                    member.group_name = group_name
                    member.status = status
                    results['updated'] += 1
                else:
                    member = Member(
                        member_no=member_no,
                        full_name=full_name,
                        phone=phone,
                        national_id=national_id,
                        group_name=group_name,
                        status=status
                    )
                    db.session.add(member)
                    results['created'] += 1

                
            db.session.commit()

            log_audit(
                'IMPORT_MEMBERS',
                'Member',
                None,
                f'Bulk member import completed. Created: {results["created"]}, Updated: {results["updated"]}, Skipped: {results["skipped"]}'
            )

            flash(
                f'Import complete. Created {results["created"]}, '
                f'updated {results["updated"]}, '
                f'skipped {results["skipped"]}.'
            )
            return render_template('members_import.html', results=results)

        except Exception as e:
            db.session.rollback()
            flash(f'Import failed: {str(e)}', 'error')

    return render_template('members_import.html', results=results)

@app.route('/contributions', methods=['GET','POST'])
@login_required
@role_required('contributions')
def contributions():
    if request.method == 'POST':
        c = Contribution(member_id=int(request.form['member_id']), month=request.form['month'], amount=money(request.form['amount']), method=request.form['method'], reference=request.form.get('reference'), paid_on=parse_date(request.form.get('paid_on')))
        db.session.add(c); db.session.commit(); log_audit('RECORD_CONTRIBUTION', 'Contribution', c.id, f'{c.member.full_name} paid {kwacha(c.amount)} for {c.month} via {c.method}'); flash('Contribution recorded.'); return redirect(url_for('contributions'))
    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = Contribution.query.order_by(
                Contribution.paid_on.desc()
            ).paginate(
                page=page,
                per_page=per_page,
                error_out=False
            )

    return render_template(
                'contributions.html',
                contributions=pagination.items,
                pagination=pagination,
                members=Member.query.order_by(Member.full_name).all()
            )

@app.route('/loans', methods=['GET','POST'])
@login_required
@role_required('loans')
def loans():
    if request.method == 'POST':
        issued = parse_date(request.form.get('issued_on'))
        due = parse_date(request.form.get('due_on')) or issued + timedelta(days=30)

        l = Loan(
            member_id=int(request.form['member_id']),
            principal=money(request.form['principal']),
            due_on=due,
            issued_on=issued,
            purpose=request.form.get('purpose'),
            status='Applied'
        )

        db.session.add(l)
        db.session.commit()

        log_audit(
            'LOAN_APPLICATION_CREATED',
            'Loan',
            l.id,
            f'{l.member.full_name} applied for {kwacha(l.principal)}; due {l.due_on}'
        )

        flash('Loan application created. It must be reviewed, approved, then disbursed before repayment can be recorded.')
        return redirect(url_for('loans'))

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = Loan.query.order_by(
        Loan.issued_on.desc(),
        Loan.id.desc()
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    return render_template(
        'loans.html',
        loans=pagination.items,
        pagination=pagination,
        members=Member.query.order_by(Member.full_name).all()
    )

@app.route('/loans/<int:loan_id>/review', methods=['POST'])
@login_required
@role_required('loans')
def loan_review(loan_id):
    loan = Loan.query.get_or_404(loan_id)
    if loan.status != 'Applied':
        flash('Only Applied loans can be marked as reviewed.', 'error')
    else:
        user = session.get('user') or {}
        loan.status = 'Reviewed'
        loan.reviewed_by = user.get('full_name') or user.get('username')
        loan.reviewed_on = date.today()
        db.session.commit(); log_audit('LOAN_REVIEWED', 'Loan', loan.id, f'{loan.member.full_name} loan reviewed by {loan.reviewed_by}'); flash('Loan application marked as reviewed.')
    return redirect(url_for('loans'))

@app.route('/loans/<int:loan_id>/approve', methods=['POST'])
@login_required
@role_required('loans')
def loan_approve(loan_id):
    loan = Loan.query.get_or_404(loan_id)
    if loan.status not in ['Applied', 'Reviewed']:
        flash('Only Applied or Reviewed loans can be approved.', 'error')
    else:
        user = session.get('user') or {}
        loan.status = 'Approved'
        loan.approved_by = request.form.get('approved_by') or user.get('full_name') or 'Management Committee'
        loan.approved_on = date.today()
        db.session.commit(); log_audit('LOAN_APPROVED', 'Loan', loan.id, f'{loan.member.full_name} loan of {kwacha(loan.principal)} approved by {loan.approved_by}'); flash('Loan approved. It is ready for disbursement.')
    return redirect(url_for('loans'))

@app.route('/loans/<int:loan_id>/reject', methods=['POST'])
@login_required
@role_required('loans')
def loan_reject(loan_id):
    loan = Loan.query.get_or_404(loan_id)
    if loan.status in ['Disbursed', 'Open', 'Closed']:
        flash('A disbursed or closed loan cannot be rejected.', 'error')
    else:
        loan.status = 'Rejected'
        loan.rejected_on = date.today()
        loan.rejection_reason = request.form.get('reason') or 'Not specified'
        db.session.commit(); log_audit('LOAN_REJECTED', 'Loan', loan.id, f'{loan.member.full_name} loan rejected. Reason: {loan.rejection_reason}'); flash('Loan application rejected.')
    return redirect(url_for('loans'))

@app.route('/loans/<int:loan_id>/disburse', methods=['POST'])
@login_required
@role_required('loans')
def loan_disburse(loan_id):
    loan = Loan.query.get_or_404(loan_id)
    if loan.status != 'Approved':
        flash('Only Approved loans can be disbursed.', 'error')
    else:
        user = session.get('user') or {}
        loan.status = 'Disbursed'
        loan.disbursed_by = user.get('full_name') or user.get('username')
        loan.disbursed_on = date.today()
        db.session.commit(); log_audit('LOAN_DISBURSED', 'Loan', loan.id, f'{loan.member.full_name} loan of {kwacha(loan.principal)} disbursed by {loan.disbursed_by}'); flash('Loan disbursed. Repayments can now be recorded.')
    return redirect(url_for('loans'))

@app.route('/repayments', methods=['POST'])
@login_required
@role_required('repayments')
def repayments():
    loan = Loan.query.get_or_404(int(request.form['loan_id']))
    if loan.status not in ['Open', 'Disbursed']:
        flash('Repayment can only be recorded for disbursed loans.', 'error')
        return redirect(url_for('loans'))
    r = Repayment(loan_id=loan.id, amount=money(request.form['amount']), method=request.form['method'], reference=request.form.get('reference'), paid_on=parse_date(request.form.get('paid_on')))
    db.session.add(r); db.session.commit(); log_audit('RECORD_REPAYMENT', 'Repayment', r.id, f'{loan.member.full_name} paid {kwacha(r.amount)} on loan #{loan.id} via {r.method}')
    if loan.balance <= 0:
        loan.status = 'Closed'; db.session.commit(); log_audit('CLOSE_LOAN', 'Loan', loan.id, f'Loan for {loan.member.full_name} fully repaid')
    flash('Repayment recorded.')
    return redirect(url_for('loans'))

@app.route('/distributions', methods=['GET','POST'])
@login_required
@role_required('distributions')
def distributions():
    if request.method == 'POST':
        d = Distribution(
            member_id=int(request.form['member_id']),
            amount=money(request.form['amount']),
            method=request.form['method'],
            reference=request.form.get('reference'),
            authorized_by=request.form.get('authorized_by'),
            paid_on=parse_date(request.form.get('paid_on'))
        )

        db.session.add(d)
        db.session.commit()

        log_audit(
            'RECORD_DISTRIBUTION',
            'Distribution',
            d.id,
            f'{d.member.full_name} received {kwacha(d.amount)} via {d.method}'
        )

        flash('Distribution recorded.')
        return redirect(url_for('distributions'))

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = Distribution.query.order_by(
        Distribution.paid_on.desc(),
        Distribution.id.desc()
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    return render_template(
        'distributions.html',
        distributions=pagination.items,
        pagination=pagination,
        members=Member.query.order_by(Member.full_name).all()
    )

@app.route('/fines', methods=['GET','POST'])
@login_required
@role_required('fines')
def fines():
    if request.method == 'POST':
        user = session.get('user') or {}
        f = FinePenalty(
            member_id=int(request.form['member_id']),
            category=request.form['category'],
            amount=money(request.form['amount']),
            reason=request.form.get('reason'),
            fine_date=parse_date(request.form.get('fine_date')),
            recorded_by=user.get('full_name') or user.get('username')
        )
        db.session.add(f); db.session.commit()
        log_audit('RECORD_FINE', 'FinePenalty', f.id, f'{f.member.full_name} fined {kwacha(f.amount)} for {f.category}')
        flash('Fine / penalty recorded successfully.')
        return redirect(url_for('fines'))
    q = request.args.get('q','').strip()
    status = request.args.get('status','').strip()
    query = FinePenalty.query.join(Member)
    if q:
        like = f'%{q}%'
        query = query.filter(db.or_(Member.full_name.like(like), Member.member_no.like(like), FinePenalty.category.like(like), FinePenalty.reason.like(like)))
    if status:
        query = query.filter(FinePenalty.status == status)
    all_matching_fines = query.order_by(
    FinePenalty.fine_date.desc(),
    FinePenalty.id.desc()
        ).all()

    total_fines = money(sum((f.amount for f in all_matching_fines), Decimal('0.00')))
    total_paid = money(sum((f.total_paid for f in all_matching_fines), Decimal('0.00')))
    total_balance = money(sum((f.balance for f in all_matching_fines if f.status != 'Waived'), Decimal('0.00')))

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = query.order_by(
        FinePenalty.fine_date.desc(),
        FinePenalty.id.desc()
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    return render_template(
        'fines.html',
        fines=pagination.items,
        pagination=pagination,
        members=Member.query.order_by(Member.full_name).all(),
        q=q,
        status=status,
        total_fines=total_fines,
        total_paid=total_paid,
        total_balance=total_balance
    )

@app.route('/fines/<int:fine_id>/pay', methods=['POST'])
@login_required
@role_required('fines')
def fine_payment(fine_id):
    fine = FinePenalty.query.get_or_404(fine_id)
    if fine.status == 'Waived':
        flash('Cannot pay a waived fine.', 'error')
        return redirect(url_for('fines'))
    payment = FinePayment(fine_id=fine.id, amount=money(request.form['amount']), method=request.form['method'], reference=request.form.get('reference'), paid_on=parse_date(request.form.get('paid_on')))
    if payment.amount <= 0:
        flash('Payment amount must be greater than zero.', 'error')
        return redirect(url_for('fines'))
    if payment.amount > fine.balance:
        flash('Payment cannot exceed the outstanding fine balance.', 'error')
        return redirect(url_for('fines'))
    db.session.add(payment); db.session.commit()
    fine.status = 'Paid' if fine.balance <= 0 else 'Partially Paid'
    db.session.commit()
    log_audit('RECORD_FINE_PAYMENT', 'FinePayment', payment.id, f'{fine.member.full_name} paid {kwacha(payment.amount)} for fine #{fine.id} via {payment.method}')
    flash('Fine payment recorded.')
    return redirect(url_for('fines'))

@app.route('/fines/<int:fine_id>/waive', methods=['POST'])
@login_required
@role_required('fines')
def fine_waive(fine_id):
    fine = FinePenalty.query.get_or_404(fine_id)
    user = session.get('user') or {}
    if fine.status == 'Paid':
        flash('Paid fines cannot be waived.', 'error')
    else:
        fine.status = 'Waived'
        fine.waived_by = user.get('full_name') or user.get('username')
        fine.waived_on = date.today()
        db.session.commit()
        log_audit('WAIVE_FINE', 'FinePenalty', fine.id, f'Fine #{fine.id} for {fine.member.full_name} waived by {fine.waived_by}. Reason: {request.form.get("waiver_reason") or "Not specified"}')
        flash('Fine waived successfully.')
    return redirect(url_for('fines'))

@app.route('/export/fines.csv')
@login_required
@role_required('fines')
def export_fines_csv():
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(['Date','Member No','Full Name','Category','Reason','Fine Amount','Paid','Balance','Status','Recorded By'])
    for f in FinePenalty.query.order_by(FinePenalty.fine_date.desc()).all():
        writer.writerow([f.fine_date, f.member.member_no, f.member.full_name, f.category, f.reason, f.amount, f.total_paid, f.balance, f.status, f.recorded_by])
    log_audit('EXPORT_FINES', 'FinePenalty', None, 'Fines and penalties exported to CSV')
    return Response(output.getvalue(), mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=fines_penalties.csv'})


@app.route('/welfare', methods=['GET','POST'])
@login_required
@role_required('welfare')
def welfare():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'contribution':
            c = WelfareContribution(
                member_id=int(request.form['member_id']),
                month=request.form['month'],
                amount=money(request.form['amount']),
                method=request.form['method'],
                reference=request.form.get('reference'),
                paid_on=parse_date(request.form.get('paid_on'))
            )
            db.session.add(c); db.session.commit()
            log_audit('RECORD_WELFARE_CONTRIBUTION', 'WelfareContribution', c.id, f'{c.member.full_name} paid welfare contribution {kwacha(c.amount)} for {c.month}')
            flash('Welfare contribution recorded.')
        elif action == 'claim':
            claim = WelfareClaim(
                member_id=int(request.form['member_id']),
                category=request.form['category'],
                amount_requested=money(request.form['amount_requested']),
                reason=request.form.get('reason'),
                requested_on=parse_date(request.form.get('requested_on'))
            )
            db.session.add(claim); db.session.commit()
            log_audit('CREATE_WELFARE_CLAIM', 'WelfareClaim', claim.id, f'{claim.member.full_name} requested {kwacha(claim.amount_requested)} for {claim.category}')
            flash('Welfare claim request recorded.')
        return redirect(url_for('welfare'))

    q = request.args.get('q','').strip()
    status = request.args.get('status','').strip()
    claim_query = WelfareClaim.query.join(Member)
    if q:
        like = f'%{q}%'
        claim_query = claim_query.filter(db.or_(Member.full_name.like(like), Member.member_no.like(like), WelfareClaim.category.like(like), WelfareClaim.reason.like(like)))
    if status:
        claim_query = claim_query.filter(WelfareClaim.status == status)
    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = claim_query.order_by(
        WelfareClaim.requested_on.desc(),
        WelfareClaim.id.desc()
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    claims = pagination.items

    contributions = WelfareContribution.query.order_by(
    WelfareContribution.paid_on.desc(),
    WelfareContribution.id.desc()
    ).limit(100).all()
    members = Member.query.order_by(Member.full_name).all()
    total_contributions = money(db.session.query(db.func.coalesce(db.func.sum(WelfareContribution.amount), 0)).scalar())
    total_paid = money(db.session.query(db.func.coalesce(db.func.sum(WelfareClaim.amount_approved), 0)).filter(WelfareClaim.status == 'Paid').scalar())
    approved_not_paid = money(db.session.query(db.func.coalesce(db.func.sum(WelfareClaim.amount_approved), 0)).filter(WelfareClaim.status == 'Approved').scalar())
    balance = money(total_contributions - total_paid)
    return render_template(
        'welfare.html',
        claims=claims,
        contributions=contributions,
        members=members,
        pagination=pagination,
        q=q,
        status=status,
        total_contributions=total_contributions,
        total_paid=total_paid,
        approved_not_paid=approved_not_paid,
        balance=balance
    )

@app.route('/welfare/claims/<int:claim_id>/review', methods=['POST'])
@login_required
@role_required('welfare')
def welfare_review(claim_id):
    claim = WelfareClaim.query.get_or_404(claim_id)
    if claim.status != 'Requested':
        flash('Only requested claims can be marked as reviewed.', 'error')
    else:
        user = session.get('user') or {}
        claim.status = 'Reviewed'
        claim.reviewed_by = user.get('full_name') or user.get('username')
        claim.reviewed_on = date.today()
        db.session.commit()
        log_audit('WELFARE_CLAIM_REVIEWED', 'WelfareClaim', claim.id, f'{claim.member.full_name} claim reviewed by {claim.reviewed_by}')
        flash('Welfare claim marked as reviewed.')
    return redirect(url_for('welfare'))

@app.route('/welfare/claims/<int:claim_id>/approve', methods=['POST'])
@login_required
@role_required('welfare')
def welfare_approve(claim_id):
    claim = WelfareClaim.query.get_or_404(claim_id)
    if claim.status not in ['Requested', 'Reviewed']:
        flash('Only requested or reviewed claims can be approved.', 'error')
    else:
        user = session.get('user') or {}
        claim.status = 'Approved'
        claim.amount_approved = money(request.form.get('amount_approved') or claim.amount_requested)
        claim.approved_by = request.form.get('approved_by') or user.get('full_name') or 'Management Committee'
        claim.approved_on = date.today()
        db.session.commit()
        log_audit('WELFARE_CLAIM_APPROVED', 'WelfareClaim', claim.id, f'{claim.member.full_name} welfare claim approved for {kwacha(claim.amount_approved)} by {claim.approved_by}')
        flash('Welfare claim approved. It is ready for payment.')
    return redirect(url_for('welfare'))

@app.route('/welfare/claims/<int:claim_id>/pay', methods=['POST'])
@login_required
@role_required('welfare')
def welfare_pay(claim_id):
    claim = WelfareClaim.query.get_or_404(claim_id)
    if claim.status != 'Approved':
        flash('Only approved welfare claims can be paid.', 'error')
    else:
        user = session.get('user') or {}
        claim.status = 'Paid'
        claim.paid_by = user.get('full_name') or user.get('username')
        claim.paid_on = parse_date(request.form.get('paid_on'))
        claim.payment_method = request.form.get('method')
        claim.reference = request.form.get('reference')
        db.session.commit()
        log_audit('WELFARE_CLAIM_PAID', 'WelfareClaim', claim.id, f'{claim.member.full_name} welfare claim paid {kwacha(claim.amount_approved)} via {claim.payment_method}')
        flash('Welfare claim paid successfully.')
    return redirect(url_for('welfare'))

@app.route('/welfare/claims/<int:claim_id>/reject', methods=['POST'])
@login_required
@role_required('welfare')
def welfare_reject(claim_id):
    claim = WelfareClaim.query.get_or_404(claim_id)
    if claim.status in ['Paid']:
        flash('Paid welfare claims cannot be rejected.', 'error')
    else:
        claim.status = 'Rejected'
        claim.rejection_reason = request.form.get('reason') or 'Not specified'
        db.session.commit()
        log_audit('WELFARE_CLAIM_REJECTED', 'WelfareClaim', claim.id, f'{claim.member.full_name} welfare claim rejected. Reason: {claim.rejection_reason}')
        flash('Welfare claim rejected.')
    return redirect(url_for('welfare'))

@app.route('/export/welfare.csv')
@login_required
@role_required('welfare')
def export_welfare_csv():
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(['Type','Date','Member No','Full Name','Category/Month','Requested/Contribution','Approved','Status','Method','Reference'])
    for c in WelfareContribution.query.order_by(WelfareContribution.paid_on.desc()).all():
        writer.writerow(['Contribution', c.paid_on, c.member.member_no, c.member.full_name, c.month, c.amount, '', 'Paid', c.method, c.reference])
    for cl in WelfareClaim.query.order_by(WelfareClaim.requested_on.desc()).all():
        writer.writerow(['Claim', cl.requested_on, cl.member.member_no, cl.member.full_name, cl.category, cl.amount_requested, cl.amount_approved, cl.status, cl.payment_method or '', cl.reference or ''])
    log_audit('EXPORT_WELFARE', 'Welfare', None, 'Welfare fund records exported to CSV')
    return Response(output.getvalue(), mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=welfare_fund.csv'})

@app.route('/meetings', methods=['GET','POST'])
@login_required
@role_required('meetings')
def meetings():
    if request.method == 'POST':
        mtg = Meeting(meeting_type=request.form['meeting_type'], meeting_date=parse_date(request.form['meeting_date']), agenda=request.form.get('agenda'), resolutions=request.form.get('resolutions'), attendance_count=int(request.form.get('attendance_count') or 0))
        db.session.add(mtg); db.session.commit(); log_audit('SAVE_MEETING', 'Meeting', mtg.id, f'{mtg.meeting_type} meeting on {mtg.meeting_date}; attendance {mtg.attendance_count}'); flash('Meeting record saved.'); return redirect(url_for('meetings'))
    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = Meeting.query.order_by(
    Meeting.meeting_date.desc(),
    Meeting.id.desc()
    ).paginate(
    page=page,
    per_page=per_page,
    error_out=False
    )

    return render_template(
    'meetings.html',
    meetings=pagination.items,
    pagination=pagination
)



@app.route('/attendance')
@login_required
@role_required('attendance')
def attendance():
    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = Meeting.query.order_by(
        Meeting.meeting_date.desc(),
        Meeting.id.desc()
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    meetings_list = pagination.items
    summary = []
    total_members = Member.query.filter_by(status='Active').count()

    for mtg in meetings_list:
        records = MeetingAttendance.query.filter_by(meeting_id=mtg.id).all()
        present = sum(1 for r in records if r.status == 'Present')
        late = sum(1 for r in records if r.status == 'Late')
        absent = sum(1 for r in records if r.status == 'Absent')
        excused = sum(1 for r in records if r.status == 'Excused')

        summary.append({
            'meeting': mtg,
            'total': len(records),
            'present': present,
            'late': late,
            'absent': absent,
            'excused': excused,
            'expected': total_members
        })

    return render_template(
        'attendance.html',
        summary=summary,
        pagination=pagination
    )

@app.route('/attendance/<int:meeting_id>', methods=['GET', 'POST'])
@login_required
@role_required('attendance')
def attendance_register(meeting_id):
    meeting = Meeting.query.get_or_404(meeting_id)
    members = Member.query.filter_by(status='Active').order_by(Member.member_no).all()
    user = session.get('user') or {}
    if request.method == 'POST':
        absent_fine = Decimal('50.00')
        late_fine = Decimal('20.00')
        create_fines = True
        saved = 0
        fines_created = 0
        for member in members:
            status = request.form.get(f'status_{member.id}') or 'Present'
            remarks = request.form.get(f'remarks_{member.id}', '').strip()
            record = MeetingAttendance.query.filter_by(meeting_id=meeting.id, member_id=member.id).first()
            if not record:
                record = MeetingAttendance(meeting_id=meeting.id, member_id=member.id)
                db.session.add(record)
            record.status = status
            record.remarks = remarks
            record.recorded_by = user.get('full_name') or user.get('username')
            record.recorded_at = datetime.utcnow()
            if create_fines and not record.fine_generated:
                fine_amount = Decimal('0.00')
                category = None
                reason = None
                if status == 'Absent' and absent_fine > 0:
                    fine_amount = absent_fine
                    category = 'Absence from Meeting'
                    reason = f'Absent from {meeting.meeting_type} meeting on {meeting.meeting_date}'
                elif status == 'Late' and late_fine > 0:
                    fine_amount = late_fine
                    category = 'Late Meeting Attendance'
                    reason = f'Late for {meeting.meeting_type} meeting on {meeting.meeting_date}'
                if fine_amount > 0:
                    fine = FinePenalty(member_id=member.id, category=category, amount=fine_amount, reason=reason, fine_date=meeting.meeting_date, recorded_by=record.recorded_by)
                    db.session.add(fine)
                    record.fine_generated = True
                    fines_created += 1
            saved += 1
        meeting.attendance_count = sum(1 for member in members if (request.form.get(f'status_{member.id}') or 'Present') in ['Present', 'Late'])
        db.session.commit()
        log_audit('SAVE_ATTENDANCE', 'Meeting', meeting.id, f'Attendance saved for {meeting.meeting_type} on {meeting.meeting_date}. Records: {saved}; fines created: {fines_created}')
        flash(f'Attendance register saved. {saved} records updated. {fines_created} fines generated.')
        return redirect(url_for('attendance_register', meeting_id=meeting.id))

    existing = {r.member_id: r for r in MeetingAttendance.query.filter_by(meeting_id=meeting.id).all()}
    stats = {'Present':0, 'Absent':0, 'Late':0, 'Excused':0}
    for r in existing.values():
        stats[r.status] = stats.get(r.status, 0) + 1
    return render_template('attendance_register.html', meeting=meeting, members=members, existing=existing, stats=stats)

@app.route('/export/attendance.csv')
@login_required
@role_required('attendance')
def export_attendance_csv():
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(['Meeting Date','Meeting Type','Member No','Full Name','Status','Remarks','Recorded By','Recorded At','Fine Generated'])
    rows = MeetingAttendance.query.join(Meeting).join(Member).order_by(Meeting.meeting_date.desc(), Member.member_no).all()
    for r in rows:
        writer.writerow([r.meeting.meeting_date, r.meeting.meeting_type, r.member.member_no, r.member.full_name, r.status, r.remarks or '', r.recorded_by or '', r.recorded_at, 'Yes' if r.fine_generated else 'No'])
    log_audit('EXPORT_ATTENDANCE', 'MeetingAttendance', None, 'Attendance register exported to CSV')
    return Response(output.getvalue(), mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=meeting_attendance.csv'})

@app.route('/reports')
@login_required
@role_required('reports')
def reports():
    month = request.args.get('month') or date.today().strftime('%Y-%m')

    contribs = Contribution.query.filter_by(month=month).all()
    paid_member_ids = {c.member_id for c in contribs}

    arrears_query = Member.query.filter(
        Member.id.notin_(paid_member_ids)
    ) if paid_member_ids else Member.query

    page = request.args.get('page', 1, type=int)
    per_page = 25

    arrears_pagination = arrears_query.order_by(
        Member.member_no
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    open_loans = Loan.query.filter(
        Loan.status.in_(['Open', 'Disbursed'])
    ).order_by(
        Loan.issued_on.desc(),
        Loan.id.desc()
    ).all()

    return render_template(
        'reports.html',
        month=month,
        contribs=contribs,
        arrears=arrears_pagination.items,
        arrears_pagination=arrears_pagination,
        open_loans=open_loans
    )


@app.route('/accounting', methods=['GET', 'POST'])
@login_required
@role_required('accounting')
def accounting():
    seed_chart_of_accounts()
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'sync':
            created = sync_operational_transactions_to_gl()
            log_audit('SYNC_ACCOUNTING', 'JournalEntry', None, f'{created} operational transaction(s) posted to the general ledger')
            flash(f'{created} operational transaction(s) posted to the General Ledger.')
            return redirect(url_for('accounting'))
        if action == 'account':
            code = request.form.get('code','').strip()
            name = request.form.get('name','').strip()
            account_type = request.form.get('account_type','Asset')
            normal_balance = 'Debit' if account_type in ['Asset','Expense'] else 'Credit'
            if not code or not name:
                flash('Account code and name are required.', 'error')
            elif Account.query.filter_by(code=code).first():
                flash('That account code already exists.', 'error')
            else:
                db.session.add(Account(code=code, name=name, account_type=account_type, normal_balance=normal_balance))
                db.session.commit()
                log_audit('CREATE_ACCOUNT', 'Account', code, f'Account {code} - {name} created')
                flash('Account created successfully.')
            return redirect(url_for('accounting'))
        if action == 'journal':
            entry_date = parse_date(request.form.get('entry_date'))
            description = request.form.get('description','').strip()
            reference = request.form.get('reference','').strip()
            debit_account = Account.query.get(int(request.form.get('debit_account_id') or 0))
            credit_account = Account.query.get(int(request.form.get('credit_account_id') or 0))
            amount = money(request.form.get('amount') or 0)
            if not description or not debit_account or not credit_account or amount <= 0:
                flash('Description, debit account, credit account and amount are required.', 'error')
            else:
                entry = post_journal(entry_date, description, reference, 'Manual', f'{datetime.utcnow().timestamp()}', [
                    {'account': debit_account, 'debit': amount},
                    {'account': credit_account, 'credit': amount},
                ])
                log_audit('POST_MANUAL_JOURNAL', 'JournalEntry', entry.id if entry else None, f'Manual journal posted: {description}')
                flash('Manual journal entry posted successfully.')
            return redirect(url_for('accounting'))

    start = request.args.get('start')
    end = request.args.get('end')
    start_date = parse_date(start) if start else None
    end_date = parse_date(end) if end else None
    accounts = Account.query.order_by(Account.code).all()

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = JournalEntry.query.order_by(
        JournalEntry.entry_date.desc(),
        JournalEntry.id.desc()
    ).paginate(
    page=page,
    per_page=per_page,
    error_out=False
    )
    entries = pagination.items
    balances = ledger_balances(start_date, end_date)
    total_debits = money(sum((b['debit'] for b in balances), Decimal('0.00')))
    total_credits = money(sum((b['credit'] for b in balances), Decimal('0.00')))
    income_total = money(sum((b['balance'] for b in balances if b['account'].account_type == 'Income'), Decimal('0.00')))
    expense_total = money(sum((b['balance'] for b in balances if b['account'].account_type == 'Expense'), Decimal('0.00')))
    net_surplus = money(income_total - expense_total)
    asset_total = money(sum((b['balance'] for b in balances if b['account'].account_type == 'Asset'), Decimal('0.00')))
    liability_total = money(sum((b['balance'] for b in balances if b['account'].account_type == 'Liability'), Decimal('0.00')))
    equity_total = money(sum((b['balance'] for b in balances if b['account'].account_type == 'Equity'), Decimal('0.00')))
    cashbook = [b for b in balances if b['account'].code in ['1000','1010','1020']]
    log_audit('VIEW_ACCOUNTING', 'Accounting', None, 'General Ledger and Accounting viewed')
    return render_template('accounting.html', **locals())

@app.route('/accounting/trial-balance')
@login_required
@role_required('accounting')
def trial_balance():
    start = request.args.get('start')
    end = request.args.get('end')

    start_date = parse_date(start) if start else None
    end_date = parse_date(end) if end else None

    balances = ledger_balances(start_date, end_date)

    total_debits = money(sum((b['debit'] for b in balances), Decimal('0.00')))
    total_credits = money(sum((b['credit'] for b in balances), Decimal('0.00')))

    return render_template(
        'trial_balance.html',
        balances=balances,
        total_debits=total_debits,
        total_credits=total_credits,
        start=start,
        end=end
    )

@app.route('/accounting/income-statement')
@login_required
@role_required('accounting')
def income_statement():
    start = request.args.get('start')
    end = request.args.get('end')

    start_date = parse_date(start) if start else None
    end_date = parse_date(end) if end else None

    loan_interest_income = money(
        db.session.query(
            db.func.coalesce(db.func.sum(LoanInterest.interest_amount), 0)
        ).scalar()
    )

    fines_income = money(
        db.session.query(
            db.func.coalesce(db.func.sum(FinePayment.amount), 0)
        ).scalar()
    )

    welfare_expense = money(
        db.session.query(
            db.func.coalesce(db.func.sum(WelfareClaim.amount_approved), 0)
        )
        .filter(WelfareClaim.status == 'Paid')
        .scalar()
    )

    shareout_expense = money(
        db.session.query(
            db.func.coalesce(db.func.sum(Distribution.amount), 0)
        ).scalar()
    )

    total_income = money(
        loan_interest_income +
        fines_income
    )

    total_expenses = money(
        welfare_expense +
        shareout_expense
    )

    net_surplus = money(total_income - total_expenses)

    return render_template(
        'income_statement.html',
        start=start,
        end=end,
        loan_interest_income=loan_interest_income,
        fines_income=fines_income,
        welfare_expense=welfare_expense,
        shareout_expense=shareout_expense,
        total_income=total_income,
        total_expenses=total_expenses,
        net_surplus=net_surplus
    )

@app.route('/accounting/balance-sheet')
@login_required
@role_required('accounting')
def balance_sheet():

    total_savings = money(
        db.session.query(
            db.func.coalesce(db.func.sum(Contribution.amount), 0)
        ).scalar()
    )

    savings_interest = money(
        db.session.query(
            db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0)
        ).scalar()
    )

    welfare_contributions = money(
        db.session.query(
            db.func.coalesce(db.func.sum(WelfareContribution.amount), 0)
        ).scalar()
    )

    welfare_paid = money(
        db.session.query(
            db.func.coalesce(db.func.sum(WelfareClaim.amount_approved), 0)
        )
        .filter(WelfareClaim.status == 'Paid')
        .scalar()
    )

    welfare_balance = money(
        welfare_contributions - welfare_paid
    )

    loans_outstanding = money(
        sum(
            (loan.balance for loan in Loan.query.all()),
            Decimal('0.00')
        )
    )

    fines_outstanding = money(
        sum(
            (
                fine.balance
                for fine in FinePenalty.query.all()
                if fine.status != 'Waived'
            ),
            Decimal('0.00')
        )
    )

    member_equity = money(
        total_savings + savings_interest
    )

    liabilities_equity = money(
        member_equity + welfare_balance
    )

    cash_balance = money(
        liabilities_equity
        - loans_outstanding
        - fines_outstanding
    )

    total_assets = money(
        cash_balance
        + loans_outstanding
        + fines_outstanding
    )

    return render_template(
        'balance_sheet.html',
        cash_balance=cash_balance,
        loans_outstanding=loans_outstanding,
        fines_outstanding=fines_outstanding,
        total_assets=total_assets,
        total_savings=total_savings,
        savings_interest=savings_interest,
        welfare_balance=welfare_balance,
        member_equity=member_equity,
        liabilities_equity=liabilities_equity
    )
@app.route('/accounting/cash-flow')
@login_required
@role_required('accounting')
def cash_flow_statement():

    contributions_cash = money(
        db.session.query(db.func.coalesce(db.func.sum(Contribution.amount), 0)).scalar()
    )

    loan_repayments_cash = money(
        db.session.query(db.func.coalesce(db.func.sum(Repayment.amount), 0)).scalar()
    )

    fine_payments_cash = money(
        db.session.query(db.func.coalesce(db.func.sum(FinePayment.amount), 0)).scalar()
    )

    welfare_contributions_cash = money(
        db.session.query(db.func.coalesce(db.func.sum(WelfareContribution.amount), 0)).scalar()
    )

    loan_disbursements_cash = money(
        db.session.query(db.func.coalesce(db.func.sum(Loan.principal), 0)).scalar()
    )

    distributions_cash = money(
        db.session.query(db.func.coalesce(db.func.sum(Distribution.amount), 0)).scalar()
    )

    welfare_payments_cash = money(
        db.session.query(db.func.coalesce(db.func.sum(WelfareClaim.amount_approved), 0))
        .filter(WelfareClaim.status == 'Paid')
        .scalar()
    )

    total_cash_in = money(
        contributions_cash
        + loan_repayments_cash
        + fine_payments_cash
        + welfare_contributions_cash
    )

    total_cash_out = money(
        loan_disbursements_cash
        + distributions_cash
        + welfare_payments_cash
    )

    net_cash_flow = money(total_cash_in - total_cash_out)

    return render_template(
        'cash_flow.html',
        contributions_cash=contributions_cash,
        loan_repayments_cash=loan_repayments_cash,
        fine_payments_cash=fine_payments_cash,
        welfare_contributions_cash=welfare_contributions_cash,
        loan_disbursements_cash=loan_disbursements_cash,
        distributions_cash=distributions_cash,
        welfare_payments_cash=welfare_payments_cash,
        total_cash_in=total_cash_in,
        total_cash_out=total_cash_out,
        net_cash_flow=net_cash_flow
    )

@app.route('/export/accounting.csv')
@login_required
@role_required('accounting')
def export_accounting_csv():
    seed_chart_of_accounts()
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(['Entry Date','Reference','Description','Account Code','Account Name','Debit','Credit','Posted By'])
    for entry in JournalEntry.query.order_by(JournalEntry.entry_date.desc(), JournalEntry.id.desc()).all():
        for line in entry.lines:
            writer.writerow([entry.entry_date, entry.reference or '', entry.description, line.account.code, line.account.name, line.debit, line.credit, entry.posted_by or ''])
    log_audit('EXPORT_ACCOUNTING', 'JournalEntry', None, 'General ledger exported to CSV')
    return Response(output.getvalue(), mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=general_ledger.csv'})

@app.route('/notifications', methods=['GET', 'POST'])
@login_required
@role_required('notifications')
def notifications():
    members = Member.query.order_by(Member.member_no).all()
    q = request.args.get('q', '').strip()

    setting = SystemSetting.query.first()
    organization_name = setting.organization_name if setting and setting.organization_name else CLIENT_NAME

    templates = {
    'Contribution Reminder': f'Dear {{name}}, your monthly contribution is due. Please pay through bank transfer, mobile money, or cash. {organization_name}.',
    'Loan Repayment Reminder': f'Dear {{name}}, this is a reminder to make your loan repayment by the due date. {organization_name}.',
    'Meeting Reminder': f'Dear {{name}}, please remember the upcoming Village Banking meeting. Your attendance is important. {organization_name}.',
    'Welfare Notification': f'Dear {{name}}, your welfare fund update has been recorded. Contact the committee for details. {organization_name}.',
    'Share-Out Notification': f'Dear {{name}}, your share-out/dividend information is ready. Contact the treasurer for your statement. {organization_name}.',
    'General Notice': f'Dear {{name}}, this is a notice from {organization_name}.'
}
    if request.method == 'POST':
        notification_type = request.form.get('notification_type') or 'General Notice'
        channel = request.form.get('channel') or 'SMS'
        subject = request.form.get('subject') or notification_type
        message = request.form.get('message') or ''
        recipient_mode = request.form.get('recipient_mode') or 'selected'
        selected_ids = request.form.getlist('member_ids')

        if recipient_mode == 'all':
            recipients = members
        else:
            ids = [int(x) for x in selected_ids if x.isdigit()]
            recipients = Member.query.filter(Member.id.in_(ids)).all() if ids else []

        if not recipients:
            flash('Please select at least one recipient or choose all members.', 'error')
            return redirect(url_for('notifications'))

        created = 0
        user = session.get('user') or {}

        for m in recipients:
            personalized = (
                message
                .replace('{name}', m.full_name)
                .replace('{member_no}', m.member_no)
                .replace('{phone}', m.phone or '')
            )

            n = NotificationLog(
                channel=channel,
                notification_type=notification_type,
                recipient_type='All Members' if recipient_mode == 'all' else 'Selected Members',
                member_id=m.id,
                phone=m.phone,
                subject=subject,
                message=personalized,
                status='Prepared',
                created_by=user.get('username')
            )

            db.session.add(n)
            created += 1

        db.session.commit()

        log_audit(
            'PREPARE_NOTIFICATIONS',
            'NotificationLog',
            None,
            f'{created} {channel} notification(s) prepared for {notification_type}'
        )

        flash(f'{created} notification(s) prepared. Provider sending can be connected later.')
        return redirect(url_for('notifications'))

    query = NotificationLog.query

    if q:
        query = query.outerjoin(Member).filter(
            (Member.full_name.contains(q)) |
            (Member.member_no.contains(q)) |
            (NotificationLog.phone.contains(q)) |
            (NotificationLog.message.contains(q)) |
            (NotificationLog.notification_type.contains(q))
        )

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = query.order_by(
        NotificationLog.created_at.desc(),
        NotificationLog.id.desc()
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    logs = pagination.items

    return render_template(
        'notifications.html',
        members=members,
        logs=logs,
        pagination=pagination,
        q=q,
        templates=templates
    )
@app.route('/export/notifications.csv')
@login_required
@role_required('notifications')
def export_notifications_csv():
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(['Created At','Channel','Type','Member No','Member','Phone','Subject','Status','Message'])
    for n in NotificationLog.query.order_by(NotificationLog.created_at.desc()).all():
        writer.writerow([n.created_at, n.channel, n.notification_type, n.member.member_no if n.member else '', n.member.full_name if n.member else '', n.phone, n.subject, n.status, n.message])
    log_audit('EXPORT_NOTIFICATIONS', 'NotificationLog', None, 'Notification log exported to CSV')
    return Response(output.getvalue(), mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=notifications.csv'})

@app.route('/users')
@login_required
@role_required('users')
def users():
    return render_template('users.html', users=User.query.order_by(User.full_name).all())

@app.route('/users/new', methods=['GET', 'POST'])
@login_required
@role_required('users')
def user_new():
    if request.method == 'POST':
        username = request.form.get('username','').strip().lower()
        password = request.form.get('password','')
        if not username or not password:
            flash('Username and password are required.', 'error')
            return redirect(url_for('user_new'))
        if User.query.filter_by(username=username).first():
            flash('That username already exists.', 'error')
            return redirect(url_for('user_new'))
        user = User(username=username, full_name=request.form.get('full_name','').strip(), role=request.form.get('role'), active=bool(request.form.get('active')), password_hash=generate_password_hash(password))
        db.session.add(user); db.session.commit(); log_audit('CREATE_USER', 'User', user.id, f'{user.username} created with role {user.role}'); flash('User account created successfully.'); return redirect(url_for('users'))
    return render_template('user_form.html', user=None)

@app.route('/users/<int:user_id>/edit', methods=['GET', 'POST'])
@login_required
@role_required('users')
def user_edit(user_id):
    user = User.query.get_or_404(user_id)
    if request.method == 'POST':
        user.full_name = request.form.get('full_name','').strip()
        user.role = request.form.get('role')
        user.active = bool(request.form.get('active'))
        new_password = request.form.get('password','').strip()
        if new_password:
            user.password_hash = generate_password_hash(new_password)
        db.session.commit(); log_audit('UPDATE_USER', 'User', user.id, f'{user.username} updated; role {user.role}; active {user.active}'); flash('User account updated successfully.'); return redirect(url_for('users'))
    return render_template('user_form.html', user=user)

@app.route('/users/<int:user_id>/toggle', methods=['POST'])
@login_required
@role_required('users')
def user_toggle(user_id):
    user = User.query.get_or_404(user_id)
    if user.id == session.get('user', {}).get('id'):
        flash('You cannot deactivate your own account while logged in.', 'error')
    else:
        user.active = not user.active
        db.session.commit()
        log_audit('TOGGLE_USER_STATUS', 'User', user.id, f'{user.username} active={user.active}')
        flash('User status updated.')
    return redirect(url_for('users'))



@app.route('/backups')
@login_required
@role_required('backups')
def backups():
    folder = backups_folder()
    disk_files = []
    for f in sorted(Path(folder).glob('*.db'), key=lambda x: x.stat().st_mtime, reverse=True):
        disk_files.append({'filename': f.name, 'size': f.stat().st_size, 'modified': datetime.fromtimestamp(f.stat().st_mtime)})
    records = BackupRecord.query.order_by(BackupRecord.created_at.desc()).limit(100).all()
    return render_template('backups.html', disk_files=disk_files, records=records)

@app.route('/backups/create', methods=['POST'])
@login_required
@role_required('backups')
def backup_create():
    try:
        notes = request.form.get('notes') or 'Manual backup created from Administration menu'
        record = create_database_backup('backup', notes=notes)
        log_audit('CREATE_BACKUP', 'BackupRecord', record.id, f'Backup created: {record.filename}')
        flash(f'Backup created successfully: {record.filename}')
    except Exception as exc:
        flash(f'Backup failed: {exc}', 'error')
    return redirect(url_for('backups'))

@app.route('/backups/download/<path:filename>')
@login_required
@role_required('backups')
def backup_download(filename):
    safe_name = secure_filename(filename)
    path = os.path.join(backups_folder(), safe_name)
    if not os.path.exists(path):
        flash('Backup file not found.', 'error')
        return redirect(url_for('backups'))
    log_audit('DOWNLOAD_BACKUP', 'Backup', safe_name, f'Backup downloaded: {safe_name}')
    return send_file(path, as_attachment=True, download_name=safe_name)

@app.route('/backups/restore', methods=['POST'])
@login_required
@role_required('backups')
def backup_restore():
    uploaded = request.files.get('backup_file')
    confirm = request.form.get('confirm_restore') == 'YES'
    if not uploaded or not uploaded.filename:
        flash('Please choose a backup file to restore.', 'error')
        return redirect(url_for('backups'))
    if not confirm:
        flash('Restore cancelled. Type YES in the confirmation box before restoring.', 'error')
        return redirect(url_for('backups'))
    filename = secure_filename(uploaded.filename)
    if not filename.lower().endswith('.db'):
        flash('Only .db backup files are allowed.', 'error')
        return redirect(url_for('backups'))
    try:
        # Safety backup before restore.
        safety = create_database_backup('pre_restore', notes=f'Safety backup before restoring {filename}')
        temp_path = os.path.join(backups_folder(), f'upload_restore_{datetime.now().strftime("%Y_%m_%d_%H%M%S")}_{filename}')
        uploaded.save(temp_path)
        db.session.remove()
        db.engine.dispose()
        shutil.copy2(temp_path, database_file_path())
        log_audit('RESTORE_BACKUP', 'Backup', filename, f'Database restored from {filename}. Safety backup: {safety.filename}')
        flash(f'Database restored from {filename}. Please stop the server with Ctrl+C and run python app.py again.')
    except Exception as exc:
        flash(f'Restore failed: {exc}', 'error')
    return redirect(url_for('backups'))

@app.route('/backups/export-json')
@login_required
@role_required('backups')
def backup_export_json():
    setting = SystemSetting.query.first()
    organization_name = setting.organization_name if setting and setting.organization_name else CLIENT_NAME

    data = {
        'created_at': datetime.utcnow().isoformat(),
        'organization': organization_name,

        'members': [
            {
                'id': m.id,
                'member_no': m.member_no,
                'full_name': m.full_name,
                'phone': m.phone,
                'national_id': m.national_id,
                'group_name': m.group_name,
                'status': m.status,
            }
            for m in Member.query.order_by(Member.id).all()
        ],

        'contributions': [
            {
                'id': c.id,
                'member_id': c.member_id,
                'month': c.month,
                'amount': str(c.amount),
                'method': c.method,
                'reference': c.reference,
                'paid_on': str(c.paid_on),
            }
            for c in Contribution.query.order_by(Contribution.id).all()
        ],

        'loans': [
            {
                'id': l.id,
                'member_id': l.member_id,
                'principal': str(l.principal),
                'interest_amount': str(l.interest_amount),
                'total_due': str(l.total_due),
                'balance': str(l.balance),
                'status': l.status,
                'issued_on': str(l.issued_on),
                'due_on': str(l.due_on),
            }
            for l in Loan.query.order_by(Loan.id).all()
        ],

        'repayments': [
            {
                'id': r.id,
                'loan_id': r.loan_id,
                'amount': str(r.amount),
                'method': r.method,
                'reference': r.reference,
                'paid_on': str(r.paid_on),
            }
            for r in Repayment.query.order_by(Repayment.id).all()
        ],

        'distributions': [
            {
                'id': d.id,
                'member_id': d.member_id,
                'amount': str(d.amount),
                'method': d.method,
                'reference': d.reference,
                'paid_on': str(d.paid_on),
            }
            for d in Distribution.query.order_by(Distribution.id).all()
        ],

        'fines': [
            {
                'id': f.id,
                'member_id': f.member_id,
                'category': f.category,
                'amount': str(f.amount),
                'balance': str(f.balance),
                'status': f.status,
            }
            for f in FinePenalty.query.order_by(FinePenalty.id).all()
        ],

        'fine_payments': [
            {
                'id': fp.id,
                'fine_id': fp.fine_id,
                'amount': str(fp.amount),
                'method': fp.method,
                'reference': fp.reference,
                'paid_on': str(fp.paid_on),
            }
            for fp in FinePayment.query.order_by(FinePayment.id).all()
        ],

        'welfare_contributions': [
            {
                'id': w.id,
                'member_id': w.member_id,
                'month': w.month,
                'amount': str(w.amount),
                'method': w.method,
                'reference': w.reference,
                'paid_on': str(w.paid_on),
            }
            for w in WelfareContribution.query.order_by(WelfareContribution.id).all()
        ],

        'welfare_claims': [
            {
                'id': wc.id,
                'member_id': wc.member_id,
                'category': wc.category,
                'amount_requested': str(wc.amount_requested),
                'amount_approved': str(wc.amount_approved),
                'status': wc.status,
                'reason': wc.reason,
                'requested_on': str(wc.requested_on),
                'paid_on': str(wc.paid_on),
            }
            for wc in WelfareClaim.query.order_by(WelfareClaim.id).all()
        ],

        'savings_interest': [
            {
                'id': s.id,
                'member_id': s.member_id,
                'month': s.month,
                'opening_balance': str(s.opening_balance),
                'interest_rate': str(s.interest_rate),
                'interest_amount': str(s.interest_amount),
                'closing_balance': str(s.closing_balance),
            }
            for s in SavingsInterest.query.order_by(SavingsInterest.id).all()
        ],

        'loan_interest': [
            {
                'id': li.id,
                'loan_id': li.loan_id,
                'member_id': li.member_id,
                'month': li.month,
                'opening_balance': str(li.opening_balance),
                'interest_rate': str(li.interest_rate),
                'interest_amount': str(li.interest_amount),
                'closing_balance': str(li.closing_balance),
            }
            for li in LoanInterest.query.order_by(LoanInterest.id).all()
        ],
    }

    buffer = io.BytesIO()
    buffer.write(json.dumps(data, indent=2).encode('utf-8'))
    buffer.seek(0)

    filename = f'cloud_backup_{datetime.utcnow().strftime("%Y%m%d_%H%M%S")}.json'

    log_audit(
        'EXPORT_JSON_BACKUP',
        'Backup',
        None,
        f'JSON backup exported: {filename}'
    )

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/json'
    )

@app.route('/audit')
@login_required
@role_required('audit')
def audit_trail():
    q = request.args.get('q', '').strip()
    action = request.args.get('action', '').strip()

    query = AuditLog.query

    if q:
        like = f'%{q}%'
        query = query.filter(
            db.or_(
                AuditLog.username.like(like),
                AuditLog.full_name.like(like),
                AuditLog.entity.like(like),
                AuditLog.details.like(like)
            )
        )

    if action:
        query = query.filter(AuditLog.action == action)

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = query.order_by(AuditLog.created_at.desc()).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    logs = pagination.items

    actions = [
        row[0]
        for row in db.session.query(AuditLog.action)
        .distinct()
        .order_by(AuditLog.action)
        .all()
    ]

    return render_template(
        'audit.html',
        logs=logs,
        pagination=pagination,
        q=q,
        action=action,
        actions=actions
    )
@app.route('/export/audit.csv')
@login_required
@role_required('audit')
def export_audit_csv():
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(['Date/Time','User','Role','Action','Entity','Entity ID','Details','IP Address'])
    for a in AuditLog.query.order_by(AuditLog.created_at.desc()).limit(2000).all():
        writer.writerow([a.created_at, a.full_name or a.username, a.role, a.action, a.entity, a.entity_id, a.details, a.ip_address])
    log_audit('EXPORT_AUDIT_TRAIL', 'AuditLog', None, 'Audit trail exported to CSV')
    return Response(output.getvalue(), mimetype='text/csv', headers={'Content-Disposition': 'attachment; filename=audit_trail.csv'})


@app.route('/statements')
@login_required
@role_required('statements')
def statements():
    q = request.args.get('q', '').strip()

    query = Member.query

    if q:
        query = query.filter(
            Member.full_name.contains(q) |
            Member.member_no.contains(q) |
            Member.phone.contains(q)
        )

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = query.order_by(
        Member.member_no
    ).paginate(
        page=page,
        per_page=per_page,
        error_out=False
    )

    return render_template(
        'statements.html',
        members=pagination.items,
        pagination=pagination,
        q=q
    )

@app.route('/statements/member/<int:member_id>.pdf')
@login_required
@role_required('statements')
def member_statement_pdf(member_id):
    member = Member.query.get_or_404(member_id)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm, topMargin=16*mm, bottomMargin=16*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('StatementTitle', parent=styles['Title'], fontSize=18, leading=22, textColor=colors.HexColor('#1f4f68'))
    small_style = ParagraphStyle('Small', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#555555'))
    normal = styles['Normal']
    story = []

    setting = SystemSetting.query.first()

    organization_name = setting.organization_name if setting and setting.organization_name else CLIENT_NAME
    organization_address = setting.organization_address if setting and setting.organization_address else ''
    organization_phone = setting.organization_phone if setting and setting.organization_phone else ''
    organization_email = setting.organization_email if setting and setting.organization_email else ''
    registration_number = setting.registration_number if setting and setting.registration_number else ''

    story.append(Paragraph(organization_name, title_style))

    if registration_number:
        story.append(Paragraph(f'Registration No: {registration_number}', small_style))

    contact_line = ' | '.join(
        x for x in [organization_address, organization_phone, organization_email]
        if x
    )

    if contact_line:
        story.append(Paragraph(contact_line, small_style))

    story.append(Paragraph(f'Member Statement | Produced by {PRODUCER_NAME}', small_style))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f'<b>Member:</b> {member.full_name} &nbsp;&nbsp; <b>Member No:</b> {member.member_no}', normal))
    story.append(Paragraph(f'<b>Phone:</b> {member.phone or "-"} &nbsp;&nbsp; <b>NRC/ID:</b> {member.national_id or "-"} &nbsp;&nbsp; <b>Group:</b> {member.group_name or "-"}', normal))
    story.append(Paragraph(f'<b>Status:</b> {member.status} &nbsp;&nbsp; <b>Generated:</b> {datetime.now().strftime("%d-%b-%Y %H:%M")}', normal))
    story.append(Spacer(1, 10))

    

    contributions = Contribution.query.filter_by(member_id=member.id).order_by(Contribution.paid_on.desc()).all()
    loans = Loan.query.filter_by(member_id=member.id).order_by(Loan.issued_on.desc()).all()
    distributions = Distribution.query.filter_by(member_id=member.id).order_by(Distribution.paid_on.desc()).all()
    fines = FinePenalty.query.filter_by(member_id=member.id).order_by(FinePenalty.fine_date.desc()).all()
    welfare_contribs = WelfareContribution.query.filter_by(member_id=member.id).order_by(WelfareContribution.paid_on.desc()).all()
    welfare_claims = WelfareClaim.query.filter_by(member_id=member.id).order_by(WelfareClaim.requested_on.desc()).all()
    total_contrib = money(sum((c.amount for c in contributions), Decimal('0.00')))
    total_principal = money(sum((l.principal for l in loans), Decimal('0.00')))
    total_interest = money(sum((l.interest_amount for l in loans), Decimal('0.00')))
    total_repaid = money(sum((l.total_paid for l in loans), Decimal('0.00')))
    total_balance = money(sum((l.balance for l in loans), Decimal('0.00')))
    total_distrib = money(sum((d.amount for d in distributions), Decimal('0.00')))
    total_fines = money(sum((f.amount for f in fines if f.status != 'Waived'), Decimal('0.00')))
    total_fine_paid = money(sum((f.total_paid for f in fines), Decimal('0.00')))
    total_fine_balance = money(sum((f.balance for f in fines if f.status != 'Waived'), Decimal('0.00')))
    total_welfare_contrib = money(sum((w.amount for w in welfare_contribs), Decimal('0.00')))
    total_welfare_paid = money(
    sum((w.amount_approved for w in welfare_claims if w.status == 'Paid'),
    Decimal('0.00'))
        )

    total_savings_interest = money(
        db.session.query(
            db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0)
        )
        .filter(SavingsInterest.member_id == member.id)
        .scalar()
    )

    gross_savings_value = money(
        total_contrib + total_savings_interest
    )

    total_loan_interest_charged = money(
        db.session.query(
            db.func.coalesce(db.func.sum(LoanInterest.interest_amount), 0)
        )
        .filter(LoanInterest.member_id == member.id)
        .scalar()
    )

    adjusted_loan_balance = money(
        total_balance + total_loan_interest_charged
    )

    member_equity = money(
        gross_savings_value
        + total_distrib
        + total_welfare_paid
        - adjusted_loan_balance
        - total_fine_balance
    )

    equity_label = (
        'Member Equity / Surplus'
        if member_equity >= 0
        else 'Member Equity / Deficit'
    )

    equity_label = 'Member Equity / Surplus' if member_equity >= 0 else 'Member Equity / Deficit'

    summary = [
    ['Total Contributions', kwacha(total_contrib)],
    ['Savings Interest Earned', kwacha(total_savings_interest)],
    ['Gross Savings Value', kwacha(gross_savings_value)],

    ['Loan Principal', kwacha(total_principal)],
    ['Initial Loan Charged', kwacha(total_interest)],
    ['Loan Interest Charged', kwacha(total_loan_interest_charged)],
    ['Loan Repayments', kwacha(total_repaid)],
    ['Outstanding Loan Balance', kwacha(total_balance)],
    ['Adjusted Loan Balance', kwacha(adjusted_loan_balance)],

    ['Distributions Received', kwacha(total_distrib)],

    ['Fines Charged', kwacha(total_fines)],
    ['Fines Paid', kwacha(total_fine_paid)],
    ['Outstanding Fines', kwacha(total_fine_balance)],

    ['Welfare Contributions', kwacha(total_welfare_contrib)],
    ['Welfare Support Paid', kwacha(total_welfare_paid)],

    [equity_label, kwacha(member_equity)],
]
    equity_color = colors.green if member_equity >= 0 else colors.red
    
    summary_table = Table(summary, colWidths=[95*mm, 65*mm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e8f2f6')),
        ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor('#cccccc')),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('PADDING', (0,0), (-1,-1), 6),
        ('TEXTCOLOR', (0, len(summary)-1), (-1, len(summary)-1), equity_color),
    ]))
    story.append(Paragraph('<b>Account Summary</b>', normal))
    story.append(summary_table)
    story.append(Spacer(1, 12))

    equity_breakdown = [
    ['Gross Savings Value', kwacha(gross_savings_value)],
    ['+ Distributions Received', kwacha(total_distrib)],
    ['+ Welfare Support Paid', kwacha(total_welfare_paid)],
    ['- Adjusted Loan Balance', f'({kwacha(adjusted_loan_balance)})'],
    ['- Outstanding Fines', f'({kwacha(total_fine_balance)})'],
    [equity_label, kwacha(member_equity)],
]

    equity_table = Table(equity_breakdown, colWidths=[95*mm, 65*mm])

    equity_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e8f2f6')),
        ('BACKGROUND', (0, 5), (-1, 5), colors.HexColor('#f3edf8')),
        ('TEXTCOLOR', (0, 5), (-1, 5), equity_color),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cccccc')),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (0, 5), (-1, 5), 'Helvetica-Bold'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))

    story.append(Paragraph('<b>Member Equity Breakdown</b>', normal))
    story.append(equity_table)
    story.append(Spacer(1, 12))

    def add_table(title, headers, rows, widths):
            story.append(Paragraph(f'<b>{title}</b>', normal))
            data = [headers] + (rows if rows else [['No records found'] + ['']*(len(headers)-1)])
            table = Table(data, colWidths=widths, repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1f4f68')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor('#dddddd')),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ALIGN', (-1,1), (-1,-1), 'RIGHT'),
                ('PADDING', (0,0), (-1,-1), 5),
            ]))
            story.append(table)
            story.append(Spacer(1, 10))

    add_table('Contributions', ['Month','Date','Method','Reference','Amount'], [[c.month, c.paid_on.strftime('%d-%b-%Y'), c.method, c.reference or '-', kwacha(c.amount)] for c in contributions], [24*mm, 28*mm, 30*mm, 48*mm, 30*mm])
    loan_rows = [[l.issued_on.strftime('%d-%b-%Y'), l.due_on.strftime('%d-%b-%Y'), kwacha(l.principal), kwacha(l.interest_amount), kwacha(l.total_due), kwacha(l.total_paid), kwacha(l.balance), l.status] for l in loans]
    add_table('Loans', ['Issued','Due','Principal','Interest','Total','Paid','Balance','Status'], loan_rows, [22*mm,22*mm,22*mm,22*mm,22*mm,22*mm,22*mm,20*mm])
    add_table('Distributions', ['Date','Method','Reference','Authorized By','Amount'], [[d.paid_on.strftime('%d-%b-%Y'), d.method, d.reference or '-', d.authorized_by or '-', kwacha(d.amount)] for d in distributions], [28*mm, 30*mm, 45*mm, 37*mm, 25*mm])
    add_table('Fines & Penalties', ['Date','Category','Amount','Paid','Balance','Status'], [[f.fine_date.strftime('%d-%b-%Y'), f.category, kwacha(f.amount), kwacha(f.total_paid), kwacha(f.balance), f.status] for f in fines], [28*mm, 45*mm, 25*mm, 25*mm, 25*mm, 25*mm])
    add_table('Welfare Contributions', ['Month','Date','Method','Reference','Amount'], [[w.month, w.paid_on.strftime('%d-%b-%Y'), w.method, w.reference or '-', kwacha(w.amount)] for w in welfare_contribs], [24*mm, 28*mm, 30*mm, 48*mm, 30*mm])
    add_table('Welfare Claims', ['Date','Category','Requested','Approved','Status','Paid On'], [[w.requested_on.strftime('%d-%b-%Y'), w.category, kwacha(w.amount_requested), kwacha(w.amount_approved), w.status, w.paid_on.strftime('%d-%b-%Y') if w.paid_on else '-'] for w in welfare_claims], [28*mm, 40*mm, 25*mm, 25*mm, 25*mm, 30*mm])

    story.append(Spacer(1, 10))
    story.append(
    Paragraph(
        f'This statement was generated from the {organization_name} system.',
        small_style
    )
        )
    doc.build(story)
    buffer.seek(0)
    log_audit('GENERATE_MEMBER_STATEMENT', 'Member', member.id, f'Member statement PDF generated for {member.member_no} - {member.full_name}')
    filename = f'{member.member_no}_{member.full_name.replace(" ", "_")}_statement.pdf'
    return send_file(buffer, as_attachment=True, download_name=secure_filename(filename), mimetype='application/pdf')

    


@app.route('/shareout', methods=['GET', 'POST'])
@login_required
@role_required('shareout')
def shareout():
    setting = SystemSetting.query.first()
    organization_name = setting.organization_name if setting and setting.organization_name else CLIENT_NAME

    today_month = date.today().strftime('%Y-%m')
    start_month = request.values.get('start_month') or f'{date.today().year}-01'
    end_month = request.values.get('end_month') or today_month

    expenses = money(request.values.get('expenses') or 0)
    other_income = money(request.values.get('other_income') or 0)

    start_date = datetime.strptime(start_month + '-01', '%Y-%m-%d').date()
    end_year, end_mon = [int(x) for x in end_month.split('-')]
    end_date = (date(end_year, end_mon, 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

    contrib_rows = db.session.query(
        Member.id,
        Member.member_no,
        Member.full_name,
        Member.group_name,
        db.func.coalesce(db.func.sum(Contribution.amount), 0)
    ).join(
        Contribution, Contribution.member_id == Member.id
    ).filter(
        Contribution.month >= start_month,
        Contribution.month <= end_month
    ).group_by(
        Member.id,
        Member.member_no,
        Member.full_name,
        Member.group_name
    ).order_by(
        Member.member_no
    ).all()

    total_contributions = money(
        sum((money(row[4]) for row in contrib_rows), Decimal('0.00'))
    )

    total_savings_interest = money(
        db.session.query(db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0))
        .filter(SavingsInterest.month >= start_month)
        .filter(SavingsInterest.month <= end_month)
        .scalar()
    )

    fines_paid_total = money(
        db.session.query(db.func.coalesce(db.func.sum(FinePayment.amount), 0))
        .filter(FinePayment.paid_on >= start_date)
        .filter(FinePayment.paid_on <= end_date)
        .scalar()
    )

    distributions_total = money(
        db.session.query(db.func.coalesce(db.func.sum(Distribution.amount), 0))
        .filter(Distribution.paid_on >= start_date)
        .filter(Distribution.paid_on <= end_date)
        .scalar()
    )

    shareout_fund = money(
        total_contributions
        + total_savings_interest
        + fines_paid_total
        + other_income
        - expenses
        - distributions_total
    )

    rows = []

    for member_id, member_no, full_name, group_name, contributed in contrib_rows:
        contributed = money(contributed)

        savings_interest = money(
            db.session.query(db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0))
            .filter(SavingsInterest.member_id == member_id)
            .filter(SavingsInterest.month >= start_month)
            .filter(SavingsInterest.month <= end_month)
            .scalar()
        )

        gross_savings_value = money(contributed + savings_interest)

        percent = Decimal('0.00') if total_contributions == 0 else (
            contributed / total_contributions * Decimal('100')
        )

        gross_shareout = Decimal('0.00') if total_contributions == 0 else money(
            contributed / total_contributions * shareout_fund
        )

        loan_principal_interest = money(
            sum(
                (
                    l.balance
                    for l in Loan.query.filter_by(member_id=member_id).all()
                    if l.status != 'Rejected'
                ),
                Decimal('0.00')
            )
        )

        compounded_loan_interest = money(
            db.session.query(db.func.coalesce(db.func.sum(LoanInterest.interest_amount), 0))
            .filter(LoanInterest.member_id == member_id)
            .filter(LoanInterest.month >= start_month)
            .filter(LoanInterest.month <= end_month)
            .scalar()
        )

        outstanding_loans = money(loan_principal_interest + compounded_loan_interest)

        fine_balance = money(
            sum(
                (
                    f.balance
                    for f in FinePenalty.query.filter_by(member_id=member_id).all()
                    if f.status != 'Waived'
                ),
                Decimal('0.00')
            )
        )

        total_deductions = money(outstanding_loans + fine_balance)

        net_payable = money(gross_shareout - total_deductions)
        net_status = 'surplus' if net_payable >= 0 else 'loss'

        rows.append({
            'member_no': member_no,
            'full_name': full_name,
            'group_name': group_name or '-',
            'contributed': contributed,
            'savings_interest': savings_interest,
            'gross_savings_value': gross_savings_value,
            'percent': percent.quantize(Decimal('0.01')),
            'gross_shareout': gross_shareout,
            'outstanding_loans': outstanding_loans,
            'fine_balance': fine_balance,
            'total_deductions': total_deductions,
            'net_payable': net_payable,
            'net_status': net_status,
        })
    page = request.args.get('page', 1, type=int)
    per_page = 25

    total_rows = len(rows)
    total_pages = (total_rows + per_page - 1) // per_page if total_rows else 1

    if page < 1:
        page = 1
    elif page > total_pages:
        page = total_pages

    start = (page - 1) * per_page
    end = start + per_page

    paged_rows = rows[start:end]    

    if request.path.endswith('.csv'):
        output = io.StringIO()
        writer = csv.writer(output)

        writer.writerow([organization_name])
        writer.writerow([f'Share-Out Report: {start_month} to {end_month}'])
        writer.writerow([])

        writer.writerow([
            'Member No',
            'Full Name',
            'Group',
            'Contributions',
            'Savings Interest',
            'Gross Savings Value',
            'Contribution %',
            'Gross Share-Out',
            'Outstanding Loans',
            'Outstanding Fines',
            'Total Deductions',
            'Net Payable'
        ])

        for r in rows:
            writer.writerow([
                r['member_no'],
                r['full_name'],
                r['group_name'],
                r['contributed'],
                r['savings_interest'],
                r['gross_savings_value'],
                r['percent'],
                r['gross_shareout'],
                r['outstanding_loans'],
                r['fine_balance'],
                r['total_deductions'],
                r['net_payable']
            ])

        log_audit(
            'EXPORT_SHAREOUT',
            'ShareOut',
            None,
            f'Share-out CSV exported for {start_month} to {end_month}'
        )

        return Response(
            output.getvalue(),
            mimetype='text/csv',
            headers={
                'Content-Disposition': f'attachment; filename=shareout_{start_month}_to_{end_month}.csv'
            }
        )

    if request.method == 'POST':
        log_audit(
            'CALCULATE_SHAREOUT',
            'ShareOut',
            None,
            f'Share-out calculated for {start_month} to {end_month}; fund {kwacha(shareout_fund)}'
        )

    return render_template('shareout.html', **locals())

@app.route('/shareout.pdf')
@login_required
@role_required('shareout')
def shareout_pdf():
    setting = SystemSetting.query.first()
    organization_name = setting.organization_name if setting and setting.organization_name else CLIENT_NAME
    organization_address = setting.organization_address if setting and setting.organization_address else ''
    organization_phone = setting.organization_phone if setting and setting.organization_phone else ''
    organization_email = setting.organization_email if setting and setting.organization_email else ''
    registration_number = setting.registration_number if setting and setting.registration_number else ''

    today_month = date.today().strftime('%Y-%m')
    start_month = request.args.get('start_month') or f'{date.today().year}-01'
    end_month = request.args.get('end_month') or today_month
    expenses = money(request.args.get('expenses') or 0)
    other_income = money(request.args.get('other_income') or 0)

    start_date = datetime.strptime(start_month + '-01', '%Y-%m-%d').date()
    end_year, end_mon = [int(x) for x in end_month.split('-')]
    end_date = (date(end_year, end_mon, 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

    contrib_rows = db.session.query(
        Member.id,
        Member.member_no,
        Member.full_name,
        Member.group_name,
        db.func.coalesce(db.func.sum(Contribution.amount), 0)
    ).join(
        Contribution, Contribution.member_id == Member.id
    ).filter(
        Contribution.month >= start_month,
        Contribution.month <= end_month
    ).group_by(
        Member.id,
        Member.member_no,
        Member.full_name,
        Member.group_name
    ).order_by(
        Member.member_no
    ).all()

    total_contributions = money(sum((money(row[4]) for row in contrib_rows), Decimal('0.00')))

    total_savings_interest = money(
        db.session.query(db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0))
        .filter(SavingsInterest.month >= start_month)
        .filter(SavingsInterest.month <= end_month)
        .scalar()
    )

    fines_paid_total = money(
        db.session.query(db.func.coalesce(db.func.sum(FinePayment.amount), 0))
        .filter(FinePayment.paid_on >= start_date)
        .filter(FinePayment.paid_on <= end_date)
        .scalar()
    )

    distributions_total = money(
        db.session.query(db.func.coalesce(db.func.sum(Distribution.amount), 0))
        .filter(Distribution.paid_on >= start_date)
        .filter(Distribution.paid_on <= end_date)
        .scalar()
    )

    shareout_fund = money(
        total_contributions
        + total_savings_interest
        + fines_paid_total
        + other_income
        - expenses
        - distributions_total
    )

    rows = []

    for member_id, member_no, full_name, group_name, contributed in contrib_rows:
        contributed = money(contributed)

        savings_interest = money(
            db.session.query(db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0))
            .filter(SavingsInterest.member_id == member_id)
            .filter(SavingsInterest.month >= start_month)
            .filter(SavingsInterest.month <= end_month)
            .scalar()
        )

        gross_shareout = Decimal('0.00') if total_contributions == 0 else money(
            contributed / total_contributions * shareout_fund
        )

        loan_principal_interest = money(
            sum(
                (
                    l.balance
                    for l in Loan.query.filter_by(member_id=member_id).all()
                    if l.status != 'Rejected'
                ),
                Decimal('0.00')
            )
        )

        compounded_loan_interest = money(
            db.session.query(db.func.coalesce(db.func.sum(LoanInterest.interest_amount), 0))
            .filter(LoanInterest.member_id == member_id)
            .filter(LoanInterest.month >= start_month)
            .filter(LoanInterest.month <= end_month)
            .scalar()
        )

        outstanding_loans = money(loan_principal_interest + compounded_loan_interest)

        fine_balance = money(
            sum(
                (
                    f.balance
                    for f in FinePenalty.query.filter_by(member_id=member_id).all()
                    if f.status != 'Waived'
                ),
                Decimal('0.00')
            )
        )

        total_deductions = money(outstanding_loans + fine_balance)
        net_payable = money(gross_shareout - total_deductions)

        rows.append([
            member_no,
            full_name,
            kwacha(contributed),
            kwacha(savings_interest),
            kwacha(gross_shareout),
            kwacha(outstanding_loans),
            kwacha(fine_balance),
            kwacha(total_deductions),
            kwacha(net_payable)
        ])

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=12*mm,
        leftMargin=12*mm,
        topMargin=14*mm,
        bottomMargin=14*mm
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'ShareOutTitle',
        parent=styles['Title'],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#1f4f68')
    )
    small_style = ParagraphStyle(
        'Small',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#555555')
    )
    normal = styles['Normal']

    story = []

    story.append(Paragraph(organization_name, title_style))

    if registration_number:
        story.append(Paragraph(f'Registration No: {registration_number}', small_style))

    contact_line = ' | '.join(
        x for x in [organization_address, organization_phone, organization_email]
        if x
    )

    if contact_line:
        story.append(Paragraph(contact_line, small_style))

    story.append(Spacer(1, 8))
    story.append(Paragraph(f'<b>Share-Out Report</b> | Period: {start_month} to {end_month}', normal))
    story.append(Paragraph(f'Produced by {PRODUCER_NAME}', small_style))
    story.append(Spacer(1, 10))

    summary = [
        ['Total Contributions', kwacha(total_contributions)],
        ['Savings Interest', kwacha(total_savings_interest)],
        ['Fines Paid', kwacha(fines_paid_total)],
        ['Other Income', kwacha(other_income)],
        ['Expenses', kwacha(expenses)],
        ['Distributions', kwacha(distributions_total)],
        ['Share-Out Fund', kwacha(shareout_fund)],
    ]

    summary_table = Table(summary, colWidths=[90*mm, 70*mm])
    summary_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cccccc')),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e8f2f6')),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))

    story.append(summary_table)
    story.append(Spacer(1, 12))

    data = [[
        'Member No',
        'Name',
        'Contrib.',
        'Interest',
        'Gross Share',
        'Loans',
        'Fines',
        'Deductions',
        'Net Payable'
    ]] + rows

    table = Table(
        data,
        repeatRows=1,
        colWidths=[18*mm, 34*mm, 22*mm, 22*mm, 24*mm, 22*mm, 20*mm, 24*mm, 24*mm]
    )

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4f68')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cccccc')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))

    story.append(table)
    story.append(Spacer(1, 10))
    story.append(Paragraph(f'This report was generated from the {organization_name} system.', small_style))

    doc.build(story)

    buffer.seek(0)

    log_audit(
        'EXPORT_SHAREOUT_PDF',
        'ShareOut',
        None,
        f'Share-out PDF exported for {start_month} to {end_month}'
    )

    return Response(
        buffer.getvalue(),
        mimetype='application/pdf',
        headers={
            'Content-Disposition': f'attachment; filename=shareout_{start_month}_to_{end_month}.pdf'
        }
    )
@app.route('/shareout.csv')
@login_required
@role_required('shareout')
def shareout_csv():
    return shareout()


@app.route('/month-end', methods=['GET', 'POST'])
@login_required
@role_required('accounting')
def month_end():
    selected_month = request.values.get('month') or date.today().strftime('%Y-%m')
    setting = SystemSetting.query.first()

    if not setting:
        setting = SystemSetting()
        db.session.add(setting)
        db.session.commit()

    rate = Decimal(setting.savings_interest_rate or 15) / Decimal('100')
    loan_rate = Decimal(setting.loan_interest_rate or 15) / Decimal('100')

    if request.method == 'POST':
        existing = MonthEndProcess.query.filter_by(month=selected_month).first()

        if existing:
            flash(f'Month-end interest has already been processed for {selected_month}. Reverse it first if you want to reprocess.', 'error')
            return redirect(url_for('month_end', month=selected_month))

        savings_total = Decimal('0.00')
        loan_interest_total = Decimal('0.00')
        members_processed = 0
        loans_processed = 0

        SavingsInterest.query.filter_by(month=selected_month).delete()
        LoanInterest.query.filter_by(month=selected_month).delete()
        db.session.commit()

        members = Member.query.filter_by(status='Active').all()

        for member in members:
            total_contributions = money(
                db.session.query(db.func.coalesce(db.func.sum(Contribution.amount), 0))
                .filter(Contribution.member_id == member.id)
                .scalar()
            )

            previous_interest = money(
                db.session.query(db.func.coalesce(db.func.sum(SavingsInterest.interest_amount), 0))
                .filter(SavingsInterest.member_id == member.id)
                .filter(SavingsInterest.month < selected_month)
                .scalar()
            )

            distributions_total = money(
                db.session.query(db.func.coalesce(db.func.sum(Distribution.amount), 0))
                .filter(Distribution.member_id == member.id)
                .scalar()
            )

            opening_balance = money(total_contributions + previous_interest - distributions_total)

            if opening_balance > 0:
                interest_amount = money(opening_balance * rate)
                closing_balance = money(opening_balance + interest_amount)

                db.session.add(SavingsInterest(
                    member_id=member.id,
                    month=selected_month,
                    opening_balance=opening_balance,
                    interest_rate=setting.savings_interest_rate,
                    interest_amount=interest_amount,
                    closing_balance=closing_balance
                ))

                savings_total += interest_amount
                members_processed += 1

        loans = Loan.query.filter(
            Loan.status.in_(['Disbursed', 'Open', 'Applied', 'Approved'])
        ).all()

        for loan in loans:
            previous_interest = money(
                db.session.query(db.func.coalesce(db.func.sum(LoanInterest.interest_amount), 0))
                .filter(LoanInterest.loan_id == loan.id)
                .filter(LoanInterest.month < selected_month)
                .scalar()
            )

            opening_balance = money(loan.balance + previous_interest)

            if opening_balance > 0:
                interest_amount = money(opening_balance * loan_rate)
                closing_balance = money(opening_balance + interest_amount)

                db.session.add(LoanInterest(
                    loan_id=loan.id,
                    member_id=loan.member_id,
                    month=selected_month,
                    opening_balance=opening_balance,
                    interest_rate=setting.loan_interest_rate,
                    interest_amount=interest_amount,
                    closing_balance=closing_balance
                ))

                loan_interest_total += interest_amount
                loans_processed += 1

        user = session.get('user') or {}

        process = MonthEndProcess(
            month=selected_month,
            savings_interest_total=savings_total,
            loan_interest_total=loan_interest_total,
            members_processed=members_processed,
            loans_processed=loans_processed,
            reversed=False,
            processed_by=user.get('full_name') or user.get('username')
        )

        db.session.add(process)
        db.session.commit()

        log_audit(
            'MONTH_END_PROCESS',
            'MonthEndProcess',
            process.id,
            f'Month-end processed for {selected_month}. Savings interest: {kwacha(savings_total)}, Loan interest: {kwacha(loan_interest_total)}, Members processed: {members_processed}, Loans processed: {loans_processed}'
        )

        flash(f'Month-end processed for {selected_month}. Members: {members_processed}, Loans: {loans_processed}.')
        return redirect(url_for('month_end', month=selected_month))

    page = request.args.get('page', 1, type=int)
    per_page = 25

    pagination = MonthEndProcess.query.order_by(
    MonthEndProcess.month.desc(),
    MonthEndProcess.id.desc()
    ).paginate(
    page=page,
    per_page=per_page,
    error_out=False
    )

    processes = pagination.items

    reversal_page = request.args.get('reversal_page', 1, type=int)

    reversal_pagination = AuditLog.query.filter_by(
        action='REVERSE_MONTH_END'
    ).order_by(
        AuditLog.created_at.desc()
    ).paginate(
        page=reversal_page,
        per_page=10,
        error_out=False
    )

    reversal_logs = reversal_pagination.items

    return render_template(
        'month_end.html',
        selected_month=selected_month,
        processes=processes,
        pagination=pagination,
        reversal_logs=reversal_logs,
        reversal_pagination=reversal_pagination
    )

@app.route('/month-end/<month>')
@login_required
@role_required('accounting')
def month_end_details(month):
    process = MonthEndProcess.query.filter_by(month=month).first_or_404()

    savings_page = request.args.get('savings_page', 1, type=int)
    loans_page = request.args.get('loans_page', 1, type=int)
    per_page = 25

    savings_pagination = SavingsInterest.query.filter_by(month=month).order_by(
        SavingsInterest.id.desc()
    ).paginate(
        page=savings_page,
        per_page=per_page,
        error_out=False
    )

    loan_pagination = LoanInterest.query.filter_by(month=month).order_by(
        LoanInterest.id.desc()
    ).paginate(
        page=loans_page,
        per_page=per_page,
        error_out=False
    )

    return render_template(
        'month_end_details.html',
        process=process,
        savings_entries=savings_pagination.items,
        loan_entries=loan_pagination.items,
        savings_pagination=savings_pagination,
        loan_pagination=loan_pagination
    )

@app.route('/month-end/<month>/reverse', methods=['POST'])
@login_required
@role_required('accounting')
def month_end_reverse(month):
    reason = request.form.get('reason', '').strip()

    if not reason:
        flash('Please provide a reason for reversing month-end.', 'error')
        return redirect(url_for('month_end_details', month=month))

    process = MonthEndProcess.query.filter_by(month=month).first_or_404()

    SavingsInterest.query.filter_by(month=month).delete()
    LoanInterest.query.filter_by(month=month).delete()
    db.session.delete(process)

    db.session.commit()

    log_audit(
        'REVERSE_MONTH_END',
        'MonthEndProcess',
        None,
        f'Month-end reversed for {month}. Reason: {reason}'
    )

    flash(f'Month-end processing for {month} has been reversed. You can now process it again.')
    return redirect(url_for('month_end'))

@app.route('/export/<kind>.csv')
@login_required
@role_required('exports')
def export_csv(kind):
    output = io.StringIO(); writer = csv.writer(output)
    if kind == 'members':
        writer.writerow(['Member No','Full Name','Phone','National ID','Group','Status'])
        for m in Member.query.order_by(Member.member_no): writer.writerow([m.member_no,m.full_name,m.phone,m.national_id,m.group_name,m.status])
    elif kind == 'loans':
        writer.writerow(['Member','Principal','Interest','Total Due','Paid','Balance','Status'])
        for l in Loan.query.all(): writer.writerow([l.member.full_name,l.principal,l.interest_amount,l.total_due,l.total_paid,l.balance,l.status])
    else:
        writer.writerow(['Unsupported export'])
    return Response(output.getvalue(), mimetype='text/csv', headers={'Content-Disposition': f'attachment; filename={kind}.csv'})

def parse_date(value):
    if not value: return date.today()
    return datetime.strptime(value, '%Y-%m-%d').date()

def next_month_end(d):
    return (d.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

def next_quarter_meeting(d):
    for m in [3,6,9,12]:
        if d.month <= m:
            return date(d.year, m, 28)
    return date(d.year + 1, 3, 28)


def ensure_schema():
    # Add workflow columns when upgrading an existing SQLite database.
    if db.engine.url.drivername != 'sqlite':
        return
    columns = {row[1] for row in db.session.execute(db.text('PRAGMA table_info(loan)')).fetchall()}
    additions = {
        'reviewed_by': 'VARCHAR(120)',
        'disbursed_by': 'VARCHAR(120)',
        'reviewed_on': 'DATE',
        'approved_on': 'DATE',
        'disbursed_on': 'DATE',
        'rejected_on': 'DATE',
        'rejection_reason': 'VARCHAR(250)',
    }
    for name, sqltype in additions.items():
        if name not in columns:
            db.session.execute(db.text(f'ALTER TABLE loan ADD COLUMN {name} {sqltype}'))
    db.session.commit()

def ensure_admin():
    if not User.query.filter_by(username='admin').first():
        db.session.add(User(username='admin', full_name='System Administrator', role='Administrator', password_hash=generate_password_hash('admin123')))
        db.session.commit()

@app.cli.command('init-db')
def init_db():
    """Initialize a clean production database. No demo members are created."""
    db.drop_all(); db.create_all()
    admin = User(username='admin', full_name='System Administrator', role='Administrator', password_hash=generate_password_hash('admin123'))
    db.session.add(admin)
    db.session.commit()
    db.session.add(AuditLog(user_id=admin.id, username='admin', full_name='System Administrator', role='Administrator', action='INIT_DB', entity='Database', details='Clean production database initialized with default admin user'))
    db.session.commit()
    print('Clean database initialized with default admin user: admin / admin123')

@app.cli.command('init-demo-db')
def init_demo_db():
    """Initialize a demo database with 250 sample members. Use for testing only."""
    db.drop_all(); db.create_all()
    for i in range(1, 251):
        db.session.add(Member(member_no=f'M{i:03d}', full_name=f'Member {i:03d}', phone=f'+260 97 {i:06d}', group_name=f'Group {((i-1)//25)+1}'))
    admin = User(username='admin', full_name='System Administrator', role='Administrator', password_hash=generate_password_hash('admin123'))
    db.session.add(admin)
    db.session.commit()
    db.session.add(AuditLog(user_id=admin.id, username='admin', full_name='System Administrator', role='Administrator', action='INIT_DEMO_DB', entity='Database', details='Demo database initialized with 250 sample members and default admin user'))
    db.session.commit()
    print('Demo database initialized with 250 members and default admin user: admin / admin123')


def initialize_database():
    with app.app_context():
        db.create_all()
        ensure_month_end_columns()
        ensure_settings_columns()
        ensure_schema()
        ensure_admin()

# Initialize tables when the application is imported by Gunicorn on Render.
initialize_database()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=os.environ.get('FLASK_DEBUG') == '1')
