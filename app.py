from datetime import date, datetime, timedelta
from decimal import Decimal
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, Response, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import csv
import io

app = Flask(__name__)
app.config['SECRET_KEY'] = 'change-this-secret-key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///village_banking.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

INTEREST_RATE = Decimal('0.15')
PAYMENT_METHODS = ['Bank Transfer', 'Mobile Money', 'Cash']
CLIENT_NAME = 'Higher Achievers'
PRODUCER_NAME = 'Excelling Foundation'

ROLES = ['Administrator', 'Chairperson', 'Treasurer', 'Secretary', 'Auditor', 'Data Clerk']
ROLE_PERMISSIONS = {
    'Administrator': ['dashboard', 'members', 'contributions', 'loans', 'repayments', 'distributions', 'meetings', 'reports', 'users', 'exports'],
    'Chairperson': ['dashboard', 'loans', 'distributions', 'meetings', 'reports', 'exports'],
    'Treasurer': ['dashboard', 'contributions', 'loans', 'repayments', 'distributions', 'reports', 'exports'],
    'Secretary': ['dashboard', 'members', 'meetings', 'reports', 'exports'],
    'Auditor': ['dashboard', 'reports', 'exports'],
    'Data Clerk': ['dashboard', 'members', 'contributions'],
}

def user_can(permission):
    user = session.get('user') or {}
    return permission in ROLE_PERMISSIONS.get(user.get('role'), [])

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    full_name = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(30), default='Administrator')
    password_hash = db.Column(db.String(255), nullable=False)
    active = db.Column(db.Boolean, default=True)

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
    status = db.Column(db.String(20), default='Open')
    approved_by = db.Column(db.String(120), default='Management Committee')
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
        return self.status == 'Open' and self.due_on < date.today() and self.balance > 0

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

class Meeting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    meeting_type = db.Column(db.String(40), nullable=False)
    meeting_date = db.Column(db.Date, nullable=False)
    agenda = db.Column(db.Text)
    resolutions = db.Column(db.Text)
    attendance_count = db.Column(db.Integer, default=0)

def money(value):
    return Decimal(value or 0).quantize(Decimal('0.01'))

@app.template_filter('kwacha')
def kwacha(value):
    return f"K {money(value):,.2f}"

@app.context_processor
def inject_globals():
    return dict(payment_methods=PAYMENT_METHODS, interest_percent=int(INTEREST_RATE * 100), client_name=CLIENT_NAME, producer_name=PRODUCER_NAME, current_year=date.today().year, current_user=session.get('user'), user_can=user_can, roles=ROLES)

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

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = User.query.filter_by(username=request.form.get('username','').strip(), active=True).first()
        if user and check_password_hash(user.password_hash, request.form.get('password','')):
            session['user'] = {'id': user.id, 'username': user.username, 'full_name': user.full_name, 'role': user.role}
            flash('Welcome back. You are logged in securely.')
            return redirect(request.args.get('next') or url_for('dashboard'))
        flash('Invalid username or password.', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.')
    return redirect(url_for('login'))

@app.route('/')
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
    open_loans = Loan.query.filter_by(status='Open').all()
    loan_balance = money(sum((l.balance for l in open_loans), Decimal('0.00')))
    interest_earned = money(sum((l.interest_amount for l in Loan.query.all()), Decimal('0.00')))
    available_fund = money(contribution_total + repayment_total - loans_total - distribution_total)
    next_committee = next_month_end(date.today())
    next_full = next_quarter_meeting(date.today())
    recent_contribs = Contribution.query.order_by(Contribution.paid_on.desc()).limit(5).all()
    recent_loans = Loan.query.order_by(Loan.issued_on.desc()).limit(5).all()
    overdue_loans = [l for l in open_loans if l.overdue]
    return render_template('dashboard.html', **locals())

@app.route('/members')
@login_required
@role_required('members')
def members():
    q = request.args.get('q','').strip()
    query = Member.query
    if q:
        query = query.filter(Member.full_name.contains(q) | Member.member_no.contains(q) | Member.phone.contains(q))
    return render_template('members.html', members=query.order_by(Member.member_no).all(), q=q)

@app.route('/members/new', methods=['GET','POST'])
@login_required
@role_required('members')
def member_new():
    if request.method == 'POST':
        m = Member(member_no=request.form['member_no'], full_name=request.form['full_name'], phone=request.form.get('phone'), national_id=request.form.get('national_id'), group_name=request.form.get('group_name'))
        db.session.add(m); db.session.commit(); flash('Member added successfully.'); return redirect(url_for('members'))
    return render_template('member_form.html')

@app.route('/contributions', methods=['GET','POST'])
@login_required
@role_required('contributions')
def contributions():
    if request.method == 'POST':
        c = Contribution(member_id=int(request.form['member_id']), month=request.form['month'], amount=money(request.form['amount']), method=request.form['method'], reference=request.form.get('reference'), paid_on=parse_date(request.form.get('paid_on')))
        db.session.add(c); db.session.commit(); flash('Contribution recorded.'); return redirect(url_for('contributions'))
    return render_template('contributions.html', contributions=Contribution.query.order_by(Contribution.paid_on.desc()).limit(100).all(), members=Member.query.order_by(Member.full_name).all())

@app.route('/loans', methods=['GET','POST'])
@login_required
@role_required('loans')
def loans():
    if request.method == 'POST':
        issued = parse_date(request.form.get('issued_on'))
        due = parse_date(request.form.get('due_on')) or issued + timedelta(days=30)
        l = Loan(member_id=int(request.form['member_id']), principal=money(request.form['principal']), due_on=due, issued_on=issued, purpose=request.form.get('purpose'), approved_by=request.form.get('approved_by') or 'Management Committee')
        db.session.add(l); db.session.commit(); flash('Loan issued at 15% interest.'); return redirect(url_for('loans'))
    all_loans = Loan.query.order_by(Loan.issued_on.desc()).all()
    return render_template('loans.html', loans=all_loans, members=Member.query.order_by(Member.full_name).all())

@app.route('/repayments', methods=['POST'])
@login_required
@role_required('repayments')
def repayments():
    loan = Loan.query.get_or_404(int(request.form['loan_id']))
    r = Repayment(loan_id=loan.id, amount=money(request.form['amount']), method=request.form['method'], reference=request.form.get('reference'), paid_on=parse_date(request.form.get('paid_on')))
    db.session.add(r); db.session.commit()
    if loan.balance <= 0:
        loan.status = 'Closed'; db.session.commit()
    flash('Repayment recorded.')
    return redirect(url_for('loans'))

@app.route('/distributions', methods=['GET','POST'])
@login_required
@role_required('distributions')
def distributions():
    if request.method == 'POST':
        d = Distribution(member_id=int(request.form['member_id']), amount=money(request.form['amount']), method=request.form['method'], reference=request.form.get('reference'), authorized_by=request.form.get('authorized_by'), paid_on=parse_date(request.form.get('paid_on')))
        db.session.add(d); db.session.commit(); flash('Distribution recorded.'); return redirect(url_for('distributions'))
    return render_template('distributions.html', distributions=Distribution.query.order_by(Distribution.paid_on.desc()).all(), members=Member.query.order_by(Member.full_name).all())

@app.route('/meetings', methods=['GET','POST'])
@login_required
@role_required('meetings')
def meetings():
    if request.method == 'POST':
        mtg = Meeting(meeting_type=request.form['meeting_type'], meeting_date=parse_date(request.form['meeting_date']), agenda=request.form.get('agenda'), resolutions=request.form.get('resolutions'), attendance_count=int(request.form.get('attendance_count') or 0))
        db.session.add(mtg); db.session.commit(); flash('Meeting record saved.'); return redirect(url_for('meetings'))
    return render_template('meetings.html', meetings=Meeting.query.order_by(Meeting.meeting_date.desc()).all())

@app.route('/reports')
@login_required
@role_required('reports')
def reports():
    month = request.args.get('month') or date.today().strftime('%Y-%m')
    contribs = Contribution.query.filter_by(month=month).all()
    paid_member_ids = {c.member_id for c in contribs}
    arrears = Member.query.filter(Member.id.notin_(paid_member_ids)).all() if paid_member_ids else Member.query.all()
    return render_template('reports.html', month=month, contribs=contribs, arrears=arrears, open_loans=Loan.query.filter_by(status='Open').all())


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
        db.session.add(user); db.session.commit(); flash('User account created successfully.'); return redirect(url_for('users'))
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
        db.session.commit(); flash('User account updated successfully.'); return redirect(url_for('users'))
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
        flash('User status updated.')
    return redirect(url_for('users'))

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

def ensure_admin():
    if not User.query.filter_by(username='admin').first():
        db.session.add(User(username='admin', full_name='System Administrator', role='Administrator', password_hash=generate_password_hash('admin123')))
        db.session.commit()

@app.cli.command('init-db')
def init_db():
    db.drop_all(); db.create_all()
    for i in range(1, 251):
        db.session.add(Member(member_no=f'M{i:03d}', full_name=f'Member {i:03d}', phone=f'+260 97 {i:06d}', group_name=f'Group {((i-1)//25)+1}'))
    db.session.add(User(username='admin', full_name='System Administrator', role='Administrator', password_hash=generate_password_hash('admin123')))
    db.session.commit()
    print('Database initialized with 250 members and default admin user: admin / admin123')

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        ensure_admin()
    app.run(debug=True)
