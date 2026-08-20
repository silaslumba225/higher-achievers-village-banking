# Village Banking System

Produced by SL Consulting for Higher Achievers.

## Run on Windows

Open Command Prompt inside the folder that contains `app.py`, then run:

```cmd
pip install -r requirements.txt
python app.py
```

Open your browser at:

```text
http://127.0.0.1:5000
```

## Default login

Username: `admin`  
Password: `admin123`

Change the administrator password after first login.

## Users and roles

Administrators can open **Users & Roles** from the menu and create users with these roles:

- Administrator: full access
- Chairperson: dashboard, loans, distributions, meetings, reports
- Treasurer: contributions, loans, repayments, distributions, reports
- Secretary: members, meetings, reports
- Auditor: reports and exports
- Data Clerk: members and contributions only

## Notes

The app uses SQLite for local testing. For multi-user production use, migrate to PostgreSQL and set a secure SECRET_KEY.

## Higher Achievers operating rules

- Annual contribution bands: Silver K10,000–K20,000; Gold K20,001–K40,000; Diamond K40,001–K80,000; Platinum K80,001–K120,000.
- Member identifiers use `HA001`, `HA002`, and so on.
- Savings earn no interest.
- Loan interest is 10% monthly on the total outstanding balance and can be charged only once per loan per month.
- Cashbook date filters recalculate the displayed cash-in, cash-out, and balance totals.

Member import requires the phrase `IMPORT MEMBERS` and creates an automatic database backup before changing rows. Member and transaction deletion is intentionally unavailable; members should be marked Inactive. Any future destructive data-management action must create a verified backup and require explicit typed confirmation before it is allowed to run.

Run `python migrate_higher_achievers.py` for a dry-run of the included data migration. Applying it requires `--apply --confirm "APPLY HA RULES"` and creates a safety backup first.
