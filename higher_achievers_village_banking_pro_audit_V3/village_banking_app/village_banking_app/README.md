# Higher Achievers Village Banking System

Produced by Excelling Foundation for Higher Achievers.

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


## Audit Trail

This version includes an Audit Trail module available to Administrator, Chairperson and Auditor roles. It records successful logins, logouts, new members, contributions, loans, repayments, distributions, meetings, user changes, database initialization and audit exports. Open **Audit Trail** from the sidebar after logging in.
