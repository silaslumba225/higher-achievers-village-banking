"""Controlled, recoverable migration for the Higher Achievers rule set."""

import argparse
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--database', default='instance/village_banking.db')
    parser.add_argument('--apply', action='store_true')
    parser.add_argument('--confirm', default='')
    args = parser.parse_args()

    database = Path(args.database).resolve()
    if not database.exists():
        raise SystemExit(f'Database not found: {database}')

    connection = sqlite3.connect(database)
    members = connection.execute(
        "SELECT id, member_no FROM member ORDER BY id"
    ).fetchall()
    changes = []
    existing = {str(number).upper() for _, number in members if number}

    for member_id, member_no in members:
        number = str(member_no or '').strip().upper()
        if number.startswith('M') and number[1:].isdigit():
            replacement = f'HA{int(number[1:]):03d}'
            if replacement in existing:
                raise SystemExit(f'Cannot migrate {number}: {replacement} already exists.')
            changes.append((member_id, number, replacement))

    print(f'Planned member-number changes: {len(changes)}')
    for _, old, new in changes:
        print(f'  {old} -> {new}')
    print('Financial settings: savings interest 0%; loan interest 10% monthly')

    if not args.apply:
        print('Dry run only. No data changed.')
        return
    if args.confirm != 'APPLY HA RULES':
        raise SystemExit('Apply cancelled. Pass --confirm "APPLY HA RULES".')

    backup_dir = database.parent / 'backups'
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / f'pre_ha_rules_{datetime.now():%Y_%m_%d_%H%M%S}.db'
    connection.close()
    shutil.copy2(database, backup)

    connection = sqlite3.connect(database)
    try:
        connection.execute('BEGIN IMMEDIATE')
        for member_id, _, replacement in changes:
            connection.execute(
                'UPDATE member SET member_no = ? WHERE id = ?',
                (replacement, member_id),
            )
        connection.execute(
            'UPDATE system_setting SET savings_interest_rate = 0, loan_interest_rate = 10'
        )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()

    print(f'Applied successfully. Safety backup: {backup.name}')


if __name__ == '__main__':
    main()
