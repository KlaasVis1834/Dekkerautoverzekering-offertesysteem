import argparse
from datetime import date
from db import init_db, connect
from importers import import_excel

def cmd_init(_):
    init_db()
    print("DB initialized -> data/app.db")

def cmd_import(args):
    n = import_excel(args.file, args.denylist)
    print(f"Imported {n} rows from {args.file}")

def cmd_list(args):
    q = """
    SELECT offer_no, klantnaam, klant_type, email, kenteken,
           delivery_method, delivery_status, is_blocked,
           call_status, decision_status, follow_up_due_at
    FROM offers
    """
    params = []
    if args.month:
        q += " WHERE month_key = ?"
        params.append(args.month)
    q += " ORDER BY created_at DESC"

    with connect() as conn:
        rows = conn.execute(q, params).fetchall()

    for r in rows:
        print(
            r["offer_no"], "|", r["klantnaam"], "|", r["klant_type"],
            "| email:", (r["email"] or "-"),
            "| kenteken:", (r["kenteken"] or "-"),
            "| levering:", r["delivery_status"],
            "| blocked:", r["is_blocked"],
            "| bellen:", (r["follow_up_due_at"] or "-")
        )

def cmd_followups(_):
    today = date.today().isoformat()
    with connect() as conn:
        rows = conn.execute("""
          SELECT offer_no, klantnaam, telefoon, call_status, decision_status, follow_up_due_at
          FROM offers
          WHERE is_blocked = 0
            AND follow_up_due_at <= ?
            AND call_status IN ('open','te_bellen')
          ORDER BY follow_up_due_at ASC
        """, (today,)).fetchall()

    for r in rows:
        print(
            r["offer_no"], "|", r["klantnaam"], "| tel:", (r["telefoon"] or "-"),
            "| status:", r["call_status"], "| besluit:", r["decision_status"],
            "| due:", r["follow_up_due_at"]
        )

def cmd_set_status(args):
    with connect() as conn:
        conn.execute("""
          UPDATE offers
          SET call_status = COALESCE(?, call_status),
              decision_status = COALESCE(?, decision_status),
              call_notes = COALESCE(?, call_notes),
              last_call_at = CASE WHEN ? IS NOT NULL THEN date('now') ELSE last_call_at END
          WHERE offer_no = ?
        """, (args.call_status, args.decision_status, args.notes, args.call_status, args.offer_no))
        conn.commit()
    print("Updated", args.offer_no)

def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers()

    s = sub.add_parser("init")
    s.set_defaults(func=cmd_init)

    s = sub.add_parser("import")
    s.add_argument("--file", required=True)
    s.add_argument("--denylist", required=True, help="Path to 'Lijstje GEEN leads.docx'")
    s.set_defaults(func=cmd_import)

    s = sub.add_parser("list")
    s.add_argument("--month", help="YYYY-MM, e.g. 2026-01")
    s.set_defaults(func=cmd_list)

    s = sub.add_parser("followups")
    s.set_defaults(func=cmd_followups)

    s = sub.add_parser("set-status")
    s.add_argument("--offer-no", required=True)
    s.add_argument("--call-status", choices=["open","te_bellen","nagebeld","niet_bellen"])
    s.add_argument("--decision-status", choices=["onbekend","akkoord","niet_akkoord"])
    s.add_argument("--notes")
    s.set_defaults(func=cmd_set_status)

    args = p.parse_args()
    if not hasattr(args, "func"):
        p.print_help()
        return
    args.func(args)

if __name__ == "__main__":
    main()
