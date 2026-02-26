import sqlite3
import os
from datetime import datetime, date, timedelta

DB_PATH = os.path.join(os.path.dirname(__file__), "workcover.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_connection()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS cases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            worker_name TEXT NOT NULL,
            state TEXT NOT NULL,
            entity TEXT,
            site TEXT,
            date_of_injury TEXT,
            injury_description TEXT,
            current_capacity TEXT DEFAULT 'Unknown',
            shift_structure TEXT,
            piawe REAL,
            reduction_rate TEXT,
            claim_number TEXT,
            claim_start_date TEXT,
            status TEXT DEFAULT 'Active',
            strategy TEXT,
            next_action TEXT,
            priority TEXT DEFAULT 'MEDIUM',
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS certificates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_id INTEGER NOT NULL,
            cert_from TEXT NOT NULL,
            cert_to TEXT NOT NULL,
            capacity TEXT,
            days_per_week INTEGER,
            hours_per_day REAL,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (case_id) REFERENCES cases(id) ON DELETE CASCADE
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS payroll_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_id INTEGER NOT NULL,
            period_from TEXT NOT NULL,
            period_to TEXT NOT NULL,
            piawe REAL,
            reduction_rate REAL,
            days_off REAL DEFAULT 0,
            hours_worked REAL DEFAULT 0,
            estimated_wages REAL DEFAULT 0,
            compensation_payable REAL DEFAULT 0,
            top_up REAL DEFAULT 0,
            back_pay_expenses REAL DEFAULT 0,
            total_payable REAL DEFAULT 0,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (case_id) REFERENCES cases(id) ON DELETE CASCADE
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_id INTEGER NOT NULL,
            doc_type TEXT NOT NULL,
            doc_name TEXT,
            is_present INTEGER DEFAULT 0,
            notes TEXT,
            FOREIGN KEY (case_id) REFERENCES cases(id) ON DELETE CASCADE
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_id INTEGER,
            action TEXT NOT NULL,
            details TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (case_id) REFERENCES cases(id) ON DELETE SET NULL
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS terminations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_id INTEGER NOT NULL UNIQUE,
            termination_type TEXT,
            approved_by TEXT,
            approved_date TEXT,
            assigned_to TEXT,
            status TEXT DEFAULT 'Pending',
            letter_drafted INTEGER DEFAULT 0,
            letter_sent INTEGER DEFAULT 0,
            response_received INTEGER DEFAULT 0,
            completed_date TEXT,
            notes TEXT,
            FOREIGN KEY (case_id) REFERENCES cases(id) ON DELETE CASCADE
        )
    """)

    conn.commit()
    conn.close()


def seed_data():
    conn = get_connection()
    c = conn.cursor()

    c.execute("SELECT COUNT(*) FROM cases")
    if c.fetchone()[0] > 0:
        conn.close()
        return

    cases = [
        ("Sayed Hadi", "VIC", "SGA", "Inghams", "2024-10-22",
         "Lower back pain - L4/L5 Disc Bulge", "No Capacity", "N/A",
         1238, "80%", "09240060334", "2024-10-22", "Active",
         "Terminate employment - inherent requirements. SGA premiums capped. Evidence of dishonesty and re-injury risk.",
         "Awaiting Henry/Mitch to action termination process", "HIGH",
         "Returned from Iran. Engaged Zaparas lawyers. MRI done but results not provided."),

        ("Gzaw Shenkute", "VIC", "SRS", "Alola", "2024-12-12",
         "Lower back pain - broad-based disc prolapse. Extensive lumbar spine degeneration.", "No Capacity", "N/A",
         None, "N/A", None, "2024-12-12", "Active",
         "Terminate employment - inherent requirements. Highly unlikely to return to cleaning duties.",
         "Awaiting Henry/Mitch to action termination. Allows insurer to pursue new employment services.", "HIGH",
         "Incident not reported initially. Communicated via son. Insurance authorised payments."),

        ("Ahmad Osmani", "VIC", "SGA", "Inghams", "2025-04-01",
         "Lower back pain - posterior disc bulge. L5/S1 disc bulge, foraminal stenosis.", "No Capacity", "N/A",
         1308, "80%", "09240069911", "2025-04-01", "Active",
         "Terminate employment - show cause or inherent requirements. SGA premiums capped.",
         "Awaiting Henry/Mitch to action termination process", "HIGH",
         "Only 23yo, worked 4 months. Engaged Zaparas lawyers. Diagnosis more like progressive degeneration."),

        ("Senait Hailu", "VIC", "Unknown", "Unknown", "2025-03-25",
         "Hand/arm injury", "Modified Duties", "Pre-injury hours",
         None, "95%", None, "2025-03-25", "Active",
         "Continue modified duties. Performing majority of role with non-injured hand.",
         "Obtain updated COC - last expired 07/12/2025", "MEDIUM",
         "Role is predominantly drying machinery with cloth. Still on modified duties completing pre-injury hours."),

        ("Tarnny Bloor", "VIC", "Buna", "Tibaldi", "2025-06-04",
         "Right arm fracture (4th metacarpal) from conveyor", "Modified Duties", "Pre-injury hours (16hrs/wk)",
         None, "95%", None, "2025-06-04", "Active",
         "Be patient until healing completes. Not premium impacting.",
         "Continue current duties. Monitor for clearance.", "LOW",
         "Working pre-injury hours, contributing well. No concerns."),

        ("Tofik Abdishekur", "VIC", "Unknown", "Unknown", None,
         "Unknown - medical certificates in file", "Unknown", "Unknown",
         None, "N/A", None, None, "Active",
         "Insufficient information. Review required.",
         "Determine current status. Latest medical cert expired 09/01/2026.", "MEDIUM",
         "Limited documentation. Has clearance certificate dated 01/01/2026 - may be closeable."),

        ("Shane Tapper", "NSW", "Myola", "Casino", "2024-04-02",
         "Right index finger crush and partial amputation", "Uncertain", "N/A",
         1081, "95%", None, "2024-04-02", "Active",
         "Re-commencing employment. ORP suggested modified role as permanent position.",
         "Confirm re-commencement details. Obtain current COC.", "MEDIUM",
         "Post-op complications. Shane declined further amputation. Longest running active case."),

        ("Shannon Kelly", "NSW", "Myola", "Casino", "2024-03-06",
         "Chronic Q-fever", "No Capacity", "N/A",
         None, "N/A", "6826218", "2024-03-06", "Active",
         "Terminate employment - show cause or inherent requirements. No chance of returning to cleaning.",
         "Awaiting Henry/Mitch to action termination. Stop leave entitlement accrual.", "HIGH",
         "Paid and managed directly by insurance. Empty case folder."),

        ("Kaleb Gaulton", "NSW", "Myola", "Casino", "2025-10-25",
         "Right middle finger crush and laceration", "Modified Duties", "5 hrs x 3 days",
         None, "95%", None, "2025-10-25", "Active",
         "Progressing well. Pain gone, just stiffness. Hoping for full capacity soon.",
         "Monitor for full capacity clearance. COC expires 13/03/2026.", "LOW",
         "Good prognosis. Employee reports pain gone, finger just a little stiff."),

        ("Damien McFarland", "NSW", "Myola", "Booyong", "2025-12-10",
         "Left thumb dislocation + wrist aggravation", "Modified Duties", "4 hrs x 4 days",
         None, "95%", "8712267", "2025-12-10", "Active",
         "Ensure compliance with doctor restrictions and RTW arrangements.",
         "COC expires 27/02/2026 - obtain new certificate.", "MEDIUM",
         "Was progressing well but re-aggravated injury lifting 8kg table outside capacity."),

        ("Ashley Hill", "NSW", "Myola", "Casino", "2026-01-14",
         "Knee injury from stepping on drainage hole. No meniscus tear.", "Modified Duties", "3 hrs x 4 days",
         779, "95%", None, "2026-01-14", "Active",
         "Provide modified duties until clearance. Strict performance reviews post-clearance.",
         "Continue modified duties. COC expires 12/03/2026.", "MEDIUM",
         "2nd workcover claim in 6 months + 2023 injury. Reports suggest possible exaggeration. High re-injury risk."),

        ("Andrew Fitzgerald", "NSW", "Myola", "Booyong", "2026-01-15",
         "Chemical burn on back", "Full Capacity", "Pre-injury",
         None, "N/A", None, "2026-01-15", "Active",
         "No further action. Only 10 days incapacity.",
         "Close claim.", "LOW",
         "Back to full duties from 17/01/2026. Simple case."),

        ("Jacob Benn", "NSW", "Myola", "Casino", None,
         "Hand injury / knee claim", "Full Capacity", "Pre-injury",
         2550, "95%", None, None, "Active",
         "Full capacity achieved. Claim open for medical expenses only.",
         "Monitor hand stiffness. Consider if claim is closeable.", "LOW",
         "Complaining of hand stiffness in mornings - suspected from overtime."),

        ("Ying Lin", "QLD", "One Harvest", "One Harvest", "2026-01-31",
         "Toe injury", "Modified Duties", "Pre-injury hours",
         None, "95%", None, "2026-01-31", "Active",
         "Continue current arrangement until clearance.",
         "COC expired 25/02/2026 - obtain new certificate or clearance URGENTLY.", "HIGH",
         "Working pre-injury hours on modified duties. Simple claim."),

        ("Asrat Bogale", "QLD", "Unknown", "Unknown", None,
         "Foot injury", "Unknown", "Unknown",
         1120, "95%", None, None, "Active",
         "Had pre-injury capacity as of 18/11. Review if closeable.",
         "Confirm clearance status. Consider closing case.", "MEDIUM",
         "Last COC expired 17/07/2025. May be ready to close."),
    ]

    for case in cases:
        c.execute("""
            INSERT INTO cases (worker_name, state, entity, site, date_of_injury,
                injury_description, current_capacity, shift_structure,
                piawe, reduction_rate, claim_number, claim_start_date, status,
                strategy, next_action, priority, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, case)

    conn.commit()

    # Seed COC data
    coc_data = [
        (1, "2025-10-29", "2025-11-26", "No Capacity", 0, 0),
        (2, "2026-01-19", "2026-02-11", "No Capacity", 0, 0),
        (3, "2026-01-20", "2026-02-16", "No Capacity", 0, 0),
        (4, "2025-11-09", "2025-12-07", "Modified Duties", None, None),
        (5, "2025-10-10", "2025-11-05", "Modified Duties", None, None),
        (6, "2025-12-12", "2026-01-09", "Unknown", None, None),
        (7, "2025-11-23", "2025-12-05", "Modified Duties", 5, 3),
        (9, "2026-02-13", "2026-03-13", "Modified Duties", 3, 5),
        (10, "2026-02-13", "2026-02-27", "Modified Duties", 4, 4),
        (11, "2026-02-16", "2026-03-12", "Modified Duties", 4, 3),
        (14, "2026-02-01", "2026-02-25", "Modified Duties", None, None),
        (15, "2025-07-03", "2025-07-17", "Unknown", None, None),
    ]

    for case_id, cfrom, cto, cap, dpw, hpd in coc_data:
        c.execute("""
            INSERT INTO certificates (case_id, cert_from, cert_to, capacity, days_per_week, hours_per_day)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (case_id, cfrom, cto, cap, dpw, hpd))

    conn.commit()

    # Seed termination data for the 4 pending cases
    termination_data = [
        (1, "Inherent Requirements", "Henry", "2026-02-19", "Mitch", "Pending"),
        (2, "Inherent Requirements", "Henry", "2026-02-19", "Mitch", "Pending"),
        (3, "Show Cause / Inherent Requirements", "Henry", "2026-02-19", "Mitch", "Pending"),
        (8, "Show Cause / Inherent Requirements", "Henry", "2026-02-19", "Mitch", "Pending"),
    ]

    for case_id, ttype, approved, adate, assigned, status in termination_data:
        c.execute("""
            INSERT INTO terminations (case_id, termination_type, approved_by, approved_date, assigned_to, status)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (case_id, ttype, approved, adate, assigned, status))

    conn.commit()

    # Seed document checklists
    doc_types = [
        "Incident Report", "Claim Form", "Payslips (12 months)",
        "PIAWE Calculation", "Certificate of Capacity (Current)",
        "RTW Plan (Current)", "Suitable Duties Plan", "Medical Certificates",
        "Insurance Correspondence", "Wage Records"
    ]

    c.execute("SELECT id FROM cases")
    case_ids = [row[0] for row in c.fetchall()]

    doc_presence = {
        1: [1, 1, 1, 1, 1, 1, 0, 1, 1, 1],   # Sayed - comprehensive
        2: [1, 0, 0, 0, 1, 1, 0, 1, 1, 0],   # Gzaw
        3: [1, 1, 1, 0, 1, 1, 0, 1, 1, 1],   # Ahmad - comprehensive
        4: [1, 1, 1, 0, 0, 1, 0, 1, 1, 1],   # Senait
        5: [1, 1, 1, 0, 0, 1, 0, 1, 1, 1],   # Tarnny
        6: [0, 0, 0, 0, 0, 1, 0, 1, 0, 1],   # Tofik
        7: [0, 0, 1, 0, 0, 0, 0, 1, 0, 0],   # Shane Tapper
        8: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],   # Shannon - empty
        9: [0, 0, 0, 0, 0, 1, 0, 0, 1, 1],   # Kaleb
        10: [0, 0, 1, 0, 1, 1, 0, 1, 1, 1],  # Damien
        11: [0, 0, 1, 0, 1, 1, 1, 1, 0, 0],  # Ashley
        12: [0, 0, 1, 0, 0, 0, 0, 0, 0, 0],  # Andrew
        13: [0, 0, 0, 0, 0, 0, 0, 0, 1, 0],  # Jacob
        14: [0, 0, 0, 0, 0, 0, 1, 0, 0, 0],  # Ying Lin
        15: [1, 0, 1, 0, 0, 0, 1, 1, 0, 1],  # Asrat
    }

    for case_id in case_ids:
        presence = doc_presence.get(case_id, [0]*10)
        for i, doc_type in enumerate(doc_types):
            c.execute("""
                INSERT INTO documents (case_id, doc_type, is_present)
                VALUES (?, ?, ?)
            """, (case_id, doc_type, presence[i] if i < len(presence) else 0))

    conn.commit()
    conn.close()


if __name__ == "__main__":
    init_db()
    seed_data()
    print("Database initialised and seeded.")
