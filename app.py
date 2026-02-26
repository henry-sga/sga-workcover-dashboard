import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import database as db

st.set_page_config(
    page_title="SGA Workcover Dashboard",
    page_icon=":shield:",
    layout="wide",
    initial_sidebar_state="expanded",
)

db.init_db()
db.seed_data()

# --- Session state ---
if "dashboard_filter" not in st.session_state:
    st.session_state.dashboard_filter = None
if "selected_case_id" not in st.session_state:
    st.session_state.selected_case_id = None

# --- Helpers ---

def get_cases_df():
    conn = db.get_connection()
    df = pd.read_sql_query("SELECT * FROM cases ORDER BY state, worker_name", conn)
    conn.close()
    return df


def get_latest_cocs():
    conn = db.get_connection()
    df = pd.read_sql_query("""
        SELECT c.case_id, c.cert_from, c.cert_to, c.capacity, c.days_per_week, c.hours_per_day,
               cs.worker_name
        FROM certificates c
        JOIN cases cs ON c.case_id = cs.id
        WHERE c.id IN (
            SELECT id FROM certificates c2
            WHERE c2.case_id = c.case_id
            ORDER BY c2.cert_to DESC
            LIMIT 1
        )
        ORDER BY c.cert_to ASC
    """, conn)
    conn.close()
    return df


def get_terminations():
    conn = db.get_connection()
    df = pd.read_sql_query("""
        SELECT t.*, c.worker_name, c.state, c.site
        FROM terminations t
        JOIN cases c ON t.case_id = c.id
        ORDER BY t.status, c.worker_name
    """, conn)
    conn.close()
    return df


def get_documents(case_id):
    conn = db.get_connection()
    df = pd.read_sql_query(
        "SELECT * FROM documents WHERE case_id = ? ORDER BY doc_type", conn, params=(case_id,)
    )
    conn.close()
    return df


def get_activity_log(case_id=None, limit=50):
    conn = db.get_connection()
    if case_id:
        df = pd.read_sql_query(
            """SELECT a.*, c.worker_name FROM activity_log a
               LEFT JOIN cases c ON a.case_id = c.id
               WHERE a.case_id = ? ORDER BY a.created_at DESC LIMIT ?""",
            conn, params=(case_id, limit)
        )
    else:
        df = pd.read_sql_query(
            """SELECT a.*, c.worker_name FROM activity_log a
               LEFT JOIN cases c ON a.case_id = c.id
               ORDER BY a.created_at DESC LIMIT ?""",
            conn, params=(limit,)
        )
    conn.close()
    return df


def log_activity(case_id, action, details=""):
    conn = db.get_connection()
    conn.execute(
        "INSERT INTO activity_log (case_id, action, details) VALUES (?, ?, ?)",
        (case_id, action, details)
    )
    conn.commit()
    conn.close()


def coc_status(cert_to_str):
    if not cert_to_str:
        return "No COC", "red"
    try:
        cert_to = datetime.strptime(cert_to_str, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return "Invalid Date", "gray"
    today = date.today()
    delta = (cert_to - today).days
    if delta < 0:
        return f"EXPIRED ({abs(delta)}d ago)", "red"
    elif delta <= 7:
        return f"EXPIRING ({delta}d)", "orange"
    else:
        return f"Current ({delta}d left)", "green"


def capacity_icon(cap):
    if not cap:
        return "\u26aa"
    cap_lower = cap.lower()
    if "no capacity" in cap_lower:
        return "\U0001f534"  # red
    elif "full" in cap_lower or "clearance" in cap_lower or "cleared" in cap_lower:
        return "\U0001f7e2"  # green
    elif "modified" in cap_lower:
        return "\U0001f7e0"  # orange
    return "\u26aa"  # white


def capacity_color(cap):
    """Return a color name for st.progress etc."""
    if not cap:
        return "gray"
    cap_lower = cap.lower()
    if "no capacity" in cap_lower:
        return "red"
    elif "full" in cap_lower or "clearance" in cap_lower or "cleared" in cap_lower:
        return "green"
    elif "modified" in cap_lower:
        return "orange"
    return "gray"


def priority_emoji(p):
    return {"HIGH": "\U0001f534", "MEDIUM": "\U0001f7e0", "LOW": "\U0001f7e2"}.get(p, "\u26aa")


def coc_icon(cert_to_str):
    _, color = coc_status(cert_to_str) if cert_to_str else ("", "red")
    return {"red": "\U0001f534", "orange": "\U0001f7e0", "green": "\U0001f7e2"}.get(color, "\u26aa")


# --- Sidebar ---

st.sidebar.title("SGA Workcover")
st.sidebar.caption(f"Today: {date.today().strftime('%d %b %Y')}")

page = st.sidebar.radio(
    "Navigate",
    ["Dashboard", "All Cases", "COC Tracker", "Terminations",
     "PIAWE Calculator", "Payroll", "Activity Log"],
    index=0
)

# Reset case selection when changing pages
if "last_page" not in st.session_state:
    st.session_state.last_page = page
if st.session_state.last_page != page:
    st.session_state.selected_case_id = None
    st.session_state.dashboard_filter = None
    st.session_state.last_page = page

st.sidebar.divider()
st.sidebar.caption("Filters")
filter_state = st.sidebar.multiselect("State", ["VIC", "NSW", "QLD"], default=["VIC", "NSW", "QLD"])
filter_capacity = st.sidebar.multiselect(
    "Capacity",
    ["No Capacity", "Modified Duties", "Full Capacity", "Uncertain", "Unknown"],
    default=["No Capacity", "Modified Duties", "Full Capacity", "Uncertain", "Unknown"]
)
filter_priority = st.sidebar.multiselect(
    "Priority", ["HIGH", "MEDIUM", "LOW"], default=["HIGH", "MEDIUM", "LOW"]
)


# --- Case detail renderer (reused across pages) ---

def render_case_detail(case_id):
    conn = db.get_connection()
    case = pd.read_sql_query("SELECT * FROM cases WHERE id = ?", conn, params=(case_id,))
    certs = pd.read_sql_query("SELECT * FROM certificates WHERE case_id = ? ORDER BY cert_to DESC", conn, params=(case_id,))
    docs = pd.read_sql_query("SELECT * FROM documents WHERE case_id = ? ORDER BY doc_type", conn, params=(case_id,))
    term = pd.read_sql_query("""
        SELECT t.* FROM terminations t WHERE t.case_id = ?
    """, conn, params=(case_id,))
    log = pd.read_sql_query("""
        SELECT * FROM activity_log WHERE case_id = ? ORDER BY created_at DESC LIMIT 20
    """, conn, params=(case_id,))
    conn.close()

    if len(case) == 0:
        st.error("Case not found")
        return

    c = case.iloc[0]
    cap_col = capacity_color(c["current_capacity"])

    # Back button
    if st.button("\u2b05\ufe0f Back to cases"):
        st.session_state.selected_case_id = None
        st.rerun()

    st.markdown(f"## :{cap_col}_circle: {c['worker_name']}")
    st.caption(f"{c['state']} | {c['entity'] or ''} - {c['site'] or ''} | Priority: {c['priority']}")

    # Key info tabs
    tab_overview, tab_medical, tab_docs, tab_payroll, tab_history = st.tabs(
        ["Overview", "Medical / COCs", "Documents", "Payroll", "History"]
    )

    with tab_overview:
        oc1, oc2 = st.columns(2)
        with oc1:
            st.markdown("#### Case Details")
            st.markdown(f"**Date of Injury:** {c['date_of_injury'] or 'N/A'}")
            st.markdown(f"**Claim #:** {c['claim_number'] or 'N/A'}")
            st.markdown(f"**Current Capacity:** {c['current_capacity']}")
            st.markdown(f"**Shift Structure:** {c['shift_structure'] or 'N/A'}")
            st.markdown(f"**PIAWE:** ${c['piawe']:,.2f}" if pd.notna(c['piawe']) else "**PIAWE:** Not recorded")
            st.markdown(f"**Reduction Rate:** {c['reduction_rate'] or 'N/A'}")

        with oc2:
            st.markdown("#### Injury")
            st.markdown(c['injury_description'] or 'N/A')

            # COC status
            if len(certs) > 0:
                latest = certs.iloc[0]
                status, color = coc_status(latest["cert_to"])
                emoji = {"red": "\U0001f534", "orange": "\U0001f7e0", "green": "\U0001f7e2"}.get(color, "\u26aa")
                st.markdown(f"#### Latest COC {emoji}")
                st.markdown(f"**Period:** {latest['cert_from']} to {latest['cert_to']}")
                st.markdown(f"**Status:** {status}")
                st.markdown(f"**Capacity:** {latest['capacity'] or 'N/A'}")
            else:
                st.markdown("#### Latest COC \U0001f534")
                st.markdown("No certificate on record")

            # Termination status
            if len(term) > 0:
                t = term.iloc[0]
                st.markdown("#### Termination")
                steps_done = sum([bool(t["letter_drafted"]), bool(t["letter_sent"]), bool(t["response_received"])])
                st.progress(steps_done / 3, text=f"{t['status']} - {steps_done}/3 steps")
                st.markdown(f"**Type:** {t['termination_type']}")
                st.markdown(f"**Assigned to:** {t['assigned_to']}")

        st.divider()
        st.markdown("#### Strategy")
        st.markdown(c['strategy'] or 'N/A')
        st.markdown("#### Next Action")
        st.markdown(c['next_action'] or 'N/A')
        st.markdown("#### Notes")
        st.markdown(c['notes'] or 'None')

        # Quick edit
        st.divider()
        with st.expander("Quick Edit"):
            with st.form(f"quick_edit_{case_id}"):
                qe1, qe2 = st.columns(2)
                cap_options = ["No Capacity", "Modified Duties", "Full Capacity", "Uncertain", "Unknown"]
                new_cap = qe1.selectbox("Capacity", cap_options,
                    index=cap_options.index(c["current_capacity"]) if c["current_capacity"] in cap_options else 4)
                pri_options = ["HIGH", "MEDIUM", "LOW"]
                new_pri = qe2.selectbox("Priority", pri_options,
                    index=pri_options.index(c["priority"]) if c["priority"] in pri_options else 1)
                new_next = st.text_area("Next Action", value=c["next_action"] or "")
                new_notes = st.text_area("Notes", value=c["notes"] or "")
                if st.form_submit_button("Save"):
                    conn = db.get_connection()
                    conn.execute("""UPDATE cases SET current_capacity=?, priority=?, next_action=?, notes=?, updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                                 (new_cap, new_pri, new_next, new_notes, case_id))
                    conn.commit()
                    conn.close()
                    log_activity(case_id, "Case Updated", f"Capacity: {new_cap}, Priority: {new_pri}")
                    st.success("Saved!")
                    st.rerun()

    with tab_medical:
        st.markdown("#### Certificate of Capacity History")
        if len(certs) > 0:
            for _, cert in certs.iterrows():
                status, color = coc_status(cert["cert_to"])
                emoji = {"red": "\U0001f534", "orange": "\U0001f7e0", "green": "\U0001f7e2"}.get(color, "\u26aa")
                with st.container(border=True):
                    mc1, mc2, mc3 = st.columns([2, 2, 2])
                    mc1.markdown(f"{emoji} **{cert['cert_from']}** to **{cert['cert_to']}**")
                    mc2.markdown(f"Capacity: {cert['capacity'] or 'N/A'}")
                    schedule = ""
                    if cert["days_per_week"]:
                        schedule += f"{cert['days_per_week']} days/wk"
                    if cert["hours_per_day"]:
                        schedule += f", {cert['hours_per_day']} hrs/day"
                    mc3.markdown(schedule or "No schedule recorded")
        else:
            st.info("No certificates recorded for this case")

        st.divider()
        st.markdown("#### Add New COC")
        with st.form(f"add_coc_case_{case_id}"):
            ac1, ac2 = st.columns(2)
            new_from = ac1.date_input("From")
            new_to = ac2.date_input("To")
            new_coc_cap = st.selectbox("Capacity", ["No Capacity", "Modified Duties", "Full Capacity", "Clearance"])
            ac3, ac4 = st.columns(2)
            new_days = ac3.number_input("Days/Week", min_value=0, max_value=7, value=0)
            new_hours = ac4.number_input("Hours/Day", min_value=0.0, max_value=24.0, value=0.0, step=0.5)
            if st.form_submit_button("Add COC"):
                conn = db.get_connection()
                conn.execute("""INSERT INTO certificates (case_id, cert_from, cert_to, capacity, days_per_week, hours_per_day)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (case_id, new_from.isoformat(), new_to.isoformat(), new_coc_cap,
                     new_days if new_days > 0 else None, new_hours if new_hours > 0 else None))
                conn.execute("UPDATE cases SET current_capacity=?, updated_at=CURRENT_TIMESTAMP WHERE id=?", (new_coc_cap, case_id))
                conn.commit()
                conn.close()
                log_activity(case_id, "COC Added", f"COC {new_from} to {new_to} - {new_coc_cap}")
                st.success("COC added!")
                st.rerun()

    with tab_docs:
        st.markdown("#### Document Checklist")
        if len(docs) > 0:
            doc_changes = {}
            dcols = st.columns(2)
            for i, (_, doc) in enumerate(docs.iterrows()):
                col = dcols[i % 2]
                check = "\u2705" if doc["is_present"] else "\u274c"
                doc_changes[doc["id"]] = col.checkbox(
                    f"{check} {doc['doc_type']}", value=bool(doc["is_present"]), key=f"detail_doc_{doc['id']}"
                )
            if st.button("Save Checklist", key=f"save_docs_{case_id}"):
                conn = db.get_connection()
                for doc_id, present in doc_changes.items():
                    conn.execute("UPDATE documents SET is_present=? WHERE id=?", (int(present), int(doc_id)))
                conn.commit()
                conn.close()
                log_activity(case_id, "Documents Updated", "Checklist updated")
                st.success("Saved!")
                st.rerun()

        present_count = len(docs[docs["is_present"] == 1]) if len(docs) > 0 else 0
        total_docs = len(docs) if len(docs) > 0 else 1
        st.progress(present_count / total_docs, text=f"{present_count}/{total_docs} documents on file")

    with tab_payroll:
        st.markdown("#### Payroll History")
        conn = db.get_connection()
        pay_hist = pd.read_sql_query(
            "SELECT * FROM payroll_entries WHERE case_id = ? ORDER BY period_to DESC", conn, params=(case_id,)
        )
        conn.close()

        if pd.notna(c["piawe"]) and c["reduction_rate"] in ("95%", "80%"):
            rate = 0.95 if c["reduction_rate"] == "95%" else 0.80
            entitled = c["piawe"] * rate
            pc1, pc2, pc3 = st.columns(3)
            pc1.metric("PIAWE", f"${c['piawe']:,.2f}")
            pc2.metric("Rate", c["reduction_rate"])
            pc3.metric("Weekly Entitlement", f"${entitled:,.2f}")
        elif pd.notna(c["piawe"]):
            st.metric("PIAWE", f"${c['piawe']:,.2f}")
        else:
            st.warning("PIAWE not recorded for this case")

        if len(pay_hist) > 0:
            st.dataframe(pay_hist[["period_from", "period_to", "piawe", "estimated_wages", "compensation_payable", "total_payable", "notes"]],
                use_container_width=True, hide_index=True)
        else:
            st.info("No payroll entries for this case")

    with tab_history:
        st.markdown("#### Activity Log")
        if len(log) > 0:
            for _, entry in log.iterrows():
                st.markdown(f"**{entry['created_at'][:16] if entry['created_at'] else ''}** - {entry['action']}: {entry['details'] or ''}")
        else:
            st.info("No activity recorded")


def render_case_list(cases_to_show, title=""):
    if title:
        st.subheader(title)

    if len(cases_to_show) == 0:
        st.info("No cases match this filter")
        return

    for _, case in cases_to_show.iterrows():
        cap_col = capacity_color(case["current_capacity"])
        emoji = priority_emoji(case["priority"])
        with st.container(border=True):
            cc1, cc2, cc3, cc4 = st.columns([3, 2, 2, 1])
            cc1.markdown(f"{emoji} **{case['worker_name']}**")
            cc2.markdown(f":{cap_col}_circle: {case['current_capacity']}")
            cc3.markdown(f"{case['state']} - {case['site'] or 'Unknown'}")
            if cc4.button("Open", key=f"open_{case['id']}"):
                st.session_state.selected_case_id = int(case["id"])
                st.rerun()


# ============================================================
# DASHBOARD PAGE
# ============================================================
if page == "Dashboard":
    # If a case is selected, show its detail view
    if st.session_state.selected_case_id:
        render_case_detail(st.session_state.selected_case_id)

    else:
        st.title("Workcover Case Management Dashboard")

        cases_df = get_cases_df()
        active = cases_df[cases_df["status"] == "Active"]
        cocs = get_latest_cocs()
        terms = get_terminations()

        # Count expired COCs
        expired_count = 0
        expired_case_ids = set()
        for _, row in cocs.iterrows():
            status, _ = coc_status(row["cert_to"])
            if "EXPIRED" in status:
                expired_count += 1
                expired_case_ids.add(row["case_id"])
        # Also count cases with no COC
        cases_with_coc = set(cocs["case_id"].tolist()) if len(cocs) > 0 else set()
        for _, case in active.iterrows():
            if case["id"] not in cases_with_coc and case["current_capacity"] not in ("Full Capacity",):
                expired_count += 1
                expired_case_ids.add(case["id"])

        pending_terms = terms[terms["status"] == "Pending"]
        term_case_ids = set(pending_terms["case_id"].tolist()) if len(pending_terms) > 0 else set()

        # Clickable metrics
        col1, col2, col3, col4, col5 = st.columns(5)

        current_filter = st.session_state.dashboard_filter

        with col1:
            active_style = "primary" if current_filter == "all" else "secondary"
            if st.button(f"**{len(active)}**\n\nActive Cases", key="btn_all", use_container_width=True, type=active_style):
                st.session_state.dashboard_filter = None if current_filter == "all" else "all"
                st.rerun()

        with col2:
            no_cap_count = len(active[active["current_capacity"] == "No Capacity"])
            active_style = "primary" if current_filter == "no_capacity" else "secondary"
            if st.button(f"**{no_cap_count}**\n\nNo Capacity", key="btn_nocap", use_container_width=True, type=active_style):
                st.session_state.dashboard_filter = None if current_filter == "no_capacity" else "no_capacity"
                st.rerun()

        with col3:
            mod_count = len(active[active["current_capacity"] == "Modified Duties"])
            active_style = "primary" if current_filter == "modified" else "secondary"
            if st.button(f"**{mod_count}**\n\nModified Duties", key="btn_mod", use_container_width=True, type=active_style):
                st.session_state.dashboard_filter = None if current_filter == "modified" else "modified"
                st.rerun()

        with col4:
            active_style = "primary" if current_filter == "terminations" else "secondary"
            if st.button(f"**{len(pending_terms)}**\n\nTerminations Pending", key="btn_term", use_container_width=True, type=active_style):
                st.session_state.dashboard_filter = None if current_filter == "terminations" else "terminations"
                st.rerun()

        with col5:
            active_style = "primary" if current_filter == "expired_coc" else "secondary"
            if st.button(f"**{expired_count}**\n\nExpired COCs", key="btn_coc", use_container_width=True, type=active_style):
                st.session_state.dashboard_filter = None if current_filter == "expired_coc" else "expired_coc"
                st.rerun()

        st.divider()

        # Show filtered cases if a metric is clicked
        if current_filter == "all":
            render_case_list(active, "All Active Cases")

        elif current_filter == "no_capacity":
            filtered = active[active["current_capacity"] == "No Capacity"]
            render_case_list(filtered, "Cases - No Capacity")

        elif current_filter == "modified":
            filtered = active[active["current_capacity"] == "Modified Duties"]
            render_case_list(filtered, "Cases - Modified Duties")

        elif current_filter == "terminations":
            st.subheader("Pending Terminations")
            for _, t in pending_terms.iterrows():
                with st.container(border=True):
                    tc1, tc2, tc3, tc4 = st.columns([2, 2, 2, 1])
                    tc1.markdown(f"\U0001f534 **{t['worker_name']}** ({t['state']})")
                    tc2.markdown(f"**Type:** {t['termination_type']}")
                    steps_done = sum([bool(t["letter_drafted"]), bool(t["letter_sent"]), bool(t["response_received"])])
                    tc3.progress(steps_done / 3, text=f"{steps_done}/3 steps")
                    case_match = active[active["worker_name"] == t["worker_name"]]
                    if len(case_match) > 0:
                        if tc4.button("Open", key=f"term_open_{t['case_id']}"):
                            st.session_state.selected_case_id = int(t["case_id"])
                            st.rerun()

        elif current_filter == "expired_coc":
            filtered = active[active["id"].isin(expired_case_ids)]
            render_case_list(filtered, "Cases with Expired / Missing COCs")

        # If no filter, show the default dashboard view
        else:
            # Alerts section
            st.subheader("Alerts & Actions Required")

            alerts = []

            for _, row in cocs.iterrows():
                status, color = coc_status(row["cert_to"])
                if color in ("red", "orange"):
                    alerts.append({
                        "type": "COC", "severity": "URGENT" if color == "red" else "WARNING",
                        "worker": row["worker_name"], "case_id": row["case_id"],
                        "message": f"COC {status}", "action": "Obtain new Certificate of Capacity"
                    })

            for _, case in active.iterrows():
                if case["id"] not in cases_with_coc and case["current_capacity"] not in ("Full Capacity",):
                    alerts.append({
                        "type": "COC", "severity": "WARNING",
                        "worker": case["worker_name"], "case_id": case["id"],
                        "message": "No COC on record", "action": "Obtain Certificate of Capacity from insurer"
                    })

            for _, t in pending_terms.iterrows():
                alerts.append({
                    "type": "TERMINATION", "severity": "ACTION",
                    "worker": t["worker_name"], "case_id": t["case_id"],
                    "message": f"Termination pending - {t['termination_type']}",
                    "action": f"Follow up with {t['assigned_to']}"
                })

            for _, case in active.iterrows():
                if pd.isna(case["piawe"]) and case["current_capacity"] not in ("Full Capacity",) and case["reduction_rate"] != "N/A":
                    alerts.append({
                        "type": "PAYROLL", "severity": "INFO",
                        "worker": case["worker_name"], "case_id": case["id"],
                        "message": "PIAWE data missing", "action": "Obtain PIAWE from insurer"
                    })

            if alerts:
                for alert in sorted(alerts, key=lambda x: {"URGENT": 0, "WARNING": 1, "ACTION": 2, "INFO": 3}[x["severity"]]):
                    icon = {"URGENT": "\U0001f6a8", "WARNING": "\u26a0\ufe0f", "ACTION": "\U0001f4cb", "INFO": "\u2139\ufe0f"}[alert["severity"]]
                    with st.container(border=True):
                        ac1, ac2, ac3, ac4 = st.columns([1, 2.5, 2, 0.5])
                        ac1.markdown(f"{icon} **{alert['severity']}**")
                        ac2.markdown(f"**{alert['worker']}** - {alert['message']}")
                        ac3.markdown(f"*{alert['action']}*")
                        if ac4.button("\u27a1\ufe0f", key=f"alert_{alert['case_id']}_{alert['type']}"):
                            st.session_state.selected_case_id = int(alert["case_id"])
                            st.rerun()
            else:
                st.success("No alerts - all cases are up to date!")

            st.divider()

            # Cases by state
            st.subheader("Cases by State")
            col1, col2, col3 = st.columns(3)

            for col, state in [(col1, "VIC"), (col2, "NSW"), (col3, "QLD")]:
                state_cases = active[active["state"] == state]
                with col:
                    st.markdown(f"### {state} ({len(state_cases)})")
                    for _, case in state_cases.iterrows():
                        cap_col = capacity_color(case["current_capacity"])
                        emoji = priority_emoji(case["priority"])
                        if st.button(
                            f"{case['worker_name']} | {case['current_capacity']}",
                            key=f"state_{case['id']}",
                            use_container_width=True
                        ):
                            st.session_state.selected_case_id = int(case["id"])
                            st.rerun()


# ============================================================
# ALL CASES PAGE
# ============================================================
elif page == "All Cases":
    if st.session_state.selected_case_id:
        render_case_detail(st.session_state.selected_case_id)
    else:
        st.title("All Cases")

        cases_df = get_cases_df()
        filtered = cases_df[
            (cases_df["state"].isin(filter_state)) &
            (cases_df["current_capacity"].isin(filter_capacity)) &
            (cases_df["priority"].isin(filter_priority))
        ]

        tab_view, tab_add, tab_edit = st.tabs(["View Cases", "Add New Case", "Edit Case"])

        with tab_view:
            render_case_list(filtered)

    with tab_add:
        st.subheader("Add New Case")
        with st.form("add_case_form"):
            ac1, ac2 = st.columns(2)
            new_name = ac1.text_input("Worker Name*")
            new_state = ac2.selectbox("State*", ["VIC", "NSW", "QLD", "TAS", "SA", "WA"])
            new_entity = ac1.text_input("Entity")
            new_site = ac2.text_input("Site")
            new_doi = ac1.date_input("Date of Injury", value=None)
            new_capacity = ac2.selectbox("Current Capacity", ["No Capacity", "Modified Duties", "Full Capacity", "Uncertain", "Unknown"])
            new_injury = st.text_area("Injury Description")
            new_shift = ac1.text_input("Shift Structure")
            new_piawe = ac2.number_input("PIAWE ($)", min_value=0.0, value=0.0, step=0.01)
            new_reduction = ac1.selectbox("Reduction Rate", ["95%", "80%", "N/A"])
            new_claim = ac2.text_input("Claim Number")
            new_priority = ac1.selectbox("Priority", ["HIGH", "MEDIUM", "LOW"])
            new_strategy = st.text_area("Strategy")
            new_next = st.text_area("Next Action Required")
            new_notes = st.text_area("Notes")

            submitted = st.form_submit_button("Add Case")
            if submitted and new_name:
                conn = db.get_connection()
                conn.execute("""
                    INSERT INTO cases (worker_name, state, entity, site, date_of_injury,
                        injury_description, current_capacity, shift_structure, piawe,
                        reduction_rate, claim_number, priority, strategy, next_action, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (new_name, new_state, new_entity, new_site,
                      new_doi.isoformat() if new_doi else None,
                      new_injury, new_capacity, new_shift,
                      new_piawe if new_piawe > 0 else None,
                      new_reduction, new_claim or None, new_priority,
                      new_strategy, new_next, new_notes))
                conn.commit()
                case_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

                # Create document checklist
                doc_types = [
                    "Incident Report", "Claim Form", "Payslips (12 months)",
                    "PIAWE Calculation", "Certificate of Capacity (Current)",
                    "RTW Plan (Current)", "Suitable Duties Plan", "Medical Certificates",
                    "Insurance Correspondence", "Wage Records"
                ]
                for dt in doc_types:
                    conn.execute("INSERT INTO documents (case_id, doc_type) VALUES (?, ?)", (case_id, dt))
                conn.commit()
                conn.close()
                log_activity(case_id, "Case Created", f"New case added for {new_name}")
                st.success(f"Case added for {new_name}!")
                st.rerun()

    with tab_edit:
        st.subheader("Edit Case")
        cases_list = cases_df["worker_name"].tolist()
        selected_name = st.selectbox("Select Case to Edit", cases_list)
        if selected_name:
            case = cases_df[cases_df["worker_name"] == selected_name].iloc[0]
            with st.form("edit_case_form"):
                ec1, ec2 = st.columns(2)
                edit_capacity = ec1.selectbox("Current Capacity",
                    ["No Capacity", "Modified Duties", "Full Capacity", "Uncertain", "Unknown"],
                    index=["No Capacity", "Modified Duties", "Full Capacity", "Uncertain", "Unknown"].index(case["current_capacity"]) if case["current_capacity"] in ["No Capacity", "Modified Duties", "Full Capacity", "Uncertain", "Unknown"] else 4
                )
                edit_shift = ec2.text_input("Shift Structure", value=case["shift_structure"] or "")
                edit_piawe = ec1.number_input("PIAWE ($)", min_value=0.0, value=float(case["piawe"]) if pd.notna(case["piawe"]) else 0.0, step=0.01)
                edit_reduction = ec2.selectbox("Reduction Rate", ["95%", "80%", "N/A"],
                    index=["95%", "80%", "N/A"].index(case["reduction_rate"]) if case["reduction_rate"] in ["95%", "80%", "N/A"] else 2
                )
                priorities = ["HIGH", "MEDIUM", "LOW"]
                edit_priority = ec1.selectbox("Priority", priorities,
                    index=priorities.index(case["priority"]) if case["priority"] in priorities else 1
                )
                statuses = ["Active", "Closed", "Pending Closure"]
                edit_status = ec2.selectbox("Status", statuses,
                    index=statuses.index(case["status"]) if case["status"] in statuses else 0
                )
                edit_strategy = st.text_area("Strategy", value=case["strategy"] or "")
                edit_next = st.text_area("Next Action", value=case["next_action"] or "")
                edit_notes = st.text_area("Notes", value=case["notes"] or "")

                save = st.form_submit_button("Save Changes")
                if save:
                    conn = db.get_connection()
                    conn.execute("""
                        UPDATE cases SET current_capacity=?, shift_structure=?, piawe=?,
                            reduction_rate=?, priority=?, status=?, strategy=?,
                            next_action=?, notes=?, updated_at=CURRENT_TIMESTAMP
                        WHERE id=?
                    """, (edit_capacity, edit_shift,
                          edit_piawe if edit_piawe > 0 else None,
                          edit_reduction, edit_priority, edit_status,
                          edit_strategy, edit_next, edit_notes, int(case["id"])))
                    conn.commit()
                    conn.close()
                    log_activity(int(case["id"]), "Case Updated", f"Updated details for {selected_name}")
                    st.success("Case updated!")
                    st.rerun()

            # Document checklist update
            st.markdown("---")
            st.markdown("**Update Document Checklist:**")
            docs = get_documents(int(case["id"]))
            if len(docs) > 0:
                doc_changes = {}
                dcols = st.columns(2)
                for i, (_, doc) in enumerate(docs.iterrows()):
                    col = dcols[i % 2]
                    doc_changes[doc["id"]] = col.checkbox(
                        doc["doc_type"], value=bool(doc["is_present"]), key=f"doc_{doc['id']}"
                    )
                if st.button("Save Document Checklist"):
                    conn = db.get_connection()
                    for doc_id, present in doc_changes.items():
                        conn.execute("UPDATE documents SET is_present=? WHERE id=?", (int(present), int(doc_id)))
                    conn.commit()
                    conn.close()
                    log_activity(int(case["id"]), "Documents Updated", f"Document checklist updated for {selected_name}")
                    st.success("Document checklist saved!")
                    st.rerun()


# ============================================================
# COC TRACKER PAGE
# ============================================================
elif page == "COC Tracker":
  if st.session_state.selected_case_id:
    render_case_detail(st.session_state.selected_case_id)
  else:
    st.title("Certificate of Capacity Tracker")

    cocs = get_latest_cocs()
    cases_df = get_cases_df()

    today = date.today()
    expired = 0
    expiring = 0
    current = 0

    for _, row in cocs.iterrows():
        status, color = coc_status(row["cert_to"])
        if color == "red":
            expired += 1
        elif color == "orange":
            expiring += 1
        elif color == "green":
            current += 1

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total COCs Tracked", len(cocs))
    c2.metric("Current", current)
    c3.metric("Expiring Soon", expiring, delta="within 7 days", delta_color="inverse")
    c4.metric("Expired", expired, delta=f"{expired} overdue", delta_color="inverse")

    st.divider()

    tab_view, tab_add = st.tabs(["COC Status", "Add New COC"])

    with tab_view:
        st.subheader("Certificate Status (sorted by expiry)")
        for _, row in cocs.iterrows():
            status, color = coc_status(row["cert_to"])
            emoji = {"red": "\U0001f534", "orange": "\U0001f7e0", "green": "\U0001f7e2"}.get(color, "\u26aa")

            with st.container(border=True):
                cc1, cc2, cc3, cc4, cc5 = st.columns([2, 2, 2, 2, 0.5])
                cc1.markdown(f"{emoji} **{row['worker_name']}**")
                cc2.markdown(f"**Period:** {row['cert_from']} to {row['cert_to']}")
                cc3.markdown(f"**Capacity:** {row['capacity'] or 'N/A'}")
                cc4.markdown(f"**Status:** {status}")
                if cc5.button("Open", key=f"coc_open_{row['case_id']}"):
                    st.session_state.selected_case_id = int(row["case_id"])
                    st.rerun()

                if row["days_per_week"] or row["hours_per_day"]:
                    st.caption(f"Schedule: {row['days_per_week'] or '?'} days/week, {row['hours_per_day'] or '?'} hrs/day")

    with tab_add:
        st.subheader("Add New Certificate of Capacity")
        with st.form("add_coc_form"):
            active_cases = cases_df[cases_df["status"] == "Active"]
            case_options = {f"{r['worker_name']} ({r['state']})": r["id"] for _, r in active_cases.iterrows()}
            selected_case = st.selectbox("Worker", list(case_options.keys()))

            cc1, cc2 = st.columns(2)
            coc_from = cc1.date_input("Certificate From")
            coc_to = cc2.date_input("Certificate To")
            coc_capacity = st.selectbox("Capacity", ["No Capacity", "Modified Duties", "Full Capacity", "Clearance"])
            cc1b, cc2b = st.columns(2)
            coc_days = cc1b.number_input("Days Per Week", min_value=0, max_value=7, value=0)
            coc_hours = cc2b.number_input("Hours Per Day", min_value=0.0, max_value=24.0, value=0.0, step=0.5)
            coc_notes = st.text_area("Notes")

            add_coc = st.form_submit_button("Add Certificate")
            if add_coc and selected_case:
                case_id = case_options[selected_case]
                conn = db.get_connection()
                conn.execute("""
                    INSERT INTO certificates (case_id, cert_from, cert_to, capacity, days_per_week, hours_per_day, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (case_id, coc_from.isoformat(), coc_to.isoformat(),
                      coc_capacity, coc_days if coc_days > 0 else None,
                      coc_hours if coc_hours > 0 else None, coc_notes))
                conn.commit()

                # Also update the case's current capacity
                conn.execute("UPDATE cases SET current_capacity=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                             (coc_capacity, case_id))
                conn.commit()
                conn.close()

                worker_name = selected_case.split(" (")[0]
                log_activity(case_id, "COC Added", f"New COC {coc_from} to {coc_to} - {coc_capacity}")
                st.success(f"Certificate added for {worker_name}!")
                st.rerun()


# ============================================================
# TERMINATIONS PAGE
# ============================================================
elif page == "Terminations":
  if st.session_state.selected_case_id:
    render_case_detail(st.session_state.selected_case_id)
  else:
    st.title("Termination Tracker")

    terms = get_terminations()
    cases_df = get_cases_df()

    pending = terms[terms["status"] == "Pending"]
    completed = terms[terms["status"] == "Completed"]

    c1, c2 = st.columns(2)
    c1.metric("Pending Terminations", len(pending))
    c2.metric("Completed", len(completed))

    st.divider()

    tab_pending, tab_add, tab_update = st.tabs(["Pending", "Initiate Termination", "Update Progress"])

    with tab_pending:
        if len(pending) == 0:
            st.info("No pending terminations")
        for _, t in pending.iterrows():
            with st.container(border=True):
                tc1, tc2, tc3 = st.columns([2, 2, 2])
                tc1.markdown(f"\U0001f534 **{t['worker_name']}** ({t['state']})")
                tc2.markdown(f"**Type:** {t['termination_type']}")
                tc3.markdown(f"**Assigned to:** {t['assigned_to']}")

                st.markdown(f"**Approved by:** {t['approved_by']} on {t['approved_date']}")

                # Progress checklist
                steps = {
                    "Letter Drafted": bool(t["letter_drafted"]),
                    "Letter Sent": bool(t["letter_sent"]),
                    "Response Received": bool(t["response_received"]),
                }
                progress = sum(steps.values())
                st.progress(progress / 3, text=f"Progress: {progress}/3 steps")

                for step, done in steps.items():
                    icon = "\u2705" if done else "\u2b1b"
                    st.markdown(f"{icon} {step}")

                if t["notes"]:
                    st.caption(f"Notes: {t['notes']}")

    with tab_add:
        st.subheader("Initiate New Termination")
        with st.form("add_termination"):
            active_cases = cases_df[cases_df["status"] == "Active"]
            existing_term_cases = set(terms["case_id"].tolist()) if len(terms) > 0 else set()
            available = active_cases[~active_cases["id"].isin(existing_term_cases)]
            case_options = {f"{r['worker_name']} ({r['state']})": r["id"] for _, r in available.iterrows()}

            if case_options:
                sel = st.selectbox("Worker", list(case_options.keys()))
                term_type = st.selectbox("Termination Type", ["Inherent Requirements", "Show Cause", "Show Cause / Inherent Requirements", "Loss of Contract", "Other"])
                approved_by = st.text_input("Approved By")
                assigned_to = st.text_input("Assigned To")
                term_notes = st.text_area("Notes")

                if st.form_submit_button("Initiate Termination"):
                    case_id = case_options[sel]
                    conn = db.get_connection()
                    conn.execute("""
                        INSERT INTO terminations (case_id, termination_type, approved_by, approved_date, assigned_to, notes)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (case_id, term_type, approved_by, date.today().isoformat(), assigned_to, term_notes))
                    conn.commit()
                    conn.close()
                    log_activity(case_id, "Termination Initiated", f"Type: {term_type}, Assigned to: {assigned_to}")
                    st.success("Termination initiated!")
                    st.rerun()
            else:
                st.info("All active cases already have termination records.")
                st.form_submit_button("Initiate Termination", disabled=True)

    with tab_update:
        st.subheader("Update Termination Progress")
        if len(terms) > 0:
            term_options = {f"{r['worker_name']} - {r['termination_type']}": r for _, r in terms.iterrows()}
            sel_term = st.selectbox("Select Termination", list(term_options.keys()))
            t = term_options[sel_term]

            with st.form("update_termination"):
                ut1, ut2 = st.columns(2)
                u_status = ut1.selectbox("Status", ["Pending", "In Progress", "Completed", "Cancelled"],
                    index=["Pending", "In Progress", "Completed", "Cancelled"].index(t["status"]) if t["status"] in ["Pending", "In Progress", "Completed", "Cancelled"] else 0
                )
                u_drafted = ut1.checkbox("Letter Drafted", value=bool(t["letter_drafted"]))
                u_sent = ut2.checkbox("Letter Sent", value=bool(t["letter_sent"]))
                u_response = ut2.checkbox("Response Received", value=bool(t["response_received"]))
                u_notes = st.text_area("Notes", value=t["notes"] or "")

                if st.form_submit_button("Update"):
                    conn = db.get_connection()
                    conn.execute("""
                        UPDATE terminations SET status=?, letter_drafted=?, letter_sent=?,
                            response_received=?, notes=?, completed_date=?
                        WHERE id=?
                    """, (u_status, int(u_drafted), int(u_sent), int(u_response), u_notes,
                          date.today().isoformat() if u_status == "Completed" else None,
                          int(t["id"])))
                    conn.commit()
                    conn.close()
                    log_activity(int(t["case_id"]), "Termination Updated", f"Status: {u_status}")
                    st.success("Updated!")
                    st.rerun()
        else:
            st.info("No termination records to update.")


# ============================================================
# PIAWE CALCULATOR PAGE
# ============================================================
elif page == "PIAWE Calculator":
    st.title("PIAWE & Compensation Calculator")

    st.info("Use this calculator to work out weekly compensation entitlements based on PIAWE, capacity, and current earnings.")

    tab_calc, tab_bulk = st.tabs(["Quick Calculator", "All Cases"])

    with tab_calc:
        with st.form("piawe_calc"):
            pc1, pc2 = st.columns(2)
            calc_piawe = pc1.number_input("PIAWE (Weekly, pre-tax)", min_value=0.0, value=0.0, step=0.01)
            calc_period = pc2.selectbox("Entitlement Period", ["Weeks 1-13 (95%)", "Weeks 14-130 (80%)"])
            calc_cwe = pc1.number_input("Current Weekly Earnings (CWE)", min_value=0.0, value=0.0, step=0.01, help="Gross amount earned by worker for working in the pay period")
            calc_days = pc2.number_input("Days in Pay Period", min_value=1, max_value=14, value=10)
            calc_backpay = pc1.number_input("Back-pay & Expenses", min_value=0.0, value=0.0, step=0.01)

            if st.form_submit_button("Calculate"):
                rate = 0.95 if "95%" in calc_period else 0.80
                entitled = calc_piawe * rate
                daily_rate = entitled / 5  # 5 working days

                if calc_cwe > 0:
                    # Worker is on modified duties earning CWE
                    compensation = max(0, entitled - (calc_cwe * rate))
                    top_up = max(0, entitled - calc_cwe) if calc_cwe < entitled else 0
                else:
                    # No capacity - full compensation
                    compensation = entitled * (calc_days / 5) if calc_days != 10 else entitled * 2
                    top_up = 0

                total = calc_cwe + compensation + calc_backpay

                st.divider()
                st.subheader("Results")
                rc1, rc2, rc3 = st.columns(3)
                rc1.metric("PIAWE Rate", f"${entitled:,.2f}/wk")
                rc1.metric("Daily Rate", f"${daily_rate:,.2f}/day")
                rc2.metric("Wages (CWE)", f"${calc_cwe:,.2f}")
                rc2.metric("Compensation", f"${compensation:,.2f}")
                rc3.metric("Total Payable", f"${total:,.2f}")
                if top_up > 0:
                    rc3.metric("Top-up Required", f"${top_up:,.2f}")

                st.caption(f"Calculation: PIAWE ${calc_piawe:,.2f} x {rate*100:.0f}% = ${entitled:,.2f} entitlement. "
                          f"CWE ${calc_cwe:,.2f}. Compensation = max(0, ${entitled:,.2f} - ${calc_cwe*rate:,.2f}) = ${compensation:,.2f}")

    with tab_bulk:
        st.subheader("PIAWE Summary - All Active Cases")
        cases_df = get_cases_df()
        active = cases_df[cases_df["status"] == "Active"]

        for _, case in active.iterrows():
            piawe = case["piawe"]
            rate_str = case["reduction_rate"]

            with st.container(border=True):
                bc1, bc2, bc3, bc4 = st.columns([2, 1, 1, 2])
                bc1.markdown(f"**{case['worker_name']}** ({case['state']})")

                if pd.notna(piawe) and rate_str in ("95%", "80%"):
                    rate = 0.95 if rate_str == "95%" else 0.80
                    entitled = piawe * rate
                    bc2.markdown(f"PIAWE: **${piawe:,.2f}**")
                    bc3.markdown(f"Rate: **{rate_str}** = ${entitled:,.2f}/wk")
                    bc4.markdown(f"Capacity: {case['current_capacity']}")
                elif pd.notna(piawe):
                    bc2.markdown(f"PIAWE: **${piawe:,.2f}**")
                    bc3.markdown(f"Rate: {rate_str}")
                    bc4.markdown(f"Capacity: {case['current_capacity']}")
                else:
                    bc2.markdown("\U0001f534 **PIAWE Missing**")
                    bc3.markdown(f"Rate: {rate_str}")
                    bc4.markdown(f"Capacity: {case['current_capacity']}")


# ============================================================
# PAYROLL PAGE
# ============================================================
elif page == "Payroll":
    st.title("Payroll - Workcover Compensation")

    cases_df = get_cases_df()
    active = cases_df[cases_df["status"] == "Active"]

    tab_entry, tab_history = st.tabs(["New Pay Period Entry", "History"])

    with tab_entry:
        st.subheader("Enter Compensation for Pay Period")

        with st.form("payroll_entry"):
            case_options = {f"{r['worker_name']} ({r['state']})": r["id"] for _, r in active.iterrows()}
            sel_case = st.selectbox("Worker", list(case_options.keys()))

            pe1, pe2 = st.columns(2)
            pay_from = pe1.date_input("Period From")
            pay_to = pe2.date_input("Period To")

            case_row = active[active["id"] == case_options[sel_case]].iloc[0]
            default_piawe = float(case_row["piawe"]) if pd.notna(case_row["piawe"]) else 0.0
            default_rate = 0.95 if case_row["reduction_rate"] == "95%" else (0.80 if case_row["reduction_rate"] == "80%" else 0.0)

            pe3, pe4 = st.columns(2)
            pay_piawe = pe3.number_input("PIAWE", value=default_piawe, step=0.01)
            pay_rate = pe4.number_input("Reduction Rate", value=default_rate, min_value=0.0, max_value=1.0, step=0.05)
            pay_days = pe3.number_input("Days Off / Light Duties", min_value=0, value=0)
            pay_hours = pe4.number_input("Hours Worked", min_value=0.0, value=0.0, step=0.5)
            pay_wages = pe3.number_input("Estimated Wages", min_value=0.0, value=0.0, step=0.01)
            pay_backpay = pe4.number_input("Back-pay & Expenses", min_value=0.0, value=0.0, step=0.01)
            pay_notes = st.text_area("Notes")

            if st.form_submit_button("Calculate & Save"):
                entitled = pay_piawe * pay_rate
                if pay_wages > 0:
                    top_up = max(0, entitled - pay_wages)
                    compensation = top_up
                else:
                    daily = entitled / 5
                    compensation = daily * pay_days
                    top_up = 0

                total = pay_wages + compensation + pay_backpay

                case_id = case_options[sel_case]
                conn = db.get_connection()
                conn.execute("""
                    INSERT INTO payroll_entries (case_id, period_from, period_to, piawe, reduction_rate,
                        days_off, hours_worked, estimated_wages, compensation_payable, top_up,
                        back_pay_expenses, total_payable, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (case_id, pay_from.isoformat(), pay_to.isoformat(), pay_piawe, pay_rate,
                      pay_days, pay_hours, pay_wages, compensation, top_up, pay_backpay, total, pay_notes))
                conn.commit()
                conn.close()
                log_activity(case_id, "Payroll Entry", f"Period {pay_from} to {pay_to}: Total ${total:,.2f}")

                st.success(f"Saved! Compensation: ${compensation:,.2f} | Wages: ${pay_wages:,.2f} | Total: ${total:,.2f}")

    with tab_history:
        st.subheader("Payroll History")
        conn = db.get_connection()
        history = pd.read_sql_query("""
            SELECT p.*, c.worker_name, c.state
            FROM payroll_entries p
            JOIN cases c ON p.case_id = c.id
            ORDER BY p.period_to DESC
        """, conn)
        conn.close()

        if len(history) > 0:
            st.dataframe(
                history[["worker_name", "state", "period_from", "period_to", "piawe",
                         "reduction_rate", "estimated_wages", "compensation_payable",
                         "top_up", "total_payable", "notes"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "worker_name": "Worker",
                    "state": "State",
                    "period_from": "From",
                    "period_to": "To",
                    "piawe": st.column_config.NumberColumn("PIAWE", format="$%.2f"),
                    "reduction_rate": st.column_config.NumberColumn("Rate", format="%.0f%%"),
                    "estimated_wages": st.column_config.NumberColumn("Wages", format="$%.2f"),
                    "compensation_payable": st.column_config.NumberColumn("Compensation", format="$%.2f"),
                    "top_up": st.column_config.NumberColumn("Top-up", format="$%.2f"),
                    "total_payable": st.column_config.NumberColumn("Total", format="$%.2f"),
                }
            )
        else:
            st.info("No payroll entries yet. Use the 'New Pay Period Entry' tab to add entries.")


# ============================================================
# ACTIVITY LOG PAGE
# ============================================================
elif page == "Activity Log":
    st.title("Activity Log")

    log = get_activity_log(limit=100)

    if len(log) > 0:
        for _, entry in log.iterrows():
            with st.container(border=True):
                lc1, lc2, lc3 = st.columns([1, 2, 3])
                lc1.caption(entry["created_at"][:16] if entry["created_at"] else "")
                lc2.markdown(f"**{entry['worker_name'] or 'System'}** - {entry['action']}")
                lc3.markdown(entry["details"] or "")
    else:
        st.info("No activity recorded yet. Actions will appear here as you use the dashboard.")
