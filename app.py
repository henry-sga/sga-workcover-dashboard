import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import database as db
import doc_generator as dg

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


def get_generated_documents(case_id):
    conn = db.get_connection()
    df = pd.read_sql_query(
        "SELECT id, case_id, doc_type, doc_name, generated_at FROM generated_documents WHERE case_id = ? ORDER BY generated_at DESC",
        conn, params=(case_id,)
    )
    conn.close()
    return df


def get_generated_doc_data(doc_id):
    conn = db.get_connection()
    row = conn.execute("SELECT doc_data, doc_name FROM generated_documents WHERE id = ?", (doc_id,)).fetchone()
    conn.close()
    if row:
        return row["doc_data"], row["doc_name"]
    return None, None


def get_doctor_details(case_id):
    conn = db.get_connection()
    row = conn.execute("SELECT * FROM doctor_details WHERE case_id = ? ORDER BY id DESC LIMIT 1", (case_id,)).fetchone()
    conn.close()
    return dict(row) if row else {}


def get_incident_details(case_id):
    conn = db.get_connection()
    row = conn.execute("SELECT * FROM incident_details WHERE case_id = ? ORDER BY id DESC LIMIT 1", (case_id,)).fetchone()
    conn.close()
    return dict(row) if row else {}


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


def build_case_data_dict(case_row):
    """Convert a case DB row/series to a dict for doc generation."""
    if isinstance(case_row, pd.Series):
        return case_row.to_dict()
    return dict(case_row)


def build_medical_data(case_id, case_data):
    """Build medical_data dict from latest COC + doctor + incident details."""
    conn = db.get_connection()
    cert = conn.execute(
        "SELECT * FROM certificates WHERE case_id = ? ORDER BY cert_to DESC LIMIT 1",
        (case_id,)
    ).fetchone()
    conn.close()

    doctor = get_doctor_details(case_id)
    incident = get_incident_details(case_id)

    med = {}
    if cert:
        med["cert_from"] = cert["cert_from"]
        med["cert_to"] = cert["cert_to"]
        med["hours_per_day"] = cert["hours_per_day"]
        med["days_per_week"] = cert["days_per_week"]
    if doctor:
        med["doctor_name"] = doctor.get("doctor_name")
        med["doctor_address"] = doctor.get("doctor_address")
        med["doctor_phone"] = doctor.get("doctor_phone")
        med["doctor_fax"] = doctor.get("doctor_fax")
    if incident:
        med["worker_dob"] = incident.get("dob")
        med["occupation"] = incident.get("occupation")
    med["restrictions"] = None  # Will show [REVIEW] marker
    return med, doctor, incident


# --- Sidebar (styled menu) ---

st.sidebar.markdown("""
<style>
div[data-testid="stSidebar"] .stRadio > label { display: none; }
div.sidebar-menu-item {
    padding: 8px 12px;
    border-radius: 8px;
    margin: 2px 0;
    cursor: pointer;
}
</style>
""", unsafe_allow_html=True)

st.sidebar.markdown("### SGA Workcover")
st.sidebar.caption(f"Today: {date.today().strftime('%d %b %Y')}")

# Build styled navigation
NAV_ITEMS = [
    ("Dashboard", "house"),
    ("New Case", "plus-circle"),
    ("All Cases", "folder"),
    ("COC Tracker", "calendar-check"),
    ("Terminations", "x-circle"),
    ("PIAWE Calculator", "calculator"),
    ("Payroll", "currency-dollar"),
    ("Activity Log", "clock-history"),
]

# Use selectbox with a cleaner look for navigation
page = st.sidebar.selectbox(
    "Navigate",
    [item[0] for item in NAV_ITEMS],
    index=0,
    label_visibility="collapsed",
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


# --- Generate Documents Dialog ---

def render_generate_documents(case_id):
    """Render the document generation UI for a case."""
    conn = db.get_connection()
    case = conn.execute("SELECT * FROM cases WHERE id = ?", (case_id,)).fetchone()
    conn.close()
    if not case:
        st.error("Case not found")
        return

    case_data = dict(case)
    medical_data, doctor_data, incident_data = build_medical_data(case_id, case_data)

    st.markdown("#### Generate Documents")
    st.caption("Select which documents to generate. They will be pre-filled with available case data.")

    # Show available documents with checkboxes
    selected_docs = {}
    for doc_key, doc_info in dg.AVAILABLE_DOCUMENTS.items():
        review_badge = ""
        if "Yes" in doc_info["needs_review"]:
            review_badge = "  :orange[Needs Review]"
        elif "Minimal" in doc_info["needs_review"]:
            review_badge = "  :blue[Minimal Review]"
        else:
            review_badge = "  :green[Ready to Use]"

        selected_docs[doc_key] = st.checkbox(
            f"{doc_info['icon']}  **{doc_info['name']}** - {doc_info['description']}{review_badge}",
            key=f"gen_doc_{case_id}_{doc_key}",
            value=False,
        )

    any_selected = any(selected_docs.values())

    if st.button("Generate Selected Documents", disabled=not any_selected, type="primary",
                 key=f"gen_btn_{case_id}"):
        docs_to_generate = [k for k, v in selected_docs.items() if v]

        with st.spinner("Generating documents..."):
            results = dg.generate_documents(
                case_data, docs_to_generate,
                medical_data=medical_data,
                doctor_data=doctor_data,
                incident_data=incident_data,
            )

        # Save to DB and provide downloads
        conn = db.get_connection()
        for doc_type, (filename, buf) in results.items():
            conn.execute(
                "INSERT INTO generated_documents (case_id, doc_type, doc_name, doc_data) VALUES (?, ?, ?, ?)",
                (case_id, doc_type, filename, buf.getvalue())
            )
        conn.commit()
        conn.close()

        log_activity(case_id, "Documents Generated",
                     f"Generated: {', '.join(dg.AVAILABLE_DOCUMENTS[k]['name'] for k in docs_to_generate)}")

        st.success(f"Generated {len(results)} document(s)!")

        # Show download buttons
        for doc_type, (filename, buf) in results.items():
            info = dg.AVAILABLE_DOCUMENTS[doc_type]
            st.download_button(
                label=f"Download {info['icon']} {info['name']}",
                data=buf.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_{case_id}_{doc_type}_{datetime.now().timestamp()}",
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
    tab_overview, tab_medical, tab_docs, tab_generate, tab_payroll, tab_history = st.tabs(
        ["Overview", "Medical / COCs", "Documents", "Generate Docs", "Payroll", "History"]
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

        # Generated documents section
        st.divider()
        st.markdown("#### Generated Documents")
        gen_docs = get_generated_documents(case_id)
        if len(gen_docs) > 0:
            for _, gdoc in gen_docs.iterrows():
                doc_info = dg.AVAILABLE_DOCUMENTS.get(gdoc["doc_type"], {})
                icon = doc_info.get("icon", "")
                with st.container(border=True):
                    gc1, gc2, gc3 = st.columns([3, 2, 1])
                    gc1.markdown(f"{icon} **{gdoc['doc_name']}**")
                    gc2.caption(f"Generated: {gdoc['generated_at'][:16] if gdoc['generated_at'] else ''}")

                    # Download button that lets user open/view the document
                    doc_data, doc_name = get_generated_doc_data(int(gdoc["id"]))
                    if doc_data:
                        gc3.download_button(
                            label="Open",
                            data=doc_data,
                            file_name=doc_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"open_gdoc_{gdoc['id']}",
                        )
        else:
            st.info("No documents generated yet. Use the 'Generate Docs' tab to create documents.")

    with tab_generate:
        render_generate_documents(case_id)

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
# NEW CASE PAGE
# ============================================================
elif page == "New Case":
    st.title("New Case Wizard")
    st.caption("Create a new workcover case and generate all required documents in one go.")

    # Use session state to track wizard steps
    if "wizard_step" not in st.session_state:
        st.session_state.wizard_step = 1

    step = st.session_state.wizard_step

    # Progress indicator
    steps_labels = ["Worker & Incident", "Medical Details", "Generate Documents"]
    sc1, sc2, sc3 = st.columns(3)
    for i, (col, label) in enumerate(zip([sc1, sc2, sc3], steps_labels), 1):
        if i < step:
            col.markdown(f":green_circle: ~~**Step {i}:** {label}~~")
        elif i == step:
            col.markdown(f":large_blue_circle: **Step {i}:** {label}")
        else:
            col.markdown(f":white_circle: Step {i}: {label}")

    st.divider()

    # ── STEP 1: Worker & Incident Details ──
    if step == 1:
        st.subheader("Step 1: Worker & Incident Details")

        with st.form("wizard_step1"):
            st.markdown("**Worker Information**")
            w1, w2 = st.columns(2)
            wiz_name = w1.text_input("Worker Name*")
            wiz_dob = w2.date_input("Date of Birth", value=None)
            wiz_address = w1.text_input("Address")
            wiz_phone = w2.text_input("Phone")
            wiz_language = w1.text_input("Language Needs (if any)")

            st.markdown("**Employer Details**")
            e1, e2, e3 = st.columns(3)
            wiz_entity = e1.text_input("Entity")
            wiz_site = e2.text_input("Site")
            wiz_state = e3.selectbox("State*", ["VIC", "NSW", "QLD", "TAS", "SA", "WA"])

            st.markdown("**Incident Details**")
            i1, i2, i3 = st.columns(3)
            wiz_doi = i1.date_input("Date of Injury*")
            wiz_time = i2.time_input("Time of Injury")
            wiz_location = i3.text_input("Location within Site")
            wiz_description = st.text_area("What Happened?*")
            wiz_witnesses = st.text_input("Witnesses")

            st.markdown("**Employment Details**")
            emp1, emp2 = st.columns(2)
            wiz_emp_type = emp1.selectbox("Employment Type", ["Permanent Employee", "Casual Employee", "Contractor"])
            wiz_tenure = emp2.text_input("Tenure (e.g. 2 years 3 months)")
            emp3, emp4 = st.columns(2)
            wiz_avg_hours = emp3.text_input("Average Hours/Days per Week (e.g. 38 hrs/5 days)")
            wiz_shift_type = emp4.selectbox("Shift", ["Day", "Afternoon", "Night"])
            wiz_shift_start = st.text_input("Shift Start Time (e.g. 6:00am)")

            st.markdown("**Injury Details**")
            inj1, inj2 = st.columns(2)
            wiz_nature = inj1.text_input("Nature of Injury (e.g. sprain, fracture)")
            wiz_body_part = inj2.text_input("Body Part Affected")
            wiz_treatment = st.selectbox("Treatment Level", ["No treatment", "First Aid", "Doctor", "Hospital"])
            wiz_pre_injury = st.text_area("Pre-injury Duties Description")

            if st.form_submit_button("Next: Medical Details", type="primary"):
                if not wiz_name:
                    st.error("Worker name is required")
                elif not wiz_description:
                    st.error("Injury description is required")
                else:
                    # Store in session state
                    st.session_state.wizard_data = {
                        "worker_name": wiz_name,
                        "dob": wiz_dob.isoformat() if wiz_dob else None,
                        "address": wiz_address,
                        "phone": wiz_phone,
                        "language": wiz_language,
                        "entity": wiz_entity,
                        "site": wiz_site,
                        "state": wiz_state,
                        "date_of_injury": wiz_doi.isoformat() if wiz_doi else None,
                        "time_of_injury": wiz_time.strftime("%H:%M") if wiz_time else None,
                        "location_detail": wiz_location,
                        "injury_description": wiz_description,
                        "witnesses": wiz_witnesses,
                        "employment_type": wiz_emp_type,
                        "tenure": wiz_tenure,
                        "avg_hours": wiz_avg_hours,
                        "shift_type": wiz_shift_type,
                        "shift_start_time": wiz_shift_start,
                        "nature_of_injury": wiz_nature,
                        "body_part": wiz_body_part,
                        "treatment_level": wiz_treatment,
                        "pre_injury_duties": wiz_pre_injury,
                    }
                    st.session_state.wizard_step = 2
                    st.rerun()

    # ── STEP 2: Medical Details ──
    elif step == 2:
        st.subheader("Step 2: Medical Details")

        if st.button("Back to Step 1"):
            st.session_state.wizard_step = 1
            st.rerun()

        with st.form("wizard_step2"):
            st.markdown("**Treating Doctor**")
            d1, d2 = st.columns(2)
            wiz_doc_name = d1.text_input("Doctor Name")
            wiz_doc_phone = d2.text_input("Doctor Phone")
            wiz_doc_address = st.text_input("Doctor Address")
            wiz_doc_fax = d1.text_input("Doctor Fax")

            st.markdown("**Initial Certificate of Capacity**")
            c1, c2 = st.columns(2)
            wiz_cert_from = c1.date_input("COC From")
            wiz_cert_to = c2.date_input("COC To")
            wiz_capacity = st.selectbox("Current Capacity", ["No Capacity", "Modified Duties", "Full Capacity", "Uncertain"])
            c3, c4 = st.columns(2)
            wiz_days_pw = c3.number_input("Days per Week", min_value=0, max_value=7, value=0)
            wiz_hrs_pd = c4.number_input("Hours per Day", min_value=0.0, max_value=24.0, value=0.0, step=0.5)
            wiz_restrictions = st.text_area("Restrictions / Constraints")

            st.markdown("**Claim Details**")
            cl1, cl2 = st.columns(2)
            wiz_claim = cl1.text_input("Claim Number (if known)")
            wiz_piawe = cl2.number_input("PIAWE ($)", min_value=0.0, value=0.0, step=0.01)
            wiz_reduction = cl1.selectbox("Reduction Rate", ["95%", "80%", "N/A"])
            wiz_shift_structure = cl2.text_input("Shift Structure (e.g. 5 hrs x 3 days)")

            st.markdown("**Strategy & Actions**")
            wiz_strategy = st.text_area("Strategy")
            wiz_next_action = st.text_area("Next Action Required")
            wiz_notes = st.text_area("Notes")

            if st.form_submit_button("Next: Generate Documents", type="primary"):
                wd = st.session_state.wizard_data
                wd.update({
                    "doctor_name": wiz_doc_name,
                    "doctor_phone": wiz_doc_phone,
                    "doctor_address": wiz_doc_address,
                    "doctor_fax": wiz_doc_fax,
                    "cert_from": wiz_cert_from.isoformat(),
                    "cert_to": wiz_cert_to.isoformat(),
                    "current_capacity": wiz_capacity,
                    "days_per_week": wiz_days_pw if wiz_days_pw > 0 else None,
                    "hours_per_day": wiz_hrs_pd if wiz_hrs_pd > 0 else None,
                    "restrictions": wiz_restrictions,
                    "claim_number": wiz_claim or None,
                    "piawe": wiz_piawe if wiz_piawe > 0 else None,
                    "reduction_rate": wiz_reduction,
                    "shift_structure": wiz_shift_structure,
                    "strategy": wiz_strategy,
                    "next_action": wiz_next_action,
                    "notes": wiz_notes,
                })
                st.session_state.wizard_step = 3
                st.rerun()

    # ── STEP 3: Generate Documents ──
    elif step == 3:
        st.subheader("Step 3: Review & Generate Documents")

        if st.button("Back to Step 2"):
            st.session_state.wizard_step = 2
            st.rerun()

        wd = st.session_state.get("wizard_data", {})

        # Show summary
        with st.expander("Case Summary", expanded=True):
            s1, s2 = st.columns(2)
            s1.markdown(f"**Worker:** {wd.get('worker_name', 'N/A')}")
            s1.markdown(f"**State:** {wd.get('state', 'N/A')}")
            s1.markdown(f"**Entity / Site:** {wd.get('entity', '')} - {wd.get('site', '')}")
            s1.markdown(f"**Date of Injury:** {wd.get('date_of_injury', 'N/A')}")
            s2.markdown(f"**Injury:** {wd.get('injury_description', 'N/A')}")
            s2.markdown(f"**Capacity:** {wd.get('current_capacity', 'N/A')}")
            s2.markdown(f"**Claim #:** {wd.get('claim_number', 'N/A')}")
            s2.markdown(f"**Doctor:** {wd.get('doctor_name', 'N/A')}")

        # Document selection
        st.markdown("#### Select Documents to Generate")
        selected_docs = {}
        for doc_key, doc_info in dg.AVAILABLE_DOCUMENTS.items():
            review_badge = ""
            if "Yes" in doc_info["needs_review"]:
                review_badge = "  :orange[Needs Review]"
            elif "Minimal" in doc_info["needs_review"]:
                review_badge = "  :blue[Minimal Review]"
            else:
                review_badge = "  :green[Ready to Use]"

            selected_docs[doc_key] = st.checkbox(
                f"{doc_info['icon']}  **{doc_info['name']}** - {doc_info['description']}{review_badge}",
                key=f"wiz_doc_{doc_key}",
                value=True,  # Default all selected for new case
            )

        st.divider()

        col_create, col_cancel = st.columns([3, 1])

        with col_create:
            if st.button("Create Case & Generate Documents", type="primary", use_container_width=True):
                # 1. Create case in DB
                conn = db.get_connection()
                conn.execute("""
                    INSERT INTO cases (worker_name, state, entity, site, date_of_injury,
                        injury_description, current_capacity, shift_structure, piawe,
                        reduction_rate, claim_number, priority, strategy, next_action, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (wd.get("worker_name"), wd.get("state"), wd.get("entity"), wd.get("site"),
                      wd.get("date_of_injury"), wd.get("injury_description"),
                      wd.get("current_capacity", "Unknown"),
                      wd.get("shift_structure"),
                      wd.get("piawe"),
                      wd.get("reduction_rate", "95%"),
                      wd.get("claim_number"),
                      "MEDIUM",
                      wd.get("strategy"), wd.get("next_action"), wd.get("notes")))
                conn.commit()
                case_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

                # 2. Create document checklist
                doc_types = [
                    "Incident Report", "Claim Form", "Payslips (12 months)",
                    "PIAWE Calculation", "Certificate of Capacity (Current)",
                    "RTW Plan (Current)", "Suitable Duties Plan", "Medical Certificates",
                    "Insurance Correspondence", "Wage Records"
                ]
                for dt in doc_types:
                    conn.execute("INSERT INTO documents (case_id, doc_type) VALUES (?, ?)", (case_id, dt))

                # 3. Save COC if provided
                if wd.get("cert_from") and wd.get("cert_to"):
                    conn.execute("""
                        INSERT INTO certificates (case_id, cert_from, cert_to, capacity, days_per_week, hours_per_day)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (case_id, wd["cert_from"], wd["cert_to"],
                          wd.get("current_capacity"), wd.get("days_per_week"), wd.get("hours_per_day")))

                # 4. Save incident details
                conn.execute("""
                    INSERT INTO incident_details (case_id, dob, occupation, date_reported,
                        task_performed, location_detail, witnesses, employment_type, tenure,
                        shift_type, shift_start_time, nature_of_injury, body_part,
                        treatment_level, pre_injury_duties, avg_hours)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (case_id, wd.get("dob"), wd.get("pre_injury_duties"),
                      wd.get("date_of_injury"), wd.get("injury_description"),
                      wd.get("location_detail"), wd.get("witnesses"),
                      wd.get("employment_type"), wd.get("tenure"),
                      wd.get("shift_type"), wd.get("shift_start_time"),
                      wd.get("nature_of_injury"), wd.get("body_part"),
                      wd.get("treatment_level"), wd.get("pre_injury_duties"),
                      wd.get("avg_hours")))

                # 5. Save doctor details
                if wd.get("doctor_name"):
                    conn.execute("""
                        INSERT INTO doctor_details (case_id, doctor_name, doctor_address, doctor_phone, doctor_fax)
                        VALUES (?, ?, ?, ?, ?)
                    """, (case_id, wd.get("doctor_name"), wd.get("doctor_address"),
                          wd.get("doctor_phone"), wd.get("doctor_fax")))

                conn.commit()
                conn.close()

                log_activity(case_id, "Case Created", f"New case via wizard for {wd.get('worker_name')}")

                # 6. Generate selected documents
                docs_to_generate = [k for k, v in selected_docs.items() if v]
                if docs_to_generate:
                    case_data = wd.copy()
                    medical_data = {
                        "cert_from": wd.get("cert_from"),
                        "cert_to": wd.get("cert_to"),
                        "hours_per_day": wd.get("hours_per_day"),
                        "days_per_week": wd.get("days_per_week"),
                        "restrictions": wd.get("restrictions"),
                        "doctor_name": wd.get("doctor_name"),
                        "doctor_address": wd.get("doctor_address"),
                        "doctor_phone": wd.get("doctor_phone"),
                        "doctor_fax": wd.get("doctor_fax"),
                        "worker_dob": wd.get("dob"),
                        "worker_address": wd.get("address"),
                        "worker_phone": wd.get("phone"),
                        "interpreter_needed": "Yes" if wd.get("language") else "No",
                        "occupation": wd.get("pre_injury_duties"),
                    }
                    doctor_data = {
                        "doctor_name": wd.get("doctor_name"),
                        "doctor_address": wd.get("doctor_address"),
                        "doctor_phone": wd.get("doctor_phone"),
                        "doctor_fax": wd.get("doctor_fax"),
                    }
                    incident_data = {
                        "dob": wd.get("dob"),
                        "occupation": wd.get("pre_injury_duties"),
                        "task_performed": wd.get("injury_description"),
                        "location_detail": wd.get("location_detail"),
                        "witnesses": wd.get("witnesses"),
                        "employment_type": wd.get("employment_type"),
                        "tenure": wd.get("tenure"),
                        "shift_type": wd.get("shift_type"),
                        "shift_start_time": wd.get("shift_start_time"),
                        "nature_of_injury": wd.get("nature_of_injury"),
                        "body_part": wd.get("body_part"),
                        "treatment_level": wd.get("treatment_level"),
                    }

                    with st.spinner("Generating documents..."):
                        results = dg.generate_documents(
                            case_data, docs_to_generate,
                            medical_data=medical_data,
                            doctor_data=doctor_data,
                            incident_data=incident_data,
                        )

                    # Save generated docs to DB
                    conn = db.get_connection()
                    for doc_type, (filename, buf) in results.items():
                        conn.execute(
                            "INSERT INTO generated_documents (case_id, doc_type, doc_name, doc_data) VALUES (?, ?, ?, ?)",
                            (case_id, doc_type, filename, buf.getvalue())
                        )
                    conn.commit()
                    conn.close()

                    log_activity(case_id, "Documents Generated",
                                 f"Generated via wizard: {', '.join(dg.AVAILABLE_DOCUMENTS[k]['name'] for k in docs_to_generate)}")

                    st.success(f"Case created and {len(results)} document(s) generated!")
                    st.balloons()

                    # Show download buttons
                    st.markdown("#### Download Generated Documents")
                    for doc_type, (filename, buf) in results.items():
                        info = dg.AVAILABLE_DOCUMENTS[doc_type]
                        st.download_button(
                            label=f"Download {info['icon']} {info['name']}",
                            data=buf.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"wiz_dl_{doc_type}",
                        )

                    st.markdown("---")
                    if st.button("Open Case", type="primary"):
                        st.session_state.selected_case_id = case_id
                        st.session_state.wizard_step = 1
                        st.session_state.last_page = "Dashboard"
                        st.rerun()
                else:
                    st.success(f"Case created for {wd.get('worker_name')}!")
                    if st.button("Open Case", type="primary"):
                        st.session_state.selected_case_id = case_id
                        st.session_state.wizard_step = 1
                        st.session_state.last_page = "Dashboard"
                        st.rerun()

        with col_cancel:
            if st.button("Cancel", use_container_width=True):
                st.session_state.wizard_step = 1
                if "wizard_data" in st.session_state:
                    del st.session_state.wizard_data
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
            st.info("For a full case setup with automatic document generation, use the **New Case** page from the sidebar.")
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
