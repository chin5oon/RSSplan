"""Streamlit edition of the RSS Plan Builder."""

from __future__ import annotations

import json
from datetime import date

import streamlit as st

from report_generator import build_docx


st.set_page_config(
    page_title="RSS Plan Builder",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

STEPS = [
    "1 · Project",
    "2 · Activities",
    "3 · People",
    "4 · Phasing",
    "5 · Technology",
    "6 · Controls",
    "7 · Records",
    "8 · Review & export",
]

WORK_TYPES = [
    "Concreting works",
    "Post-tensioning works",
    "Safety barriers",
    "Foundation works - bored pile",
    "Foundation works - displacement pile",
    "Foundation works - shallow foundation",
    "Structural steelworks",
    "Demolition",
    "Precast concrete components",
    "Curtainwall and cladding",
    "ERSS works",
    "Other / project-specific activity",
]

APPROACHES = [
    "Technology replacement",
    "Evidence-based supervision",
    "Live remote supervision",
    "In-person attendance",
]

DEFAULTS = {
    "project": {
        "reference": "",
        "description": "",
        "site_type": "Construction site",
        "address": "",
        "structural_system": "",
        "foundation_system": "",
        "challenges": "",
        "permit_date": "",
    },
    "team": {
        "qp_name": "",
        "pe_number": "",
        "company": "",
        "prepared_date": date.today().isoformat(),
        "organisation": "",
        "site_supervisors": "",
        "builder_operators": "",
        "backup_personnel": "",
        "training": "",
        "competency": "",
    },
    "phases": {
        "phase_1": "Maximum 15% remote supervision within the first 30% of works.",
        "phase_2": "Maximum 30% remote supervision from 30% to 75% of works.",
        "phase_3": "Maximum 50% remote supervision from 75% to completion.",
        "beyond": "Any proposal beyond 50% will be supported by proven processes, experienced personnel and activity-specific justification.",
        "criteria": "At least 95% system uptime; zero critical safety incidents; satisfactory QP(S) performance evaluation; evidence complete and retrievable.",
        "parallel_plan": "Conduct parallel in-person and remote supervision during the learning phase and compare outcomes before progression.",
        "review_cadence": "QP(S) review at the end of each phase and after any material change in risk, workmanship, technology or site conditions.",
    },
    "technology": {
        "live_devices": "",
        "evidence_devices": "",
        "audio": "",
        "connectivity": "",
        "backup_connectivity": "",
        "platform": "",
        "video_standard": "",
        "storage": "",
        "power_backup": "",
        "equipment_register": "",
    },
    "process": {
        "before": "Confirm approved drawings and method statements; identify elements; verify lighting, access, safety and equipment readiness; brief all parties.",
        "during": "Introduce participants; confirm location and element IDs; follow the approved checklist; maintain clear two-way communication; record decisions and non-conformities in real time.",
        "after": "Confirm inspection scope and outcome; save evidence; update the inspection register; assign follow-up actions; distribute and archive records.",
        "communication": "Use standard terminology, repeat back critical instructions, keep authorised participants visible where practicable, and record decisions with timestamps.",
        "stop_work": "Stop or suspend remote supervision when the QP(S)/site supervisor cannot clearly verify the works, cannot intervene effectively, or observes a safety or quality risk.",
        "tech_failure": "Switch to tested backup equipment/connectivity. If adequate visibility or communication cannot be restored, revert to in-person supervision and record the incident.",
        "poor_evidence": "Reject unclear, incomplete or untraceable evidence; repeat the inspection or attend in person before accepting the works.",
        "safety_incident": "Prioritise site emergency procedures, suspend RSS, notify the responsible parties and QP(S), and preserve the incident record.",
        "non_conformity": "Record the issue, notify the QP(S), prevent concealment or continuation where required, track rectification, and verify closure before acceptance.",
    },
    "records": {
        "naming": "Use a unique session reference linked to project, activity, location/element, date and revision.",
        "access": "Role-based access for authorised project personnel, with periodic access reviews and an audit trail.",
        "backups": "Automated backup to a separate protected location, with routine restore checks.",
        "retention": "RSS reports and signed logbooks: minimum 5 years after TOP; video and supporting digital evidence: minimum 2 years after TOP; calibration records: minimum 3 years; training records: employment plus 2 years; incident/non-conformity/escalation records: minimum 5 years, or longer where required.",
        "verification": "Conduct parallel in-person inspection for a 10% sample and check evidence completeness before acceptance.",
        "audits": "Internal process and record audits, management review at planned intervals, peer review, and independent review where applicable.",
        "performance": "Track uptime, evidence rejection, defect detection, interventions, safety incidents, response time and phase acceptance.",
        "traceability": "Each record identifies the inspected element/location, date and time, result, supervising personnel, supporting evidence and conformity decision.",
    },
    "activities": [],
    "signoff": {
        "confirm_professional_judgement": False,
        "confirm_annex_d": False,
        "confirm_fallback": False,
        "confirm_records": False,
        "qp_signature": "",
        "sign_date": "",
    },
}


def new_activity() -> dict:
    return {
        "work_type": "Concreting works",
        "description": "",
        "location": "",
        "complexity": "Simple",
        "frequency": "Periodic",
        "approach": "Technology replacement",
        "phase": "Phase 1",
        "extent": "Maximum 15% of the first 30% of works",
        "evidence": "",
        "equipment": "",
        "annex_d_reviewed": False,
        "deviation": "",
    }


def recommended_approach(complexity: str, frequency: str) -> str:
    if complexity == "Simple" and frequency == "Periodic":
        return "Technology replacement"
    if complexity == "Complex" and frequency == "Periodic":
        return "Evidence-based supervision"
    if complexity == "Simple" and frequency == "Continuous":
        return "Live remote supervision"
    return "In-person attendance"


def initialise() -> None:
    if "plan" not in st.session_state:
        st.session_state.plan = json.loads(json.dumps(DEFAULTS))
        st.session_state.plan["activities"] = [new_activity()]
    if "current_step" not in st.session_state:
        st.session_state.current_step = STEPS[0]
    if "site_plan_bytes" not in st.session_state:
        st.session_state.site_plan_bytes = None


def text_area(section: str, key: str, label: str, height: int = 100) -> None:
    plan = st.session_state.plan
    plan[section][key] = st.text_area(
        label,
        value=plan[section].get(key, ""),
        height=height,
        key=f"{section}_{key}",
    )


def validate(plan: dict) -> list[tuple[int, str]]:
    issues: list[tuple[int, str]] = []
    p = plan["project"]
    for key, label in (
        ("reference", "project reference"),
        ("description", "project description"),
        ("address", "supervision location"),
        ("structural_system", "structural system"),
    ):
        if not str(p.get(key, "")).strip():
            issues.append((0, f"Complete the {label}."))
    if not plan["activities"]:
        issues.append((1, "Add at least one structural activity."))
    for index, activity in enumerate(plan["activities"], start=1):
        if not activity["description"].strip() or not activity["location"].strip():
            issues.append((1, f"Complete scope and location for activity {index}."))
        if not activity["evidence"].strip() or not activity["equipment"].strip():
            issues.append((1, f"Complete evidence and equipment for activity {index}."))
        if not activity["annex_d_reviewed"]:
            issues.append((1, f"Confirm Annex D review for activity {index}."))
        if (
            activity["approach"]
            != recommended_approach(activity["complexity"], activity["frequency"])
            and not activity["deviation"].strip()
        ):
            issues.append((1, f"Justify the approach for activity {index}."))
    team = plan["team"]
    if not team["qp_name"].strip() or not team["pe_number"].strip():
        issues.append((2, "Enter the QP(S) name and PE registration number."))
    if not team["site_supervisors"].strip() or not team["backup_personnel"].strip():
        issues.append((2, "Complete site supervision and backup arrangements."))
    if not team["training"].strip() or not team["competency"].strip():
        issues.append((2, "Complete training and competency verification."))
    tech = plan["technology"]
    if any(
        not tech[key].strip()
        for key in ("live_devices", "connectivity", "backup_connectivity", "storage")
    ):
        issues.append((4, "Complete devices, connectivity, backup and storage."))
    sign = plan["signoff"]
    if not all(
        sign[key]
        for key in (
            "confirm_professional_judgement",
            "confirm_annex_d",
            "confirm_fallback",
            "confirm_records",
        )
    ):
        issues.append((7, "Complete all QP(S) declarations."))
    if not sign["qp_signature"].strip() or not sign["sign_date"]:
        issues.append((7, "Enter the QP(S) signatory name and date."))
    return issues


def render_project() -> None:
    st.subheader("Set the supervision context")
    st.info(
        "RSS applies only where the intended supervision outcome can be achieved. "
        "Define the project, site constraints and structural systems clearly."
    )
    plan = st.session_state.plan
    p = plan["project"]
    col1, col2 = st.columns(2)
    with col1:
        p["reference"] = st.text_input(
            "Project reference number *", p["reference"], key="project_reference"
        )
        p["site_type"] = st.selectbox(
            "Supervision location type",
            ["Construction site", "Fabrication yard"],
            index=["Construction site", "Fabrication yard"].index(p["site_type"]),
        )
        p["structural_system"] = st.text_area(
            "Structural system *", p["structural_system"], key="project_structural"
        )
    with col2:
        p["permit_date"] = st.text_input(
            "Permit date", p["permit_date"], placeholder="YYYY-MM-DD"
        )
        p["address"] = st.text_input(
            "Address / location *", p["address"], key="project_address"
        )
        p["foundation_system"] = st.text_area(
            "Foundation system", p["foundation_system"], key="project_foundation"
        )
    p["description"] = st.text_area(
        "Project description *", p["description"], key="project_description"
    )
    p["challenges"] = st.text_area(
        "RSS constraints and challenges", p["challenges"], key="project_challenges"
    )
    uploaded = st.file_uploader(
        "Overall site / location plan (optional)", type=["png", "jpg", "jpeg"]
    )
    if uploaded:
        st.session_state.site_plan_bytes = uploaded.getvalue()
        st.image(uploaded, caption="Overall site / location plan", width=560)


def render_activities() -> None:
    st.subheader("Classify every structural activity")
    st.warning(
        "The matrix is a starting point. Upgrade to in-person attendance whenever "
        "workmanship, site conditions, visibility or intervention capability is inadequate."
    )
    plan = st.session_state.plan
    for index, activity in enumerate(plan["activities"]):
        with st.expander(
            f"Activity {index + 1:02d} · {activity['work_type']}", expanded=True
        ):
            col1, col2 = st.columns(2)
            with col1:
                activity["work_type"] = st.selectbox(
                    "Structural activity",
                    WORK_TYPES,
                    index=WORK_TYPES.index(activity["work_type"]),
                    key=f"work_type_{index}",
                )
                activity["location"] = st.text_input(
                    "Location / element IDs *",
                    activity["location"],
                    key=f"activity_location_{index}",
                )
            with col2:
                activity["complexity"] = st.selectbox(
                    "Inspection complexity",
                    ["Simple", "Complex"],
                    index=["Simple", "Complex"].index(activity["complexity"]),
                    key=f"complexity_{index}",
                )
                activity["frequency"] = st.selectbox(
                    "Supervision frequency",
                    ["Periodic", "Continuous"],
                    index=["Periodic", "Continuous"].index(activity["frequency"]),
                    key=f"frequency_{index}",
                )
            activity["description"] = st.text_area(
                "Scope of supervision *",
                activity["description"],
                key=f"activity_description_{index}",
            )
            recommendation = recommended_approach(
                activity["complexity"], activity["frequency"]
            )
            st.caption(f"Matrix starting point: **{recommendation}**")
            activity["approach"] = st.selectbox(
                "Selected approach",
                APPROACHES,
                index=APPROACHES.index(activity["approach"]),
                key=f"approach_{index}",
            )
            if activity["approach"] != recommendation:
                activity["deviation"] = st.text_area(
                    "Professional justification for alternative approach *",
                    activity["deviation"],
                    key=f"deviation_{index}",
                )
            phase_col, extent_col = st.columns(2)
            with phase_col:
                activity["phase"] = st.selectbox(
                    "Implementation phase",
                    ["Phase 1", "Phase 2", "Phase 3", "Beyond Phase 3"],
                    index=["Phase 1", "Phase 2", "Phase 3", "Beyond Phase 3"].index(
                        activity["phase"]
                    ),
                    key=f"phase_{index}",
                )
            with extent_col:
                activity["extent"] = st.text_input(
                    "Extent of RSS", activity["extent"], key=f"extent_{index}"
                )
            evidence_col, equipment_col = st.columns(2)
            with evidence_col:
                activity["evidence"] = st.text_area(
                    "Evidence and minimum capture *",
                    activity["evidence"],
                    key=f"evidence_{index}",
                )
            with equipment_col:
                activity["equipment"] = st.text_area(
                    "Equipment / software *",
                    activity["equipment"],
                    key=f"equipment_{index}",
                )
            activity["annex_d_reviewed"] = st.checkbox(
                "Applicable Annex D baseline and limitations reviewed",
                value=activity["annex_d_reviewed"],
                key=f"annex_d_{index}",
            )
            if len(plan["activities"]) > 1 and st.button(
                "Remove activity", key=f"remove_{index}"
            ):
                plan["activities"].pop(index)
                st.rerun()
    if st.button("＋ Add another activity", type="secondary"):
        plan["activities"].append(new_activity())
        st.rerun()


def render_people() -> None:
    st.subheader("Make accountability explicit")
    st.info(
        "Remote tools support - but do not transfer - the QP(S)'s professional responsibility."
    )
    team = st.session_state.plan["team"]
    c1, c2, c3 = st.columns(3)
    with c1:
        team["qp_name"] = st.text_input("QP(S) name *", team["qp_name"])
    with c2:
        team["pe_number"] = st.text_input(
            "PE registration number *", team["pe_number"]
        )
    with c3:
        team["company"] = st.text_input("Company", team["company"])
    for key, label in (
        ("organisation", "Organisation and reporting lines"),
        ("site_supervisors", "Site supervisors (RE/RTO) *"),
        ("builder_operators", "Builder-side RSS operators"),
        ("backup_personnel", "Backup personnel and handover *"),
        ("training", "Training programme *"),
        ("competency", "Competency verification *"),
    ):
        team[key] = st.text_area(label, team[key], key=f"team_{key}")


def render_phasing() -> None:
    st.subheader("Build confidence in controlled phases")
    c1, c2, c3 = st.columns(3)
    c1.metric("Phase 1", "15%", "of first 30% of works")
    c2.metric("Phase 2", "30%", "30-75% of works")
    c3.metric("Phase 3", "50%", "75-100% of works")
    st.warning(
        "During the first 30% of works, conduct parallel in-person and remote "
        "supervision to compare outcomes and permit immediate intervention."
    )
    for key, label in (
        ("phase_1", "Phase 1 approach"),
        ("phase_2", "Phase 2 approach"),
        ("phase_3", "Phase 3 approach"),
        ("beyond", "Proposal beyond Phase 3"),
        ("criteria", "Phase acceptance criteria *"),
        ("parallel_plan", "Parallel supervision plan *"),
        ("review_cadence", "Review and adjustment cadence"),
    ):
        text_area("phases", key, label)


def render_technology() -> None:
    st.subheader("Specify the complete supervision chain")
    st.info(
        "The QP(S) should be able to see, hear, measure, intervene, record and "
        "retrieve evidence. Each critical link needs a tested backup."
    )
    labels = {
        "live_devices": "Live streaming devices *",
        "evidence_devices": "Evidence and measurement devices",
        "audio": "Two-way audio",
        "connectivity": "Primary connectivity *",
        "backup_connectivity": "Backup connectivity *",
        "platform": "RSS platform",
        "video_standard": "Video and recording standard",
        "storage": "Secure storage *",
        "power_backup": "Power and battery backup",
        "equipment_register": "Equipment register and calibration",
    }
    keys = list(labels)
    for start in range(0, len(keys), 2):
        cols = st.columns(2)
        for col, key in zip(cols, keys[start : start + 2]):
            with col:
                text_area("technology", key, labels[key])


def render_controls() -> None:
    st.subheader("Define what happens before, during and after")
    st.caption("PREPARE → VERIFY LIVE → DECIDE → RECORD → CLOSE")
    for key, label in (
        ("before", "Before remote supervision"),
        ("during", "During remote supervision"),
        ("after", "After remote supervision"),
        ("communication", "Communication protocol"),
    ):
        text_area("process", key, label, 120)
    st.markdown("#### Hard stops and fallback")
    for key, label in (
        ("stop_work", "Stop-work / revert-to-site trigger *"),
        ("tech_failure", "Technology failure *"),
        ("poor_evidence", "Poor or incomplete evidence"),
        ("safety_incident", "Safety incident"),
        ("non_conformity", "Non-conformity and escalation *"),
    ):
        text_area("process", key, label)


def render_records() -> None:
    st.subheader("Make every decision traceable")
    st.info(
        "Images and recordings should show scale, orientation, location and "
        "element identification. Evidence without context may not support verification."
    )
    for key, label in (
        ("naming", "Naming and indexing"),
        ("access", "Access and data security"),
        ("backups", "Backup and recovery"),
        ("retention", "Retention schedule *"),
        ("verification", "Verification plan *"),
        ("audits", "Audits and management review"),
        ("performance", "Performance monitoring"),
        ("traceability", "Session traceability *"),
    ):
        text_area("records", key, label)


def render_review() -> None:
    st.subheader("Resolve gaps, sign and export")
    plan = st.session_state.plan
    issues = validate(plan)
    if issues:
        st.warning(
            f"{len(issues)} readiness item(s) remain. The report can be generated "
            "for review, but should not be treated as submission-ready."
        )
        for step_index, issue in issues:
            st.write(f"**{STEPS[step_index]}** — {issue}")
    else:
        st.success(
            "All builder checks are complete. Conduct the final professional and "
            "current-requirements review before submission."
        )

    st.markdown("#### QP(S) declarations")
    sign = plan["signoff"]
    sign["confirm_professional_judgement"] = st.checkbox(
        "I applied professional judgement to the suitability of RSS for each activity.",
        sign["confirm_professional_judgement"],
    )
    sign["confirm_annex_d"] = st.checkbox(
        "I reviewed the applicable Annex D minimum requirements and documented alternatives.",
        sign["confirm_annex_d"],
    )
    sign["confirm_fallback"] = st.checkbox(
        "I will require in-person supervision whenever the intended outcome cannot be achieved remotely.",
        sign["confirm_fallback"],
    )
    sign["confirm_records"] = st.checkbox(
        "RSS records will be controlled, protected, retrievable and retained as stated.",
        sign["confirm_records"],
    )
    c1, c2 = st.columns(2)
    with c1:
        sign["qp_signature"] = st.text_input(
            "QP(S) signatory name *", sign["qp_signature"]
        )
    with c2:
        sign["sign_date"] = st.text_input(
            "Sign date *", sign["sign_date"], placeholder="YYYY-MM-DD"
        )

    st.markdown("#### Submission package")
    docx_bytes = build_docx(plan, st.session_state.site_plan_bytes)
    safe_ref = "".join(
        c if c.isalnum() or c in "-_" else "-"
        for c in plan["project"]["reference"]
    ).strip("-") or "RSS-Plan"
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download native Word report (.docx)",
            data=docx_bytes,
            file_name=f"{safe_ref}-Remote-Site-Supervision-Plan.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "Export editable working file (.json)",
            data=json.dumps(plan, indent=2),
            file_name=f"{safe_ref}-working-file.json",
            mime="application/json",
            use_container_width=True,
        )


def render_sidebar() -> str:
    with st.sidebar:
        st.markdown("## RSS Plan Builder")
        st.caption("For Qualified Persons")
        selected = st.radio(
            "PLAN PROGRESS",
            STEPS,
            index=STEPS.index(st.session_state.current_step),
            label_visibility="visible",
        )
        st.session_state.current_step = selected
        issues = validate(st.session_state.plan)
        completion = max(0, round(100 - min(len(issues), 20) / 20 * 100))
        st.progress(completion / 100, text=f"{completion}% builder complete")
        st.divider()
        st.caption("Guide basis")
        st.write("RSS Guidebook v2.0 · June 2026")
        st.caption(
            "Statutory requirements and the latest BCA guidance prevail."
        )
        working_upload = st.file_uploader(
            "Import working file", type=["json"], key="working_file_upload"
        )
        if working_upload and st.button("Load imported file"):
            try:
                imported = json.load(working_upload)
                if not all(
                    key in imported
                    for key in ("project", "activities", "team", "signoff")
                ):
                    raise ValueError("Missing sections")
                st.session_state.plan = imported
                st.success("Working file loaded.")
                st.rerun()
            except Exception:
                st.error("This is not a valid RSS working file.")
    return selected


st.markdown(
    """
<style>
    :root { --ink:#112536; --teal:#087e8b; --wash:#f3f7f6; }
    .stApp { background: #f2f6f5; color: var(--ink); }
    [data-testid="stSidebar"] { background: var(--ink); }
    [data-testid="stSidebar"] * { color: #e8eff2; }
    [data-testid="stSidebar"] [data-testid="stProgressBar"] div div { background:#11a7a3; }
    h1, h2, h3 { color:var(--ink); letter-spacing:-.02em; }
    div[data-testid="stExpander"] { border-color:#cfdbde; background:white; }
    .block-container { max-width: 1120px; padding-top: 2.6rem; }
    .rss-hero { background:white; padding:28px 32px; border:1px solid #d6e0e3;
      border-radius:14px; margin-bottom:26px; box-shadow:0 16px 42px rgba(20,49,62,.07); }
    .rss-eyebrow { color:#087e8b; font-size:11px; letter-spacing:.15em;
      font-weight:800; text-transform:uppercase; }
    .rss-hero h1 { font-family:Georgia,serif; font-size:46px; font-weight:500;
      line-height:1.04; margin:8px 0 12px; }
    .rss-hero p { max-width:760px; color:#536976; line-height:1.6; margin:0; }
</style>
<div class="rss-hero">
  <div class="rss-eyebrow">BCA-aligned guided workflow</div>
  <h1>Build a defensible remote supervision plan.</h1>
  <p>Move from activity assessment to a structured submission report, with
  professional judgement and fallback controls visible at every step.</p>
</div>
""",
    unsafe_allow_html=True,
)

initialise()
selected_step = render_sidebar()

renderers = {
    STEPS[0]: render_project,
    STEPS[1]: render_activities,
    STEPS[2]: render_people,
    STEPS[3]: render_phasing,
    STEPS[4]: render_technology,
    STEPS[5]: render_controls,
    STEPS[6]: render_records,
    STEPS[7]: render_review,
}
renderers[selected_step]()

st.divider()
index = STEPS.index(selected_step)
left, _, right = st.columns([1, 4, 1])
with left:
    if index > 0 and st.button("← Back", use_container_width=True):
        st.session_state.current_step = STEPS[index - 1]
        st.rerun()
with right:
    if index < len(STEPS) - 1 and st.button(
        "Continue →", type="primary", use_container_width=True
    ):
        st.session_state.current_step = STEPS[index + 1]
        st.rerun()

st.caption(
    "This tool supports preparation; it does not replace the QP(S)'s statutory "
    "duties, professional judgement or review of current BCA requirements."
)

