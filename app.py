"""Streamlit edition of the RSS Plan Builder."""

from __future__ import annotations

import json
from datetime import date
from uuid import uuid4

import streamlit as st

from report_generator import build_docx


st.set_page_config(
    page_title="RSS Plan Builder",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

STEPS = [
    "1 - Project",
    "2 - People",
    "3 - Phasing framework",
    "4 - Technology",
    "5 - Controls",
    "6 - Records",
    "7 - Activities",
    "8 - Review & export",
]

ACTIVITY_GROUPS = {
    "Structural supervision checklists": [
        "Concreting works",
        "Post-tensioning works",
        "Safety barriers",
        "Foundation works (Bored Pile)",
        "Foundation works (Displacement Pile)",
        "Foundation works (Shallow Foundation)",
        "Structural Steelworks",
        "Mass Engineered Timber",
        "Reinforcement",
        "Demolition",
        "Precast Concrete Components",
        "Curtainwall & Cladding (Stick System)",
        "Curtainwall & Cladding (Unitised System)",
        "PPVC (Steel)",
        "PPVC (Concrete)",
        "Post-Installed Anchors/Rebars",
        "ERSS Works",
    ],
    "Bored tunnelling": [
        "Bored Tunnelling Works",
    ],
    "Material tests - structural concrete": [
        "Initial test for concrete (normal)",
        "Initial test for concrete (waterproof)",
        "Cement",
        "Aggregate",
        "Admixture",
        "Concrete",
        "Steel reinforcement",
        "Coupler for mechanical splices of steel reinforcement",
        "Welded steel fabric reinforcement",
        "Wire loop / slim box (concrete PPVC)",
    ],
    "Material tests - post-tensioning": [
        "Steel for Prestressing Concrete",
        "Grout for Prestressing Concrete",
    ],
    "Material tests - structural steelworks": [
        "Structural steel",
        "Structural steel welding (NDT)",
        "Shear studs (weld)",
        "Bolts, screws and studs (Carbon steel & alloy steel)",
        "Bolts, screws and studs (Stainless steel)",
        "High friction grip bolts",
    ],
    "Fabrication-yard supervision": [
        "Fabrication of Precast Concrete",
        "Fabrication of Structural Steelworks",
    ],
    "Project-specific": [
        "Other / project-specific supervision activity",
    ],
}

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
        "competency_provider_training": False,
        "competency_trial": False,
        "competency_registration": False,
        "competency_upgrade_training": False,
        "competency_evidence": "",
        "competency_verifier": "",
        "competency_date": "",
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
    "profiles": {
        "people": [
            {
                "id": "P1",
                "name": "Project default RSS team",
                "default": True,
                "site_supervisors": "",
                "builder_operators": "",
                "backup_personnel": "",
                "notes": "",
            }
        ],
        "technology": [
            {
                "id": "T1",
                "name": "Project default technology set",
                "default": True,
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
            }
        ],
        "controls": [
            {
                "id": "C1",
                "name": "Project default control profile",
                "default": True,
                "before": "",
                "during": "",
                "after": "",
                "communication": "",
                "stop_work": "",
                "tech_failure": "",
                "poor_evidence": "",
                "safety_incident": "",
                "non_conformity": "",
            }
        ],
        "records": [
            {
                "id": "R1",
                "name": "Project default record profile",
                "default": True,
                "naming": "",
                "access": "",
                "backups": "",
                "retention": "",
                "verification": "",
                "audits": "",
                "performance": "",
                "traceability": "",
            }
        ],
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


PHASE_PRESETS = {
    "Phase 1": {
        "progress_range": "First 30% of the activity",
        "rss_extent": "Maximum 15% remote supervision",
        "parallel_supervision": "Use the project parallel-supervision plan.",
    },
    "Phase 2": {
        "progress_range": "30% to 75% of the activity",
        "rss_extent": "Maximum 30% remote supervision",
        "parallel_supervision": "Targeted in-person verification as required.",
    },
    "Phase 3": {
        "progress_range": "75% to 100% of the activity",
        "rss_extent": "Maximum 50% remote supervision",
        "parallel_supervision": "Targeted in-person verification as required.",
    },
    "Beyond Phase 3": {
        "progress_range": "After satisfactory completion of Phase 3",
        "rss_extent": "State the professionally justified proposed extent",
        "parallel_supervision": "State the retained in-person verification.",
    },
    "Custom": {
        "progress_range": "",
        "rss_extent": "",
        "parallel_supervision": "",
    },
}

PROFILE_PREFIXES = {
    "people": "P",
    "technology": "T",
    "controls": "C",
    "records": "R",
}

PROFILE_ASSIGNMENT_KEYS = {
    "people": "people_profile_id",
    "technology": "technology_profile_id",
    "controls": "control_profile_id",
    "records": "record_profile_id",
}


def new_implementation_phase(position: int = 0) -> dict:
    names = list(PHASE_PRESETS)
    name = names[min(position, len(names) - 1)]
    preset = PHASE_PRESETS[name]
    return {
        "uid": uuid4().hex[:10],
        "name": name,
        "custom_name": "",
        "progress_range": preset["progress_range"],
        "rss_extent": preset["rss_extent"],
        "parallel_supervision": preset["parallel_supervision"],
        "acceptance_criteria": "Use the project phase-acceptance criteria.",
        "review_point": "QP(S) review before progression to the next phase.",
        "remarks": "",
    }


def new_activity() -> dict:
    return {
        "uid": uuid4().hex[:10],
        "category": "",
        "work_type": "",
        "description": "",
        "location": "",
        "complexity": "Simple",
        "frequency": "Periodic",
        "approach": "Technology replacement",
        "implementation_phases": [new_implementation_phase()],
        "people_profile_id": "P1",
        "technology_profile_id": "T1",
        "control_profile_id": "C1",
        "record_profile_id": "R1",
        "personnel_requirements": "",
        "evidence": "",
        "equipment": "",
        "control_overrides": "",
        "record_overrides": "",
        "annex_d_reviewed": False,
        "deviation": "",
    }


def _merge_missing(target: dict, defaults: dict) -> dict:
    for key, value in defaults.items():
        if key not in target:
            target[key] = json.loads(json.dumps(value))
        elif isinstance(value, dict) and isinstance(target.get(key), dict):
            _merge_missing(target[key], value)
    return target


def migrate_plan(plan: dict) -> dict:
    """Upgrade older working files without discarding user-entered content."""
    _merge_missing(plan, DEFAULTS)
    profiles = plan.setdefault("profiles", {})
    for kind, default_profiles in DEFAULTS["profiles"].items():
        existing = profiles.setdefault(kind, [])
        default_id = default_profiles[0]["id"]
        if not any(profile.get("id") == default_id for profile in existing):
            existing.insert(0, json.loads(json.dumps(default_profiles[0])))
        for profile in existing:
            profile.setdefault("name", f"{profile.get('id', '')} profile")
            profile.setdefault("default", profile.get("id") == default_id)
            _merge_missing(profile, default_profiles[0])

    for activity in plan.setdefault("activities", []):
        legacy_phase = activity.get("phase", "Phase 1")
        legacy_extent = activity.get("extent", "")
        had_implementation_phases = bool(activity.get("implementation_phases"))
        _merge_missing(activity, new_activity())
        if not had_implementation_phases:
            position = (
                list(PHASE_PRESETS).index(legacy_phase)
                if legacy_phase in PHASE_PRESETS
                else 0
            )
            phase = new_implementation_phase(position)
            phase["name"] = legacy_phase or "Phase 1"
            if legacy_extent:
                phase["rss_extent"] = legacy_extent
            activity["implementation_phases"] = [phase]
        for position, phase in enumerate(activity["implementation_phases"]):
            _merge_missing(phase, new_implementation_phase(position))
        activity["category"] = activity.get("category") or category_for_activity(
            activity.get("work_type", "")
        )
        for kind, assignment_key in PROFILE_ASSIGNMENT_KEYS.items():
            available_ids = {
                profile.get("id") for profile in profiles.get(kind, [])
            }
            if activity.get(assignment_key) not in available_ids:
                activity[assignment_key] = DEFAULTS["profiles"][kind][0]["id"]
    return plan


def new_profile(kind: str, plan: dict) -> dict:
    template = json.loads(json.dumps(DEFAULTS["profiles"][kind][0]))
    used = {
        profile.get("id", "")
        for profile in plan.setdefault("profiles", {}).setdefault(kind, [])
    }
    prefix = PROFILE_PREFIXES[kind]
    number = 2
    while f"{prefix}{number}" in used:
        number += 1
    template["id"] = f"{prefix}{number}"
    template["name"] = f"Alternative {kind.title()} profile {number}"
    template["default"] = False
    return template


def category_for_activity(work_type: str) -> str:
    for category, activities in ACTIVITY_GROUPS.items():
        if work_type in activities:
            return category
    return "Project-specific" if work_type else ""


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
    st.session_state.plan = migrate_plan(st.session_state.plan)
    if (
        "current_step" not in st.session_state
        or st.session_state.current_step not in STEPS
    ):
        st.session_state.current_step = STEPS[0]
    if "site_plan_bytes" not in st.session_state:
        st.session_state.site_plan_bytes = None
    if "org_chart_bytes" not in st.session_state:
        st.session_state.org_chart_bytes = None


def text_area(section: str, key: str, label: str, height: int = 100) -> None:
    plan = st.session_state.plan
    plan[section][key] = st.text_area(
        label,
        value=plan[section].get(key, ""),
        height=height,
        key=f"{section}_{key}",
    )


PROFILE_FIELD_LABELS = {
    "people": [
        ("site_supervisors", "Site supervisors assigned"),
        ("builder_operators", "Builder-side operators assigned"),
        ("backup_personnel", "Backup personnel assigned"),
        ("notes", "Deployment / handover notes"),
    ],
    "technology": [
        ("live_devices", "Live-streaming devices"),
        ("evidence_devices", "Evidence / measurement devices"),
        ("audio", "Two-way audio"),
        ("connectivity", "Primary connectivity"),
        ("backup_connectivity", "Backup connectivity"),
        ("platform", "RSS platform / software"),
        ("video_standard", "Video / recording standard"),
        ("storage", "Secure storage"),
        ("power_backup", "Power / battery backup"),
        ("equipment_register", "Equipment register / calibration"),
    ],
    "controls": [
        ("before", "Preparation variation"),
        ("during", "Live-supervision variation"),
        ("after", "Close-out variation"),
        ("communication", "Communication variation"),
        ("stop_work", "Stop-work / in-person trigger variation"),
        ("tech_failure", "Technology-failure variation"),
        ("poor_evidence", "Poor-evidence variation"),
        ("safety_incident", "Safety-incident variation"),
        ("non_conformity", "Non-conformity variation"),
    ],
    "records": [
        ("naming", "Naming / indexing variation"),
        ("access", "Access / data-security variation"),
        ("backups", "Backup / recovery variation"),
        ("retention", "Retention variation"),
        ("verification", "Verification variation"),
        ("audits", "Audit variation"),
        ("performance", "Performance-monitoring variation"),
        ("traceability", "Traceability variation"),
    ],
}


def profile_label(plan: dict, kind: str, profile_id: str) -> str:
    for profile in plan.get("profiles", {}).get(kind, []):
        if profile.get("id") == profile_id:
            return f"{profile_id} - {profile.get('name', 'Unnamed profile')}"
    return profile_id


def phase_label(phase: dict) -> str:
    if phase.get("name") == "Custom":
        return phase.get("custom_name", "").strip() or "Custom phase"
    return phase.get("name", "").strip() or "Unnamed phase"


def render_profile_manager(kind: str, description: str) -> None:
    plan = st.session_state.plan
    profiles = plan["profiles"][kind]
    st.markdown(f"#### Reusable {kind.title()} profiles")
    st.caption(description)
    for profile in list(profiles):
        profile_id = profile["id"]
        if profile.get("default"):
            st.info(
                f"**{profile_id} - {profile['name']}** uses the project-wide "
                f"{kind} information entered above."
            )
            continue
        with st.expander(
            f"{profile_id} - {profile.get('name', 'Unnamed profile')}",
            expanded=False,
        ):
            profile["name"] = st.text_input(
                "Profile name *",
                profile.get("name", ""),
                key=f"profile_{kind}_{profile_id}_name",
            )
            fields = PROFILE_FIELD_LABELS[kind]
            for start in range(0, len(fields), 2):
                columns = st.columns(2)
                for column, (field, label) in zip(columns, fields[start : start + 2]):
                    with column:
                        profile[field] = st.text_area(
                            label,
                            profile.get(field, ""),
                            key=f"profile_{kind}_{profile_id}_{field}",
                            help=(
                                "Leave blank to inherit the project-wide default. "
                                "Enter only the variation for this profile."
                            ),
                        )
            if st.button(
                f"Remove {profile_id}",
                key=f"remove_profile_{kind}_{profile_id}",
            ):
                profiles.remove(profile)
                default_id = DEFAULTS["profiles"][kind][0]["id"]
                assignment_key = PROFILE_ASSIGNMENT_KEYS[kind]
                for activity in plan["activities"]:
                    if activity.get(assignment_key) == profile_id:
                        activity[assignment_key] = default_id
                st.rerun()
    if st.button(
        f"+ Add {kind} profile",
        key=f"add_profile_{kind}",
        type="secondary",
    ):
        profiles.append(new_profile(kind, plan))
        st.rerun()


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
    activity_step = 6
    if not plan["activities"]:
        issues.append((activity_step, "Add at least one structural activity."))
    for index, activity in enumerate(plan["activities"], start=1):
        if not activity.get("work_type", "").strip():
            issues.append(
                (activity_step, f"Select the supervision activity for activity {index}.")
            )
        if not activity["description"].strip() or not activity["location"].strip():
            issues.append(
                (activity_step, f"Complete scope and location for activity {index}.")
            )
        if not activity["evidence"].strip():
            issues.append(
                (activity_step, f"Complete evidence requirements for activity {index}.")
            )
        phases = activity.get("implementation_phases") or []
        if not phases:
            issues.append(
                (activity_step, f"Add an implementation phase for activity {index}.")
            )
        for phase_index, phase in enumerate(phases, start=1):
            if any(
                not str(phase.get(key, "")).strip()
                for key in (
                    "name",
                    "progress_range",
                    "rss_extent",
                    "acceptance_criteria",
                    "review_point",
                )
            ):
                issues.append(
                    (
                        activity_step,
                        f"Complete phase {phase_index} for activity {index}.",
                    )
                )
            if (
                phase.get("name") == "Custom"
                and not phase.get("custom_name", "").strip()
            ):
                issues.append(
                    (
                        activity_step,
                        f"Name custom phase {phase_index} for activity {index}.",
                    )
                )
        for kind, assignment_key in PROFILE_ASSIGNMENT_KEYS.items():
            available_ids = {
                profile.get("id")
                for profile in plan.get("profiles", {}).get(kind, [])
            }
            if activity.get(assignment_key) not in available_ids:
                issues.append(
                    (
                        activity_step,
                        f"Select a valid {kind} profile for activity {index}.",
                    )
                )
        if not activity["annex_d_reviewed"]:
            issues.append(
                (activity_step, f"Confirm Annex D review for activity {index}.")
            )
        if (
            activity["approach"]
            != recommended_approach(activity["complexity"], activity["frequency"])
            and not activity["deviation"].strip()
        ):
            issues.append((activity_step, f"Justify the approach for activity {index}."))
    team = plan["team"]
    if not team["qp_name"].strip() or not team["pe_number"].strip():
        issues.append((1, "Enter the QP(S) name and PE registration number."))
    if not team["site_supervisors"].strip() or not team["backup_personnel"].strip():
        issues.append((1, "Complete site supervision and backup arrangements."))
    if not team["training"].strip():
        issues.append((1, "Complete the RSS training programme."))
    competency_checks = (
        "competency_provider_training",
        "competency_trial",
        "competency_registration",
        "competency_upgrade_training",
    )
    if not all(team.get(key, False) for key in competency_checks):
        issues.append((1, "Confirm all competency-verification checks."))
    if (
        not team.get("competency_evidence", "").strip()
        or not team.get("competency_verifier", "").strip()
        or not team.get("competency_date", "").strip()
    ):
        issues.append(
            (1, "Record competency evidence, the verifier and verification date.")
        )
    phases = plan["phases"]
    if not phases.get("criteria", "").strip() or not phases.get(
        "parallel_plan", ""
    ).strip():
        issues.append((2, "Complete the project phase gates and parallel plan."))
    tech = plan["technology"]
    if any(
        not tech[key].strip()
        for key in ("live_devices", "connectivity", "backup_connectivity", "storage")
    ):
        issues.append((3, "Complete devices, connectivity, backup and storage."))
    for kind, step_index in (
        ("people", 1),
        ("technology", 3),
        ("controls", 4),
        ("records", 5),
    ):
        for profile in plan.get("profiles", {}).get(kind, []):
            if not profile.get("name", "").strip():
                issues.append(
                    (step_index, f"Name the {profile.get('id', kind)} profile.")
                )
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
    st.subheader("Apply the project arrangements to each activity")
    st.warning(
        "The matrix is a starting point. Upgrade to in-person attendance whenever "
        "workmanship, site conditions, visibility or intervention capability is inadequate."
    )
    st.caption(
        "Select reusable project profiles, then record only activity-specific "
        "requirements and deviations. Each activity can have several implementation phases."
    )
    plan = st.session_state.plan
    for index, activity in enumerate(plan["activities"]):
        activity_id = activity["uid"]
        activity["category"] = activity.get("category") or category_for_activity(
            activity.get("work_type", "")
        )
        selected_name = activity.get("work_type") or "Select an activity"
        with st.expander(
            f"Activity {index + 1:02d} - {selected_name}", expanded=True
        ):
            col1, col2 = st.columns(2)
            with col1:
                categories = list(ACTIVITY_GROUPS)
                current_category = activity.get("category", "")
                category_index = (
                    categories.index(current_category)
                    if current_category in categories
                    else None
                )
                selected_category = st.selectbox(
                    "Activity group *",
                    categories,
                    index=category_index,
                    placeholder="Select a guide category",
                    key=f"activity_category_{activity_id}",
                )
                if selected_category != current_category:
                    activity["category"] = selected_category or ""
                    activity["work_type"] = ""
                choices = ACTIVITY_GROUPS.get(activity.get("category", ""), [])
                current_work_type = activity.get("work_type", "")
                work_type_index = (
                    choices.index(current_work_type)
                    if current_work_type in choices
                    else None
                )
                activity["work_type"] = (
                    st.selectbox(
                        "Supervision activity *",
                        choices,
                        index=work_type_index,
                        placeholder="Select the activity",
                        key=f"work_type_{activity_id}_{activity.get('category', 'none')}",
                        disabled=not choices,
                    )
                    or ""
                )
                activity["location"] = st.text_input(
                    "Location / element IDs *",
                    activity["location"],
                    key=f"activity_location_{activity_id}",
                )
            with col2:
                activity["complexity"] = st.selectbox(
                    "Inspection complexity",
                    ["Simple", "Complex"],
                    index=["Simple", "Complex"].index(activity["complexity"]),
                    key=f"complexity_{activity_id}",
                )
                activity["frequency"] = st.selectbox(
                    "Supervision frequency",
                    ["Periodic", "Continuous"],
                    index=["Periodic", "Continuous"].index(activity["frequency"]),
                    key=f"frequency_{activity_id}",
                )
            activity["description"] = st.text_area(
                "Scope of supervision *",
                activity["description"],
                key=f"activity_description_{activity_id}",
            )
            recommendation = recommended_approach(
                activity["complexity"], activity["frequency"]
            )
            st.caption(f"Matrix starting point: **{recommendation}**")
            activity["approach"] = st.selectbox(
                "Selected approach",
                APPROACHES,
                index=(
                    APPROACHES.index(activity["approach"])
                    if activity.get("approach") in APPROACHES
                    else APPROACHES.index(recommendation)
                ),
                key=f"approach_{activity_id}",
            )
            if activity["approach"] != recommendation:
                activity["deviation"] = st.text_area(
                    "Professional justification for alternative approach *",
                    activity["deviation"],
                    key=f"deviation_{activity_id}",
                )

            st.markdown("##### Project-profile assignments")
            assignment_columns = st.columns(4)
            assignment_labels = {
                "people": "People",
                "technology": "Technology",
                "controls": "Controls",
                "records": "Records",
            }
            for column, kind in zip(assignment_columns, assignment_labels):
                profiles = plan["profiles"][kind]
                options = [profile["id"] for profile in profiles]
                assignment_key = PROFILE_ASSIGNMENT_KEYS[kind]
                current_id = activity.get(assignment_key, options[0])
                if current_id not in options:
                    current_id = options[0]
                with column:
                    activity[assignment_key] = st.selectbox(
                        assignment_labels[kind],
                        options,
                        index=options.index(current_id),
                        format_func=lambda value, k=kind: profile_label(plan, k, value),
                        key=f"{assignment_key}_{activity_id}",
                    )

            st.markdown("##### Activity-specific requirements")
            st.caption(
                "Do not repeat the profile text. Enter the evidence needed for this "
                "activity and only the variations from the selected profiles."
            )
            req1, req2 = st.columns(2)
            with req1:
                activity["personnel_requirements"] = st.text_area(
                    "Personnel assignment / competency variation",
                    activity.get("personnel_requirements", ""),
                    key=f"personnel_requirements_{activity_id}",
                    placeholder="Optional - leave blank when the selected people profile applies.",
                )
                activity["evidence"] = st.text_area(
                    "Evidence and minimum capture *",
                    activity["evidence"],
                    key=f"evidence_{activity_id}",
                )
            with req2:
                activity["equipment"] = st.text_area(
                    "Equipment / software variation",
                    activity["equipment"],
                    key=f"equipment_{activity_id}",
                    placeholder="Optional - leave blank when the selected technology profile applies.",
                )
                activity["control_overrides"] = st.text_area(
                    "Activity-specific hold points / control variation",
                    activity.get("control_overrides", ""),
                    key=f"control_overrides_{activity_id}",
                )
            activity["record_overrides"] = st.text_area(
                "Record, retention or verification variation",
                activity.get("record_overrides", ""),
                key=f"record_overrides_{activity_id}",
                placeholder="Optional - leave blank when the selected record profile applies.",
            )

            st.markdown("##### Implementation phases")
            st.caption(
                "Add every phase that applies to this activity. Progression is "
                "activity-specific and remains subject to the project-wide gates."
            )
            phases = activity["implementation_phases"]
            for phase_index, phase in enumerate(list(phases)):
                phase_id = phase["uid"]
                with st.container(border=True):
                    phase_title, remove_column = st.columns([5, 1])
                    with phase_title:
                        st.markdown(
                            f"**Phase entry {phase_index + 1}: "
                            f"{phase_label(phase)}**"
                        )
                    with remove_column:
                        if len(phases) > 1 and st.button(
                            "Remove",
                            key=f"remove_phase_{activity_id}_{phase_id}",
                            use_container_width=True,
                        ):
                            phases.remove(phase)
                            st.rerun()
                    row1, row2, row3 = st.columns(3)
                    with row1:
                        phase["name"] = st.selectbox(
                            "Phase",
                            list(PHASE_PRESETS),
                            index=list(PHASE_PRESETS).index(phase["name"])
                            if phase["name"] in PHASE_PRESETS
                            else 0,
                            key=f"phase_name_{activity_id}_{phase_id}",
                        )
                        if phase["name"] == "Custom":
                            phase["custom_name"] = st.text_input(
                                "Custom phase name *",
                                phase.get("custom_name", ""),
                                key=f"phase_custom_name_{activity_id}_{phase_id}",
                            )
                        if st.button(
                            "Apply selected phase guide defaults",
                            key=f"phase_defaults_{activity_id}_{phase_id}",
                            help=(
                                "Replaces the progress range, RSS extent and "
                                "parallel-supervision entry for this phase."
                            ),
                        ):
                            preset = PHASE_PRESETS[phase["name"]]
                            phase["progress_range"] = preset["progress_range"]
                            phase["rss_extent"] = preset["rss_extent"]
                            phase["parallel_supervision"] = preset[
                                "parallel_supervision"
                            ]
                            for widget_key in (
                                f"phase_progress_{activity_id}_{phase_id}",
                                f"phase_extent_{activity_id}_{phase_id}",
                                f"phase_parallel_{activity_id}_{phase_id}",
                            ):
                                st.session_state.pop(widget_key, None)
                            st.rerun()
                    with row2:
                        phase["progress_range"] = st.text_input(
                            "Activity progress range *",
                            phase.get("progress_range", ""),
                            key=f"phase_progress_{activity_id}_{phase_id}",
                        )
                    with row3:
                        phase["rss_extent"] = st.text_input(
                            "Extent of RSS *",
                            phase.get("rss_extent", ""),
                            key=f"phase_extent_{activity_id}_{phase_id}",
                        )
                    row4, row5 = st.columns(2)
                    with row4:
                        phase["parallel_supervision"] = st.text_area(
                            "Parallel / in-person verification",
                            phase.get("parallel_supervision", ""),
                            key=f"phase_parallel_{activity_id}_{phase_id}",
                        )
                        phase["acceptance_criteria"] = st.text_area(
                            "Acceptance / progression criteria *",
                            phase.get("acceptance_criteria", ""),
                            key=f"phase_acceptance_{activity_id}_{phase_id}",
                        )
                    with row5:
                        phase["review_point"] = st.text_area(
                            "QP(S) review point *",
                            phase.get("review_point", ""),
                            key=f"phase_review_{activity_id}_{phase_id}",
                        )
                        phase["remarks"] = st.text_area(
                            "Phase remarks",
                            phase.get("remarks", ""),
                            key=f"phase_remarks_{activity_id}_{phase_id}",
                        )
            if st.button(
                "+ Add implementation phase",
                key=f"add_phase_{activity_id}",
                type="secondary",
            ):
                phases.append(new_implementation_phase(len(phases)))
                st.rerun()

            activity["annex_d_reviewed"] = st.checkbox(
                "Applicable Annex D baseline and limitations reviewed",
                value=activity["annex_d_reviewed"],
                key=f"annex_d_{activity_id}",
            )
            if len(plan["activities"]) > 1 and st.button(
                "Remove activity", key=f"remove_{activity_id}"
            ):
                plan["activities"].remove(activity)
                st.rerun()
    if st.button("+ Add another activity", type="secondary"):
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
    team["organisation"] = st.text_area(
        "Organisation and reporting lines",
        team["organisation"],
        key="team_organisation",
        help=(
            "Describe who reports to whom, who operates RSS equipment, who can "
            "stop the activity and how issues escalate to the QP(S)."
        ),
    )
    org_chart = st.file_uploader(
        "Organisation / reporting-line chart (optional image)",
        type=["png", "jpg", "jpeg"],
        key="org_chart_upload",
        help=(
            "The image will appear with the written description in Section 3 "
            "of the Word plan."
        ),
    )
    if org_chart:
        st.session_state.org_chart_bytes = org_chart.getvalue()
        st.image(
            org_chart,
            caption="RSS team organisation and reporting lines",
            width=560,
        )

    for key, label in (
        ("site_supervisors", "Site supervisors (RE/RTO) *"),
        ("builder_operators", "Builder-side RSS operators"),
        ("backup_personnel", "Backup personnel and handover *"),
        ("training", "Training programme *"),
    ):
        team[key] = st.text_area(label, team.get(key, ""), key=f"team_{key}")

    st.markdown("#### Competency verification")
    st.caption(
        "This should not be only one tick. The checks confirm the required basis, "
        "while the evidence reference, verifier and date make the assessment auditable."
    )
    check_col1, check_col2 = st.columns(2)
    with check_col1:
        team["competency_provider_training"] = st.checkbox(
            "Relevant technology-provider training completed",
            value=team.get("competency_provider_training", False),
            help="For the actual hardware/software used on this project.",
        )
        team["competency_trial"] = st.checkbox(
            "Proficiency demonstrated during an RSS trial",
            value=team.get("competency_trial", False),
            help=(
                "The RE/RTO should demonstrate the complete workflow, not only "
                "attend training."
            ),
        )
    with check_col2:
        team["competency_registration"] = st.checkbox(
            "Professional registration / project appointment checked",
            value=team.get("competency_registration", False),
        )
        team["competency_upgrade_training"] = st.checkbox(
            "Refresher training arranged for software/hardware upgrades",
            value=team.get("competency_upgrade_training", False),
        )
    team["competency_evidence"] = st.text_area(
        "Evidence / record reference *",
        team.get("competency_evidence", team.get("competency", "")),
        key="team_competency_evidence",
        placeholder=(
            "e.g. training certificates TR-001/TR-002; trial record RSS-TRIAL-01; "
            "QP(S) observation checklist dated 15 Aug 2026"
        ),
        help="Reference retained records rather than pasting certificates into the plan.",
    )
    verifier_col, date_col = st.columns(2)
    with verifier_col:
        team["competency_verifier"] = st.text_input(
            "Verified by *",
            team.get("competency_verifier", ""),
            placeholder="Name and designation",
        )
    with date_col:
        team["competency_date"] = st.text_input(
            "Verification date *",
            team.get("competency_date", ""),
            placeholder="YYYY-MM-DD",
        )
    render_profile_manager(
        "people",
        "P1 uses the master personnel and competency information above. Add a "
        "profile only when an activity uses a different deployment team; blank "
        "fields inherit the project default.",
    )


def render_phasing() -> None:
    st.subheader("Define the project-wide progression framework")
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
    st.markdown("#### Activity phase coverage")
    st.caption(
        "The detailed phase entries are maintained within each activity. This "
        "summary shows whether the project framework has been applied."
    )
    rows = []
    for index, activity in enumerate(st.session_state.plan["activities"], start=1):
        rows.append(
            {
                "Activity": activity.get("work_type") or f"Activity {index}",
                "Phases": ", ".join(
                    phase_label(phase)
                    for phase in activity.get("implementation_phases", [])
                )
                or "None",
            }
        )
    if rows:
        st.dataframe(rows, use_container_width=True, hide_index=True)


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
    render_profile_manager(
        "technology",
        "T1 uses the complete project technology chain above. Add another "
        "profile for a different camera, connectivity, measurement or storage "
        "arrangement; blank fields inherit T1.",
    )


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
    render_profile_manager(
        "controls",
        "C1 uses the project-wide procedures above. Alternative profiles should "
        "contain only changed hold points, fallbacks or operating controls.",
    )


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
    render_profile_manager(
        "records",
        "R1 uses the project-wide information-governance rules above. Add a "
        "profile only where an activity needs different evidence handling, "
        "verification or retention.",
    )


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

    st.markdown("#### Activity allocation summary")
    allocation_rows = []
    for index, activity in enumerate(plan["activities"], start=1):
        allocation_rows.append(
            {
                "Activity": activity.get("work_type") or f"Activity {index}",
                "People": profile_label(
                    plan, "people", activity.get("people_profile_id", "P1")
                ),
                "Technology": profile_label(
                    plan,
                    "technology",
                    activity.get("technology_profile_id", "T1"),
                ),
                "Controls": profile_label(
                    plan, "controls", activity.get("control_profile_id", "C1")
                ),
                "Records": profile_label(
                    plan, "records", activity.get("record_profile_id", "R1")
                ),
                "Phases": ", ".join(
                    phase_label(phase)
                    for phase in activity.get("implementation_phases", [])
                ),
            }
        )
    if allocation_rows:
        st.dataframe(allocation_rows, use_container_width=True, hide_index=True)

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
    docx_bytes = build_docx(
        plan,
        site_plan_bytes=st.session_state.site_plan_bytes,
        org_chart_bytes=st.session_state.org_chart_bytes,
    )
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
                st.session_state.plan = migrate_plan(imported)
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
    STEPS[1]: render_people,
    STEPS[2]: render_phasing,
    STEPS[3]: render_technology,
    STEPS[4]: render_controls,
    STEPS[5]: render_records,
    STEPS[6]: render_activities,
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
