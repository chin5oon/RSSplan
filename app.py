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
    "2 - Manpower & organisation",
    "3 - Phasing framework",
    "4 - Technology",
    "5 - Controls",
    "6 - Records",
    "7 - Activities",
    "8 - Review & export",
]

LEGACY_ACTIVITY_GROUPS = {
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

STRUCTURAL_SUPERVISION_ACTIVITIES = LEGACY_ACTIVITY_GROUPS[
    "Structural supervision checklists"
]

SUPERVISION_ACTIVITIES = list(
    dict.fromkeys(
        activity
        for group, activities in LEGACY_ACTIVITY_GROUPS.items()
        if group != "Project-specific"
        for activity in activities
    )
)

APPROACHES = [
    "Technology Replacement",
    "Evidence-Based Supervision",
    "Live Remote Supervision",
    "In-Person Attendance",
]

LEGACY_APPROACHES = {
    "Technology replacement": "Technology Replacement",
    "Evidence-based supervision": "Evidence-Based Supervision",
    "Live remote supervision": "Live Remote Supervision",
    "In-person attendance": "In-Person Attendance",
}

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
        "qp_responsibilities": "",
        "site_supervisor_responsibilities": "",
        "rss_operator_responsibilities": "",
        "safety_oversight_responsibilities": "",
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
        "tech_failure": "Switch to tested backup equipment. If adequate supervision cannot be restored, revert to in-person supervision and record the incident.",
        "poor_connectivity": "Switch to backup connectivity or another suitable supervision approach. Reschedule or revert to in-person supervision if communication quality is inadequate, and document the disruption.",
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
                "poor_connectivity": "",
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
    "use_shared_activity_phasing": False,
    "shared_phasing_source": "default",
    "shared_implementation_phases": [],
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
        "work_type": "",
        "custom_activity_name": "",
        "description": "",
        "location": "",
        "complexity": "Simple",
        "frequency": "Periodic",
        "approach": "Technology Replacement",
        "phasing_source": "default",
        "implementation_phases": [],
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
    if plan.get("shared_phasing_source") not in ("default", "new"):
        plan["shared_phasing_source"] = "default"
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

    for position, phase in enumerate(plan.get("shared_implementation_phases", [])):
        _merge_missing(phase, new_implementation_phase(position))

    for activity in plan.setdefault("activities", []):
        legacy_phase = activity.get("phase", "Phase 1")
        legacy_extent = activity.get("extent", "")
        had_implementation_phases = bool(activity.get("implementation_phases"))
        had_phasing_source = "phasing_source" in activity
        _merge_missing(activity, new_activity())
        if had_implementation_phases and not had_phasing_source:
            activity["phasing_source"] = "new"
        if not had_implementation_phases:
            if (activity.get("phase") or activity.get("extent")) and not had_phasing_source:
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
                activity["phasing_source"] = "new"
        for position, phase in enumerate(activity["implementation_phases"]):
            _merge_missing(phase, new_implementation_phase(position))
        activity["work_type"] = flatten_legacy_activity(activity)
        activity["approach"] = LEGACY_APPROACHES.get(
            activity.get("approach"), activity.get("approach", APPROACHES[0])
        )
        activity.pop("category", None)
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
    profile_kind = {
        "people": "manpower deployment",
        "technology": "technology arrangement",
        "controls": "process and contingency",
        "records": "documentation and records",
    }[kind]
    template["name"] = f"Alternative {profile_kind} profile {number}"
    template["default"] = False
    return template


def flatten_legacy_activity(activity: dict) -> str:
    """Map the earlier group + child selection into the flat activity list."""
    work_type = str(activity.get("work_type", "")).strip()
    category = str(activity.get("category", "")).strip()
    if work_type in SUPERVISION_ACTIVITIES:
        return work_type
    if (
        category == "Structural supervision checklists"
        and work_type in STRUCTURAL_SUPERVISION_ACTIVITIES
    ):
        return work_type
    for legacy_group, legacy_activities in LEGACY_ACTIVITY_GROUPS.items():
        if work_type in legacy_activities:
            if legacy_group != "Project-specific":
                return work_type
    # Older working files may contain a broad group or a project-specific value.
    # Keep the text in the JSON, but require the user to choose a supported list item.
    if work_type and work_type not in SUPERVISION_ACTIVITIES:
        activity["legacy_work_type"] = work_type
    return ""


def activity_display_name(activity: dict) -> str:
    return activity.get("work_type", "").strip()


def recommended_approach(complexity: str, frequency: str) -> str:
    if complexity == "Simple" and frequency == "Periodic":
        return "Technology Replacement"
    if complexity == "Complex" and frequency == "Periodic":
        return "Evidence-Based Supervision"
    if complexity == "Simple" and frequency == "Continuous":
        return "Live Remote Supervision"
    return "In-Person Attendance"


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
    destination = st.session_state.pop("_navigation_destination", None)
    if destination in STEPS:
        st.session_state.current_step = destination
        st.session_state.sidebar_step = destination
    if (
        "sidebar_step" not in st.session_state
        or st.session_state.sidebar_step not in STEPS
    ):
        st.session_state.sidebar_step = st.session_state.current_step
    if "navigation_alert" not in st.session_state:
        st.session_state.navigation_alert = []
    if "pending_step" not in st.session_state:
        st.session_state.pending_step = None
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
        ("builder_operators", "RSS operators (workers from Builder side) assigned"),
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
        ("poor_connectivity", "Connectivity / communication variation"),
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

PROFILE_TITLES = {
    "people": "manpower deployment",
    "technology": "technology arrangement",
    "controls": "process and contingency",
    "records": "documentation and records",
}

PROFILE_EXAMPLES = {
    "people": "a different RE/RTO and RSS operator team for precast-yard work",
    "technology": "a different camera, connectivity or measurement setup",
    "controls": "different preparation steps, hold points or contingency procedures",
    "records": "different evidence handling, verification or retention rules",
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


def guidebook_default_phases() -> list[dict]:
    """Return the recommended Phase 1-3 progression from the RSS Guidebook."""
    return [new_implementation_phase(position) for position in range(3)]


def resolve_implementation_phases(
    plan: dict, activity_index: int, seen: set[str] | None = None
) -> list[dict]:
    """Resolve shared, default, new or preceding-activity phasing without copying."""
    if plan.get("use_shared_activity_phasing", False):
        if plan.get("shared_phasing_source") == "new":
            return plan.get("shared_implementation_phases", [])
        return guidebook_default_phases()

    activities = plan.get("activities", [])
    if activity_index < 0 or activity_index >= len(activities):
        return []
    activity = activities[activity_index]
    source = activity.get("phasing_source", "default")
    if source == "default":
        return guidebook_default_phases()
    if source == "new":
        return activity.get("implementation_phases", [])
    if source.startswith("activity:"):
        source_uid = source.split(":", 1)[1]
        seen = set() if seen is None else set(seen)
        if source_uid in seen:
            return []
        seen.add(source_uid)
        for preceding_index, preceding in enumerate(activities[:activity_index]):
            if preceding.get("uid") == source_uid:
                return resolve_implementation_phases(plan, preceding_index, seen)
    return []


def implementation_phasing_source_label(plan: dict, activity_index: int) -> str:
    if plan.get("use_shared_activity_phasing", False):
        return (
            "Shared new implementation phasing"
            if plan.get("shared_phasing_source") == "new"
            else "Shared Guidebook default"
        )
    activity = plan["activities"][activity_index]
    source = activity.get("phasing_source", "default")
    if source == "default":
        return "Default - Guidebook recommended Phases 1-3"
    if source == "new":
        return "New - activity-specific implementation phasing"
    if source.startswith("activity:"):
        source_uid = source.split(":", 1)[1]
        for index, preceding in enumerate(plan["activities"][:activity_index]):
            if preceding.get("uid") == source_uid:
                name = activity_display_name(preceding) or "Unnamed activity"
                return f"Activity {index + 1:02d} - {name}"
    return "Invalid phasing source"


def render_implementation_phase_summary(phases: list[dict]) -> None:
    """Show the full effective phasing when no editable fields are displayed."""
    if not phases:
        st.warning("No implementation phases are currently defined.")
        return
    st.markdown("###### What this implementation phasing contains")
    rows = [
        {
            "Phase": phase_label(phase),
            "Activity progress range": phase.get("progress_range", "")
            or "Not specified",
            "Extent of RSS": phase.get("rss_extent", "") or "Not specified",
            "Parallel / in-person verification": phase.get(
                "parallel_supervision", ""
            )
            or "Not specified",
            "Acceptance / progression criteria": phase.get(
                "acceptance_criteria", ""
            )
            or "Not specified",
            "QP(S) review point": phase.get("review_point", "")
            or "Not specified",
        }
        for phase in phases
    ]
    st.dataframe(rows, use_container_width=True, hide_index=True)


def materialize_activity_phasing(plan: dict) -> dict:
    """Create a report-ready copy with every activity's effective phases expanded."""
    materialized = json.loads(json.dumps(plan))
    for index, activity in enumerate(materialized.get("activities", [])):
        activity["implementation_phases"] = json.loads(
            json.dumps(resolve_implementation_phases(plan, index))
        )
    return materialized


def render_implementation_phase_editor(phases: list[dict], owner_id: str) -> None:
    """Render one editable implementation-phasing definition."""
    if not phases:
        phases.append(new_implementation_phase())
    for phase_index, phase in enumerate(list(phases)):
        phase_id = phase["uid"]
        with st.container(border=True):
            phase_title, remove_column = st.columns([5, 1])
            with phase_title:
                st.markdown(
                    f"**Phase entry {phase_index + 1}: {phase_label(phase)}**"
                )
            with remove_column:
                if len(phases) > 1 and st.button(
                    "Remove",
                    key=f"remove_phase_{owner_id}_{phase_id}",
                    use_container_width=True,
                ):
                    phases.remove(phase)
                    st.rerun()
            row1, row2, row3 = st.columns(3)
            with row1:
                phase["name"] = st.selectbox(
                    "Phase",
                    list(PHASE_PRESETS),
                    index=(
                        list(PHASE_PRESETS).index(phase["name"])
                        if phase["name"] in PHASE_PRESETS
                        else 0
                    ),
                    key=f"phase_name_{owner_id}_{phase_id}",
                )
                if phase["name"] == "Custom":
                    phase["custom_name"] = st.text_input(
                        "Custom phase name *",
                        phase.get("custom_name", ""),
                        key=f"phase_custom_name_{owner_id}_{phase_id}",
                    )
                if st.button(
                    "Apply selected phase guide defaults",
                    key=f"phase_defaults_{owner_id}_{phase_id}",
                    help=(
                        "Replaces the progress range, RSS extent and "
                        "parallel-supervision entry for this phase."
                    ),
                ):
                    preset = PHASE_PRESETS[phase["name"]]
                    phase["progress_range"] = preset["progress_range"]
                    phase["rss_extent"] = preset["rss_extent"]
                    phase["parallel_supervision"] = preset["parallel_supervision"]
                    for widget_key in (
                        f"phase_progress_{owner_id}_{phase_id}",
                        f"phase_extent_{owner_id}_{phase_id}",
                        f"phase_parallel_{owner_id}_{phase_id}",
                    ):
                        st.session_state.pop(widget_key, None)
                    st.rerun()
            with row2:
                phase["progress_range"] = st.text_input(
                    "Activity progress range *",
                    phase.get("progress_range", ""),
                    key=f"phase_progress_{owner_id}_{phase_id}",
                )
            with row3:
                phase["rss_extent"] = st.text_input(
                    "Extent of RSS *",
                    phase.get("rss_extent", ""),
                    key=f"phase_extent_{owner_id}_{phase_id}",
                )
            row4, row5 = st.columns(2)
            with row4:
                phase["parallel_supervision"] = st.text_area(
                    "Parallel / in-person verification",
                    phase.get("parallel_supervision", ""),
                    key=f"phase_parallel_{owner_id}_{phase_id}",
                )
                phase["acceptance_criteria"] = st.text_area(
                    "Acceptance / progression criteria *",
                    phase.get("acceptance_criteria", ""),
                    key=f"phase_acceptance_{owner_id}_{phase_id}",
                )
            with row5:
                phase["review_point"] = st.text_area(
                    "QP(S) review point *",
                    phase.get("review_point", ""),
                    key=f"phase_review_{owner_id}_{phase_id}",
                )
                phase["remarks"] = st.text_area(
                    "Phase remarks",
                    phase.get("remarks", ""),
                    key=f"phase_remarks_{owner_id}_{phase_id}",
                )
    if st.button(
        "+ Add implementation phase",
        key=f"add_phase_{owner_id}",
        type="secondary",
    ):
        phases.append(new_implementation_phase(len(phases)))
        st.rerun()


def render_profile_manager(kind: str, description: str) -> None:
    plan = st.session_state.plan
    profiles = plan["profiles"][kind]
    title = PROFILE_TITLES[kind]
    st.markdown(f"#### Reusable {title} profiles")
    st.info(
        "**What is a reusable profile?** It is a saved set of project arrangements "
        "that can be assigned to one or more activities in Step 7. Enter the common "
        "arrangement once, then create another profile only when an activity needs "
        f"{PROFILE_EXAMPLES[kind]}. This avoids repeating the same information."
    )
    st.caption(description)
    for profile in list(profiles):
        profile_id = profile["id"]
        if profile.get("default"):
            st.info(
                f"**{profile_id} - {profile['name']}** is the default profile. "
                "Activities use the project-wide information entered above unless "
                "you assign them a different profile in Step 7."
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
        f"+ Add {title} profile",
        key=f"add_profile_{kind}",
        type="secondary",
    ):
        profiles.append(new_profile(kind, plan))
        st.rerun()


def issues_for_step(plan: dict, step_index: int) -> list[str]:
    return [message for issue_step, message in validate(plan) if issue_step == step_index]


def request_navigation(target_step: str, *, check_required: bool = True) -> bool:
    """Queue navigation or show an explicit incomplete-required-fields prompt."""
    current_step = st.session_state.current_step
    current_index = STEPS.index(current_step)
    target_index = STEPS.index(target_step)
    issues = issues_for_step(st.session_state.plan, current_index)
    if check_required and target_index > current_index and issues:
        st.session_state.pending_step = target_step
        st.session_state.navigation_alert = issues
        return False
    st.session_state.pending_step = None
    st.session_state.navigation_alert = []
    st.session_state._navigation_destination = target_step
    return True


def handle_sidebar_navigation() -> None:
    requested = st.session_state.sidebar_step
    current = st.session_state.current_step
    if requested == current:
        return
    if STEPS.index(requested) > STEPS.index(current):
        issues = issues_for_step(st.session_state.plan, STEPS.index(current))
        if issues:
            st.session_state.pending_step = requested
            st.session_state.navigation_alert = issues
            st.session_state.sidebar_step = current
            return
    st.session_state.current_step = requested
    st.session_state.pending_step = None
    st.session_state.navigation_alert = []


def render_navigation_alert() -> None:
    issues = st.session_state.get("navigation_alert", [])
    target = st.session_state.get("pending_step")
    if not issues or target not in STEPS:
        return
    st.warning(
        f"Complete the compulsory fields marked * before moving to **{target}**, "
        "or explicitly continue with an incomplete draft."
    )
    for issue in issues:
        st.write(f"- {issue}")
    stay, proceed = st.columns(2)
    with stay:
        if st.button("Stay and complete the fields", use_container_width=True):
            st.session_state.pending_step = None
            st.session_state.navigation_alert = []
            st.rerun()
    with proceed:
        if st.button(
            "Continue with incomplete fields",
            type="secondary",
            use_container_width=True,
        ):
            request_navigation(target, check_required=False)
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
        if activity.get("work_type", "").strip() not in SUPERVISION_ACTIVITIES:
            issues.append(
                (
                    activity_step,
                    f"Select a listed supervision activity for activity {index}.",
                )
            )
        if not activity["description"].strip() or not activity["location"].strip():
            issues.append(
                (activity_step, f"Complete scope and location for activity {index}.")
            )
        if not activity["evidence"].strip():
            issues.append(
                (activity_step, f"Complete evidence requirements for activity {index}.")
            )
        phases = resolve_implementation_phases(plan, index - 1)
        if not phases:
            issues.append(
                (
                    activity_step,
                    f"Select or add valid implementation phasing for activity {index}.",
                )
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
                profile_name = PROFILE_TITLES[kind]
                issues.append(
                    (
                        activity_step,
                        f"Select a valid {profile_name} profile for activity {index}.",
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
    if not team.get("organisation", "").strip():
        issues.append((1, "Describe the manpower hierarchy and reporting lines."))
    if any(
        not team.get(key, "").strip()
        for key in (
            "site_supervisors",
            "builder_operators",
            "backup_personnel",
        )
    ):
        issues.append((1, "Complete the manpower and backup assignments."))
    if any(
        not team.get(key, "").strip()
        for key in (
            "qp_responsibilities",
            "site_supervisor_responsibilities",
            "rss_operator_responsibilities",
        )
    ):
        issues.append((1, "Complete the required roles and responsibilities."))
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
    process = plan["process"]
    if any(
        not process.get(key, "").strip()
        for key in (
            "stop_work",
            "tech_failure",
            "poor_connectivity",
            "safety_incident",
            "non_conformity",
        )
    ):
        issues.append((4, "Complete all compulsory contingency procedures."))
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
    plan = st.session_state.plan
    st.markdown("#### Implementation phasing setup")
    plan["use_shared_activity_phasing"] = st.checkbox(
        "Use the same implementation phasing for all activities",
        value=plan.get("use_shared_activity_phasing", False),
        key="use_shared_activity_phasing",
        help=(
            "When selected, define the implementation phasing once here. Individual "
            "activities will inherit it and will not show separate phasing fields."
        ),
    )
    if plan["use_shared_activity_phasing"]:
        shared_options = ["default", "new"]
        current_shared_source = plan.get("shared_phasing_source", "default")
        if current_shared_source not in shared_options:
            current_shared_source = "default"
        plan["shared_phasing_source"] = st.selectbox(
            "Implementation phasing for all activities",
            shared_options,
            index=shared_options.index(current_shared_source),
            format_func=lambda value: (
                "Default - Guidebook recommended Phases 1-3"
                if value == "default"
                else "New - enter one shared implementation phasing"
            ),
            key="shared_phasing_source_selector",
        )
        if plan["shared_phasing_source"] == "new":
            st.caption(
                "Enter the shared implementation phasing once. It will apply to every "
                "activity in this plan."
            )
            render_implementation_phase_editor(
                plan["shared_implementation_phases"], "shared"
            )
        else:
            st.info(
                "All activities use the Guidebook default: Phase 1 (first 30%, maximum "
                "15% RSS), Phase 2 (30-75%, maximum 30% RSS), and Phase 3 (75-100%, "
                "maximum 50% RSS). No additional phasing fields are required."
            )
            render_implementation_phase_summary(guidebook_default_phases())
    st.warning(
        "The matrix is a starting point. Upgrade to in-person attendance whenever "
        "workmanship, site conditions, visibility or intervention capability is inadequate."
    )
    st.caption(
        "Select one supervision activity from the Guidebook-based list, assign "
        "reusable project arrangements, and record only activity-specific variations. "
        "Each activity can have several implementation phases."
    )
    for index, activity in enumerate(plan["activities"]):
        activity_id = activity["uid"]
        selected_name = activity_display_name(activity) or "Select an activity"
        with st.expander(
            f"Activity {index + 1:02d} - {selected_name}", expanded=True
        ):
            col1, col2 = st.columns(2)
            with col1:
                current_work_type = activity.get("work_type", "")
                work_type_index = (
                    SUPERVISION_ACTIVITIES.index(current_work_type)
                    if current_work_type in SUPERVISION_ACTIVITIES
                    else None
                )
                activity["work_type"] = (
                    st.selectbox(
                        "Supervision activity *",
                        SUPERVISION_ACTIVITIES,
                        index=work_type_index,
                        placeholder="Select the activity",
                        key=f"work_type_{activity_id}",
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
                    "Complexity of supervision",
                    ["Simple", "Complex"],
                    index=["Simple", "Complex"].index(activity["complexity"]),
                    key=f"complexity_{activity_id}",
                    format_func=lambda value: f"{value} inspection",
                )
                activity["frequency"] = st.selectbox(
                    "Frequency of supervision",
                    ["Periodic", "Continuous"],
                    index=["Periodic", "Continuous"].index(activity["frequency"]),
                    key=f"frequency_{activity_id}",
                    format_func=lambda value: f"{value} inspection",
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

            st.markdown("##### Reusable arrangement assignments")
            st.caption(
                "Each selection points to information entered once in the relevant "
                "earlier step. Choose an alternative profile only when this activity "
                "uses a different arrangement."
            )
            assignment_columns = st.columns(4)
            assignment_labels = {
                "people": "Manpower deployment",
                "technology": "Technology arrangement",
                "controls": "Process & contingency",
                "records": "Documentation & records",
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
                    "Manpower assignment / competency variation",
                    activity.get("personnel_requirements", ""),
                    key=f"personnel_requirements_{activity_id}",
                    placeholder=(
                        "Optional - leave blank when the selected manpower deployment "
                        "profile applies."
                    ),
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
            if plan["use_shared_activity_phasing"]:
                effective_phases = resolve_implementation_phases(plan, index)
                st.info(
                    f"Uses **{implementation_phasing_source_label(plan, index)}**. "
                    "The implementation phasing is defined once at the top of this page; "
                    "no activity-specific phasing fields are required."
                )
            else:
                preceding_sources = [
                    f"activity:{preceding.get('uid')}"
                    for preceding in plan["activities"][:index]
                ]
                source_options = ["default", *preceding_sources, "new"]
                current_source = activity.get("phasing_source", "default")
                if current_source not in source_options:
                    current_source = "default"

                def format_phasing_source(value: str) -> str:
                    if value == "default":
                        return "Default - Guidebook recommended Phases 1-3"
                    if value == "new":
                        return "New - enter a different implementation phasing"
                    source_uid = value.split(":", 1)[1]
                    for preceding_index, preceding in enumerate(
                        plan["activities"][:index]
                    ):
                        if preceding.get("uid") == source_uid:
                            name = activity_display_name(preceding) or "Unnamed activity"
                            return f"Activity {preceding_index + 1:02d} - {name}"
                    return "Unavailable preceding activity"

                activity["phasing_source"] = st.selectbox(
                    "Implementation phasing source",
                    source_options,
                    index=source_options.index(current_source),
                    format_func=format_phasing_source,
                    key=f"phasing_source_{activity_id}",
                    help=(
                        "Use the Guidebook default, link to any preceding activity's "
                        "implementation phasing, or choose New to enter different details."
                    ),
                )
                if activity["phasing_source"] == "new":
                    st.caption(
                        "Only this activity uses the implementation phasing entered below."
                    )
                    render_implementation_phase_editor(
                        activity["implementation_phases"], activity_id
                    )
                    effective_phases = activity["implementation_phases"]
                else:
                    effective_phases = resolve_implementation_phases(plan, index)
                    st.info(
                        f"Uses **{implementation_phasing_source_label(plan, index)}**. "
                        "No additional phasing fields are required for this activity."
                    )
                    render_implementation_phase_summary(effective_phases)

            activity["annex_d_reviewed"] = st.checkbox(
                "Applicable Annex D baseline and limitations reviewed *",
                value=activity["annex_d_reviewed"],
                key=f"annex_d_{activity_id}",
            )
            if len(plan["activities"]) > 1 and st.button(
                "Remove activity", key=f"remove_{activity_id}"
            ):
                removed_source = f"activity:{activity_id}"
                plan["activities"].remove(activity)
                for remaining_activity in plan["activities"]:
                    if remaining_activity.get("phasing_source") == removed_source:
                        remaining_activity["phasing_source"] = "default"
                st.rerun()
    if st.button("+ Add another activity", type="secondary"):
        plan["activities"].append(new_activity())
        st.rerun()


def render_people() -> None:
    st.subheader("Define the manpower organisation, roles and competency")
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

    st.markdown("#### Manpower organisation chart")
    st.caption(
        "Include the QP(S), Site Supervisors (RE/RTO), RSS operators (workers from "
        "Builder side), backup personnel and the reporting/escalation path."
    )
    team["organisation"] = st.text_area(
        "Manpower hierarchy and reporting lines *",
        team["organisation"],
        key="team_organisation",
        help=(
            "Describe who reports to whom, who operates RSS equipment, who can "
            "stop the activity and how issues escalate to the QP(S)."
        ),
    )
    org_chart = st.file_uploader(
        "Manpower organisation chart (optional image)",
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
            caption="RSS manpower organisation chart",
            width=560,
        )

    st.markdown("#### Manpower assignments")
    for key, label in (
        ("site_supervisors", "Site supervisors (RE/RTO) *"),
        ("builder_operators", "RSS operators (workers from Builder side) *"),
        ("backup_personnel", "Backup personnel and handover *"),
    ):
        team[key] = st.text_area(label, team.get(key, ""), key=f"team_{key}")

    st.markdown("#### Roles and responsibilities")
    responsibility_fields = (
        (
            "qp_responsibilities",
            "Qualified Person (Supervision) roles and responsibilities *",
            "e.g. overall RSS responsibility, approval, effectiveness review, final sign-off and liaison with authorities",
        ),
        (
            "site_supervisor_responsibilities",
            "Site Supervisor (RE/RTO) roles and responsibilities *",
            "e.g. conduct supervision, direct the RSS operator, intervene, assess conformity and complete records",
        ),
        (
            "rss_operator_responsibilities",
            "RSS operator roles and responsibilities *",
            "e.g. operate assigned equipment, follow instructions, identify elements and capture required evidence",
        ),
        (
            "safety_oversight_responsibilities",
            "Safety oversight and escalation responsibilities",
            "e.g. stop-work authority, emergency response and escalation path to the QP(S)",
        ),
    )
    for start in range(0, len(responsibility_fields), 2):
        columns = st.columns(2)
        for column, (key, label, placeholder) in zip(
            columns, responsibility_fields[start : start + 2]
        ):
            with column:
                team[key] = st.text_area(
                    label,
                    team.get(key, ""),
                    key=f"team_{key}",
                    placeholder=placeholder,
                )

    st.markdown("#### Training and competency requirements")
    team["training"] = st.text_area(
        "Site supervision team training programme *",
        team.get("training", ""),
        key="team_training",
        placeholder=(
            "Cover digital tools and platforms, RSS procedures, safety, assigned "
            "equipment, and documentation/evidence collection."
        ),
    )
    st.markdown("#### Competency verification")
    st.caption(
        "This should not be only one tick. The checks confirm the required basis, "
        "while the evidence reference, verifier and date make the assessment auditable."
    )
    check_col1, check_col2 = st.columns(2)
    with check_col1:
        team["competency_provider_training"] = st.checkbox(
            "Relevant technology-provider training completed *",
            value=team.get("competency_provider_training", False),
            help="For the actual hardware/software used on this project.",
        )
        team["competency_trial"] = st.checkbox(
            "Proficiency demonstrated during an RSS trial *",
            value=team.get("competency_trial", False),
            help=(
                "The RE/RTO should demonstrate the complete workflow, not only "
                "attend training."
            ),
        )
    with check_col2:
        team["competency_registration"] = st.checkbox(
            "Professional registration / project appointment checked *",
            value=team.get("competency_registration", False),
        )
        team["competency_upgrade_training"] = st.checkbox(
            "Refresher training arranged for software/hardware upgrades *",
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
        "P1 uses the project-wide manpower and competency information above. Add a "
        "profile only when an activity uses a different deployment team; blank fields "
        "inherit P1.",
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
    st.markdown("#### Activity phase coverage - automatic summary")
    st.info(
        "**This table is not filled in here.** It is populated automatically from "
        "the implementation phasing selections in Step 7, whether they use the shared "
        "setup, the Guidebook default, a preceding activity, or a new definition."
    )
    rows = []
    for index, activity in enumerate(st.session_state.plan["activities"], start=1):
        effective_phases = resolve_implementation_phases(
            st.session_state.plan, index - 1
        )
        rows.append(
            {
                "Activity": activity_display_name(activity) or f"Activity {index}",
                "Phases": ", ".join(
                    phase_label(phase)
                    for phase in effective_phases
                )
                or "Not yet configured in Step 7",
                "Source": implementation_phasing_source_label(
                    st.session_state.plan, index - 1
                ),
            }
        )
    if rows:
        st.dataframe(rows, use_container_width=True, hide_index=True)
    if st.button("Open activity implementation phases", type="secondary"):
        request_navigation(STEPS[6])
        st.rerun()


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
    st.subheader("Define RSS procedures and contingency planning")
    st.caption("PREPARE → VERIFY LIVE → DECIDE → RECORD → CLOSE")
    for key, label in (
        ("before", "Before remote supervision"),
        ("during", "During remote supervision"),
        ("after", "After remote supervision"),
        ("communication", "Communication protocol"),
    ):
        text_area("process", key, label, 120)
    st.markdown("#### Quality assurance and risk management")
    st.caption(
        "Contingency planning follows Section 6.7.1 of the RSS Guidebook. State how "
        "the team will respond, escalate, document the event and revert to in-person "
        "supervision when remote supervision is no longer effective."
    )
    for key, label in (
        ("stop_work", "Conditions for stopping RSS / reverting to in-person supervision *"),
        ("tech_failure", "Technology failure and equipment malfunction *"),
        ("poor_connectivity", "Poor connectivity or communication breakdown *"),
        ("poor_evidence", "Poor or incomplete evidence"),
        ("safety_incident", "Safety incidents during remote supervision *"),
        ("non_conformity", "Non-conformity detection and escalation procedures *"),
    ):
        text_area("process", key, label)
    render_profile_manager(
        "controls",
        "C1 uses the project-wide procedures above. Alternative profiles should "
        "contain only changed hold points, contingency procedures or operating controls.",
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
        effective_phases = resolve_implementation_phases(plan, index - 1)
        allocation_rows.append(
            {
                "Activity": activity_display_name(activity) or f"Activity {index}",
                "Manpower": profile_label(
                    plan, "people", activity.get("people_profile_id", "P1")
                ),
                "Technology": profile_label(
                    plan,
                    "technology",
                    activity.get("technology_profile_id", "T1"),
                ),
                "Process / contingency": profile_label(
                    plan, "controls", activity.get("control_profile_id", "C1")
                ),
                "Records": profile_label(
                    plan, "records", activity.get("record_profile_id", "R1")
                ),
                "Phases": ", ".join(
                    phase_label(phase)
                    for phase in effective_phases
                ),
                "Phasing source": implementation_phasing_source_label(
                    plan, index - 1
                ),
            }
        )
    if allocation_rows:
        st.dataframe(allocation_rows, use_container_width=True, hide_index=True)

    st.markdown("#### QP(S) declarations")
    sign = plan["signoff"]
    sign["confirm_professional_judgement"] = st.checkbox(
        "I applied professional judgement to the suitability of RSS for each activity. *",
        sign["confirm_professional_judgement"],
    )
    sign["confirm_annex_d"] = st.checkbox(
        "I reviewed the applicable Annex D minimum requirements and documented alternatives. *",
        sign["confirm_annex_d"],
    )
    sign["confirm_fallback"] = st.checkbox(
        "I will require in-person supervision whenever the intended outcome cannot be achieved remotely. *",
        sign["confirm_fallback"],
    )
    sign["confirm_records"] = st.checkbox(
        "RSS records will be controlled, protected, retrievable and retained as stated. *",
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
        materialize_activity_phasing(plan),
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
        st.radio(
            "PLAN PROGRESS",
            STEPS,
            label_visibility="visible",
            key="sidebar_step",
            on_change=handle_sidebar_navigation,
        )
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
    return st.session_state.current_step


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
  professional judgement and contingency controls visible at every step.</p>
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

render_navigation_alert()

st.divider()
index = STEPS.index(selected_step)
left, _, right = st.columns([1, 4, 1])
with left:
    if index > 0 and st.button("← Back", use_container_width=True):
        request_navigation(STEPS[index - 1], check_required=False)
        st.rerun()
with right:
    if index < len(STEPS) - 1 and st.button(
        "Continue →", type="primary", use_container_width=True
    ):
        request_navigation(STEPS[index + 1])
        st.rerun()

st.caption(
    "This tool supports preparation; it does not replace the QP(S)'s statutory "
    "duties, professional judgement or review of current BCA requirements."
)
