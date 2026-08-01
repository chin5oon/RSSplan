"""Native DOCX report generation for the RSS Plan Builder."""

from __future__ import annotations

from io import BytesIO
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


INK = "112536"
TEAL = "087E8B"
PALE = "E8F6F4"
LINE = "C7D4D9"
MUTED = "667985"


def _text(value: Any) -> str:
    value = "" if value is None else str(value).strip()
    return value or "Not specified"


def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def _set_cell_margins(cell, top=120, start=140, bottom=120, end=140) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (
        ("top", top),
        ("start", start),
        ("bottom", bottom),
        ("end", end),
    ):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def _add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Page ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string(MUTED)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def _configure_document(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.12

    for name, size, color, before, after in (
        ("Title", 30, INK, 0, 12),
        ("Heading 1", 18, INK, 18, 8),
        ("Heading 2", 13, TEAL, 12, 5),
        ("Heading 3", 11, INK, 9, 4),
    ):
        style = styles[name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = name != "Title"
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    header = section.header.paragraphs[0]
    header.text = "REMOTE SITE SUPERVISION PLAN"
    header.style = styles["Normal"]
    header.runs[0].font.size = Pt(8)
    header.runs[0].font.bold = True
    header.runs[0].font.color.rgb = RGBColor.from_string(TEAL)
    header.paragraph_format.space_after = Pt(0)

    _add_page_number(section.footer.paragraphs[0])


def _add_kv_table(doc: Document, rows: list[tuple[str, Any]]) -> None:
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(4.6)
    table.columns[1].width = Cm(12.4)
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].width = Cm(4.6)
        cells[1].width = Cm(12.4)
        cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _set_cell_shading(cells[0], PALE)
        for cell in cells:
            _set_cell_margins(cell)
        label_p = cells[0].paragraphs[0]
        label_p.paragraph_format.space_after = Pt(0)
        label_run = label_p.add_run(label.upper())
        label_run.bold = True
        label_run.font.size = Pt(8)
        label_run.font.color.rgb = RGBColor.from_string(TEAL)
        value_p = cells[1].paragraphs[0]
        value_p.paragraph_format.space_after = Pt(0)
        value_p.add_run(_text(value))
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def _add_callout(doc: Document, title: str, body: str) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    _set_cell_shading(cell, PALE)
    _set_cell_margins(cell, top=170, start=220, bottom=170, end=220)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(f"{title} ")
    run.bold = True
    run.font.color.rgb = RGBColor.from_string(TEAL)
    p.add_run(body)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def _add_activity(doc: Document, activity: dict[str, Any], index: int) -> None:
    heading = doc.add_heading(
        f"4.{index} {_text(activity.get('work_type'))}", level=2
    )
    heading.paragraph_format.keep_with_next = True
    _add_kv_table(
        doc,
        [
            ("Location / element IDs", activity.get("location")),
            ("Scope of supervision", activity.get("description")),
            (
                "Assessment",
                f"{_text(activity.get('complexity'))} inspection / "
                f"{_text(activity.get('frequency'))} supervision",
            ),
            ("Selected approach", activity.get("approach")),
            ("Implementation phase", activity.get("phase")),
            ("Extent of RSS", activity.get("extent")),
            ("Evidence requirements", activity.get("evidence")),
            ("Equipment / software", activity.get("equipment")),
            ("Professional justification", activity.get("deviation")),
            (
                "Annex D review",
                "Confirmed"
                if activity.get("annex_d_reviewed")
                else "Not yet confirmed",
            ),
        ],
    )


def build_docx(data: dict[str, Any], site_plan_bytes: bytes | None = None) -> bytes:
    """Return a polished native DOCX containing the completed RSS plan."""
    doc = Document()
    _configure_document(doc)
    project = data.get("project", {})
    team = data.get("team", {})
    phases = data.get("phases", {})
    technology = data.get("technology", {})
    process = data.get("process", {})
    records = data.get("records", {})
    signoff = data.get("signoff", {})

    cover = doc.add_paragraph()
    cover.paragraph_format.space_before = Pt(34)
    eyebrow = cover.add_run("REMOTE SITE SUPERVISION")
    eyebrow.bold = True
    eyebrow.font.size = Pt(9)
    eyebrow.font.color.rgb = RGBColor.from_string(TEAL)
    eyebrow.font.letter_spacing = Pt(1.5)

    title = doc.add_paragraph(style="Title")
    title.add_run("Remote Site\nSupervision Plan")
    title.paragraph_format.space_after = Pt(28)

    _add_kv_table(
        doc,
        [
            ("Project reference", project.get("reference")),
            ("Project description", project.get("description")),
            (
                "Location",
                f"{_text(project.get('site_type'))}: {_text(project.get('address'))}",
            ),
            (
                "Prepared by",
                f"{_text(team.get('qp_name'))}, PE Reg. No. "
                f"{_text(team.get('pe_number'))}",
            ),
            ("Company", team.get("company")),
            ("Date", team.get("prepared_date")),
        ],
    )
    note = doc.add_paragraph()
    note.paragraph_format.space_before = Pt(28)
    note_run = note.add_run(
        "Prepared with reference to the Guidebook for Remote Site Supervision, "
        "Version 2.0 (June 2026), and the Guide Book for Site Supervision Plan, "
        "Version 1.1 (October 2023). The Building Control Act and Regulations "
        "prevail. Verify current BCA requirements before submission."
    )
    note_run.italic = True
    note_run.font.size = Pt(8)
    note_run.font.color.rgb = RGBColor.from_string(MUTED)
    doc.add_page_break()

    doc.add_heading("1. Project background", level=1)
    _add_kv_table(
        doc,
        [
            ("Project description", project.get("description")),
            ("Site location", project.get("address")),
            ("Structural system", project.get("structural_system")),
            ("Foundation system", project.get("foundation_system")),
            ("RSS challenges", project.get("challenges")),
            ("Permit date", project.get("permit_date")),
        ],
    )
    if site_plan_bytes:
        doc.add_heading("Overall site / location plan", level=2)
        try:
            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run().add_picture(BytesIO(site_plan_bytes), width=Cm(15.5))
            caption = doc.add_paragraph("Figure 1: Overall site plan or location plan")
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption.runs[0].italic = True
            caption.runs[0].font.size = Pt(8)
            caption.runs[0].font.color.rgb = RGBColor.from_string(MUTED)
        except Exception:
            _add_callout(
                doc,
                "Image note.",
                "The uploaded site plan could not be embedded. Insert it before submission.",
            )
    else:
        _add_callout(
            doc,
            "Required before submission.",
            "Insert the overall site plan or location plan.",
        )

    doc.add_heading("2. Implementation plan for remote supervision", level=1)
    _add_callout(
        doc,
        "Implementation basis.",
        "RSS complements in-person supervision. Suitability is assessed activity "
        "by activity using risk, complexity, frequency, technology capability "
        "and site conditions.",
    )
    _add_kv_table(
        doc,
        [
            ("Phase 1", phases.get("phase_1")),
            ("Phase 2", phases.get("phase_2")),
            ("Phase 3", phases.get("phase_3")),
            ("Beyond Phase 3", phases.get("beyond")),
            ("Acceptance criteria", phases.get("criteria")),
            ("Parallel supervision", phases.get("parallel_plan")),
            ("Review cadence", phases.get("review_cadence")),
        ],
    )

    doc.add_heading("3. Personnel and organisational structure", level=1)
    _add_kv_table(
        doc,
        [
            ("QP(S)", f"{_text(team.get('qp_name'))} / {_text(team.get('pe_number'))}"),
            ("Organisation / reporting lines", team.get("organisation")),
            ("Site supervisors", team.get("site_supervisors")),
            ("Builder-side RSS operators", team.get("builder_operators")),
            ("Backup personnel", team.get("backup_personnel")),
            ("Training programme", team.get("training")),
            ("Competency verification", team.get("competency")),
        ],
    )

    doc.add_heading(
        "4. Structural activity classification, supervision and evidence", level=1
    )
    activities = data.get("activities") or []
    if not activities:
        _add_callout(doc, "Required.", "Add at least one structural activity.")
    for index, activity in enumerate(activities, start=1):
        _add_activity(doc, activity, index)

    doc.add_heading("5. Devices for remote supervision", level=1)
    _add_kv_table(
        doc,
        [
            ("Live streaming devices", technology.get("live_devices")),
            ("Evidence / measurement devices", technology.get("evidence_devices")),
            ("Two-way audio", technology.get("audio")),
            ("Power / battery backup", technology.get("power_backup")),
            ("Equipment register / calibration", technology.get("equipment_register")),
        ],
    )

    doc.add_heading("6. Infrastructure requirements", level=1)
    _add_kv_table(
        doc,
        [
            ("Primary connectivity", technology.get("connectivity")),
            ("Backup connectivity", technology.get("backup_connectivity")),
            ("RSS platform", technology.get("platform")),
            ("Video / recording standard", technology.get("video_standard")),
            ("Storage", technology.get("storage")),
        ],
    )

    doc.add_heading("7. Process for conducting remote supervision", level=1)
    for title, key in (
        ("Before remote supervision", "before"),
        ("During remote supervision", "during"),
        ("After remote supervision", "after"),
        ("Communication protocol", "communication"),
    ):
        doc.add_heading(title, level=2)
        doc.add_paragraph(_text(process.get(key)))

    doc.add_heading("8. Quality assurance and contingency procedures", level=1)
    _add_kv_table(
        doc,
        [
            ("Stop-work / in-person trigger", process.get("stop_work")),
            ("Technology failure", process.get("tech_failure")),
            ("Poor or incomplete evidence", process.get("poor_evidence")),
            ("Safety incident", process.get("safety_incident")),
            ("Non-conformity", process.get("non_conformity")),
            ("Verification", records.get("verification")),
            ("Audits / management review", records.get("audits")),
            ("Performance monitoring", records.get("performance")),
        ],
    )

    doc.add_heading("9. Documentation and records management", level=1)
    _add_kv_table(
        doc,
        [
            ("Naming / indexing", records.get("naming")),
            ("Access and security", records.get("access")),
            ("Backup and recovery", records.get("backups")),
            ("Retention", records.get("retention")),
            ("Traceability", records.get("traceability")),
        ],
    )

    doc.add_heading("QP(S) declaration", level=1)
    declaration = (
        "I confirm that this Remote Site Supervision Plan has been prepared "
        "using professional judgement for the project and activities described; "
        "applicable minimum requirements have been reviewed; clear fallback to "
        "in-person supervision is provided where effective remote supervision "
        "cannot be achieved; and records will be controlled and retained in "
        "accordance with this plan and prevailing requirements."
    )
    _add_callout(doc, "Declaration.", declaration)
    doc.add_paragraph()
    sign_table = doc.add_table(rows=1, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_table.autofit = False
    sign_table.columns[0].width = Cm(10.5)
    sign_table.columns[1].width = Cm(5.5)
    values = [
        (
            f"{_text(signoff.get('qp_signature'))}\nQualified Person "
            f"(Supervision)\nPE Reg. No. {_text(team.get('pe_number'))}"
        ),
        f"{_text(signoff.get('sign_date'))}\nDate",
    ]
    for cell, value in zip(sign_table.rows[0].cells, values):
        cell.width = Cm(10.5 if cell == sign_table.rows[0].cells[0] else 5.5)
        _set_cell_margins(cell, top=300, start=100, bottom=100, end=100)
        cell.text = value
        cell.paragraphs[0].paragraph_format.space_before = Pt(18)
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        cell.paragraphs[0].runs[0].bold = True

    output = BytesIO()
    doc.save(output)
    return output.getvalue()

