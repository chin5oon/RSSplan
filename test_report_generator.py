import io
import unittest

from docx import Document

from report_generator import build_docx


class ReportGeneratorTest(unittest.TestCase):
    def test_generates_valid_docx_with_expected_sections(self):
        data = {
            "project": {
                "reference": "A1234-TEST",
                "description": "Worked example",
                "site_type": "Construction site",
                "address": "1 Example Road, Singapore",
                "structural_system": "Reinforced-concrete frame",
                "foundation_system": "Bored piles",
                "challenges": "None identified",
                "permit_date": "2026-07-01",
            },
            "team": {
                "qp_name": "Er Example",
                "pe_number": "12345",
                "company": "Example Consultants",
                "prepared_date": "2026-08-01",
                "organisation": "QP(S) → RE → RTO → Builder operator",
                "site_supervisors": "Er Resident (RE)",
                "builder_operators": "Site operator",
                "backup_personnel": "Named trained backup",
                "training": "Provider and project RSS training completed",
                "competency": "Demonstrated during trial and accepted by QP(S)",
            },
            "phases": {
                "phase_1": "15% of the first 30%",
                "phase_2": "30% from 30-75%",
                "phase_3": "50% from 75-100%",
                "beyond": "Not proposed",
                "criteria": "95% uptime and satisfactory performance",
                "parallel_plan": "10% sample and parallel learning phase",
                "review_cadence": "End of each phase",
            },
            "technology": {
                "live_devices": "Two HD cameras and one backup",
                "evidence_devices": "Digital calliper",
                "audio": "Noise-cancelling headset",
                "connectivity": "Primary 5G",
                "backup_connectivity": "Independent 4G router",
                "platform": "Controlled project platform",
                "video_standard": "1080p, timestamped",
                "storage": "Role-controlled project server",
                "power_backup": "Two spare batteries",
                "equipment_register": "Serialised and calibrated",
            },
            "process": {
                "before": "Verify drawings, method and equipment.",
                "during": "Inspect systematically and record decisions.",
                "after": "Save evidence and close actions.",
                "communication": "Repeat back critical instructions.",
                "stop_work": "Stop if visibility or intervention is inadequate.",
                "tech_failure": "Use backup, then revert in person.",
                "poor_evidence": "Reject and repeat.",
                "safety_incident": "Suspend and follow emergency plan.",
                "non_conformity": "Record, rectify and verify closure.",
            },
            "records": {
                "naming": "Project-Activity-Date-Revision",
                "access": "Role-based access",
                "backups": "Daily protected backup",
                "retention": "Reports 5 years; video 2 years after TOP",
                "verification": "10% in-person sample",
                "audits": "Planned internal audit",
                "performance": "Uptime and detection rate",
                "traceability": "Element, personnel, time, result and evidence",
            },
            "activities": [
                {
                    "work_type": "Concreting works",
                    "description": "Concrete placement supervision",
                    "location": "Block A, Level 2",
                    "complexity": "Simple",
                    "frequency": "Continuous",
                    "approach": "Live remote supervision",
                    "phase": "Phase 1",
                    "extent": "15% of the first 30%",
                    "evidence": "Full recording and identified screenshots",
                    "equipment": "HD cameras and two-way audio",
                    "annex_d_reviewed": True,
                    "deviation": "",
                }
            ],
            "signoff": {
                "qp_signature": "Er Example",
                "sign_date": "2026-08-01",
            },
        }
        payload = build_docx(data)
        self.assertGreater(len(payload), 20_000)
        doc = Document(io.BytesIO(payload))
        all_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("Remote Site", all_text)
        self.assertIn("1. Project background", all_text)
        self.assertIn("QP(S) declaration", all_text)
        self.assertGreaterEqual(len(doc.tables), 8)


if __name__ == "__main__":
    unittest.main()
