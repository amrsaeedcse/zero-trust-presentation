"""Part 09 — Data Security and Classification.

Owner: Mahmoud Mohamed Ragab Mohamed Hassan
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Data as the core Zero Trust protect surface.
- University data classification tiers (Public, Internal/Academic, Confidential/Student PII,
  Restricted/Research & Financial).
- Data-centric protection: encryption at rest/in transit, DLP, sensitivity tagging/IRM,
  database access controls & exfiltration prevention.
- Unifying technical finale tying all 9 pillars back to campus data protection.

Out of Scope (DO NOT re-explain):
- Network segmentation (P5), remote VPN/ZTNA (P6), SOC playbooks (P7), or adoption KPIs/budget (P8).

Narrative Bridge:
- Builds on Part 08 (mature campus ZTA) -> Closes the technical story at the innermost asset: University Data.
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_09(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 09  ·  CORE PROTECT SURFACE",
        title="Data Security &\nClassification",
        summary=(
            "Protecting the ultimate target at the center of campus Zero Trust: classifying academic, "
            "student PII, and research data with encryption, DLP, and strict governance."
        ),
        presenter="Mahmoud Mohamed Ragab Mohamed Hassan",
        motif="nested_arch",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 09. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
