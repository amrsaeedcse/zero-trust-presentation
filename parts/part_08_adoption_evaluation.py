"""Part 08 — Adoption and Evaluation.

Owner: Amr Khaled Husseini Mohamed Tawfik
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Phased Zero Trust migration strategy for brownfield campuses (CISA Maturity Model / NIST phases).
- University-specific challenges (academic openness, budget, BYOD pushback, legacy lab equipment).
- Evaluation metrics & KPIs (blast-radius reduction, MTTR, auth/proxy latency overhead, usability, TCO/ROI).

Out of Scope (DO NOT re-explain):
- Re-explaining technical controls of Parts 01–07; deep data classification & encryption architectures (P9).

Narrative Bridge:
- Builds on Parts 01–07 (complete technical stack) -> Bridges to Data Security & Classification (Part 09).
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_08(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 08  ·  STRATEGY & METRICS",
        title="Adoption &\nEvaluation",
        summary=(
            "Migrating legacy university infrastructure through phased Zero Trust maturity milestones "
            "while measuring security ROI, latency overhead, and campus usability."
        ),
        presenter="Amr Khaled Husseini Mohamed Tawfik",
        motif="crescents",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 08. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
