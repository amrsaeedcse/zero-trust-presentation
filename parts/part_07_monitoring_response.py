"""Part 07 — Monitoring and Response.

Owner: Omar Mohamed Fawzy Fahmy Morsi Aboelenein
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Continuous visibility & telemetry ingestion (IdP logs, EDR, PEP/ZTNA flows, east-west traffic).
- SIEM, UEBA (behavioral anomaly detection), XDR, and SOAR automated playbooks.
- University threat scenarios (compromised student account, credential stuffing, lab exfiltration).
- Closed-loop dynamic trust feedback back into the Policy Engine.

Out of Scope (DO NOT re-explain):
- Initial auth/MDM mechanics (P2–P3), ZTNA proxy tunneling (P6),
  migration ROI/maturity roadmap (P8), data classification tiers & DLP policy taxonomy (P9).

Narrative Bridge:
- Builds on Parts 01–06 (all controls emit telemetry) -> Bridges to Adoption & Evaluation (Part 08).
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_07(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 07  ·  VISIBILITY & SOC",
        title="Monitoring &\nAutomated Response",
        summary=(
            "Assuming breach across local and remote sessions through SIEM, UEBA anomaly detection, "
            "SOAR playbooks, and closed-loop dynamic trust revocation."
        ),
        presenter="Omar Mohamed Fawzy Fahmy Morsi Aboelenein",
        motif="arc",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 07. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
