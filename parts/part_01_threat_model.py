"""Part 01 — Threat Model and Boundaries.

Owner: Ahmed Hassan Mohamed Mostafa
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Why traditional university campus perimeter ("castle-and-moat") fails
  (open campus, BYOD, guests, smart-campus IoT, lateral movement).
- Campus threat landscape & attacker profiles.
- Core Zero Trust tenets (Never Trust Always Verify, Least Privilege, Assume Breach).
- Defining protect surfaces & trust boundaries.

Out of Scope (DO NOT re-explain):
- MFA/SSO protocols (P2), device posture/MDM (P3), policy engine internals (P4),
  VLAN/SDN segmentation plumbing (P5), VPN vs ZTNA (P6), SIEM/SOAR (P7),
  migration roadmap (P8), data classification tiers (P9).

Narrative Bridge:
- Opens the technical deck -> Bridges to Identity (Part 02) as the first verification pillar.
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_01(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 01  ·  FOUNDATION",
        title="Threat Model &\nTrust Boundaries",
        summary=(
            "Why the traditional castle-and-moat university perimeter breaks down under BYOD, "
            "open research, and IoT — and how Zero Trust redefines campus protect surfaces."
        ),
        presenter="Ahmed Hassan Mohamed Mostafa",
        motif="arc",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 01. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
