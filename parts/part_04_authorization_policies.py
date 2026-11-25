"""Part 04 — Authorization Policies.

Owner: Omar Ahmed Abdelaziz Mashhour
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- NIST SP 800-207 core logical components (Policy Engine, Policy Administrator, Policy Enforcement Point).
- RBAC vs. ABAC vs. context/risk-aware PBAC.
- Dynamic least-privilege, Just-In-Time (JIT) access grants, per-session continuous evaluation & revocation.

Out of Scope (DO NOT re-explain):
- How credentials/MFA work (P2), how device telemetry is gathered (P3),
  L2/L3 micro-segmentation mechanics (P5), remote ZTNA tunneling (P6), data labeling schemas (P9).

Narrative Bridge:
- Builds on Parts 02+03 (identity + device signals) -> Bridges to Network Segmentation (Part 05).
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_04(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 04  ·  POLICY ENGINE",
        title="Authorization\nPolicies",
        summary=(
            "Turning live identity and device posture signals into dynamic, per-session, "
            "least-privilege access decisions via NIST SP 800-207 Policy Engines."
        ),
        presenter="Omar Ahmed Abdelaziz Mashhour",
        motif="prism",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 04. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
