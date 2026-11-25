"""Part 06 — Remote Resource Access.

Owner: Amr Abdelazeem Saeed Mohamed (Repo Lead)
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Off-campus access to university resources (LMS, grading/ERP, research GPU clusters, internal apps).
- Architectural comparison: Legacy VPN (network-level trust, concentrator bottleneck, lateral exposure)
  vs. Zero Trust Network Access (ZTNA) / Identity-Aware Proxy (IAP).
- Client-based vs. clientless ZTNA, outbound-only app connectors, dark/invisible internal IPs.

Out of Scope (DO NOT re-explain):
- General MFA/SSO setup (P2), MDM/EDR enrollment (P3),
  internal campus VLAN micro-segmentation (P5), SIEM/SOAR response (P7).

Narrative Bridge:
- Builds on Part 05 (on-campus micro-segmentation) -> Bridges to Monitoring & Response (Part 07).
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_06(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 06  ·  REMOTE ACCESS",
        title="Remote Resource\nAccess (ZTNA vs. VPN)",
        summary=(
            "Extending Zero Trust beyond campus walls: replacing network-level VPN tunnels with "
            "application-layer Identity-Aware Proxies and outbound-only connectors."
        ),
        presenter="Amr Abdelazeem Saeed Mohamed",
        motif="signal_waves",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 06. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
