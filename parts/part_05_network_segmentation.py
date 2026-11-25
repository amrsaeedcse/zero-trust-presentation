"""Part 05 — Network Segmentation.

Owner: Ali Mohamed Ali Abdallah Omara
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Macro- vs. micro-segmentation for campus zones (dorms/student Wi-Fi, labs, research clusters,
  admin/ERP, IoT/BMS, guest Wi-Fi).
- Software-Defined Perimeter (SDP), SDN/VXLAN/host-based micro-segmentation.
- East-west traffic inspection & blast-radius containment inside the campus network.

Out of Scope (DO NOT re-explain):
- Re-explaining castle-and-moat (P1), PE/PA decision math (P4),
  off-campus remote access/ZTNA vs VPN (P6), SOC monitoring pipelines (P7).

Narrative Bridge:
- Builds on Part 04 (policy enforcement) -> Bridges to Remote Resource Access (Part 06).
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_05(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 05  ·  CAMPUS NETWORK",
        title="Network\nSegmentation",
        summary=(
            "Isolating student Wi-Fi, academic labs, research clusters, ERP, and smart IoT "
            "to inspect east-west traffic and shrink lateral blast radius."
        ),
        presenter="Ali Mohamed Ali Abdallah Omara",
        motif="signal_waves",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 05. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
