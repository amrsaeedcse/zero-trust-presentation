"""Part 10 — Unified Conclusion & Q&A (Coordinator-Owned).

DO NOT modify inside individual member part tasks.
"""

from core.layouts import L01_cover, L12_split_checklist
from core.theme import set_notes


def build_part_10(prs) -> list:
    slides = []

    s1, _, _ = L12_split_checklist(
        prs,
        title="Campus Zero Trust\nIn One View",
        subtitle=(
            "Security moves from a brittle campus firewall to continuous verification "
            "across identities, endpoints, policies, micro-segments, remote sessions, and data."
        ),
        checklist_title="Key Architectural Takeaways",
        items=[
            "Eliminate implicit campus-LAN & VPN trust — verify every session explicitly",
            "Bind user identity + live endpoint posture into dynamic NIST PE/PA/PEP policies",
            "Contain lateral movement via campus micro-segmentation & app-level ZTNA",
            "Feed continuous SIEM/UEBA telemetry back into real-time access decisions",
            "Anchor every control around classified university data and phased adoption",
        ],
        part_label="Conclusion  ·  Summary",
        transition="fade",
    )
    set_notes(s1, "Unified conclusion slide tying together all 9 pillars of University Campus Zero Trust Architecture.")
    slides.append(s1)

    s2, _ = L01_cover(
        prs,
        title="Thank You\n& Questions",
        subtitle="Zero Trust Architecture for University Campus Networks  ·  9-Member Research Team",
        top_label="Computer Networks  ·  References: NIST SP 800-207  ·  CISA ZTMM v2.0",
        badge_text="Q & A  ·  Discussion",
        transition="fade",
    )
    set_notes(s2, "Closing thank-you and Q&A slide.")
    slides.append(s2)

    return slides
