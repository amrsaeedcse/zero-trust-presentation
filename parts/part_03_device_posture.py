"""Part 03 — Device Security Posture.

Owner: Omar Mounir Mohamed Arafa
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Campus endpoint heterogeneity (managed lab PCs, faculty laptops, student BYOD, smart IoT).
- Device health signals (OS patching, disk encryption, EDR/AV, MDM/UEM, TPM/attestation, device certs).
- Continuous posture scoring, quarantine, and remediation of untrusted devices.

Out of Scope (DO NOT re-explain):
- User MFA/SSO mechanics (P2), PDP/PEP policy evaluation logic (P4),
  network VLAN/SDN isolation (P5), SIEM/SOAR incident playbooks (P7).

Narrative Bridge:
- Builds on Part 02 (verified user) -> Bridges to Authorization Policies (Part 04).
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_03(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 03  ·  ENDPOINT PILLAR",
        title="Device Security\nPosture",
        summary=(
            "Continuously assessing managed lab workstations, student BYOD laptops, and campus IoT "
            "before granting trust to any connecting endpoint."
        ),
        presenter="Omar Mounir Mohamed Arafa",
        motif="crescents",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 03. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
