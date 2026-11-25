"""Part 02 — Identity and Authentication.

Owner: Omar Abdelsalam Abdelfattah Mohamed Afifi
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Identity as the new perimeter across campus personas
  (students, faculty, staff, researchers, guests, service accounts).
- Centralized IdP & SSO (SAML 2.0, OIDC, OAuth 2.0).
- Phishing-resistant MFA (FIDO2 / WebAuthn).
- Campus identity lifecycle (admission -> graduation/offboarding) & continuous re-auth.

Out of Scope (DO NOT re-explain):
- Why the perimeter failed (P1), endpoint posture/MDM (P3),
  RBAC/ABAC policy engine (P4), LAN segmentation (P5), remote proxy/ZTNA (P6).

Narrative Bridge:
- Builds on Part 01 (no implicit perimeter trust) -> Bridges to Device Posture (Part 03).
"""

from core.layouts import L03_section_divider
from core.theme import set_notes


def build_part_02(prs) -> list:
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 02  ·  IDENTITY PILLAR",
        title="Identity &\nAuthentication",
        summary=(
            "Establishing strong, phishing-resistant identity verification across students, "
            "faculty, guests, and non-human accounts throughout the academic lifecycle."
        ),
        presenter="Omar Abdelsalam Abdelfattah Mohamed Afifi",
        motif="nested_arch",
        transition="push_left",
    )
    set_notes(s1, "Starter section divider for Part 02. Expand to 5–6 slides (max 7) following AGENTS.md.")
    slides.append(s1)

    return slides
