# AGENTS.md — Zero Trust Architecture for University Campus Networks

> **Universal Agent & Contributor Contract** (Antigravity, Claude Code, Cursor, Codex, and all AI/human contributors).
> Read this entire file **before** creating or modifying any file in this repository.

---

## 1. Project Mission & Context

- **Course / Subject:** Computer Networks — Team Research Presentation
- **Research Topic:** **Zero Trust Architecture for University Campus Networks**
- **Team Size:** 9 Members (each member owns **one** dedicated Python part module)
- **Why this repo exists:** Previously, each teammate built an isolated slide deck separately, causing broken visual consistency, clashing themes, and overlapping/disconnected content. This repository solves that by using a **Vibe-Coding Python Architecture**:
  - **One unified visual template & theme** (shared across all 9 parts — zero ad-hoc styling).
  - **One continuous narrative flow** (each part picks up directly where the previous part ends — zero repetition).
  - **Animation-First (#1 Priority):** Every slide uses smooth, intentional, native PowerPoint OOXML animations and slide transitions written directly in Python.
  - **Strict part isolation:** Every teammate (and their AI agent) edits **only their own part file**.

---

## 2. Mandatory Agent Startup Protocol (Do This BEFORE Writing Any Code)

Every AI agent (Antigravity, Claude Code, Cursor, Codex) **MUST** do these steps before creating or editing any slide code:

1. **Read & Follow the Project Skills First:** Read `.agents/skills/pptx/SKILL.md` in full (and `.agents/skills/pdf/SKILL.md` whenever reading team PDFs) and follow all of their design, inspection, and QA standards.
   - *Stack Override:* Use **Python (`python-pptx` + `core/`)** instead of `pptxgenjs` so all 9 parts compile into one unified animated `.pptx` via `build.py`.
2. **Inspect the Master Template & Shared Core:** Reference `templates/master_template.pptx` plus `core/theme.py`, `core/layouts.py`, and `core/animations.py` so every slide strictly obeys the locked colors, 4 gradients, frosted-glass styling, fonts (`Space Grotesk` + `Work Sans` in `fonts/`), and native OOXML animations. The 12 layouts (`L01`–`L12`) and 7 motifs are optional building blocks — custom layouts/diagrams are welcome as long as they stay 100% on-theme.
3. **Respect Your Scope, Neighbors & Slide Cap:** Confirm your single file (`parts/part_0N_*.py`) in Section 3, obey the non-overlap & handoff rules in Section 5, and stay within **5–6 slides target (7 absolute max)**.
4. **Build & Run Skill QA:** Build your part (`python build.py --part N`) and complete the QA checks defined in `.agents/skills/pptx/SKILL.md` before finishing.

---

## 3. Team Roster & Strict Part Ownership

| Part # | Team Member | Assigned Topic | Owned File (ONLY edit this file) |
| :---: | :--- | :--- | :--- |
| **01** | **Ahmed Hassan Mohamed Mostafa** | Threat model and boundaries | `parts/part_01_threat_model.py` |
| **02** | **Omar Abdelsalam Abdelfattah Mohamed Afifi** | Identity and authentication | `parts/part_02_identity_auth.py` |
| **03** | **Omar Mounir Mohamed Arafa** | Device security posture | `parts/part_03_device_posture.py` |
| **04** | **Omar Ahmed Abdelaziz Mashhour** | Authorization policies | `parts/part_04_authorization_policies.py` |
| **05** | **Ali Mohamed Ali Abdallah Omara** | Network segmentation | `parts/part_05_network_segmentation.py` |
| **06** | **Amr Abdelazeem Saeed Mohamed** *(Repo Lead)* | Remote resource access | `parts/part_06_remote_access.py` |
| **07** | **Omar Mohamed Fawzy Fahmy Morsi Aboelenein** | Monitoring and response | `parts/part_07_monitoring_response.py` |
| **08** | **Amr Khaled Husseini Mohamed Tawfik** | Adoption and evaluation | `parts/part_08_adoption_evaluation.py` |
| **09** | **Mahmoud Mohamed Ragab Mohamed Hassan** | Data Security and Classification | `parts/part_09_data_security.py` |

### Shared / Coordinator-Only Files (DO NOT modify inside individual part tasks)
- `core/theme.py` — Locked color palette, fonts, typography sizes, and spacing tokens extracted from the chosen template.
- `core/layouts.py` — Reusable template-matched slide frames, cards, diagrams, tables, and headers/footers.
- `core/animations.py` — Native PowerPoint OOXML entrance/emphasis animations & slide transitions.
- `parts/part_00_intro.py` & `parts/part_10_conclusion.py` — Shared cover, agenda, conclusion, and references.
- `build.py` — Master deck compiler (supports full build or single-part preview).
- `AGENTS.md` — This contract file.

---

## 4. Non-Negotiable Rules (Hard Constraints)

### Rule 1 — Strict File & Ownership Isolation
- When asked to work on Member *N*'s part, **ONLY create or edit `parts/part_0N_*.py`** (and assets inside `assets/part_0N/` if needed).
- **NEVER** edit, reformat, or "improve" another teammate's part file (`parts/part_0M_*.py`).
- **NEVER** change `core/theme.py`, `core/layouts.py`, `core/animations.py`, or `build.py` during an individual part task. If a new reusable visual helper is genuinely needed, propose it to the repo lead instead of silently changing shared code.

### Rule 2 — Hard Slide Count Cap (6 Target / 7 Absolute Max)
- **Target:** **5 to 6 slides** per member (including your opening section divider slide if used).
- **Absolute Hard Cap:** **7 slides maximum per member** — `build_part_0N(prs)` **MUST NOT** return more than 7 slides under any circumstance. Never produce 8+ slides; tighten copy, use diagrams/tables, or move detailed narration into `slide.notes_slide` instead of adding slides.

### Rule 3 — Single Locked Theme & Visual Language (Flexible Layouts, Zero Off-Brand Styling)
- **`templates/master_template.pptx` is our visual Single Source of Truth.** Every slide must look unmistakably native to `templates/master_template.pptx` — strictly locked to its exact **color palette, 4 OOXML gradients, frosted-glass card aesthetic, table styling, and font pair (`Space Grotesk` + `Work Sans` in `fonts/`)**.
- **Do NOT invent custom color palettes, random fonts, or off-brand visual styles.** Always import colors, fonts, backgrounds, motifs, and primitive helpers from `core/theme.py` and `core/layouts.py` (never hardcode random hexes or foreign fonts).
- **Custom Layouts & Diagrams ARE Allowed (On-Theme Only):** The 12 layout builders (`L01`–`L12`) are your ready-to-use foundation, not a rigid cage. You **may** design custom slide layouts or add custom shapes, network diagrams, icons, or callouts inside your own `parts/part_0N_*.py` when your content needs a specialized visual — **as long as** every added shape/diagram strictly follows the locked theme, palette, frosted-glass surfaces, typography, and margins so it blends seamlessly with the rest of the deck.
- Keep slides clean, scannable, and visual — **no wall-of-text bullet dumps**, no decorative AI accent stripes/underlines under titles, strong text/icon contrast, generous whitespace (≥ `0.75"` margins on the `20" × 11.25"` canvas), and varied layouts across consecutive slides (never repeat the exact same layout twice in a row).

### Rule 4 — Animation-First Policy (Priority #1)
- **Every content slide MUST include purposeful native PowerPoint animations and a slide transition** via `core/animations.py`!
- Follow cinematic, clean pacing — never chaotic spinning/bouncing gimmicks:
  - **Slide Transitions:** Smooth `fade`, `push`, or `morph` transitions between slides.
  - **Progressive Disclosure:** Animate titles/headers on slide entry (`with_previous`), then reveal cards, diagram nodes, flow arrows, and key callouts in logical reading order (`on_click` or clean staggered `after_previous`, 250–450ms duration).
  - **Diagram Storytelling:** For architecture flows (e.g., Request → IdP → Device Check → PEP → Resource, or VPN vs ZTNA comparison), animate elements step-by-step so the presenter can talk through each stage live!
  - **Valid OOXML Only:** Never hand-hack broken XML; always use `core/animations.py` helpers and verify the output with `validate.py` so PowerPoint never prompts to "repair" the file.

### Rule 5 — Zero Content Overlap & Seamless Narrative Flow
- Treat the 9 parts as **one continuous, cohesive presentation**, NOT 9 disconnected mini-presentations!
- **NEVER re-introduce "What is Zero Trust?"** after Part 1.
- **NEVER re-explain a teammate's domain.** If your slide briefly references a concept owned by another part (e.g., Part 6 mentioning that ZTNA consumes identity & device posture signals), treat it as an **already-established input** in 3–6 words — do **not** explain how MFA or MDM works.
- **Follow the Handoff Chain:** Start your part by naturally continuing from the previous part's closing thread, and end your final slide with a clean narrative bridge into the next part (without mentioning "as Part 5 will say" literally on the slide — keep transitions organic and conceptual!).

---

## 5. Content Scope & Non-Overlap Matrix (All 9 Parts)

Every contributor and AI agent MUST check this matrix before drafting slide outlines:

| Part | Owner & Topic | MUST Cover (In Scope) | MUST NOT Re-Explain (Out of Scope) | Narrative Bridge (From → To) |
| :---: | :--- | :--- | :--- | :--- |
| **01** | **Ahmed Hassan**<br>*Threat Model & Boundaries* | Why traditional university campus perimeter ("castle-and-moat") fails (open campus, BYOD, guests, IoT, lateral movement); campus threat landscape & attackers; Zero Trust core tenets (*Never Trust Always Verify*, *Least Privilege*, *Assume Breach*); defining protect surfaces & trust boundaries. | MFA/SSO protocols (P2); device compliance checks (P3); policy engine internals (P4); micro-segmentation plumbing (P5); VPN vs ZTNA (P6); SIEM/SOAR (P7); migration roadmap (P8); data classification tiers (P9). | **Opens the deck** → Bridges to **Identity** as the first verification pillar once the perimeter is gone. |
| **02** | **Omar Abdelsalam**<br>*Identity & Authentication* | Identity as the new perimeter across campus personas (students, faculty, staff, researchers, guests, service accounts); centralized IdP & SSO (SAML 2.0, OIDC, OAuth 2.0); phishing-resistant MFA (FIDO2/WebAuthn); identity lifecycle (admission → graduation/offboarding); continuous re-authentication. | Why perimeter failed (P1); endpoint posture/MDM (P3); RBAC/ABAC policy engine (P4); LAN segmentation (P5); remote proxy architecture (P6). | Builds on **P1** (no implicit trust) → Bridges to **Device Posture** (verifying *what* device the verified user is on). |
| **03** | **Omar Mounir**<br>*Device Security Posture* | Campus endpoint diversity (managed lab PCs, faculty laptops, student BYOD, smart campus IoT); device health signals (OS patching, disk encryption, EDR/AV, MDM/UEM enrollment, TPM/hardware attestation, device certificates); continuous posture scoring & quarantine/remediation of untrusted endpoints. | User MFA/SSO mechanics (P2); PDP/PEP policy evaluation logic (P4); network VLAN/SDN isolation (P5); SIEM/SOAR incident playbooks (P7). | Builds on **P2** (verified user) → Bridges to **Authorization Policies** (combining user + device signals into access decisions). |
| **04** | **Omar Ahmed**<br>*Authorization Policies* | NIST SP 800-207 core logical components (Policy Engine, Policy Administrator, Policy Enforcement Point); RBAC vs. ABAC vs. context/risk-aware PBAC; dynamic least-privilege & Just-In-Time (JIT) access grants; per-session continuous evaluation & real-time access revocation. | How credentials/MFA work (P2); how device telemetry is gathered (P3); L2/L3 micro-segmentation mechanics (P5); remote ZTNA tunneling (P6); data labeling schemas (P9). | Builds on **P2+P3** (identity + device signals) → Bridges to **Network Segmentation** (enforcing policy boundaries across campus traffic). |
| **05** | **Ali Mohamed**<br>*Network Segmentation* | Macro- vs. micro-segmentation for campus zones (dorms/student Wi-Fi, labs, research clusters, admin/ERP, IoT/BMS, guest Wi-Fi); Software-Defined Perimeter (SDP), SDN/VXLAN/host-based micro-segmentation; east-west traffic inspection & blast-radius containment inside the campus network. | Re-explaining castle-and-moat (P1); PE/PA decision math (P4); off-campus remote access/ZTNA vs VPN (P6); SOC monitoring pipelines (P7). | Builds on **P4** (policy enforcement) → Bridges to **Remote Resource Access** (extending Zero Trust beyond campus walls). |
| **06** | **Amr Abdelazeem**<br>*Remote Resource Access* | Off-campus access to university resources (LMS, grading/ERP, research GPU clusters, internal apps); deep architectural comparison of legacy VPN (network-level trust, concentrator bottleneck, lateral exposure) vs. Zero Trust Network Access (ZTNA) / Identity-Aware Proxy (IAP); client-based vs. clientless ZTNA, outbound-only app connectors, dark/invisible internal IPs. | General MFA/SSO setup (P2); MDM/EDR enrollment (P3); internal campus VLAN micro-segmentation (P5); SIEM/SOAR response (P7). | Builds on **P5** (on-campus micro-segmentation) → Bridges to **Monitoring & Response** (assuming breach across all local & remote sessions). |
| **07** | **Omar Mohamed Fawzy**<br>*Monitoring & Response* | Continuous visibility & telemetry ingestion (IdP logs, endpoint EDR, PEP/ZTNA flows, east-west traffic); SIEM, UEBA (behavioral anomaly detection), XDR, and SOAR automated playbooks; university threat scenarios (compromised student account, credential stuffing, anomalous lab exfiltration); closed-loop dynamic trust feedback into the Policy Engine. | Initial auth/MDM mechanics (P2–P3); ZTNA proxy tunneling (P6); migration ROI/maturity roadmap (P8); data classification tiers & DLP policy taxonomy (P9). | Builds on **P1–P6** (all controls emit telemetry) → Bridges to **Adoption & Evaluation** (how a university deploys and measures this architecture). |
| **08** | **Amr Khaled**<br>*Adoption & Evaluation* | Phased Zero Trust migration strategy for brownfield campuses (CISA Zero Trust Maturity Model / NIST phases); university-specific challenges (academic openness, budget, BYOD pushback, legacy lab equipment); evaluation metrics & KPIs (blast-radius reduction, MTTR, auth/proxy latency overhead, usability, TCO/ROI). | Re-explaining technical controls of P1–P7; deep data classification & encryption architectures (P9). | Builds on **P1–P7** (full technical stack) → Bridges to **Data Security & Classification** (everything ultimately protects campus data at the center). |
| **09** | **Mahmoud Mohamed Ragab**<br>*Data Security & Classification* | Data as the core Zero Trust protect surface; university data classification tiers (Public, Internal/Academic, Confidential/Student PII, Restricted/Research & Financial); data-centric protection — encryption at rest/in transit, Data Loss Prevention (DLP), sensitivity tagging/IRM, database access controls & exfiltration prevention; unifying finale tying all pillars back to data protection. | Re-explaining network segmentation (P5), remote VPN/ZTNA (P6), SOC playbooks (P7), or adoption KPIs/budget (P8). | Builds on **P8** (mature campus ZTA) → **Closes the technical story** at the innermost Zero Trust asset: University Data. |

---

## 6. Locked Visual Template & Theme (`LOCKED` ✅)

- **Master Reference File:** `templates/master_template.pptx` (*Cyber Security Seminar* — 15 slides, 16:9 High-DPI canvas).
- **How `build.py` Uses It (Zero Duplication):**
  - The 7 reusable transparent glassmorphic motif PNGs are read **directly from inside `templates/master_template.pptx` in memory** (`zipfile` + `BytesIO`) — do NOT dump extracted template media into the repo.
  - All gradients, frosted-glass cards, table styles, and layouts match `templates/master_template.pptx` natively in OOXML.
  - Sample stock photos (`image7, 11, 12, 15, 16, 17.jpeg`), static sample charts (`image9.png`), and SlidesCarnival credits (`image24.png`) are **never** reused; replace them with native editable PowerPoint charts, architecture diagrams, and Zero Trust visuals built in the exact same palette.

### 6.1 Canvas & Locked Open-Source Fonts (`20.00" × 11.25"` 16:9)
> **Canvas & Font Note:** Because `master_template.pptx` uses a `20.00" × 11.25"` canvas (1.5× standard `13.33" × 7.5"`), all point sizes and margins below are calibrated to `20" × 11.25"`. Because the template's display font `Glock Grotesk` is commercial and subsetted (missing glyphs like `Z`), the team standardizes on **`Space Grotesk`** (free geometric neo-grotesk matching its exact visual character) + **`Work Sans`** (the template's exact original body font). All `.ttf` files are bundled in `fonts/` — teammates install them once before editing, and the final deck is saved with **Embed All Characters** for the lecture hall PC.

| Role | Locked Font (`fonts/`) | Size (on 20"×11.25") | Color |
| :--- | :--- | :---: | :--- |
| Cover Title | `Space Grotesk` (Bold) | `90–95pt` | `#FFFFFF` |
| Slide Title (Standard) | `Space Grotesk` (Bold) | `56–64pt` | `#FFFFFF` |
| Split / Compact Title | `Space Grotesk` (Bold) | `38–42pt` | `#FFFFFF` |
| Card / Table / Node Heading | `Work Sans` (Bold) | `24–30pt` | `#FFFFFF` (or `#000000` on light panels) |
| Body / Descriptions | `Work Sans` (Regular) | `20–26pt` *(floor: `18pt`)* | `#FFFFFF` / `#D3DDD3` (dark bg) or `#000000` (light sage) |
| Subtle Footers / Captions | `Work Sans` (Regular) | `15–17pt` | `#D3DDD3` |

### 6.2 Locked Color Tokens (Extracted Directly from `master_template.pptx` XML)

| Token Name | Hex | Role in Template |
| :--- | :---: | :--- |
| `OBSIDIAN` | `#070908` | Deepest dark gradient anchor |
| `DARK_TEAL` | `#233D40` | Deepest swatch / dark chart bar / accent surface |
| `DEEP_PINE` | `#29464A` | Core dark-teal gradient midpoint |
| `MUTED_TEAL` | `#638079` | Mid-tone swatch / chart bars |
| `TABLE_HEADER` | `#6F8384` | Table header fill (with bold `#FFFFFF` text) & gradient stop `#6D8080` |
| `SAGE_GLASS` | `#9EB49F` | Frosted card hairline border (`0.75pt` @ `29.8%` alpha) |
| `SOFT_SAGE` | `#B7C6B7` | Table body row fill (with `#000000` text) & light gradient edge (`#B9C8B6` / `#BDCBBB`) |
| `ICE_MIST` | `#D3DDD3` | Lightest swatch / light checklist card (`L12`) / muted highlights |
| `WHITE` | `#FFFFFF` | Primary text on dark, flowchart arrows (`1.5pt`), frosted card fill (`29.8%` alpha) |
| `BLACK` | `#000000` | High-contrast body text inside `#B7C6B7` / `#D3DDD3` light surfaces |
| `BORDER_GRAY` | `#CCCCCC` | Subtle table cell gridlines |
| `VERIFIED_MINT` *(diagram-only)* | `#6FD3A8` | Subtle semantic highlight for **Verified / Allow / ZTNA** badges or arrows |
| `ALERT_CORAL` *(diagram-only)* | `#E07A6B` | Subtle semantic highlight for **Blocked / Breach / Legacy VPN Risk** badges or arrows |

### 6.3 The 4 Native OOXML Gradient Backgrounds
- **`BG_HORIZONTAL`** (`0°`: `0% #070908` → `33% #29464A` → `67% #6D8080` → `100% #BDCBBB`) — Cover (`L01`), Chart (`L06`), Split Panel (`L12`).
- **`BG_VERTICAL`** (`90°`: `0% #070908` → `33% #29464A` → `67% #6D8080` → `100% #B9C8B6`) — 2×2 Grid (`L02`), Wide Visual (`L04`), Credits.
- **`BG_DIAGONAL`** (`45°`: `0% #070908` → `50% #29464A` → `100% #6D8080`) — Section Divider (`L03`), Table (`L05`), Stepped Stack (`L11`).
- **`BG_RADIAL`** (`radial`: `0% #070908` → `33% #6D8080` → `67% #29464A` → `100% #BDCBBB`) — 6-Node Flowchart (`L07`), Architecture Diagrams.

### 6.4 Signature Component Recipes & 7 Embedded Glassmorphic Motifs
- **Frosted Glass Card:** `#FFFFFF` fill @ `29.8%` opacity (`alpha=29804`), `0.75pt` solid `#9EB49F` border @ `29.8%` opacity, softly rounded corners.
- **Stepped-Opacity Stack (`L11`):** 4 stacked cards ramping fill & border opacity `11% → 21% → 31% → 46%`.
- **Sage Comparison Table (`L05`):** Header `#6F8384` (`Work Sans Bold 24pt #FFFFFF`), body `#B7C6B7` (`Work Sans 24pt #000000`), `#CCCCCC` hairlines.
- **Flowchart Connectors (`L07`):** `1.5pt` (`w=19050`) crisp `#FFFFFF` arrows between frosted nodes.
- **7 Embedded Motifs — Including ALL Slide 13 Resource Page Elements (loaded in memory from `templates/master_template.pptx`):** Any member can freely reuse, scale, position, rotate, or layer any of these 7 native transparent glassmorphic elements on any slide via `core/theme.py` (`add_motif(slide, name, ...)`):
  1. `prism` (`image1.png`) — 8-blade geometric prism star (cover / finale / architectural focal point)
  2. `logo_wave` (`image3.png`) — metallic triple-ribbon emblem (cover header / brand mark)
  3. `arc` (`image5.png`) — sweeping curved glass fan (section dividers / hooks)
  4. `glow` (`image8.png`) — soft radial luminous bloom (ambient highlight behind tables, cards, or diagrams)
  5. `nested_arch` (`image13.png`) — concentric arch layers (defense-in-depth / Zero Trust layers / 2-column compare)
  6. `crescents` (`image18.png`) — stacked concentric crescent orbs (review / maturity phases / stepped stacks)
  7. `signal_waves` (`image22.png`) — curved network signal bands from Slide 13 Resource Page (campus Wi-Fi, segmentation, remote ZTNA, telemetry)

### 6.5 Template-Matched Starter Layouts (`L01`–`L12`) & Custom On-Theme Layouts
Use these 12 template-matched layouts in `core/layouts.py` as your primary toolkit (and never repeat the exact same layout back-to-back). **Need a custom layout or specialized diagram/shapes** (e.g., VPN vs. ZTNA topology, NIST PE/PA/PEP loop, campus VLAN zones)? Build it directly inside your own `parts/part_0N_*.py` using the locked primitives from `core/theme.py` (exact palette, gradients, frosted-glass cards, arrows, and `Space Grotesk` / `Work Sans` typography) so it looks 100% part of `master_template.pptx`:

| ID | Template Slide | Layout Archetype | Best Use Across Parts 1–9 |
| :---: | :---: | :--- | :--- |
| **`L01`** | Slide 01 | Cover / Hero (`prism` + `logo_wave` + pill badge) | Deck cover (`part_00`), optional closing hero |
| **`L02`** | Slide 02 | Centered Title + 2×2 Frosted Glass Grid | Agenda, 4 threats, 4 identity pillars, 4 controls |
| **`L03`** | Slide 03 | Bold Statement / Section Divider + `arc` | Opening hook or section divider for a part |
| **`L04`** | Slide 04 | Split Top Header + Full-Width Bottom Visual Stage | Large custom Zero Trust architecture / topology diagrams |
| **`L05`** | Slide 05 | Split Top Header + Sage Comparison Table + `glow` | VPN vs ZTNA, RBAC vs ABAC, Data Classification Matrix |
| **`L06`** | Slide 06 | Native Chart Left (65%) + Key Takeaways Right (35%) | Campus metrics, latency benchmarks, risk reduction KPIs |
| **`L07`** | Slide 07 | 6-Node Connected Architecture Flowchart | ZTNA request lifecycle, NIST PE/PA/PEP loop, attack kill chain |
| **`L08`** | Slide 08 | Left Title + 4 Floating Translucent Glass Circles | Campus personas, posture signals, threat vectors |
| **`L09`** | Slide 09 | Left Title + `nested_arch` + 2 Tall Comparison Columns | Legacy vs Zero Trust, Client vs Clientless ZTNA, Managed vs BYOD |
| **`L10`** | Slide 10 | Centered Title + 3 Vertical Showcase Cards | 3 Zero Trust tenets, 3 segmentation layers, 3 SOC stages |
| **`L11`** | Slide 11 | Left Title + `crescents` + 4 Stepped-Opacity Rows | Maturity phases (P8), 4 data tiers (P9), part takeaways |
| **`L12`** | Slide 12 | Frosted Dark Summary Panel + Light `#D3DDD3` Checklist Card | Policy rules, hardening checklist, narrative handoff bridge |

---

## 7. Repository Layout & Module Contract

```text
zero_trust_presentation/
├── AGENTS.md                         # Universal rules (this file)
├── .agents/skills/                   # Latest Anthropic pptx & pdf skills
│   ├── pptx/
│   └── pdf/
├── fonts/                            # Bundled Space Grotesk & Work Sans .ttf files (install once)
├── templates/
│   └── master_template.pptx          # Single source of truth (theme, layouts, 7 embedded motifs)
├── core/                             # Locked shared framework (Coordinator-owned)
│   ├── theme.py                      # Palette, fonts, gradients, motifs, spacing (20"×11.25")
│   ├── layouts.py                    # Template-matched L01–L12 layout builders & components
│   └── animations.py                 # Native PowerPoint OOXML animations & transitions
├── parts/                            # 9 isolated member modules + intro/outro
│   ├── part_00_intro.py              # Cover slide, team roster, architecture agenda
│   ├── part_01_threat_model.py       # Ahmed Hassan
│   ├── part_02_identity_auth.py      # Omar Abdelsalam
│   ├── part_03_device_posture.py     # Omar Mounir
│   ├── part_04_authorization_policies.py # Omar Ahmed
│   ├── part_05_network_segmentation.py   # Ali Mohamed
│   ├── part_06_remote_access.py      # Amr Abdelazeem
│   ├── part_07_monitoring_response.py    # Omar Mohamed Fawzy
│   ├── part_08_adoption_evaluation.py    # Amr Khaled
│   ├── part_09_data_security.py      # Mahmoud Mohamed Ragab
│   └── part_10_conclusion.py         # Unified summary, references, Q&A
├── assets/                           # Optional per-part custom visuals (assets/part_01/ ... part_09/)
├── scripts/                          # Local PowerPoint render & QA scripts
├── build.py                          # Build full deck (`python build.py`) or one part (`--part 6`)
└── dist/                             # Generated .pptx & QA renders (gitignored)
```

### Standard Part Entrypoint Signature
Every member module `parts/part_0N_<slug>.py` exposes one function:
```python
def build_part_0N(prs) -> list:
    """Append 5–6 slides (absolute max 7) for Part N to `prs` and return the created slide list."""
```
`build.py` automatically enforces the ≤ 7 slide cap per part at compile time.

---

## 8. Before Finishing Your Part

1. **Fonts Installed:** Install the `.ttf` files in `fonts/` once locally (`Space Grotesk` + `Work Sans`).
2. **Boundary & Budget Check:** Only your `parts/part_0N_*.py` was edited, zero overlap with Section 5 neighbors, and total slides are **5–6 target (7 max)**.
3. **Build Your Part:** Run `python build.py --part <N>` to generate `dist/part_0N_preview.pptx`.
4. **Complete Skill QA & Animation Check:** Run the full QA workflow from `.agents/skills/pptx/SKILL.md` (file validation + visual slide inspection) and confirm every content slide has smooth transitions and animations.
