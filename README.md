# Zero Trust Architecture for University Campus Networks

Unified 9-member Computer Networks presentation built with **Vibe-Coding Python (`python-pptx` + native OOXML animations)** locked to `templates/master_template.pptx`.

> **Every contributor & AI agent MUST read [`AGENTS.md`](./AGENTS.md) and `.agents/skills/` before editing any file.**

---

## Quick Start for Team Members (5 Steps)

1. **Clone & install Python dependencies:**
   ```bash
   git clone https://github.com/amrsaeedcse/zero-trust-presentatio.git
   cd zero-trust-presentatio
   pip install -r requirements.txt
   ```
2. **Install bundled fonts (once):**
   Open `fonts/`, select all 7 `.ttf` files (`Space Grotesk` + `Work Sans`), right-click → **Install**.
3. **Create your own branch (zero merge conflicts!):**
   ```bash
   git checkout -b part-0N-your-topic
   ```
4. **Edit ONLY your owned file (`parts/part_0N_*.py`)** — or prompt your AI coding agent (Antigravity / Cursor / Claude Code / Codex):
   > *"Read `AGENTS.md` and `.agents/skills/pptx/SKILL.md` (plus `pdf/SKILL.md` if reading a PDF), inspect `core/`, and build Part N (`parts/part_0N_*.py`) with 5–6 animated slides (max 7) following the locked theme and non-overlap rules."*
5. **Build, Validate & Visually Inspect your part:**
   ```bash
   python build.py --part N --validate
   # On Windows with PowerPoint installed:
   .\scripts\render_qa.ps1 -InputPptx dist\part_0N_preview.pptx -OutDir dist\qa_part_0N
   ```
   Then push your branch and open a Pull Request into `main`!

---

## Team Part Ownership

| Part | Member | Topic | File |
| :---: | :--- | :--- | :--- |
| **01** | Ahmed Hassan Mohamed Mostafa | Threat model and boundaries | `parts/part_01_threat_model.py` |
| **02** | Omar Abdelsalam Abdelfattah Mohamed Afifi | Identity and authentication | `parts/part_02_identity_auth.py` |
| **03** | Omar Mounir Mohamed Arafa | Device security posture | `parts/part_03_device_posture.py` |
| **04** | Omar Ahmed Abdelaziz Mashhour | Authorization policies | `parts/part_04_authorization_policies.py` |
| **05** | Ali Mohamed Ali Abdallah Omara | Network segmentation | `parts/part_05_network_segmentation.py` |
| **06** | Amr Abdelazeem Saeed Mohamed *(Lead)* | Remote resource access | `parts/part_06_remote_access.py` |
| **07** | Omar Mohamed Fawzy Fahmy Morsi Aboelenein | Monitoring and response | `parts/part_07_monitoring_response.py` |
| **08** | Amr Khaled Husseini Mohamed Tawfik | Adoption and evaluation | `parts/part_08_adoption_evaluation.py` |
| **09** | Mahmoud Mohamed Ragab Mohamed Hassan | Data Security and Classification | `parts/part_09_data_security.py` |

## Full Deck Compilation (Coordinator)
```bash
python build.py --validate
```
Produces `dist/Zero_Trust_Campus_Networks.pptx` combining Intro + Parts 01–09 + Conclusion, automatically enforcing the **≤ 7 slides per member** cap.
