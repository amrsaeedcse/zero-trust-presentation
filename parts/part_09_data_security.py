"""Part 09 — Data Security and Classification.

Owner: Mahmoud Mohamed Ragab Mohamed Hassan
Slide Budget: 6 slides (Target 5–6 slides, 7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Data as the core Zero Trust protect surface (the crown jewels at the center of the architecture).
- University data classification tiers (Public, Internal/Academic, Confidential/Student PII,
  Restricted/Research & Financial).
- Data-centric protection: encryption at rest/in transit, DLP, sensitivity tagging/IRM,
  database access controls & exfiltration prevention.
- Unifying technical finale tying all 9 pillars back to campus data protection.

Out of Scope (DO NOT re-explain):
- Network segmentation plumbing (P5), remote VPN/ZTNA (P6), SOC playbooks (P7), or adoption KPIs/budget (P8).

Narrative Bridge:
- Builds on Part 08 (mature campus ZTA with operational governance) -> Closes the technical story
  at the innermost asset: University Data.
"""

from __future__ import annotations

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import (
    L03_section_divider,
    L04_wide_visual,
    L05_table_compare,
    L07_flowchart_6node,
    L09_two_columns,
    L10_three_cards,
)
from core.theme import (
    COLORS,
    FONT_BODY,
    FONT_TITLE,
    MARGIN_X,
    MARGIN_Y,
    add_arrow_connector,
    add_badge,
    add_footer,
    add_frosted_card,
    add_motif,
    add_solid_card,
    add_textbox,
    rgb,
    set_notes,
)


def build_part_09(prs) -> list:
    """Compile Part 09 (Data Security & Classification) — exactly 6 animated slides."""
    slides = []

    # =========================================================================
    # SLIDE 1 (L03): Section Divider
    # =========================================================================
    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 09  ·  CORE PROTECT SURFACE",
        title="Data Security &\nClassification",
        summary=(
            "Every perimeter barrier, identity challenge, and network segment exists for one "
            "ultimate objective: safeguarding high-value university data assets at rest, "
            "in transit, and in use."
        ),
        presenter="Mahmoud Mohamed Ragab Mohamed Hassan",
        motif="nested_arch",
        transition="push_left",
    )
    set_notes(
        s1,
        "Welcome to Part 09: Data Security and Classification. Presented by Mahmoud Mohamed Ragab.\n\n"
        "Up to this point in our presentation, we have examined identity verification, endpoint health, "
        "dynamic policy engines, campus micro-segmentation, remote access proxies, telemetry monitoring, "
        "and adoption roadmaps. But what are all these systems actually protecting? The answer is University Data. "
        "In a Zero Trust Architecture, data is the ultimate protect surface—the crown jewels. Today we explore how "
        "a university categorizes, encrypts, tags, and governs its information assets from creation to archival.",
    )
    slides.append(s1)

    # =========================================================================
    # SLIDE 2 (L05): University Data Classification Tiers
    # =========================================================================
    s2, _, _ = L05_table_compare(
        prs,
        title="Campus Data Classification Framework",
        intro=(
            "A university cannot protect all information equally. Establishing four distinct classification "
            "tiers aligns campus data sensitivity with strict cryptographic, identity, and access controls."
        ),
        headers=[
            "Tier & Sensitivity",
            "Campus Examples & Assets",
            "Access & Identity Criteria",
            "Security & Storage Controls",
        ],
        rows=[
            [
                "Tier 1: Public",
                "Course catalogs, faculty directories, press releases, published academic research",
                "Open to all campus guests and unauthenticated web visitors; public read-only",
                "Integrity monitoring, CDN edge caching, standard HTTPS in transit",
            ],
            [
                "Tier 2: Internal / Academic",
                "Course syllabi, lecture slides, departmental meeting minutes, LMS course content",
                "Campus SSO (SAML 2.0 / OIDC); enrolled students, active faculty, and staff",
                "TLS 1.3 in transit, role-based boundaries, standard server-side storage encryption",
            ],
            [
                "Tier 3: Confidential / PII",
                "Student transcripts, admissions files, financial aid records, FERPA data, staff HR/payroll",
                "Phishing-resistant MFA (FIDO2) + compliant managed device posture check",
                "AES-256 at rest, strict DLP inspection, dynamic data masking for query views",
            ],
            [
                "Tier 4: Restricted / Regulated",
                "Export-controlled defense research (ITAR), genomic data, IP patents, clinical HIPAA records",
                "Just-In-Time (JIT) access, hardware TPM token, verified isolated enclave boundary",
                "Envelope encryption (KMS/HSM), air-gapped enclaves, IRM persistent rights",
            ],
        ],
        col_widths=[3.8, 4.8, 4.6, 4.56],
        part_label="Part 09: Data Security & Classification  ·  Slide 2/6",
        transition="fade",
    )
    set_notes(
        s2,
        "Slide 2 establishes the university data classification framework. In an academic institution, openness "
        "is essential for learning, but sensitive records must be rigorously guarded. We establish four tiers:\n\n"
        "1. Public: Freely accessible info like course catalogs.\n"
        "2. Internal/Academic: Day-to-day educational material accessible to all active students and faculty.\n"
        "3. Confidential / Student PII: Governed by federal regulations like FERPA. Requires verified identity and device compliance.\n"
        "4. Restricted / Regulated: High-value IP, defense-sponsored research (ITAR), and clinical healthcare data. Requires "
        "isolated enclaves, hardware keys, and envelope encryption.",
    )
    slides.append(s2)

    # =========================================================================
    # SLIDE 3 (L10): Data-Centric Protection: Cryptography & Database Defense
    # =========================================================================
    s3, _, _ = L10_three_cards(
        prs,
        title="Data-Centric Protection: Cryptography & Database Defense",
        cards_data=[
            (
                "CRYPTOGRAPHY",
                "Envelope Encryption & KMS",
                "Protects data across campus datacenters and multi-cloud repositories. Uses unique Data Encryption "
                "Keys (DEKs) wrapped by master Key Encryption Keys (KEKs) stored in FIPS 140-2 Level 3 HSMs. "
                "Enforces TLS 1.3 and mutual TLS (mTLS) across all data-in-flight communication paths.",
            ),
            (
                "DATABASE DEFENSE",
                "Micro-Perimeters & Query Proxies",
                "Databases storing student PII and research are placed behind Zero Trust database proxies. "
                "The proxy validates individual SQL queries, enforces fine-grained row- and column-level access controls, "
                "and blocks unauthorized direct database connections from compromised application servers.",
            ),
            (
                "EXFILTRATION SHIELD",
                "Dynamic Masking & Tokenization",
                "Sensitive student identifiers (SSNs, banking details, grades) are dynamically masked or tokenized based "
                "on user role. Teaching assistants view masked student IDs while registrars see full records. "
                "Rate-limiting halts mass SQL query dumps and automated exfiltration attempts.",
            ),
        ],
        part_label="Part 09: Data Security & Classification  ·  Slide 3/6",
        transition="fade",
    )
    set_notes(
        s3,
        "Slide 3 focuses on foundational data-centric technical controls: cryptography and database defense.\n\n"
        "First, Envelope Encryption: Rather than relying on a single static key, individual datasets use ephemeral Data "
        "Encryption Keys (DEKs) wrapped by Key Encryption Keys (KEKs) in hardware security modules.\n"
        "Second, Database Micro-Perimeters: Instead of granting direct network connectivity to a database server, applications "
        "must query through an intelligent Zero Trust proxy that inspects the SQL syntax and verifies row-level authorization.\n"
        "Third, Dynamic Masking & Tokenization: Mitigates breach impact by redacting sensitive fields in real time.",
    )
    slides.append(s3)

    # =========================================================================
    # SLIDE 4 (L09): Content Awareness: DLP & Persistent IRM
    # =========================================================================
    s4, _, _ = L09_two_columns(
        prs,
        title="Content Awareness: DLP & Persistent IRM",
        subtitle=(
            "Campus Data Loss Prevention gates network exit points, while Information Rights "
            "Management embeds cryptographic security directly inside files wherever they travel."
        ),
        left_col=(
            "Endpoint & Network DLP",
            [
                "Deep Content Inspection: Scans regex, hashes, and document fingerprints for FERPA identifiers, payment cards, and research formulas.",
                "Multi-Vector Egress Controls: Blocks unauthorized file uploads to personal cloud storage (Dropbox, personal Drive), USB drives, and unencrypted emails.",
                "Context-Sensitive Alerting: Triggers real-time educational coaching popups for users, administrative alerts, or immediate block actions upon violation.",
                "Clipboard & Print Quarantine: Restricts clipboard copy/paste, screen captures, and printing on unmanaged student BYOD devices.",
            ],
        ),
        right_col=(
            "Information Rights Management (IRM)",
            [
                "Persistent Cryptographic Binding: Security metadata and access policies are bound directly into the file binary (PDF, DOCX, datasets).",
                "Zero Trust In Transit & Off-Campus: Protection travels with the document even if emailed, copied to flash storage, or stored off-campus.",
                "Granular Rights Enforcement: Decouples file possession from usage; selectively revokes print, copy/paste, screen recording, and export rights.",
                "Instant Remote Revocation: Revokes access permissions immediately when a researcher leaves a grant or a student graduates, without deleting the file.",
            ],
        ),
        left_badge="EGRESS INSPECTION",
        right_badge="DATA-CENTRIC ENVELOPE",
        left_accent="SAGE_GLASS",
        right_accent="VERIFIED_MINT",
        part_label="Part 09: Data Security & Classification  ·  Slide 4/6",
        transition="fade",
    )
    set_notes(
        s4,
        "Slide 4 contrasts two complementary layers of data protection: DLP versus IRM.\n\n"
        "DLP acts as the border guard: It inspects data moving across endpoints, USB drives, emails, and cloud storage "
        "sync tools. If an employee tries to upload unencrypted FERPA records to a personal cloud drive, DLP blocks the transfer.\n\n"
        "IRM, on the other hand, is the embedded bodyguard: Even if a document leaves the campus network, the file itself "
        "remains an encrypted container. It calls home to the policy engine to check permissions before opening. If a student "
        "withdraws or a grant ends, access is revoked instantly.",
    )
    slides.append(s4)

    # =========================================================================
    # SLIDE 5 (L07): End-to-End Zero Trust Data Access Lifecycle
    # =========================================================================
    s5, _, _ = L07_flowchart_6node(
        prs,
        title="End-to-End Zero Trust Data Access Lifecycle",
        subtitle="Six sequential verification and enforcement checkpoints governing sensitive campus data interactions",
        steps=[
            (
                "Classification & Tagging",
                "Data is ingested, classified into tiers, and stamped with persistent cryptographic sensitivity tags at creation.",
            ),
            (
                "Context-Aware Request",
                "User requests data access; IdP authenticates persona (P2) while endpoint posture engine verifies device health (P3).",
            ),
            (
                "Dynamic Policy Evaluation",
                "Policy Engine (P4) evaluates sensitivity tag against user attributes, grant permissions, and environmental risk.",
            ),
            (
                "Key Release & Masking",
                "KMS releases ephemeral DEK to decrypt authorized records; database proxy applies dynamic masking to sensitive fields.",
            ),
            (
                "Monitored Interaction",
                "Endpoint DLP monitors live clipboard, print, and export actions; suspicious exfiltration attempts trigger instant quarantine.",
            ),
            (
                "Revocation & Audit Trail",
                "Ephemeral session token expires; full cryptographic access record is streamed into campus SIEM for continuous monitoring (P7).",
            ),
        ],
        part_label="Part 09: Data Security & Classification  ·  Slide 5/6",
        click_reveal=True,
        transition="fade",
    )
    set_notes(
        s5,
        "Slide 5 traces the full lifecycle of a sensitive data access request under Zero Trust.\n\n"
        "1. Classification: Data is tagged upon ingestion.\n"
        "2. Request: The user attempts access; identity and device health are verified.\n"
        "3. Policy Check: The Policy Engine assesses contextual risk.\n"
        "4. Decryption & Masking: Ephemeral keys decrypt only the necessary fields, applying masking.\n"
        "5. Monitored Session: DLP actively monitors for cut/paste or mass export.\n"
        "6. Revocation & Audit: Session keys expire, and full telemetry streams to the SIEM.",
    )
    slides.append(s5)

    # =========================================================================
    # SLIDE 6 (L04 Stage): Unifying Finale: The 9-Pillar Convergence on Data
    # =========================================================================
    s6, anim6, stage_bounds = L04_wide_visual(
        prs,
        title="The Zero Trust Finale: Convergence on Data",
        intro=(
            "Every preceding pillar—identity, endpoint posture, dynamic policy, network segmentation, "
            "secure remote access, telemetry monitoring, and phased adoption—exists to defend this single "
            "central asset: University Data."
        ),
        part_label="Part 09: Data Security & Classification  ·  Slide 6/6",
        stage_card=True,
        transition="fade",
    )

    # Custom On-Theme Visual Inside Stage Bounds
    # Left Side: 4 Concentric Defense Layers + Core Data Bullseye (width: 9.8 in)
    # Right Side: 2 Frosted Synthesis Cards (width: 6.9 in)
    sx, sy, sw, sh = stage_bounds

    # 1. Concentric Defense Rings / Horizontal Stack on Left
    left_x = sx + 0.35
    left_w = 10.2
    top_y = sy + 0.35

    pillars_stack = [
        (
            "PILLARS 02 & 03  ·  FOUNDATION",
            "Verified Identity & Endpoint Health",
            "Phishing-resistant MFA, FIDO2 credentials, and continuous MDM/EDR posture validation",
            "SAGE_GLASS",
            16.0,
        ),
        (
            "PILLARS 04 & 05  ·  BOUNDARY",
            "Dynamic Policy & Campus Micro-Segmentation",
            "Context-aware PBAC rules, zone containment (labs, dorms, ERP), and east-west traffic blocks",
            "MUTED_TEAL",
            22.0,
        ),
        (
            "PILLARS 06 & 07  ·  ENFORCEMENT",
            "Secure Remote ZTNA & Closed-Loop Telemetry",
            "Dark application proxies (no open ports), outbound connectors, and continuous UEBA/SIEM feedback",
            "DEEP_PINE",
            30.0,
        ),
        (
            "PILLAR 08  ·  GOVERNANCE",
            "Phased Maturity & Brownfield Governance",
            "CISA Zero Trust maturity alignment, stakeholder buy-in, and continuous operational verification",
            "DARK_TEAL",
            38.0,
        ),
    ]

    layer_cards = []
    layer_h = 0.98
    gap_layer = 0.16
    for idx, (p_tag, p_head, p_sub, b_color, alpha_val) in enumerate(pillars_stack):
        ly = top_y + idx * (layer_h + gap_layer)
        c_shape = add_frosted_card(
            s6,
            left_x,
            ly,
            left_w,
            layer_h,
            fill_alpha=alpha_val,
            border_color=b_color,
            border_alpha=55.0,
            radius=0.04,
            name=f"DefenseLayer_{idx+1}",
        )
        tf = c_shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.24)
        tf.margin_right = Inches(0.24)
        tf.margin_top = Inches(0.06)
        tf.margin_bottom = Inches(0.06)

        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.LEFT
        p1.line_spacing = 1.05
        r_tag = p1.add_run()
        r_tag.text = f"{p_tag}  ▸  "
        r_tag.font.name = FONT_BODY
        r_tag.font.size = Pt(13.5)
        r_tag.font.bold = True
        r_tag.font.color.rgb = rgb("SOFT_SAGE")

        r_head = p1.add_run()
        r_head.text = p_head
        r_head.font.name = FONT_TITLE
        r_head.font.size = Pt(17.0)
        r_head.font.bold = True
        r_head.font.color.rgb = rgb("WHITE")

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        p2.line_spacing = 1.08
        r_sub = p2.add_run()
        r_sub.text = p_sub
        r_sub.font.name = FONT_BODY
        r_sub.font.size = Pt(14.0)
        r_sub.font.color.rgb = rgb("ICE_MIST")

        layer_cards.append(c_shape)

    # Core Protect Surface Box (The Bullseye at Bottom Left)
    core_y = top_y + 4 * (layer_h + gap_layer) + 0.08
    core_h = 1.48
    core_box = add_frosted_card(
        s6,
        left_x,
        core_y,
        left_w,
        core_h,
        fill_color="DARK_TEAL",
        fill_alpha=92.0,
        border_color="VERIFIED_MINT",
        border_alpha=85.0,
        border_pt=1.5,
        radius=0.06,
        name="CoreDataBullseye",
    )
    c_tf = core_box.text_frame
    c_tf.word_wrap = True
    c_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    c_tf.margin_left = Inches(0.28)
    c_tf.margin_right = Inches(0.28)
    c_tf.margin_top = Inches(0.10)
    c_tf.margin_bottom = Inches(0.10)

    cp1 = c_tf.paragraphs[0]
    cp1.alignment = PP_ALIGN.LEFT
    cr1 = cp1.add_run()
    cr1.text = "★  THE CORE PROTECT SURFACE  ·  PART 09: UNIVERSITY DATA"
    cr1.font.name = FONT_BODY
    cr1.font.size = Pt(14.5)
    cr1.font.bold = True
    cr1.font.color.rgb = rgb("VERIFIED_MINT")

    cp2 = c_tf.add_paragraph()
    cp2.alignment = PP_ALIGN.LEFT
    cp2.space_before = Pt(4)
    cp2.line_spacing = 1.14
    cr2 = cp2.add_run()
    cr2.text = (
        "Classified Academic Records (FERPA)  ·  Proprietary Research & IP  ·  Clinical Patient Data (HIPAA)  ·  "
        "Financial & Payroll Systems  ·  Encrypted Keys & Credentials"
    )
    cr2.font.name = FONT_TITLE
    cr2.font.size = Pt(16.0)
    cr2.font.bold = True
    cr2.font.color.rgb = rgb("WHITE")

    # Right Side: 2 Synthesis Takeaway Cards (clean vertical fit inside stage!)
    right_x = left_x + left_w + 0.42
    right_w = sw - (left_w + 1.12)
    card_r1_h = 2.96
    card_r2_h = 2.96

    card_r1 = add_frosted_card(
        s6,
        right_x,
        top_y,
        right_w,
        card_r1_h,
        fill_alpha=26.0,
        border_color="SAGE_GLASS",
        border_alpha=40.0,
        radius=0.05,
        name="SynthesisCard_1",
    )
    tf_r1 = card_r1.text_frame
    tf_r1.word_wrap = True
    tf_r1.vertical_anchor = MSO_ANCHOR.TOP
    tf_r1.margin_left = Inches(0.28)
    tf_r1.margin_right = Inches(0.28)
    tf_r1.margin_top = Inches(0.22)
    tf_r1.margin_bottom = Inches(0.16)

    pr1_head = tf_r1.paragraphs[0]
    pr1_head.alignment = PP_ALIGN.LEFT
    pr1_head.space_after = Pt(6)
    run_r1_h = pr1_head.add_run()
    run_r1_h.text = "Never Trust the Path, Protect the Payload"
    run_r1_h.font.name = FONT_TITLE
    run_r1_h.font.size = Pt(19.5)
    run_r1_h.font.bold = True
    run_r1_h.font.color.rgb = rgb("WHITE")

    pr1_body = tf_r1.add_paragraph()
    pr1_body.alignment = PP_ALIGN.LEFT
    pr1_body.line_spacing = 1.18
    run_r1_b = pr1_body.add_run()
    run_r1_b.text = (
        "In an open campus network, perimeters will inevitably be breached, credentials will be phished, "
        "and untrusted student laptops will connect to Wi-Fi. By enforcing envelope encryption, dynamic masking, "
        "and persistent IRM metadata, university data remains completely undecipherable and unexfiltratable even "
        "if an attacker breaches the network segment."
    )
    run_r1_b.font.name = FONT_BODY
    run_r1_b.font.size = Pt(15.5)
    run_r1_b.font.color.rgb = rgb("ICE_MIST")

    card_r2 = add_frosted_card(
        s6,
        right_x,
        top_y + card_r1_h + 0.18,
        right_w,
        card_r2_h,
        fill_alpha=26.0,
        border_color="VERIFIED_MINT",
        border_alpha=45.0,
        radius=0.05,
        name="SynthesisCard_2",
    )
    tf_r2 = card_r2.text_frame
    tf_r2.word_wrap = True
    tf_r2.vertical_anchor = MSO_ANCHOR.TOP
    tf_r2.margin_left = Inches(0.28)
    tf_r2.margin_right = Inches(0.28)
    tf_r2.margin_top = Inches(0.22)
    tf_r2.margin_bottom = Inches(0.16)

    pr2_head = tf_r2.paragraphs[0]
    pr2_head.alignment = PP_ALIGN.LEFT
    pr2_head.space_after = Pt(6)
    run_r2_h = pr2_head.add_run()
    run_r2_h.text = "The Closed-Loop Assurance Architecture"
    run_r2_h.font.name = FONT_TITLE
    run_r2_h.font.size = Pt(19.5)
    run_r2_h.font.bold = True
    run_r2_h.font.color.rgb = rgb("WHITE")

    pr2_body = tf_r2.add_paragraph()
    pr2_body.alignment = PP_ALIGN.LEFT
    pr2_body.line_spacing = 1.18
    run_r2_b = pr2_body.add_run()
    run_r2_b.text = (
        "Zero Trust is not a static gate but a continuous feedback loop. Every data access event feeds telemetry "
        "to the SIEM/UEBA engine (P7). Anomalous download volumes or off-hours queries dynamically increase the "
        "session risk score, instantly revoking cryptographic keys and isolating endpoints before data can be stolen."
    )
    run_r2_b.font.name = FONT_BODY
    run_r2_b.font.size = Pt(15.5)
    run_r2_b.font.color.rgb = rgb("ICE_MIST")

    # Queue animations for custom stage elements
    anim6.stagger(layer_cards, effect="fly_left", step_ms=110, on_click_first=False, base_delay_ms=80)
    anim6.add(core_box, effect="zoom", trigger="after_previous", duration_ms=450, delay_ms=100)
    anim6.add(card_r1, effect="fly_right", trigger="after_previous", duration_ms=400, delay_ms=80)
    anim6.add(card_r2, effect="fly_right", trigger="after_previous", duration_ms=400, delay_ms=80)
    anim6.apply()

    set_notes(
        s6,
        "Slide 6 represents the technical climax and synthesis of our entire presentation: The Convergence on Data.\n\n"
        "Notice how the concentric layers work together:\n"
        "- Outer Foundation: Verified identities and healthy endpoints (Parts 02 & 03).\n"
        "- Policy & Micro-segmentation: Dynamic evaluation and network zone containment (Parts 04 & 05).\n"
        "- Enforcement & Monitoring: Dark ZTNA proxies and closed-loop telemetry feedback (Parts 06 & 07).\n"
        "- Governance: Phased maturity models ensuring operational sustainability (Part 08).\n\n"
        "At the center of it all is University Data. We do not trust the network path; we protect the data payload. "
        "This concludes Part 09 and seamlessly leads into our team's unified conclusion and open Q&A in Part 10.",
    )
    slides.append(s6)

    return slides
