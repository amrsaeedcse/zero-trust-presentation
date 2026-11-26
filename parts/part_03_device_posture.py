"""Part 03 — Device Security Posture.

Owner: Omar Mounir Mohamed Arafa
Slide Budget: 7 slides (including opening section divider; hard cap <= 7).

In Scope:
- Campus endpoint heterogeneity (Tier 1 staff PCs, Tier 2 shared labs, Tier 3 student BYOD, Tier 4 IoT & unknown hosts).
- Multi-signal health checks (OS/app patching, EDR health, BitLocker/host config, Secure Boot & TPM) and Tri-State classification.
- Cryptographic device certificates, tiered freshness windows (15-min exam portals), and mandatory alert gates.
- Privacy-safe student BYOD vs. disposable lab VMs, passive IoT controls, and expiring exceptions.
- Closed-loop Quarantine VLAN repair, live mid-session reassessment, and 3-phase campus pilot (90% fresh coverage).

Narrative Bridge:
- Builds on Part 02 (verified user identity) -> Bridges to Part 04 (Authorization Policies).
"""

from __future__ import annotations

import re
from typing import Sequence

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import (
    L03_section_divider,
    L07_flowchart_6node,
    L09_two_columns,
    _style_table_cell,
)
from core.theme import (
    COLORS,
    FONT_BODY,
    FONT_TITLE,
    MARGIN_X,
    _set_fill_alpha,
    add_arrow_connector,
    add_badge,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_motif,
    add_textbox,
    rgb,
)

PART_LABEL = "Part 03  ·  Device Security Posture"
_EN_SPAN_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9\.\-/% ]*[A-Za-z0-9%]|[A-Za-z0-9%]")


def _set_rtl_notes(slide, lines: Sequence[str]) -> None:
    """Write native RTL speaker notes with isolated ar-EG / en-US runs so mixed English terms never scramble."""
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.clear()
    tf.word_wrap = True

    for idx, raw_line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.RIGHT
        p.space_after = Pt(8)
        pPr = p._p.get_or_add_pPr()
        pPr.set("rtl", "1")
        pPr.set("algn", "r")

        line = "\u200f• " + raw_line.strip()
        pos = 0
        for m in _EN_SPAN_RE.finditer(line):
            if m.start() > pos:
                r_ar = p.add_run()
                r_ar.text = line[pos:m.start()]
                r_ar.font.size = Pt(16)
                r_ar._r.get_or_add_rPr().set("lang", "ar-EG")

            en_chunk = m.group(0)
            r_en = p.add_run()
            r_en.text = f"\u200e{en_chunk}\u200e\u200f"
            r_en.font.size = Pt(16)
            r_en._r.get_or_add_rPr().set("lang", "en-US")
            pos = m.end()

        if pos < len(line):
            r_tail = p.add_run()
            r_tail.text = line[pos:]
            r_tail.font.size = Pt(16)
            r_tail._r.get_or_add_rPr().set("lang", "ar-EG")


def _add_card_heading_body(
    card,
    heading: str,
    body: str,
    head_size: float = 25.0,
    body_size: float = 19.0,
    head_color: str = "WHITE",
    body_color: str = "ICE_MIST",
    top_margin_in: float = 0.24,
) -> None:
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.36)
    tf.margin_right = Inches(0.34)
    tf.margin_top = Inches(top_margin_in)
    tf.margin_bottom = Inches(0.22)

    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    p0.space_after = Pt(7)
    p0.line_spacing = 1.12
    r0 = p0.add_run()
    r0.text = heading
    r0.font.name = FONT_TITLE
    r0.font.size = Pt(head_size)
    r0.font.bold = True
    r0.font.color.rgb = rgb(head_color)

    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.LEFT
    p1.line_spacing = 1.22
    r1 = p1.add_run()
    r1.text = body
    r1.font.name = FONT_BODY
    r1.font.size = Pt(body_size)
    r1.font.color.rgb = rgb(body_color)


def _slide_2_device_tiers(prs):
    """Slide 2: Why Identity Is Not Enough + Four Campus Endpoint Tiers (2x2 + Scenario Bar)."""
    slide = add_blank_slide(prs, bg="vertical")
    motif = add_motif(slide, "signal_waves", left=15.35, top=-1.05, width=5.3)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.88,
        width=20.0 - 2 * MARGIN_X,
        height=0.95,
        text="Four Campus Device Tiers Need Distinct Baselines",
        font_name=FONT_TITLE,
        font_size=52.0,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="TiersTitle",
    )
    sub = add_textbox(
        slide,
        left=1.8,
        top=1.88,
        width=16.4,
        height=0.50,
        text="Network location never proves safety — valid credentials verify the user, not endpoint patch or defense health",
        font_name=FONT_BODY,
        font_size=21.0,
        color="ICE_MIST",
        align=PP_ALIGN.CENTER,
        name="TiersSubtitle",
    )

    tiers = [
        (
            "TIER 01  ·  FULL UEM + EDR",
            "VERIFIED_MINT",
            "Managed Staff & Admin PCs",
            "Handle grading records and exam files; enforce UEM enrollment, BitLocker encryption, and active EDR.",
        ),
        (
            "TIER 02  ·  SHARED LABS",
            "SAGE_GLASS",
            "Shared Teaching Laboratories",
            "Run course software across many students; reconcile live network scans with department asset registers.",
        ),
        (
            "TIER 03  ·  SCOPED PRIVACY",
            "VERIFIED_MINT",
            "Student BYOD Phones & Laptops",
            "Personal devices connect with limited visibility; route learning access through restricted browser paths.",
        ),
        (
            "TIER 04  ·  OWNER TRACKED",
            "SAGE_GLASS",
            "Campus IoT & Unknown Hosts",
            "Printers, cameras, and lab tools tracked by owner; unregistered hosts receive only guest or ID access.",
        ),
    ]

    cw, ch = 8.58, 2.62
    gx, gy = 0.60, 0.38
    sx, sy = MARGIN_X, 2.58

    tier_groups = []
    for idx, (btxt, accent, heading, desc) in enumerate(tiers):
        r, c = divmod(idx, 2)
        x = sx + c * (cw + gx)
        y = sy + r * (ch + gy)
        card = add_frosted_card(
            slide,
            x,
            y,
            cw,
            ch,
            fill_color="DARK_TEAL",
            fill_alpha=48.0,
            border_color=accent,
            border_alpha=58.0,
            border_pt=1.15,
            radius=0.05,
            name=f"TierCard_{idx+1}",
        )
        _add_card_heading_body(
            card,
            heading=heading,
            body=desc,
            head_size=25.0,
            body_size=18.5,
            head_color="WHITE",
            body_color="ICE_MIST",
            top_margin_in=0.82,
        )
        badge = add_badge(
            slide,
            left=x + 0.36,
            top=y + 0.22,
            width=3.80,
            height=0.46,
            text=btxt,
            fill_color=accent,
            fill_alpha=24.0,
            border_color=accent,
            border_alpha=70.0,
            text_color="WHITE",
            font_size=15.0,
            name=f"TierBadge_{idx+1}",
        )
        tier_groups.extend([card, badge])

    # Lecturer exam-leak callout bar
    callout = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=8.62,
        width=20.0 - 2 * MARGIN_X,
        height=1.42,
        fill_color="DARK_TEAL",
        fill_alpha=68.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        border_pt=1.25,
        radius=0.06,
        name="LecturerScenarioBar",
    )
    sc_badge = add_badge(
        slide,
        left=MARGIN_X + 0.34,
        top=9.05,
        width=3.15,
        height=0.56,
        text="LECTURER SCENARIO",
        fill_color="ALERT_CORAL",
        fill_alpha=32.0,
        border_color="ALERT_CORAL",
        border_alpha=82.0,
        text_color="WHITE",
        font_size=16.0,
        name="ScenarioBadge",
    )
    sc_text = add_textbox(
        slide,
        left=MARGIN_X + 3.74,
        top=8.78,
        width=13.68,
        height=1.10,
        text=(
            "A lecturer signing in with valid credentials from an unprotected laptop with disabled security "
            "exposes exam papers — posture checks block risky endpoints before sensitive records open."
        ),
        font_name=FONT_BODY,
        font_size=20.0,
        color="WHITE",
        valign=MSO_ANCHOR.MIDDLE,
        line_spacing=1.22,
        name="ScenarioText",
    )

    add_footer(slide, PART_LABEL, slide_num=2)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(motif, trigger="on_load", duration_ms=420)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.stagger(tier_groups, effect="fly_bottom", step_ms=75, on_click_first=False)
    anim.fly_in(callout, direction="bottom", trigger="after_previous", duration_ms=360)
    anim.fade_in(sc_badge, trigger="with_previous", duration_ms=300)
    anim.fade_in(sc_text, trigger="with_previous", duration_ms=300, delay_ms=60)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "المبدأ الأساسي إن وجود اللابتوب جوة الكلية مش دليل إنه آمن؛ كلمة السر والـ MFA بيثبتوا هوية المستخدم بس، مش إن جهازه سليم.",
            "مثال عملي: لو دكتور دخل بحساب سليم يرصد درجات من لابتوب الحماية فيه مقفولة أو فيه ثغرة، أوراق الامتحانات هتكون في خطر تسريب فوري.",
            "عشان كده قسمنا أجهزة الجامعة لـ 4 فئات لكل واحدة Baseline مناسب:",
            "الأولى أجهزة الدكاترة والإدارة بتحكم كامل وتشفير BitLocker وحماية EDR، والتانية أجهزة المعامل المشتركة بمطابقة فحص الشبكة مع السجلات.",
            "التالتة موبايلات ولابتوبات الطلبة BYOD بمسار متصفح محدود يحمي خصوصيتهم، والرابعة أجهزة الـ IoT والأبحاث مربوطة بمسؤول محدد، وأي جهاز مجهول بيدخل Guest بس.",
        ],
    )
    return slide


def _slide_3_signals_and_states(prs):
    """Slide 3: Multi-Signal Assessment Table + Semantic Tri-State Classification Cards."""
    slide = add_blank_slide(prs, bg="diagonal")
    glow = add_motif(slide, "glow", left=14.4, top=-3.2, width=8.8)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.88,
        width=9.6,
        height=1.85,
        text="Multi-Signal Checks &\nTri-State Verdict",
        font_name=FONT_TITLE,
        font_size=50.0,
        color="WHITE",
        bold=True,
        line_spacing=1.06,
        name="SignalsTitle",
    )
    desc = add_textbox(
        slide,
        left=11.10,
        top=1.02,
        width=7.78,
        height=1.60,
        text=(
            "Four telemetry layers evaluate software, active defense, storage encryption, "
            "and boot integrity to classify every endpoint into three explicit states."
        ),
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        line_spacing=1.24,
        name="SignalsIntro",
    )

    headers = ["Signal Category", "What the Assessment Verifies", "Specific Risk Addressed"]
    rows = [
        ["OS & App Patching", "Supported OS build, browser, PDF reader & firmware fixes", "Exploitable software & reader flaws"],
        ["EDR Agent Health", "Active sensor service, recent check-in & zero high alerts", "Disabled protection or active threat"],
        ["Host Configuration", "Host firewall, auto screen lock, BitLocker & least privilege", "Data exposure on lost/stolen laptops"],
        ["Platform Integrity", "UEFI Secure Boot, TPM hardware attestation & root/jailbreak", "Boot-level tampering or rooting"],
    ]
    col_widths = [3.65, 8.35, 5.76]
    tbl_w = sum(col_widths)
    tbl_h = 4.65
    tbl_shape = slide.shapes.add_table(5, 3, Inches(MARGIN_X), Inches(2.90), Inches(tbl_w), Inches(tbl_h))
    tbl_shape.name = "SignalsTable"
    table = tbl_shape.table
    for ci, w_in in enumerate(col_widths):
        table.columns[ci].width = Inches(w_in)

    for ci, h_txt in enumerate(headers):
        _style_table_cell(
            table.cell(0, ci),
            h_txt,
            fill_hex=COLORS["TABLE_HEADER"],
            text_hex=COLORS["WHITE"],
            bold=True,
            font_size=21.0,
            align=PP_ALIGN.CENTER,
        )
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            _style_table_cell(
                table.cell(ri, ci),
                val,
                fill_hex=COLORS["SOFT_SAGE"],
                text_hex=COLORS["BLACK"],
                bold=(ci == 0),
                font_size=19.5,
                align=PP_ALIGN.CENTER if ci == 0 else PP_ALIGN.LEFT,
            )

    # Bottom 3 Tri-State Result Cards + Golden Rule
    states = [
        (
            "STATE 1  ·  COMPLIANT",
            "VERIFIED_MINT",
            "Meets target baseline with fresh evidence → eligible for access.",
        ),
        (
            "STATE 2  ·  NONCOMPLIANT",
            "ALERT_CORAL",
            "Failed check or active alert → block sensitive action & repair.",
        ),
        (
            "STATE 3  ·  UNKNOWN (GOLDEN RULE)",
            "ICE_MIST",
            "Missing, stale, or unsupported data → NEVER treat Unknown as Compliant.",
        ),
    ]
    sw, sh, sgap = 5.68, 2.18, 0.36
    sy = 7.95
    state_cards = []
    for idx, (st_head, accent, st_body) in enumerate(states):
        sx = MARGIN_X + idx * (sw + sgap)
        sc = add_frosted_card(
            slide,
            sx,
            sy,
            sw,
            sh,
            fill_color="DARK_TEAL",
            fill_alpha=60.0 if idx == 2 else 46.0,
            border_color=accent,
            border_alpha=78.0,
            border_pt=1.35,
            radius=0.06,
            name=f"StateCard_{idx+1}",
        )
        _add_card_heading_body(
            sc,
            heading=st_head,
            body=st_body,
            head_size=20.5,
            body_size=18.5,
            head_color=accent,
            body_color="WHITE",
            top_margin_in=0.28,
        )
        state_cards.append(sc)

    add_footer(slide, PART_LABEL, slide_num=3)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(glow, trigger="on_load", duration_ms=380)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    anim.fade_in(desc, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.fly_in(tbl_shape, direction="bottom", trigger="after_previous", duration_ms=420)
    anim.stagger(state_cards, effect="zoom", step_ms=130, on_click_first=False)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "بنفحص الجهاز من خلال 4 طبقات إشارات متكاملة مع بعض:",
            "أولاً تحديثات الويندوز والمتصفح وقارئ الـ PDF لمنع الثغرات، وتانياً صحة عميل الـ EDR ونبضه اللحظي وخلوه من الإنذارات الخطيرة.",
            "تالتاً إعدادات الأمان وتشفير الهارد بالـ BitLocker والـ Firewall، ورابعاً سلامة الإقلاع بالـ Secure Boot وشريحة الـ TPM وخلو الموبايل من الـ Root.",
            "والفحص بيصنف الجهاز لـ 3 حالات: Compliant محقق الشروط وجاهز للصلاحية، أو Noncompliant يتمنع عن السيستم الحساس ويتحول للإصلاح.",
            "والحالة التالتة Unknown، والقاعدة الذهبية عندنا هي Never treat Unknown as Compliant — لو البيانات قديمة أو مش واصلة بنعتبره غير آمن فوراً.",
        ],
    )
    return slide


def _slide_4_certificates_and_gates(prs):
    """Slide 4: Anti-Spoofing Device Certificates, Tiered Freshness & Mandatory Gates Diagram."""
    slide = add_blank_slide(prs, bg="radial")
    prism = add_motif(slide, "prism", left=15.9, top=-1.2, width=4.8)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.88,
        width=17.76,
        height=0.95,
        text="Cryptographic Certificates, Freshness & Mandatory Gates",
        font_name=FONT_TITLE,
        font_size=48.0,
        color="WHITE",
        bold=True,
        name="CertsTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.88,
        width=16.8,
        height=0.50,
        text="Authenticated telemetry links and TPM-bound keys prevent MAC spoofing and forged compliance records",
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        name="CertsSubtitle",
    )

    # Column 1: 3 Evidence Sources (x=1.12, w=5.05)
    col1_x, col1_w = MARGIN_X, 5.05
    sources = [
        ("1 · Multi-Source Telemetry", "Timestamped UEM, MDM, EDR alerts, and scans over authenticated links."),
        ("2 · X.509 Device Certificates", "Bind posture to cryptographic TPM keys rather than spoofable MACs."),
        ("3 · Signed Audit Channels", "Versioned compliance logs stop attackers forging endpoint health."),
    ]
    src_cards = []
    src_arrows = []
    for i, (st, sb) in enumerate(sources):
        sy = 2.75 + i * (2.15 + 0.34)
        c = add_frosted_card(
            slide,
            col1_x,
            sy,
            col1_w,
            2.15,
            fill_color="DARK_TEAL",
            fill_alpha=52.0,
            radius=0.06,
            name=f"SrcCard_{i+1}",
        )
        _add_card_heading_body(c, st, sb, head_size=22.0, body_size=18.0, top_margin_in=0.28)
        src_cards.append(c)
        cy = sy + 1.075
        arr = add_arrow_connector(slide, col1_x + col1_w + 0.10, cy, 7.02, cy, color="WHITE", name=f"ArrSrc_{i+1}")
        src_arrows.append(arr)

    # Column 2: Central Posture Verifier Hub (x=7.12, y=2.75, w=5.65, h=7.13)
    hub_x, hub_y, hub_w, hub_h = 7.12, 2.75, 5.65, 7.13
    hub = add_frosted_card(
        slide,
        hub_x,
        hub_y,
        hub_w,
        hub_h,
        fill_color="DARK_TEAL",
        fill_alpha=62.0,
        border_color="VERIFIED_MINT",
        border_alpha=70.0,
        border_pt=1.3,
        radius=0.05,
        name="PostureHub",
    )
    hub_badge = add_badge(
        slide,
        left=hub_x + 0.36,
        top=hub_y + 0.28,
        width=4.93,
        height=0.50,
        text="POSTURE VERIFICATION & GATE RULES",
        fill_color="VERIFIED_MINT",
        fill_alpha=24.0,
        border_color="VERIFIED_MINT",
        border_alpha=72.0,
        font_size=15.0,
        name="HubBadge",
    )

    fresh_box = add_frosted_card(
        slide,
        hub_x + 0.32,
        hub_y + 1.02,
        hub_w - 0.64,
        2.78,
        fill_alpha=20.0,
        border_color="SAGE_GLASS",
        border_alpha=55.0,
        radius=0.06,
        name="FreshnessCard",
    )
    _add_card_heading_body(
        fresh_box,
        heading="Tiered Freshness SLA",
        body=(
            "▸ Exam & grading portals: ≤ 15m\n"
            "▸ Coursework LMS: wider window\n"
            "▸ Offline hosts revert to Unknown"
        ),
        head_size=22.0,
        body_size=18.0,
        head_color="VERIFIED_MINT",
        body_color="WHITE",
        top_margin_in=0.24,
    )

    gate_box = add_frosted_card(
        slide,
        hub_x + 0.32,
        hub_y + 4.04,
        hub_w - 0.64,
        2.78,
        fill_alpha=20.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        border_pt=1.2,
        radius=0.06,
        name="MandatoryGateCard",
    )
    _add_card_heading_body(
        gate_box,
        heading="Mandatory Gate Rule",
        body=(
            "Minor passes never mask an active high-severity alert — "
            "critical threats act as mandatory binary gates."
        ),
        head_size=22.0,
        body_size=18.0,
        head_color="ALERT_CORAL",
        body_color="WHITE",
        top_margin_in=0.24,
    )

    mid_arrow = add_arrow_connector(slide, hub_x + hub_w + 0.10, 6.31, 13.65, 6.31, color="VERIFIED_MINT", width_pt=2.0, name="ArrHubOut")

    # Column 3: 4 Enforcement Outcomes Panel (x=13.75, y=2.75, w=5.13, h=7.13)
    out_x, out_w = 13.75, 5.13
    out_panel = add_frosted_card(
        slide,
        out_x,
        hub_y,
        out_w,
        hub_h,
        fill_color="DARK_TEAL",
        fill_alpha=60.0,
        radius=0.05,
        name="OutcomesPanel",
    )
    out_head = add_textbox(
        slide,
        left=out_x + 0.34,
        top=hub_y + 0.30,
        width=out_w - 0.68,
        height=0.92,
        text="4 Posture-Driven\nAccess Outcomes",
        font_name=FONT_TITLE,
        font_size=25.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="OutcomesHeader",
    )
    outcomes = [
        ("1 · Normal Access", "Compliant managed PC + MFA", "VERIFIED_MINT"),
        ("2 · Limited Session", "Scoped browser / VDI route", "ICE_MIST"),
        ("3 · Repair VLAN", "802.1X quarantine for updates", "SOFT_SAGE"),
        ("4 · Full Isolation", "Contain active compromise", "ALERT_CORAL"),
    ]
    out_items = []
    for j, (ot, od, ocol) in enumerate(outcomes):
        oy = hub_y + 1.46 + j * (1.22 + 0.16)
        ocard = add_frosted_card(
            slide,
            out_x + 0.28,
            oy,
            out_w - 0.56,
            1.22,
            fill_color="OBSIDIAN",
            fill_alpha=38.0,
            border_color=ocol,
            border_alpha=72.0,
            border_pt=1.1,
            radius=0.08,
            name=f"OutcomeRow_{j+1}",
        )
        _add_card_heading_body(
            ocard,
            heading=ot,
            body=od,
            head_size=19.5,
            body_size=17.5,
            head_color=ocol,
            body_color="WHITE",
            top_margin_in=0.16,
        )
        out_items.append(ocard)

    add_footer(slide, PART_LABEL, slide_num=4)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(prism, trigger="on_load", duration_ms=420)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.stagger(src_cards, effect="fly_left", step_ms=110, on_click_first=False)
    anim.stagger(src_arrows, effect="wipe_right", step_ms=70, on_click_first=False)
    anim.zoom_in(hub, trigger="after_previous", duration_ms=360)
    anim.fade_in(hub_badge, trigger="with_previous", duration_ms=280)
    anim.fade_in(fresh_box, trigger="with_previous", duration_ms=320, delay_ms=80)
    anim.fade_in(gate_box, trigger="with_previous", duration_ms=320, delay_ms=160)
    anim.wipe_in(mid_arrow, direction="right", trigger="after_previous", duration_ms=300)
    anim.fade_in(out_panel, trigger="with_previous", duration_ms=320)
    anim.fade_in(out_head, trigger="with_previous", duration_ms=300)
    anim.stagger(out_items, effect="fly_right", step_ms=95, on_click_first=False)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "عشان نمنع التزوير، مبنعتمدش على الـ MAC Address لأنه سهل يتزور، لكن بنربط حالة الجهاز بشهادات رقمية X.509 مشفرة جوة شريحة الـ TPM.",
            "كمان بنحدد صلاحية زمنية للأدلة بمبدأ الـ Tiered Freshness: بوابات رصد الدرجات والامتحانات لازم فحص طازة خلال آخر 15 دقيقة، بينما الكورسات العادية بتسمح بوقت أطول.",
            "الأدلة دي بتخضع لقاعدة الـ Mandatory Gate: لو فيه إنذار خطر نشط بيمنع الدخول فوراً حتى لو الفحوصات البسيطة سليمة.",
            "ومع معيار 802.1X بيطلع واحد من 4 مسارات: وصول كامل، أو جلسة محدودة بالمتصفح، أو Repair VLAN للتحديث، أو عزل تام.",
        ],
    )
    return slide


def _slide_5_byod_and_labs(prs):
    """Slide 5: Student BYOD Privacy vs. Shared Labs & Exceptional Research Gear (L09)."""
    slide, _, cards = L09_two_columns(
        prs,
        title="BYOD Privacy &\nLab Exceptions",
        subtitle=(
            "Balancing privacy-conscious student access against multi-user teaching labs "
            "and fragile scientific equipment that cannot run endpoint agents."
        ),
        left_col=(
            "Privacy-Safe Student Access",
            [
                "Service Tiers: Public pages require no enrollment; sensitive admin tasks require managed PCs.",
                "Protected Sessions: Use restricted browsers, VDI streams, or mobile work profiles isolating personal data.",
                "Lab Alternatives: Campus computer labs and loaner PCs keep coursework open for older student phones.",
            ],
        ),
        right_col=(
            "Shared Labs & Legacy Gear",
            [
                "Disposable Lab VMs: Isolate student experiments in resettable VMs while patching golden host images.",
                "Agentless Sensors: Protect printers and fragile lab instruments via firewalls and passive scans.",
                "Expiring Exceptions: Record compensating controls, accountable owners, and mandatory review dates.",
            ],
        ),
        left_badge="STUDENT BYOD",
        right_badge="LABS & LEGACY IOT",
        left_accent="VERIFIED_MINT",
        right_accent="SAGE_GLASS",
        part_label=f"{PART_LABEL}  ·  05",
        transition="fade",
    )
    for col_card in cards:
        _set_fill_alpha(col_card, "DARK_TEAL", 58.0)
        col_card.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

    _set_rtl_notes(
        slide,
        [
            "بالنسبة لأجهزة الطلبة الشخصية BYOD: مينفعش ننتهك خصوصيتهم ونفرض سيطرة كاملة على موبايلاتهم.",
            "عشان كده الصفحات العامة مفتوحة بدون تسجيل، والخدمات الحساسة بنشغلها بمتصفح مقيد أو VDI أو Work Profile يفصل ملفات الكلية عن صورهم الشخصية، مع توفير معامل ولابتوبات إعارة للأجهزة القديمة.",
            "أما معامل الكلية المشتركة فبنشغل تجارب الطلبة على Disposable VMs بيتعملها Reset بعد كل سكشن مع تحديث النسخة الأصلية Golden Image.",
            "وأجهزة الأبحاث والـ IoT اللي مبينزلش عليها Agent بنحميها بـ Passive Scans وجدران نارية، وأي استثناء بيتسجل بمسؤول وتاريخ انتهاء للمراجعة.",
        ],
    )
    return slide


def _slide_6_quarantine_loop(prs):
    """Slide 6: Closed-Loop Quarantine Repair & Live Reassessment (L07 6-Node Serpentine)."""
    slide, _, _ = L07_flowchart_6node(
        prs,
        title="Quarantine Repair & Live Reassessment Loop",
        subtitle="Linking posture failure to isolated repair, verified patching, and mid-session revocation",
        steps=[
            ("Detect Failure", "Outdated OS/reader, stopped EDR, or stale check triggers block"),
            ("Quarantine VLAN", "Isolate host via 802.1X; block internal records & lateral paths"),
            ("Scoped Repair", "Allow patch mirrors, DNS, NTP & support only"),
            ("Test & Verify Fix", "Test lab compatibility & confirm fixes are applied, not just downloaded"),
            ("Fresh Re-Check", "Re-verify device cert and live posture before restoring access"),
            ("Live Session Watch", "Mid-session severe alert revokes access & triggers incident reset"),
        ],
        part_label=f"{PART_LABEL}  ·  06",
        click_reveal=False,
        transition="fade",
    )
    _set_rtl_notes(
        slide,
        [
            "اكتشاف جهاز غير متوافق ملوش قيمة بدون مسار إصلاح مغلق؛ أول ما الفحص يفشل بنعزل الجهاز في Quarantine VLAN تمنعه من درجات وسجلات الكلية.",
            "الشبكة دي بتسمح له يوصل بس لسيرفرات التحديث والـ DNS ومزامنة الوقت والدعم الفني عشان يصلح المشكلة.",
            "وبنتأكد إن الـ Patch اتثبت فعلياً مش مجرد اتحمل قبل ما نعيد فحص الشهادة والحالة ونرجع الصلاحية.",
            "والمراقبة شغالة لحظياً وسط الجلسة بالـ Live Reassessment: لو ظهر تهديد خطير أثناء الشغل، الصلاحية بتتسحب والجهاز بيتعزل فوراً.",
        ],
    )
    return slide


def _slide_7_pilot_and_handoff(prs):
    """Slide 7: Phased Campus Pilot, Native 90% Coverage Chart & Handoff to Part 04."""
    slide = add_blank_slide(prs, bg="horizontal")
    glow = add_motif(slide, "glow", left=-2.6, top=-3.0, width=8.2)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.88,
        width=17.76,
        height=0.95,
        text="Phased Campus Pilot, Coverage KPIs & Policy Handoff",
        font_name=FONT_TITLE,
        font_size=48.0,
        color="WHITE",
        bold=True,
        name="PilotTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.88,
        width=16.8,
        height=0.50,
        text="Validating enforcement accuracy and educational usability before campus-wide rollout",
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        name="PilotSubtitle",
    )

    # Left Panel: Native Editable PowerPoint Column Chart (180/200 = 90% Fresh Coverage)
    chart_card = add_frosted_card(
        slide,
        MARGIN_X,
        2.68,
        9.25,
        7.30,
        fill_color="DARK_TEAL",
        fill_alpha=54.0,
        radius=0.04,
        name="PilotChartCard",
    )
    chart_data = CategoryChartData()
    chart_data.categories = ["In-Scope Fleet", "Fresh (<15 min)", "Stale / Investigate"]
    chart_data.add_series("Pilot Endpoints", [200, 180, 20])

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(MARGIN_X + 0.34),
        Inches(2.96),
        Inches(8.56),
        Inches(6.72),
        chart_data,
    )
    chart_shape.name = "PilotTelemetryChart"
    chart = chart_shape.chart
    chart.has_legend = False
    chart.has_title = True
    chart.chart_title.text_frame.text = "Illustrative Staff Pilot Coverage (90% Freshness)"
    for p in chart.chart_title.text_frame.paragraphs:
        p.font.name = FONT_BODY
        p.font.size = Pt(19)
        p.font.bold = True
        p.font.color.rgb = rgb("WHITE")

    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plot.data_labels.font.name = FONT_BODY
    plot.data_labels.font.size = Pt(17)
    plot.data_labels.font.bold = True
    plot.data_labels.font.color.rgb = rgb("WHITE")

    bar_colors = ["MUTED_TEAL", "VERIFIED_MINT", "ALERT_CORAL"]
    series = plot.series[0]
    for i, col_key in enumerate(bar_colors):
        pt = series.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(col_key)

    for axis in (chart.category_axis, chart.value_axis):
        axis.tick_labels.font.name = FONT_BODY
        axis.tick_labels.font.size = Pt(15)
        axis.tick_labels.font.color.rgb = rgb("ICE_MIST")
        axis.format.line.color.rgb = rgb("SAGE_GLASS")
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = rgb("SAGE_GLASS")

    # Right Stack: 3 High-Contrast Dark-Teal Glass Cards (x=10.78, w=8.10)
    rx, rw = 10.78, 8.10
    right_specs = [
        (
            "1 · Three-Phase Pilot Rollout",
            "Start with asset & posture inventory → test rules in shadow observation mode → enforce baseline with clear repair guidance.",
            "SAGE_GLASS",
            58.0,
        ),
        (
            "2 · Proportionate Failure Policy",
            "Fail-closed on sensitive grading & exam operations when posture is unavailable; keep low-risk LMS coursework open.",
            "SAGE_GLASS",
            58.0,
        ),
        (
            "3 · Posture → Authorization Bridge",
            "Track 90% fresh coverage (180/200) and repair speed — streaming verified endpoint states into dynamic Authorization Policies.",
            "VERIFIED_MINT",
            68.0,
        ),
    ]
    rcards = []
    for k, (rt, rb, raccent, ralpha) in enumerate(right_specs):
        ry = 2.68 + k * (2.22 + 0.32)
        rc = add_frosted_card(
            slide,
            rx,
            ry,
            rw,
            2.22,
            fill_color="DARK_TEAL",
            fill_alpha=ralpha,
            border_color=raccent,
            border_alpha=75.0,
            border_pt=1.25,
            radius=0.06,
            name=f"PilotRightCard_{k+1}",
        )
        _add_card_heading_body(
            rc,
            heading=rt,
            body=rb,
            head_size=22.5,
            body_size=18.5,
            head_color="VERIFIED_MINT" if k == 2 else "WHITE",
            body_color="WHITE",
            top_margin_in=0.26,
        )
        rcards.append(rc)

    add_footer(slide, PART_LABEL, slide_num=7)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(glow, trigger="on_load", duration_ms=380)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.fade_in(chart_card, trigger="after_previous", duration_ms=320)
    anim.wipe_in(chart_shape, direction="up", trigger="with_previous", duration_ms=480, delay_ms=80)
    anim.stagger(rcards, effect="fly_right", step_ms=130, on_click_first=False)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "لتطبيق النظام في الجامعة بنمشي على Pilot من 3 مراحل: جرد ومراقبة الأول، بعدين اختبار في Observation Mode بدون قفل، وأخيراً التطبيق الإلزامي تدريجياً.",
            "وفي الأعطال بنطبق Fail-Closed على رصد الدرجات والامتحانات الحساسة، مع إبقاء الكورسات مفتوحة عشان الدراسة متقفش.",
            "وبنقيس النجاح بتحقيق 90% تغطية بأدلة طازة يعني 180/200 جهاز في الـ Pilot وسرعة حل الثغرات.",
            "وبكده بنسلم حالة الجهاز الموثقة لحظياً لمحرك سياسات الصلاحيات Authorization Policies في الجزء الرابع عشان ياخد قرار الوصول النهائي. شكراً لحضرتك يا دكتورة.",
        ],
    )
    return slide


def build_part_03(prs) -> list:
    """Append 7 animated, template-locked slides for Part 03 to `prs`."""
    slides = []

    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 03  ·  ENDPOINT PILLAR",
        title="Device Security\nPosture",
        summary=(
            "Verified identity confirms who is signing in — not whether their endpoint is patched, "
            "encrypted, or compromise-free before accessing protected university services."
        ),
        presenter="Omar Mounir Mohamed Arafa  ·  Section 2 (BN 16)",
        motif="crescents",
        transition="push_left",
    )
    _set_rtl_notes(
        s1,
        [
            "السلام عليكم يا دكتورة، أنا عمر منير محمد عرفه، وهتكلم في الجزء التالت عن الـ Device Security Posture.",
            "وإزاي بعد ما زميلي اتأكد من هوية المستخدم، بنتحقق من صحة وأمان الجهاز نفسه قبل ما نسمح له يوصل لسيرفرات وشبكة الكلية.",
        ],
    )
    slides.append(s1)

    slides.append(_slide_2_device_tiers(prs))
    slides.append(_slide_3_signals_and_states(prs))
    slides.append(_slide_4_certificates_and_gates(prs))
    slides.append(_slide_5_byod_and_labs(prs))
    slides.append(_slide_6_quarantine_loop(prs))
    slides.append(_slide_7_pilot_and_handoff(prs))

    return slides
