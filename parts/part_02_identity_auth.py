"""Part 02 — Identity and Authentication.

Owner: Omar Abdelsalam Abdelfattah Mohamed Afifi
Slide Budget: 7 slides (including opening section divider; hard cap <= 7).

Slide Arc (7 Distinct On-Theme Layouts, Alternating Gradients, Auto-Play OOXML Animations):
1. L03 Section Divider       (bg=diagonal,   motif=nested_arch)  — Part 02 Opening & Bridge from Part 01
2. Custom Split Visual       (bg=vertical,   motif=signal_waves) — Why Addresses Cannot Prove Identity + 4 Campus Personas
3. L05 Table + Token Strip   (bg=diagonal,   motif=glow)         — Centralized Identity Provider & Single Sign-On Comparison
4. Custom Tiered + Edge Flow (bg=horizontal)                     — Tiered Multi-Factor Authentication & 802.1X Network Binding
5. Custom 8-Node Serpentine  (bg=radial)                         — Eight-Step Zero Trust Login & Verification Flow
6. Custom Lifecycle + Kill   (bg=vertical)                       — Account Lifecycle & Sub-Minute Revocation (< 60 Seconds)
7. Custom 3-Case + Bridge    (bg=horizontal)                     — Why Valid Identity Alone Never Grants Trust (Bridge to Part 03)

Clarity Standard:
- All heavy technical acronyms are written in full, speakable English for live presentation
  (keeping only foundational networking terms: Wi-Fi, IP, TCP, HTTP, DNS, 802.1X, 802.11, JSON Web Token).
- Native RTL Egyptian Arabic Speaker Notes with isolated English runs for clean Presenter View reading.

Out of Scope (treated as inputs/handoffs — NOT re-explained):
- Why perimeter failed (Part 01), endpoint posture / device health internals (Part 03),
  authorization policy engine math (Part 04), network segmentation plumbing (Part 05), remote access proxy (Part 06).
"""

from __future__ import annotations

import re
from typing import Sequence

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import (
    L03_section_divider,
    L05_table_compare,
    _style_table_cell,
)
from core.theme import (
    COLORS,
    FONT_BODY,
    FONT_DISPLAY,
    FONT_TITLE,
    MARGIN_X,
    MARGIN_Y,
    add_arrow_connector,
    add_badge,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_motif,
    add_solid_card,
    add_textbox,
    rgb,
)

PART_LABEL = "Part 02  ·  Identity & Authentication"
_EN_SPAN_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9\.\-/%& ]*[A-Za-z0-9%]|[A-Za-z0-9%]")


def _set_rtl_notes(slide, lines: Sequence[str]) -> None:
    """Write clean native RTL Egyptian Arabic speaker notes with isolated English runs."""
    for raw in lines:
        if "\u0627\u064b" in raw:
            raise ValueError(f"Invalid Arabic Tanween Fath on Alef in notes: {raw}")

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


def _tight_pill_margins(pill, mx: float = 0.06, my: float = 0.02) -> None:
    """Tighten horizontal/vertical text padding inside compact badges so text stays single-line."""
    tf = pill.text_frame
    tf.margin_left = Inches(mx)
    tf.margin_right = Inches(mx)
    tf.margin_top = Inches(my)
    tf.margin_bottom = Inches(my)


def _fill_card_paragraphs(
    card,
    heading: str,
    body: str,
    head_size: float = 22.0,
    body_size: float = 18.0,
    head_color: str = "WHITE",
    body_color: str = "ICE_MIST",
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    margin_x: float = 0.28,
    margin_y: float = 0.18,
    head_space_after: float = 6.0,
    line_spacing: float = 1.18,
) -> None:
    """Populate a frosted/solid card with crisp Space Grotesk title + Work Sans body."""
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margin_x)
    tf.margin_right = Inches(margin_x)
    tf.margin_top = Inches(margin_y)
    tf.margin_bottom = Inches(margin_y)

    p0 = tf.paragraphs[0]
    p0.alignment = align
    p0.space_after = Pt(head_space_after if body else 0)
    p0.line_spacing = 1.10
    r0 = p0.add_run()
    r0.text = heading
    r0.font.name = FONT_TITLE
    r0.font.size = Pt(head_size)
    r0.font.bold = True
    r0.font.color.rgb = rgb(head_color)

    if body:
        for line in body.split("\n"):
            p = tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
            r = p.add_run()
            r.text = line
            r.font.name = FONT_BODY
            r.font.size = Pt(body_size)
            r.font.bold = False
            r.font.color.rgb = rgb(body_color)


# ===========================================================================
# SLIDE 1 — Section Divider (L03, bg=diagonal, motif=nested_arch)
# ===========================================================================
def _slide_01_divider(prs):
    slide, _ = L03_section_divider(
        prs,
        part_badge="PART 02  ·  IDENTITY PILLAR",
        title="Identity &\nAuthentication",
        summary=(
            "Once network location stops proving trust, verified user identity becomes the "
            "campus's first security pillar — uniting centralized accounts, phishing-resistant "
            "Multi-Factor Authentication, network binding, and sub-minute revocation."
        ),
        presenter="Omar Abdelsalam Abdelfattah Mohamed Afifi",
        motif="nested_arch",
        transition="push_left",
    )
    _set_rtl_notes(
        slide,
        [
            "السلام عليكم يا دكتورة، أنا عمر عبدالسلام عبدالفتاح، وهكمل من بعد زميلي أحمد حسن في جزء Identity and Authentication.",
            "بعد ما شفنا في Threat Model and Boundaries إن حدود الشبكة التقليدية انتهت ومبقاش ينفع نثق في الجهاز لمجرد إنه جوه الجامعة، أول سؤال لازم نجاوب عليه في Zero Trust هو: مين الشخص اللي بيطلب الوصول بالظبط؟",
            "في الـ ٦ سلايدز الجايين هنشوف ليه عناوين الشبكة متثبتش الهوية، وإزاي بنبني Central Identity Provider، وMulti-Factor Authentication متدرج، وربط الواي فاي بـ 802.1X، وإزاي نقفل أي حساب مخترق في أقل من ٦٠ ثانية.",
        ],
    )
    return slide


# ===========================================================================
# SLIDE 2 — Why Network Addresses Fail + Four Campus Personas (bg=vertical)
# ===========================================================================
def _slide_02_addresses_and_personas(prs):
    slide = add_blank_slide(prs, bg="vertical")
    motif = add_motif(slide, "signal_waves", left=15.45, top=-1.15, width=5.1)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.80,
        width=20.0 - 2 * MARGIN_X,
        height=0.88,
        text="Identity Replaces Network Addresses Across Four Roles",
        font_name=FONT_DISPLAY,
        font_size=41.0,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.LEFT,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.72,
        width=16.2,
        height=0.52,
        text="Network identifiers track hardware or temporary leases — never the person behind the session.",
        font_name=FONT_BODY,
        font_size=21.0,
        color="ICE_MIST",
        align=PP_ALIGN.LEFT,
        name="SlideSubtitle",
    )

    # Left Column: Why Network Layers Cannot Prove Identity (Dark frosted glass for crisp contrast)
    left_label = add_badge(
        slide,
        left=MARGIN_X,
        top=2.46,
        width=5.45,
        height=0.50,
        text="WHY NETWORK ADDRESSES FAIL",
        fill_color="DARK_TEAL",
        fill_alpha=82.0,
        border_color="ALERT_CORAL",
        border_alpha=80.0,
        text_color="WHITE",
        font_size=15.5,
        bold=True,
        name="LeftSectionBadge",
    )

    layer_specs = [
        (
            "Layer 2",
            "Data Link Layer (Media Access Control)",
            "A 48-bit Media Access Control address identifies the Network Interface Card hardware, "
            "not the student, and is easily spoofed. Address Resolution Protocol caches expire in about 20 minutes.",
        ),
        (
            "Layer 3",
            "Network Layer (Changing IP Leases)",
            "Dynamic Host Configuration Protocol assigns changing IP addresses as users walk across "
            "campus buildings; a single IP address never proves who the user is.",
        ),
        (
            "Layers 4–5",
            "Transport & Application (Legacy Cookies)",
            "TCP sockets prove port reachability only. Classic 8-hour HTTP cookies remain active "
            "even after an administrator resets the user's password.",
        ),
    ]

    left_cards = []
    card_x = MARGIN_X
    card_w = 8.55
    card_h = 2.08
    gap_y = 0.30
    start_y = 3.16
    for idx, (layer_tag, title_txt, desc_txt) in enumerate(layer_specs):
        cy = start_y + idx * (card_h + gap_y)
        card = add_frosted_card(
            slide,
            left=card_x,
            top=cy,
            width=card_w,
            height=card_h,
            fill_color="DARK_TEAL",
            fill_alpha=74.0,
            border_color="SAGE_GLASS",
            border_alpha=58.0,
            border_pt=1.1,
            radius=0.05,
            name=f"LayerCard_{idx+1}",
        )
        badge = add_badge(
            slide,
            left=card_x + 0.22,
            top=cy + 0.76,
            width=1.68,
            height=0.56,
            text=layer_tag,
            fill_color="ALERT_CORAL",
            fill_alpha=38.0,
            border_color="ALERT_CORAL",
            border_alpha=90.0,
            text_color="WHITE",
            font_size=15.5,
            bold=True,
            name=f"LayerBadge_{idx+1}",
        )
        _tight_pill_margins(badge, mx=0.06)
        tb = add_textbox(
            slide,
            left=card_x + 2.08,
            top=cy + 0.16,
            width=card_w - 2.30,
            height=card_h - 0.28,
            text=title_txt,
            font_name=FONT_TITLE,
            font_size=21.0,
            color="WHITE",
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
            line_spacing=1.10,
            name=f"LayerText_{idx+1}",
        )
        p_body = tb.text_frame.add_paragraph()
        p_body.space_before = Pt(6)
        p_body.line_spacing = 1.18
        r_b = p_body.add_run()
        r_b.text = desc_txt
        r_b.font.name = FONT_BODY
        r_b.font.size = Pt(18.0)
        r_b.font.color.rgb = rgb("ICE_MIST")

        left_cards.append((card, badge, tb))

    # Right Column: Four Campus Identity Personas (2x2 High-Contrast Dark Frosted Grid)
    right_x = 10.18
    right_label = add_badge(
        slide,
        left=right_x,
        top=2.46,
        width=5.65,
        height=0.50,
        text="FOUR CAMPUS IDENTITY PERSONAS",
        fill_color="DARK_TEAL",
        fill_alpha=82.0,
        border_color="VERIFIED_MINT",
        border_alpha=80.0,
        text_color="WHITE",
        font_size=15.5,
        bold=True,
        name="RightSectionBadge",
    )

    personas = [
        (
            "01 · Student",
            "University account + Time-Based One-Time Password for Learning Portal and Labs.",
            "Role: Student",
            "VERIFIED_MINT",
        ),
        (
            "02 · Lecturer & Researcher",
            "Faculty account + Hardware Passkey for Labs, Course Grading, and Research.",
            "Role: Lecturer",
            "SOFT_SAGE",
        ),
        (
            "03 · Administrator & Service",
            "Privileged account + Hardware Key (or short-lived service certificates) for Core Systems.",
            "Role: Administrator",
            "ALERT_CORAL",
        ),
        (
            "04 · Sponsored Guest",
            "Faculty sponsor + automatic expiration timer restricted to Public Campus Wi-Fi.",
            "Expires: 24 Hours – 7 Days",
            "ICE_MIST",
        ),
    ]

    pw = 4.22
    ph = 3.26
    pgx = 0.26
    pgy = 0.32
    persona_groups = []
    for idx, (ptitle, pdesc, pill_txt, accent) in enumerate(personas):
        r, c = divmod(idx, 2)
        px = right_x + c * (pw + pgx)
        py = start_y + r * (ph + pgy)
        pcard = add_frosted_card(
            slide,
            left=px,
            top=py,
            width=pw,
            height=ph,
            fill_color="DARK_TEAL",
            fill_alpha=76.0,
            border_color=accent,
            border_alpha=68.0,
            border_pt=1.2,
            radius=0.05,
            name=f"PersonaCard_{idx+1}",
        )
        _fill_card_paragraphs(
            pcard,
            heading=ptitle,
            body=pdesc,
            head_size=21.0,
            body_size=18.5,
            head_color="WHITE",
            body_color="ICE_MIST",
            margin_x=0.24,
            margin_y=0.18,
            head_space_after=6.0,
            line_spacing=1.18,
        )
        pill = add_badge(
            slide,
            left=px + 0.24,
            top=py + ph - 0.72,
            width=pw - 0.48,
            height=0.50,
            text=pill_txt,
            fill_color="OBSIDIAN",
            fill_alpha=70.0,
            border_color=accent,
            border_alpha=85.0,
            text_color="WHITE",
            font_size=14.5,
            bold=True,
            name=f"PersonaPill_{idx+1}",
        )
        _tight_pill_margins(pill, mx=0.06)
        persona_groups.append((pcard, pill))

    add_footer(slide, PART_LABEL)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(motif, trigger="on_load", duration_ms=480)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.fade_in(left_label, trigger="after_previous", duration_ms=280)
    for idx, (card, badge, tb) in enumerate(left_cards):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 120
        anim.fly_in(card, direction="left", trigger=trig, duration_ms=360, delay_ms=dly)
        anim.fade_in(badge, trigger="with_previous", duration_ms=320, delay_ms=dly + 40)
        anim.fade_in(tb, trigger="with_previous", duration_ms=320, delay_ms=dly + 60)
    anim.fade_in(right_label, trigger="after_previous", duration_ms=280)
    for idx, (pcard, pill) in enumerate(persona_groups):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 110
        anim.fly_in(pcard, direction="bottom", trigger=trig, duration_ms=360, delay_ms=dly)
        anim.fade_in(pill, trigger="with_previous", duration_ms=320, delay_ms=dly + 50)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "لو ربطنا تصميم الهوية بمحتوى كورس الشبكات، هنلاقي إن عناوين الشبكة متقدرش أبدًا تثبت هوية الشخص:",
            "في Layer 2، عنوان Media Access Control بيعرّف كارت الشبكة مش الطالب وبيتزور بسهولة، وجدول Address Resolution Protocol بينتهي كل ٢٠ دقيقة تقريبًا.",
            "في Layer 3، بروتوكول Dynamic Host Configuration Protocol بيدي IP مختلف كل ما الطالب يتحرك بين المدرجات، فـ IP واحد عمره ما يعرّف شخص واحد.",
            "وفي Layers 4–5، الـ TCP بيثبت إن البورت شغال بس، والـ HTTP Cookies التقليدية اللي مدتها ٨ ساعات بتفضل شغالة حتى لو غيّرنا الباسورد.",
            "عشان كده بنبني الهوية على ٤ أدوار واضحة في الجامعة: الطالب (Time-Based One-Time Password)، الدكتور والباحث (Hardware Passkey)، مدير النظام وحسابات الخدمات (Hardware Key أو شهادات قصيرة)، والضيف بضمان دكتور ومدة انتهاء تلقائية على الواي فاي العام بس.",
        ],
    )
    return slide


# ===========================================================================
# SLIDE 3 — Centralized Identity Provider & Single Sign-On (L05, bg=diagonal)
# ===========================================================================
def _slide_03_sso_protocol_comparison(prs):
    slide, anim, tbl_shape = L05_table_compare(
        prs,
        title="Centralized Identity &\nSingle Sign-On Choice",
        intro=(
            "Replacing fragmented per-system logins with one Campus Identity Provider "
            "issuing short-lived cryptographic tokens across every web, mobile, and lab service."
        ),
        headers=["Architecture Option", "Mechanism", "Campus Impact", "Decision"],
        rows=[
            [
                "Local Per-System Accounts",
                "Separate user tables on Portal, Lab Secure Shell, and Records.",
                "Password reuse; locking a hacked student misses other systems; forgotten guests stay open.",
                "Rejected\n(Emergency Backup)",
            ],
            [
                "Central Identity Provider +\nSecurity Assertion Markup Language 2.0",
                "Signed Extensible Markup Language assertions via browser redirects.",
                "Mature web support, but heavy message parsing and weak mobile or application token rotation.",
                "Fallback for\nLegacy Portals",
            ],
            [
                "Central Identity Provider +\nOpenID Connect (Open Authorization 2.0)",
                "Signed short-lived JSON Web Tokens (5–10 minutes) + central refresh tokens.",
                "Carries verified role and authentication strength to Policy Decision Point; instant logout everywhere.",
                "PRIMARY STANDARD",
            ],
        ],
        col_widths=[4.45, 4.45, 5.56, 3.30],
        part_label=PART_LABEL,
        transition="fade",
    )

    # Tighten table vertical footprint cleanly above bottom Token Anatomy Strip (ends by y ~ 8.15)
    tbl_shape.top = Inches(3.26)
    tbl_shape.height = Inches(4.78)
    tbl = tbl_shape.table
    tbl.rows[0].height = Inches(0.68)
    row_heights = [1.28, 1.40, 1.40]
    for r_idx, rh in enumerate(row_heights, start=1):
        tbl.rows[r_idx].height = Inches(rh)
        for c_idx in range(len(tbl.columns)):
            cell = tbl.cell(r_idx, c_idx)
            cell.margin_top = Inches(0.08)
            cell.margin_bottom = Inches(0.08)
            cell.margin_left = Inches(0.18)
            cell.margin_right = Inches(0.18)
            for p in cell.text_frame.paragraphs:
                p.line_spacing = 1.12
                for run in p.runs:
                    run.font.size = Pt(18.0)

    # Clear each Decision cell BEFORE restyling so text never duplicates!
    decision_styles = [
        (1, "Rejected\n(Emergency Backup)", COLORS["SOFT_SAGE"], COLORS["ALERT_CORAL"], 18.0),
        (2, "Fallback for\nLegacy Portals", COLORS["SOFT_SAGE"], COLORS["DARK_TEAL"], 18.0),
        (3, "PRIMARY STANDARD", COLORS["DARK_TEAL"], COLORS["VERIFIED_MINT"], 19.0),
    ]
    for row_i, text_val, bg_hex, fg_hex, fsz in decision_styles:
        c = tbl.cell(row_i, 3)
        c.text_frame.clear()
        _style_table_cell(
            c,
            text_val,
            fill_hex=bg_hex,
            text_hex=fg_hex,
            bold=True,
            font_size=fsz,
            align=PP_ALIGN.CENTER,
        )
        c.margin_top = Inches(0.08)
        c.margin_bottom = Inches(0.08)

    # Bottom Signed JSON Web Token Anatomy Strip (clean gap below table: y=8.48 .. 10.14)
    token_bar = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=8.48,
        width=20.0 - 2 * MARGIN_X,
        height=1.66,
        fill_color="DARK_TEAL",
        fill_alpha=72.0,
        border_color="VERIFIED_MINT",
        border_alpha=78.0,
        border_pt=1.25,
        radius=0.06,
        name="TokenAnatomyBar",
    )
    token_hdr = add_textbox(
        slide,
        left=MARGIN_X + 0.30,
        top=8.60,
        width=17.0,
        height=0.38,
        text="WHAT THE SIGNED JSON WEB TOKEN CARRIES TO EVERY CAMPUS SERVICE:",
        font_name=FONT_BODY,
        font_size=15.0,
        color="VERIFIED_MINT",
        bold=True,
        name="TokenHeader",
    )

    claims = [
        ("User Role", "Student · Lecturer · Admin · Guest"),
        ("Authentication Strength", "Time-Based Code or Hardware Passkey"),
        ("Account Status", "Active · Graduated · Suspended"),
        ("Short Expiration", "5 to 10 Minutes Maximum"),
    ]
    cw = 4.08
    cgap = 0.26
    cx0 = MARGIN_X + 0.30
    claim_cards = []
    for idx, (cl_k, cl_v) in enumerate(claims):
        cx = cx0 + idx * (cw + cgap)
        cc = add_frosted_card(
            slide,
            left=cx,
            top=9.06,
            width=cw,
            height=0.92,
            fill_color="OBSIDIAN",
            fill_alpha=52.0,
            border_color="SAGE_GLASS",
            border_alpha=55.0,
            radius=0.10,
            name=f"TokenClaim_{idx+1}",
        )
        _fill_card_paragraphs(
            cc,
            heading=cl_k,
            body=cl_v,
            head_size=16.5,
            body_size=14.5,
            head_color="WHITE",
            body_color="ICE_MIST",
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
            margin_x=0.10,
            margin_y=0.04,
            head_space_after=3.0,
            line_spacing=1.10,
        )
        claim_cards.append(cc)

    anim.fade_in(token_bar, trigger="after_previous", duration_ms=340)
    anim.fade_in(token_hdr, trigger="with_previous", duration_ms=300, delay_ms=60)
    anim.stagger(claim_cards, effect="fly_bottom", duration_ms=320, step_ms=90, on_click_first=False, base_delay_ms=60)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "عشان ندير حسابات الجامعة صح، قارنّا بين ٣ معماريات أساسية:",
            "أولًا: الحسابات المنفصلة على كل سيستم لوحده (Local Per-System Accounts) — ودي مرفوضة تمامًا إلا كحسابات طوارئ مقفولة، لأنها بتخلي الطالب يكرر الباسورد ولو قفلنا حسابه على البوابة بيفضل مفتوح على المعامل.",
            "ثانيًا: Central Identity Provider مع بروتوكول Security Assertion Markup Language 2.0 — كويس ومجرّب للبوابات الأكاديمية القديمة فخليناه Fallback بس، لأنه تقيل في الـ Extensible Markup Language وضعيف في تجديد توكنز الموبايل والتطبيقات.",
            "ثالثًا وهو اختيارنا الأساسي: Central Identity Provider مع OpenID Connect المبني على Open Authorization 2.0، لأنه بيطلع JSON Web Token موقّع مدته قصيرة من ٥ لـ ١٠ دقائق بس، شايل جواه ٤ حاجات واضحة: دور المستخدم، قوة التحقق، حالة الحساب، ووقت الانتهاء القصير، وبيسمح بخروج فوري من كل أنظمة الجامعة.",
        ],
    )
    return slide


# ===========================================================================
# SLIDE 4 — Tiered Multi-Factor Authentication & 802.1X Edge (bg=horizontal)
# ===========================================================================
def _slide_04_tiered_mfa_and_8021x(prs):
    slide = add_blank_slide(prs, bg="horizontal")

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.80,
        width=20.0 - 2 * MARGIN_X,
        height=0.90,
        text="Tiered Multi-Factor Authentication & 802.1X Binding",
        font_name=FONT_DISPLAY,
        font_size=48.0,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.72,
        width=17.5,
        height=0.52,
        text="Matching verification strength to campus risk (Digital Identity Guidelines 800-63B) and enforcing identity at the Wi-Fi edge.",
        font_name=FONT_BODY,
        font_size=20.5,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    # Left Column: 3 High-Contrast Dark Frosted Multi-Factor Cards
    left_w = 10.35
    mfa_hdr = add_badge(
        slide,
        left=MARGIN_X,
        top=2.46,
        width=5.85,
        height=0.48,
        text="MULTI-FACTOR AUTHENTICATION TIERS",
        fill_color="DARK_TEAL",
        fill_alpha=80.0,
        border_color="SAGE_GLASS",
        border_alpha=75.0,
        text_color="WHITE",
        font_size=15.5,
        bold=True,
        name="MfaSectionBadge",
    )

    mfa_specs = [
        (
            "Text Message / Email One-Time Password",
            "REJECT AS PRIMARY",
            "ALERT_CORAL",
            "Fails inside concrete basement labs without cellular signal; vulnerable to "
            "phone-number hijacking and real-time phishing proxies.",
            2.95,
        ),
        (
            "Time-Based One-Time Password (Standard 6238)",
            "STUDENTS & GUESTS",
            "VERIFIED_MINT",
            "Generates 30-second offline codes directly on smartphones; zero hardware cost; "
            "works reliably inside basement labs and signal dead zones.",
            2.72,
        ),
        (
            "Hardware Passkey / Web Authentication",
            "FACULTY & ADMINISTRATORS",
            "VERIFIED_MINT",
            "Domain-bound public-key cryptography (Fast Identity Online 2) defeats phishing "
            "completely; mandatory before grading or accessing Student Records.",
            3.55,
        ),
    ]

    mfa_groups = []
    mh = 2.08
    mgap = 0.28
    my0 = 3.14
    for idx, (mtitle, mbadge, accent, mbody, bw_pill) in enumerate(mfa_specs):
        my = my0 + idx * (mh + mgap)
        card = add_frosted_card(
            slide,
            left=MARGIN_X,
            top=my,
            width=left_w,
            height=mh,
            fill_color="DARK_TEAL",
            fill_alpha=68.0 if idx < 2 else 82.0,
            border_color=accent,
            border_alpha=82.0,
            border_pt=1.35,
            radius=0.05,
            name=f"MfaCard_{idx+1}",
        )
        t_box = add_textbox(
            slide,
            left=MARGIN_X + 0.30,
            top=my + 0.22,
            width=left_w - bw_pill - 0.68,
            height=0.52,
            text=mtitle,
            font_name=FONT_TITLE,
            font_size=19.0,
            color="WHITE",
            bold=True,
            name=f"MfaTitle_{idx+1}",
        )
        b_pill = add_badge(
            slide,
            left=MARGIN_X + left_w - bw_pill - 0.26,
            top=my + 0.20,
            width=bw_pill,
            height=0.50,
            text=mbadge,
            fill_color=accent if idx == 0 else "OBSIDIAN",
            fill_alpha=40.0 if idx == 0 else 72.0,
            border_color=accent,
            border_alpha=92.0,
            text_color="WHITE" if idx == 0 else "VERIFIED_MINT",
            font_size=14.0,
            bold=True,
            name=f"MfaBadge_{idx+1}",
        )
        _tight_pill_margins(b_pill, mx=0.08)
        d_box = add_textbox(
            slide,
            left=MARGIN_X + 0.30,
            top=my + 0.84,
            width=left_w - 0.60,
            height=1.10,
            text=mbody,
            font_name=FONT_BODY,
            font_size=19.0,
            color="ICE_MIST",
            line_spacing=1.22,
            name=f"MfaDesc_{idx+1}",
        )
        mfa_groups.append((card, t_box, b_pill, d_box))

    # Right Panel: 802.1X + Remote Authentication Dial-In User Service
    rx = 11.92
    rw = 20.0 - MARGIN_X - rx  # 6.96"
    rpanel = add_frosted_card(
        slide,
        left=rx,
        top=2.46,
        width=rw,
        height=7.50,
        fill_color="DARK_TEAL",
        fill_alpha=78.0,
        border_color="SAGE_GLASS",
        border_alpha=72.0,
        border_pt=1.25,
        radius=0.024,
        name="EdgePanel",
    )
    rhdr = add_textbox(
        slide,
        left=rx + 0.24,
        top=2.64,
        width=rw - 0.48,
        height=0.92,
        text="Identity at the Wi-Fi & Switch Edge",
        font_name=FONT_TITLE,
        font_size=23.5,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="EdgeHeader",
    )
    p_rsub = rhdr.text_frame.add_paragraph()
    p_rsub.alignment = PP_ALIGN.CENTER
    p_rsub.space_before = Pt(4)
    rr = p_rsub.add_run()
    rr.text = "(802.1X + Remote Authentication Dial-In User Service)"
    rr.font.name = FONT_BODY
    rr.font.size = Pt(15.5)
    rr.font.bold = True
    rr.font.color.rgb = rgb("VERIFIED_MINT")

    edge_steps = [
        (
            "01 · Device Joins Campus 802.11 Wi-Fi",
            "Wireless Access Point blocks lab and internet traffic; allows only Extensible Authentication Protocol frames.",
        ),
        (
            "02 · Server Checks Identity Provider",
            "Remote Authentication Server verifies account status (Student, Guest, or Staff) before granting any IP.",
        ),
        (
            "03 · Identity-Bound IP Assignment",
            "Dynamic Host Configuration Protocol issues an IP; Network Segmentation maps the matching Virtual Local Area Network.",
        ),
    ]

    edge_cards = []
    edge_arrows = []
    eh = 1.72
    egap = 0.42
    ey0 = 3.76
    for idx, (etitle, edesc) in enumerate(edge_steps):
        ey = ey0 + idx * (eh + egap)
        ec = add_frosted_card(
            slide,
            left=rx + 0.30,
            top=ey,
            width=rw - 0.60,
            height=eh,
            fill_color="OBSIDIAN",
            fill_alpha=46.0,
            border_color="VERIFIED_MINT" if idx == 2 else "SAGE_GLASS",
            border_alpha=78.0 if idx == 2 else 55.0,
            border_pt=1.2,
            radius=0.07,
            name=f"EdgeStep_{idx+1}",
        )
        _fill_card_paragraphs(
            ec,
            heading=etitle,
            body=edesc,
            head_size=19.5,
            body_size=17.5,
            head_color="WHITE",
            body_color="ICE_MIST",
            margin_x=0.24,
            margin_y=0.16,
            head_space_after=5.0,
            line_spacing=1.18,
        )
        edge_cards.append(ec)
        if idx < 2:
            arr = add_arrow_connector(
                slide,
                x1=rx + rw / 2.0,
                y1=ey + eh + 0.04,
                x2=rx + rw / 2.0,
                y2=ey + eh + egap - 0.04,
                color="VERIFIED_MINT",
                width_pt=2.2,
                name=f"EdgeArr_{idx+1}",
            )
            edge_arrows.append(arr)

    add_footer(slide, PART_LABEL)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.fade_in(mfa_hdr, trigger="after_previous", duration_ms=260)
    for idx, (card, t_box, b_pill, d_box) in enumerate(mfa_groups):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 130
        anim.fly_in(card, direction="left", trigger=trig, duration_ms=360, delay_ms=dly)
        anim.fade_in(t_box, trigger="with_previous", duration_ms=300, delay_ms=dly + 40)
        anim.fade_in(b_pill, trigger="with_previous", duration_ms=300, delay_ms=dly + 60)
        anim.fade_in(d_box, trigger="with_previous", duration_ms=300, delay_ms=dly + 80)

    anim.fade_in(rpanel, trigger="after_previous", duration_ms=340)
    anim.fade_in(rhdr, trigger="with_previous", duration_ms=300, delay_ms=60)
    for idx, ec in enumerate(edge_cards):
        if idx > 0:
            anim.wipe_in(edge_arrows[idx - 1], direction="down", trigger="after_previous", duration_ms=240)
        anim.fly_in(ec, direction="bottom", trigger="after_previous" if idx == 0 else "with_previous", duration_ms=340, delay_ms=60)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "بالنسبة للـ Multi-Factor Authentication وحسب Digital Identity Guidelines رقم 800-63B، عملنا تدرج عملي يناسب الجامعة:",
            "أول اختيار وهو رسائل الموبايل والإيميل (Text Message / Email One-Time Password) رفضناه كحل أساسي لأنه بيقع جوه المعامل اللي في البدروم ومفيهاش شبكة موبايل، وكمان معرض لسرقة رقم الخط والتصيد اللحظي.",
            "فالحل العملي للطلبة والضيوف هو تطبيق Time-Based One-Time Password (معيار 6238)، بيولّد كود كل ٣٠ ثانية أوفلاين من غير شبكة موبايل ومن غير تكلفة أجهزة.",
            "أما للدكاترة ومديري النظام، فإلزامي Hardware Passkey / Web Authentication (معيار Fast Identity Online 2) لأنه مربوط بالدومين بالتشفير ومستحيل يتسرق بالـ Phishing.",
            "وعلى طرف الواي فاي والسويتشات بنربط الاتصال بالهوية باستخدام 802.1X مع خادم Remote Authentication Dial-In User Service: أول ما الجهاز يدخل على 802.11، الـ Wireless Access Point بتمنع أي ترافيك ماعدا Extensible Authentication Protocol لحد ما الخادم يتأكد من الحساب في الـ Identity Provider، وبعدها الـ Dynamic Host Configuration Protocol يديه IP وجزء Network Segmentation يوجّهه للـ Virtual Local Area Network المناسبة.",
        ],
    )
    return slide


# ===========================================================================
# SLIDE 5 — Eight-Step Zero Trust Login & Handoff Flow (bg=radial)
# ===========================================================================
def _slide_05_login_decision_flow(prs):
    slide = add_blank_slide(prs, bg="radial")

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.80,
        width=20.0 - 2 * MARGIN_X,
        height=0.88,
        text="Eight-Step Zero Trust Login & Verification Flow",
        font_name=FONT_DISPLAY,
        font_size=48.0,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.70,
        width=17.5,
        height=0.50,
        text="Every campus request proves identity and authentication strength before reaching the Policy Decision Point.",
        font_name=FONT_BODY,
        font_size=20.5,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    bw = 4.02
    bh = 2.48
    gap_x = 0.56
    xs = [MARGIN_X + i * (bw + gap_x) for i in range(4)]
    y_top = 2.50
    y_bot = 5.84

    # High-contrast dark frosted cards along serpentine path: 1->2->3->4 -> down -> 5->6->7->8
    step_defs = [
        (xs[0], y_top, "01", "Access Request", "User opens Learning Portal, Lab Secure Shell, or Student Records web address.", "DARK_TEAL", 72.0, "SAGE_GLASS"),
        (xs[1], y_top, "02", "DNS & Secure Channel", "DNS lookup (Port 53), TCP 3-way handshake, and Transport Layer Security.", "DARK_TEAL", 72.0, "SAGE_GLASS"),
        (xs[2], y_top, "03", "Enforcement Check", "Identity-aware proxy checks for a valid 10-minute JSON Web Token.", "DARK_TEAL", 72.0, "SAGE_GLASS"),
        (xs[3], y_top, "04", "Identity Provider Login", "Missing or expired token triggers OpenID Connect redirect to Identity Provider.", "DARK_TEAL", 82.0, "VERIFIED_MINT"),
        (xs[3], y_bot, "05", "Lifecycle Gate", "Active? Graduated, expired, or suspended → Deny & alert Monitoring and Response.", "DARK_TEAL", 78.0, "ALERT_CORAL"),
        (xs[2], y_bot, "06", "Multi-Factor Check", "Time-Based One-Time Password (Students) or Hardware Passkey (Staff).", "DARK_TEAL", 82.0, "VERIFIED_MINT"),
        (xs[1], y_bot, "07", "Issue Signed Token", "Signed JSON Web Token: role, authentication strength, status, 5–10 minute expiry.", "DARK_TEAL", 84.0, "VERIFIED_MINT"),
        (xs[0], y_bot, "08", "Policy Handoff", "Policy Decision Point merges Identity, Device Security Posture & Authorization Policies.", "OBSIDIAN", 86.0, "VERIFIED_MINT"),
    ]

    nodes = []
    for idx, (x, y, num, stitle, sbody, fcol, falpha, bcol) in enumerate(step_defs):
        card = add_frosted_card(
            slide,
            left=x,
            top=y,
            width=bw,
            height=bh,
            fill_color=fcol,
            fill_alpha=falpha,
            border_color=bcol,
            border_alpha=82.0 if bcol != "SAGE_GLASS" else 62.0,
            border_pt=1.35 if bcol != "SAGE_GLASS" else 1.0,
            radius=0.06,
            name=f"FlowStep_{num}",
        )
        num_pill = add_badge(
            slide,
            left=x + 0.20,
            top=y + 0.20,
            width=0.72,
            height=0.46,
            text=num,
            fill_color=bcol if bcol != "SAGE_GLASS" else "OBSIDIAN",
            fill_alpha=36.0 if bcol != "SAGE_GLASS" else 76.0,
            border_color=bcol,
            border_alpha=90.0,
            text_color="WHITE",
            font_size=15.0,
            bold=True,
            name=f"StepNum_{num}",
        )
        _tight_pill_margins(num_pill, mx=0.04, my=0.02)
        ttl = add_textbox(
            slide,
            left=x + 1.00,
            top=y + 0.21,
            width=bw - 1.14,
            height=0.46,
            text=stitle,
            font_name=FONT_TITLE,
            font_size=17.5,
            color="WHITE",
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
            name=f"StepTitle_{num}",
        )
        bdy = add_textbox(
            slide,
            left=x + 0.22,
            top=y + 0.82,
            width=bw - 0.44,
            height=bh - 0.96,
            text=sbody,
            font_name=FONT_BODY,
            font_size=18.0,
            color="ICE_MIST",
            line_spacing=1.20,
            name=f"StepBody_{num}",
        )
        nodes.append((card, num_pill, ttl, bdy))

    # Arrows connecting 1->2->3->4 -> down -> 5->6->7->8
    cy_top = y_top + bh / 2.0
    cy_bot = y_bot + bh / 2.0
    arrows = [
        add_arrow_connector(slide, xs[0] + bw + 0.06, cy_top, xs[1] - 0.06, cy_top, color="WHITE", width_pt=2.0, name="Arr_1_2"),
        add_arrow_connector(slide, xs[1] + bw + 0.06, cy_top, xs[2] - 0.06, cy_top, color="WHITE", width_pt=2.0, name="Arr_2_3"),
        add_arrow_connector(slide, xs[2] + bw + 0.06, cy_top, xs[3] - 0.06, cy_top, color="VERIFIED_MINT", width_pt=2.2, name="Arr_3_4"),
        add_arrow_connector(slide, xs[3] + bw / 2.0, y_top + bh + 0.06, xs[3] + bw / 2.0, y_bot - 0.06, color="VERIFIED_MINT", width_pt=2.2, name="Arr_4_5"),
        add_arrow_connector(slide, xs[3] - 0.06, cy_bot, xs[2] + bw + 0.06, cy_bot, color="VERIFIED_MINT", width_pt=2.2, name="Arr_5_6"),
        add_arrow_connector(slide, xs[2] - 0.06, cy_bot, xs[1] + bw + 0.06, cy_bot, color="VERIFIED_MINT", width_pt=2.2, name="Arr_6_7"),
        add_arrow_connector(slide, xs[1] - 0.06, cy_bot, xs[0] + bw + 0.06, cy_bot, color="VERIFIED_MINT", width_pt=2.2, name="Arr_7_8"),
    ]
    cred_lbl = add_badge(
        slide,
        left=xs[3] + bw / 2.0 - 2.68,
        top=y_top + bh + 0.19,
        width=2.44,
        height=0.46,
        text="Credentials Valid",
        fill_color="DARK_TEAL",
        fill_alpha=90.0,
        border_color="VERIFIED_MINT",
        border_alpha=90.0,
        text_color="VERIFIED_MINT",
        font_size=14.5,
        bold=True,
        name="CredValidBadge",
    )
    _tight_pill_margins(cred_lbl, mx=0.06, my=0.02)

    # Bottom Scope Boundary Banner
    scope_bar = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=8.78,
        width=20.0 - 2 * MARGIN_X,
        height=1.34,
        fill_color="DARK_TEAL",
        fill_alpha=78.0,
        border_color="VERIFIED_MINT",
        border_alpha=82.0,
        border_pt=1.3,
        radius=0.06,
        name="ScopeBoundaryBanner",
    )
    tf_sb = scope_bar.text_frame
    tf_sb.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_sb.margin_left = Inches(0.36)
    tf_sb.margin_right = Inches(0.36)
    p_sb = tf_sb.paragraphs[0]
    p_sb.alignment = PP_ALIGN.LEFT
    p_sb.line_spacing = 1.20
    r_sb1 = p_sb.add_run()
    r_sb1.text = "Architecture Boundary:  "
    r_sb1.font.name = FONT_TITLE
    r_sb1.font.size = Pt(19.5)
    r_sb1.font.bold = True
    r_sb1.font.color.rgb = rgb("VERIFIED_MINT")
    r_sb2 = p_sb.add_run()
    r_sb2.text = (
        "Identity and Authentication completes at Step 7 by delivering verified cryptographic identity claims. "
        "Step 8 hands off to Device Security Posture and Authorization Policies before any traffic passes the Policy Enforcement Point."
    )
    r_sb2.font.name = FONT_BODY
    r_sb2.font.size = Pt(18.5)
    r_sb2.font.color.rgb = rgb("WHITE")

    add_footer(slide, PART_LABEL)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    for idx, (card, num_pill, ttl, bdy) in enumerate(nodes):
        if idx > 0:
            wdir = "right" if idx <= 3 else ("down" if idx == 4 else "left")
            anim.wipe_in(arrows[idx - 1], direction=wdir, trigger="after_previous", duration_ms=190)
            if idx == 4:
                anim.fade_in(cred_lbl, trigger="with_previous", duration_ms=220)
        anim.fade_in(card, trigger="after_previous" if idx == 0 else "with_previous", duration_ms=260, delay_ms=30)
        anim.fade_in(num_pill, trigger="with_previous", duration_ms=240, delay_ms=40)
        anim.fade_in(ttl, trigger="with_previous", duration_ms=240, delay_ms=50)
        anim.fade_in(bdy, trigger="with_previous", duration_ms=240, delay_ms=60)
    anim.fly_in(scope_bar, direction="bottom", trigger="after_previous", duration_ms=340)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "الرسمة دي بتوضح مسار تسجيل الدخول الكامل في ٨ خطوات متتالية:",
            "١- المستخدم بيفتح رابط البوابة أو معمل الـ Secure Shell أو سجلات الطلبة. ٢- بيحصل DNS على بورت 53 وبعده TCP 3-way handshake وقناة تشفير Transport Layer Security.",
            "٣- الـ Policy Enforcement Point بيفحص هل فيه JSON Web Token ساري مدته ١٠ دقائق. ٤- لو مفيش أو منتهي، بيحوّله عبر OpenID Connect للـ Campus Identity Provider.",
            "٥- أول ما الباسورد يطلع صح، بندخل بوابة دورة حياة الحساب (Lifecycle Gate): لو الحساب متخرج أو منتهي أو موقوف بنرفضه فورًا ونبلّغ Monitoring and Response.",
            "٦- لو الحساب نشط، بيعمل Multi-Factor Authentication المناسب لدوره. ٧- الـ Identity Provider بيصدر JSON Web Token موقّع مدته ٥–١٠ دقائق شايل الدور وقوة التحقق والحالة.",
            "وهنا بينتهي دوري في Identity and Authentication عند الخطوة السابعة، وفي الخطوة التامنة بنسلّم التوكن لـ Policy Decision Point عشان يدمجه مع Device Security Posture وAuthorization Policies.",
        ],
    )
    return slide


# ===========================================================================
# SLIDE 6 — Account Lifecycle & Sub-Minute Revocation (< 60s) (bg=vertical)
# ===========================================================================
def _slide_06_lifecycle_and_revocation(prs):
    slide = add_blank_slide(prs, bg="vertical")

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.80,
        width=20.0 - 2 * MARGIN_X,
        height=0.88,
        text="Account Lifecycle & Sub-Minute Revocation",
        font_name=FONT_DISPLAY,
        font_size=48.0,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.68,
        width=17.5,
        height=0.48,
        text="Automated lifecycle sync from admission to graduation — plus full compromise shutdown under 60 seconds.",
        font_name=FONT_BODY,
        font_size=20.0,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    # Top Row: 4 High-Contrast Dark Frosted Lifecycle Stages
    lc_w = 4.02
    lc_h = 1.78
    lc_gap = 0.56
    lc_y = 2.38
    lc_specs = [
        ("01 · Provisioning", "Human Resources & Registrar sync plus automatic guest expiration.", "SAGE_GLASS"),
        ("02 · Active Use", "Short 10-minute JSON Web Token plus 802.1X Wi-Fi session.", "VERIFIED_MINT"),
        ("03 · Role Transition", "Student to Teaching Assistant role refresh on next 10-minute token.", "SOFT_SAGE"),
        ("04 · Offboard / Alert", "Graduated, expired, or security alert triggers instant shutdown.", "ALERT_CORAL"),
    ]
    lc_cards = []
    lc_arrows = []
    for idx, (lt, lb, acc) in enumerate(lc_specs):
        lx = MARGIN_X + idx * (lc_w + lc_gap)
        lc = add_frosted_card(
            slide,
            left=lx,
            top=lc_y,
            width=lc_w,
            height=lc_h,
            fill_color="DARK_TEAL",
            fill_alpha=74.0,
            border_color=acc,
            border_alpha=82.0 if idx in (1, 3) else 60.0,
            border_pt=1.25,
            radius=0.06,
            name=f"LifecycleStage_{idx+1}",
        )
        _fill_card_paragraphs(
            lc,
            heading=lt,
            body=lb,
            head_size=20.0,
            body_size=18.0,
            head_color="WHITE",
            body_color="ICE_MIST",
            margin_x=0.22,
            margin_y=0.16,
            head_space_after=6.0,
            line_spacing=1.16,
        )
        lc_cards.append(lc)
        if idx < 3:
            arr = add_arrow_connector(
                slide,
                x1=lx + lc_w + 0.08,
                y1=lc_y + lc_h / 2.0,
                x2=lx + lc_w + lc_gap - 0.08,
                y2=lc_y + lc_h / 2.0,
                color="WHITE",
                width_pt=2.0,
                name=f"LcArr_{idx+1}",
            )
            lc_arrows.append(arr)

    # Walkthrough Heading + Kill SLA Pill
    walk_hdr = add_textbox(
        slide,
        left=MARGIN_X,
        top=4.44,
        width=12.60,
        height=0.50,
        text="Walkthrough: Phished Student Account (Kareem) — Legacy Cookie vs. Zero Trust",
        font_name=FONT_TITLE,
        font_size=20.0,
        color="WHITE",
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="WalkthroughHeading",
    )
    sla_pill = add_badge(
        slide,
        left=13.88,
        top=4.42,
        width=5.00,
        height=0.52,
        text="FULL REVOCATION UNDER 60 SECONDS",
        fill_color="DARK_TEAL",
        fill_alpha=90.0,
        border_color="VERIFIED_MINT",
        border_alpha=92.0,
        text_color="VERIFIED_MINT",
        font_size=15.0,
        bold=True,
        name="RevocationSlaBadge",
    )

    # Bottom Left: Legacy 8-Hour Cookie Gap (Dark frosted glass for high legibility over sage bottom)
    leg_y = 5.14
    leg_h = 4.96
    leg_w = 5.45
    leg_card = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=leg_y,
        width=leg_w,
        height=leg_h,
        fill_color="ALERT_CORAL",
        fill_alpha=100.0,
        border_color="ALERT_CORAL",
        border_alpha=85.0,
        border_pt=1.5,
        radius=0.028,
        name="LegacyCookieCard",
    )
    tf_leg = leg_card.text_frame
    tf_leg.vertical_anchor = MSO_ANCHOR.TOP
    tf_leg.margin_left = Inches(0.32)
    tf_leg.margin_right = Inches(0.32)
    tf_leg.margin_top = Inches(0.28)

    p_lh = tf_leg.paragraphs[0]
    p_lh.alignment = PP_ALIGN.LEFT
    p_lh.space_after = Pt(14)
    r_lh = p_lh.add_run()
    r_lh.text = "Legacy 8-Hour Cookie Gap"
    r_lh.font.name = FONT_TITLE
    r_lh.font.size = Pt(23.0)
    r_lh.font.bold = True
    r_lh.font.color.rgb = rgb("ALERT_CORAL")

    legacy_bullets = [
        ("10:00 AM:  ", "Kareem clicks a fake portal link; attacker steals his password and 8-hour HTTP cookie."),
        ("10:05 AM:  ", "Information Technology disables Kareem's password inside the database."),
        ("RESULT:  ", "The stolen cookie STILL works until 6:00 PM because the portal never re-checks!"),
    ]
    for b_lead, b_txt in legacy_bullets:
        pb = tf_leg.add_paragraph()
        pb.alignment = PP_ALIGN.LEFT
        pb.space_after = Pt(14)
        pb.line_spacing = 1.22
        r1 = pb.add_run()
        r1.text = f"▸  {b_lead}"
        r1.font.name = FONT_BODY
        r1.font.size = Pt(18.5)
        r1.font.bold = True
        r1.font.color.rgb = rgb("ALERT_CORAL" if b_lead.startswith("RESULT") else "WHITE")
        r2 = pb.add_run()
        r2.text = b_txt
        r2.font.name = FONT_BODY
        r2.font.size = Pt(18.5)
        r2.font.color.rgb = rgb("WHITE" if b_lead.startswith("RESULT") else "ICE_MIST")

    mid_arr = add_arrow_connector(
        slide,
        x1=MARGIN_X + leg_w + 0.14,
        y1=leg_y + leg_h / 2.0,
        x2=MARGIN_X + leg_w + 0.68,
        y2=leg_y + leg_h / 2.0,
        color="DARK_TEAL",
        width_pt=2.8,
        name="LegacyToZtArrow",
    )

    # Bottom Right: 2x2 High-Contrast Sub-Minute Revocation Kill Chain
    rx0 = MARGIN_X + leg_w + 0.82
    rw_card = 5.58
    rh_card = 2.34
    rgx = 0.33
    rgy = 0.28
    rev_specs = [
        (
            "Step 1 · Directory Suspension",
            "Identity Provider marks Kareem's account as SUSPENDED immediately upon receiving an alert from Monitoring and Response.",
        ),
        (
            "Step 2 · Refresh & Single Sign-On Kill",
            "Identity Provider voids all OpenID Connect refresh tokens and closes the central Single Sign-On session, blocking new 10-minute tokens.",
        ),
        (
            "Step 3 · Instant Back-Channel Logout",
            "Identity Provider pushes a signed OpenID Connect logout to the Learning Portal and Lab Policy Enforcement Points so caches flush.",
        ),
        (
            "Step 4 · Instant Wireless Disconnect",
            "Remote Authentication Server sends a Change of Authorization disconnect to the Wireless Access Point, kicking the attacker off Wi-Fi at Layer 2.",
        ),
    ]
    rev_cards = []
    for idx, (rt, rb) in enumerate(rev_specs):
        r, c = divmod(idx, 2)
        cx = rx0 + c * (rw_card + rgx)
        cy = leg_y + r * (rh_card + rgy)
        rc = add_frosted_card(
            slide,
            left=cx,
            top=cy,
            width=rw_card,
            height=rh_card,
            fill_color="DARK_TEAL",
            fill_alpha=82.0,
            border_color="VERIFIED_MINT",
            border_alpha=82.0,
            border_pt=1.35,
            radius=0.05,
            name=f"RevokeStep_{idx+1}",
        )
        _fill_card_paragraphs(
            rc,
            heading=rt,
            body=rb,
            head_size=20.0,
            body_size=18.0,
            head_color="VERIFIED_MINT",
            body_color="WHITE",
            margin_x=0.24,
            margin_y=0.16,
            head_space_after=6.0,
            line_spacing=1.16,
        )
        rev_cards.append(rc)

    add_footer(slide, PART_LABEL)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    for idx, lc in enumerate(lc_cards):
        if idx > 0:
            anim.wipe_in(lc_arrows[idx - 1], direction="right", trigger="with_previous", duration_ms=200, delay_ms=idx * 90)
        anim.fly_in(lc, direction="bottom", trigger="after_previous" if idx == 0 else "with_previous", duration_ms=320, delay_ms=idx * 90)
    anim.fade_in(walk_hdr, trigger="after_previous", duration_ms=280)
    anim.fade_in(sla_pill, trigger="with_previous", duration_ms=280, delay_ms=60)
    anim.fly_in(leg_card, direction="left", trigger="after_previous", duration_ms=360)
    anim.wipe_in(mid_arr, direction="right", trigger="after_previous", duration_ms=240)
    anim.stagger(rev_cards, effect="fly_bottom", duration_ms=340, step_ms=110, on_click_first=False, base_delay_ms=40)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "إزاي نضمن إن مفيش حساب منتهي أو مخترق يفضل شغال؟ ربطنا دورة حياة الحساب بشئون الطلبة وHuman Resources عشان التخرج أو الترقية لـ Teaching Assistant تتحدث تلقائيًا مع أول JSON Web Token جديد.",
            "تخيلوا طالب اسمه كريم اتنصب عليه بلينك مزيف الساعة ١٠ الصبح والهاكر سرق الباسورد وHTTP Cookie مدتها ٨ ساعات. في النظام التقليدي، حتى لو قسم Information Technology قفل الباسورد الساعة ١٠ و٥ دقائق، الكوكي بتفضل شغالة لحد ٦ المغرب!",
            "أما في تصميمنا، بمجرد تنبيه من Monitoring and Response بنقفل كل حاجة في أقل من ٦٠ ثانية في ٤ خطوات: أولًا إيقاف الحساب في الـ Identity Provider؛ ثانيًا إلغاء كل OpenID Connect Refresh Tokens وجلسة Single Sign-On؛ ثالثًا إرسال Back-Channel Logout موقّع للبوابة وPolicy Enforcement Points؛ ورابعًا رسالة Change of Authorization من خادم Remote Authentication للـ Wireless Access Point عشان تفصل الهاكر من الواي فاي على Layer 2.",
        ],
    )
    return slide


# ===========================================================================
# SLIDE 7 — 3-Case Contrast, Resilience & Bridge to Part 03 (bg=horizontal)
# ===========================================================================
def _slide_07_scenarios_and_bridge(prs):
    slide = add_blank_slide(prs, bg="horizontal")

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.78,
        width=20.0 - 2 * MARGIN_X,
        height=0.86,
        text="Why Valid Identity Alone Never Grants Automatic Access",
        font_name=FONT_DISPLAY,
        font_size=46.0,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.64,
        width=17.5,
        height=0.48,
        text="Three campus access scenarios proving that authentication identifies WHO the user is — never blind trust.",
        font_name=FONT_BODY,
        font_size=20.0,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    cw = 5.62
    ch = 5.90
    gap = 0.45
    y_cards = 2.30

    cases = [
        (
            "CASE A  ·  AUTHORIZED LECTURER",
            "FULL ACCESS GRANTED",
            "VERIFIED_MINT",
            [
                ("Identity: ", "Valid faculty password, active account, and Hardware Passkey."),
                ("Claims Issued: ", "Role is Lecturer; verified with Hardware Passkey."),
                ("Connected Signals: ", "Device Security Posture passes; Authorization Policies match."),
                ("Outcome: ", "Opens Learning Portal, Lab Servers, and course grade sheets."),
            ],
        ),
        (
            "CASE B  ·  SPONSORED GUEST",
            "STRICTLY SCOPED ACCESS",
            "SOFT_SAGE",
            [
                ("Identity: ", "Faculty-sponsored login + Time-Based One-Time Password (24-hour window)."),
                ("Claims Issued: ", "Role is Guest; 14 hours remaining before automatic expiration."),
                ("Connected Signals: ", "Unmanaged personal device; isolated guest network zone."),
                ("Outcome: ", "Public Wi-Fi and seminar info only; Labs and Records blocked."),
            ],
        ),
        (
            "CASE C  ·  UNHEALTHY DEVICE",
            "IDENTITY PASS  →  DENIED",
            "ALERT_CORAL",
            [
                ("Identity: ", "Lecturer enters a valid password and a valid Hardware Passkey!"),
                ("Claims Issued: ", "Identity Provider proves identity honestly — never fakes failure."),
                ("Connected Signals: ", "Device Security Posture flags outdated Operating System & no encryption."),
                ("Outcome: ", "Policy Decision Point denies access — Valid Login Is Not Trust."),
            ],
        ),
    ]

    case_groups = []
    for idx, (ctag, verdict, accent, bullets) in enumerate(cases):
        cx = MARGIN_X + idx * (cw + gap)
        card = add_frosted_card(
            slide,
            left=cx,
            top=y_cards,
            width=cw,
            height=ch,
            fill_color="DARK_TEAL",
            fill_alpha=74.0 if idx < 2 else 86.0,
            border_color=accent,
            border_alpha=86.0 if idx != 1 else 74.0,
            border_pt=1.5 if idx == 2 else 1.25,
            radius=0.028,
            name=f"CaseCard_{idx+1}",
        )
        hdr = add_textbox(
            slide,
            left=cx + 0.30,
            top=y_cards + 0.22,
            width=cw - 0.60,
            height=0.42,
            text=ctag,
            font_name=FONT_TITLE,
            font_size=19.0,
            color="WHITE",
            bold=True,
            name=f"CaseHeader_{idx+1}",
        )
        pill = add_badge(
            slide,
            left=cx + 0.30,
            top=y_cards + 0.74,
            width=cw - 0.60,
            height=0.50,
            text=verdict,
            fill_color=accent if idx == 2 else "OBSIDIAN",
            fill_alpha=42.0 if idx == 2 else 74.0,
            border_color=accent,
            border_alpha=92.0,
            text_color="WHITE" if idx != 0 else "VERIFIED_MINT",
            font_size=15.0,
            bold=True,
            name=f"CaseVerdict_{idx+1}",
        )
        body_tb = add_textbox(
            slide,
            left=cx + 0.30,
            top=y_cards + 1.38,
            width=cw - 0.60,
            height=ch - 1.56,
            text="",
            font_name=FONT_BODY,
            font_size=18.0,
            color="ICE_MIST",
            name=f"CaseBullets_{idx+1}",
        )
        tf_b = body_tb.text_frame
        tf_b.clear()
        for b_idx, (lead, txt) in enumerate(bullets):
            p = tf_b.paragraphs[0] if b_idx == 0 else tf_b.add_paragraph()
            p.space_after = Pt(8)
            p.line_spacing = 1.18
            r1 = p.add_run()
            r1.text = f"▸  {lead}"
            r1.font.name = FONT_BODY
            r1.font.size = Pt(18.0)
            r1.font.bold = True
            r1.font.color.rgb = rgb("WHITE")
            r2 = p.add_run()
            r2.text = txt
            r2.font.name = FONT_BODY
            r2.font.size = Pt(18.0)
            r2.font.color.rgb = rgb("ICE_MIST" if idx != 2 or b_idx < 3 else "WHITE")
        case_groups.append((card, hdr, pill, body_tb))

    # Bottom Dual Bar: Resilience Safeguards (Left) + Handoff Bridge to Part 03 (Right)
    bot_y = 8.46
    bot_h = 1.72
    res_card = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=bot_y,
        width=8.85,
        height=bot_h,
        fill_color="DARK_TEAL",
        fill_alpha=82.0,
        border_color="SAGE_GLASS",
        border_alpha=72.0,
        border_pt=1.2,
        radius=0.06,
        name="ResilienceCard",
    )
    _fill_card_paragraphs(
        res_card,
        heading="Operational Resilience & Recovery",
        body=(
            "Replicated Identity Providers behind DNS aliases + synchronized Network Time Protocol clocks "
            "keep tokens valid; lost phones recover via printed backup codes or verified 24-hour bypass."
        ),
        head_size=19.5,
        body_size=17.5,
        head_color="VERIFIED_MINT",
        body_color="WHITE",
        valign=MSO_ANCHOR.MIDDLE,
        margin_x=0.28,
        margin_y=0.12,
        head_space_after=4.0,
        line_spacing=1.15,
    )

    bridge_card = add_solid_card(
        slide,
        left=MARGIN_X + 9.15,
        top=bot_y,
        width=8.61,
        height=bot_h,
        fill_color="ICE_MIST",
        border_color="VERIFIED_MINT",
        border_pt=1.5,
        radius=0.06,
        name="BridgeToPart03Card",
    )
    _fill_card_paragraphs(
        bridge_card,
        heading="Next Pillar  →  Device Security Posture",
        body=(
            "Case C proves verified identity is only half the equation: next, we inspect the security health "
            "of the endpoint itself before granting access to any campus resource."
        ),
        head_size=21.0,
        body_size=17.5,
        head_color="DARK_TEAL",
        body_color="BLACK",
        valign=MSO_ANCHOR.MIDDLE,
        margin_x=0.30,
        margin_y=0.12,
        head_space_after=4.0,
        line_spacing=1.15,
    )

    add_footer(slide, PART_LABEL)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    for idx, (card, hdr, pill, body_tb) in enumerate(case_groups):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 140
        anim.fly_in(card, direction="bottom", trigger=trig, duration_ms=380, delay_ms=dly)
        anim.fade_in(hdr, trigger="with_previous", duration_ms=300, delay_ms=dly + 40)
        anim.fade_in(pill, trigger="with_previous", duration_ms=300, delay_ms=dly + 60)
        anim.fade_in(body_tb, trigger="with_previous", duration_ms=320, delay_ms=dly + 80)
    anim.fade_in(res_card, trigger="after_previous", duration_ms=340)
    anim.fly_in(bridge_card, direction="right", trigger="with_previous", duration_ms=380, delay_ms=100)
    anim.apply()

    _set_rtl_notes(
        slide,
        [
            "مقارنة الحالات التلاتة دي بتثبت جوهر Zero Trust ببساطة جدًا:",
            "في Case A الدكتور سليم ومعاه Hardware Passkey وجهازه سليم فبياخد صلاحياته كاملة على البوابة والمعامل والدرجات.",
            "في Case B الضيف داخل بضمان دكتور وTime-Based One-Time Password ضمن ٢٤ ساعة متبقي منها ١٤ ساعة، لكن جهازه شخصي وغير مُدار فبيتقيد في الواي فاي العام بس.",
            "الأهم هي Case C: هنا الدكتور دخل صح ١٠٠٪ بالباسورد والـ Hardware Passkey، والـ Identity Provider بيقول الحقيقة إنه فعلًا الدكتور وعمره ما يزوّر فشل في الهوية! لكن لأن جزء Device Security Posture اكتشف إن الـ Operating System قديم وتشفير الهارد مقفول، الـ Policy Decision Point رفض دخوله.",
            "وطبعًا عشان نضمن استمرارية الخدمة وقت الامتحانات، بنشغّل نسخ احتياطية من الـ Identity Provider ورا DNS Aliases ومزامنة ساعات بـ Network Time Protocol، ولو طالب ضاع موبايله بيرجع بأكواد طوارئ مطبوعة أو تصريح ٢٤ ساعة بعد إثبات شخصيته بالكارنيه.",
            "وبكده Case C بتوصلنا بسلاسة للركيزة اللي بعدها مع زميلي عمر منير في Device Security Posture عشان نشوف إزاي بنفحص الجهاز نفسه قبل السماح بالوصول. شكرًا يا دكتورة.",
        ],
    )
    return slide


def build_part_02(prs) -> list:
    """Append 7 animated, on-theme slides for Part 02 (Identity & Authentication)."""
    return [
        _slide_01_divider(prs),
        _slide_02_addresses_and_personas(prs),
        _slide_03_sso_protocol_comparison(prs),
        _slide_04_tiered_mfa_and_8021x(prs),
        _slide_05_login_decision_flow(prs),
        _slide_06_lifecycle_and_revocation(prs),
        _slide_07_scenarios_and_bridge(prs),
    ]
