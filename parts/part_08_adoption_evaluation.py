"""Part 08 — Adoption and Evaluation.

Owner: Amr Khaled Hesseny (Amr Khaled Husseini Mohamed Tawfik)
Slide Budget: 6 slides (within 5–6 target, 7 absolute max).
Target Speaking Time: ~4 min 00 sec total (safely under 5 min cap).

In Scope (MUST cover):
- Phased Zero Trust migration strategy for brownfield campuses (4-phase roadmap & CISA ZTMM v2.0).
- University-specific challenges & engineering solutions (academic freedom, BYOD, legacy labs/IoT,
  decentralized IT, limited budgets).
- Evaluation metrics & KPIs (attack surface, lateral containment, MTTD/cutoff, latency, throughput,
  false rejections, usability, TCO/ROI) and pilot trial validation.

Out of Scope (DO NOT re-explain):
- Technical controls of Parts 01–07; deep data classification & encryption architectures (Part 09).

Narrative Bridge:
- Builds on Part 07 (continuous telemetry & automated response) -> Bridges to Part 09 (Data Security).
"""

from __future__ import annotations

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import (
    L03_section_divider,
    L04_wide_visual,
    L05_table_compare,
    L06_chart_insight,
    L09_two_columns,
)
from core.theme import (
    FONT_BODY,
    FONT_TITLE,
    MARGIN_X,
    MARGIN_Y,
    add_arrow_connector,
    add_badge,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_motif,
    add_textbox,
    rgb,
    set_notes,
)

PART_LABEL = "Part 08  ·  Adoption & Evaluation"


def _add_roadmap_slide(prs):
    """Slide 2 (L04): Four-Phase Brownfield Campus Adoption Roadmap."""
    slide, anim, bounds = L04_wide_visual(
        prs,
        title="Four-Phase Campus\nAdoption Roadmap",
        intro=(
            "Incremental rollout keeps lectures, labs, and exam portals uninterrupted — "
            "start with crown-jewel assets, validate KPIs, then expand faculty by faculty."
        ),
        part_label=PART_LABEL,
        stage_card=False,
        transition="fade",
    )

    sx, sy, sw, _ = bounds
    card_w = 3.90
    gap = 0.72
    card_y = sy + 0.10
    card_h = 5.15

    phases = [
        (
            "PHASE 01",
            "Discover &\nClassify",
            "• Inventory users, devices & apps\n"
            "• Map normal campus traffic\n"
            "• Prioritize critical targets",
            18.0,
        ),
        (
            "PHASE 02",
            "Unify Baseline\nTrust",
            "• Consolidate campus SSO\n"
            "• Adaptive MFA for high risk\n"
            "• Baseline managed endpoints",
            24.0,
        ),
        (
            "PHASE 03",
            "Segment &\nShift Access",
            "• Enclave labs & admin ERP\n"
            "• Replace broad VPN with ZTNA\n"
            "• Enforce per-app policies",
            31.0,
        ),
        (
            "PHASE 04",
            "Automate &\nOptimize",
            "• Link telemetry to cutoffs\n"
            "• Tune risk thresholds live\n"
            "• Scale across all faculties",
            38.0,
        ),
    ]

    cards = []
    arrows = []
    cy_mid = card_y + card_h / 2.0

    for idx, (badge_txt, heading, bullets, alpha) in enumerate(phases):
        x = sx + idx * (card_w + gap)
        card = add_frosted_card(
            slide,
            left=x,
            top=card_y,
            width=card_w,
            height=card_h,
            fill_alpha=alpha,
            border_color="VERIFIED_MINT" if idx == 3 else "SAGE_GLASS",
            border_alpha=52.0 if idx == 3 else 34.0,
            radius=0.05,
            name=f"PhaseCard_{idx + 1}",
        )
        badge = add_badge(
            slide,
            left=x + 0.30,
            top=card_y + 0.32,
            width=1.78,
            height=0.48,
            text=badge_txt,
            fill_color="VERIFIED_MINT" if idx == 3 else "WHITE",
            fill_alpha=24.0,
            border_color="VERIFIED_MINT" if idx == 3 else "SAGE_GLASS",
            border_alpha=55.0,
            font_size=16.0,
            name=f"PhaseBadge_{idx + 1}",
        )
        head_tb = add_textbox(
            slide,
            left=x + 0.30,
            top=card_y + 0.98,
            width=card_w - 0.60,
            height=1.05,
            text=heading,
            font_name=FONT_TITLE,
            font_size=26.0,
            color="WHITE",
            bold=True,
            line_spacing=1.10,
            name=f"PhaseTitle_{idx + 1}",
        )
        body_tb = add_textbox(
            slide,
            left=x + 0.30,
            top=card_y + 2.18,
            width=card_w - 0.60,
            height=2.70,
            text=bullets,
            font_name=FONT_BODY,
            font_size=19.5,
            color="WHITE" if idx >= 2 else "ICE_MIST",
            line_spacing=1.35,
            name=f"PhaseBody_{idx + 1}",
        )
        cards.append([card, badge, head_tb, body_tb])

        if idx > 0:
            prev_right = sx + (idx - 1) * (card_w + gap) + card_w
            arr = add_arrow_connector(
                slide,
                x1=prev_right + 0.10,
                y1=cy_mid,
                x2=x - 0.10,
                y2=cy_mid,
                color="VERIFIED_MINT" if idx == 3 else "WHITE",
                width_pt=2.0,
                name=f"PhaseArrow_{idx}",
            )
            arrows.append(arr)

    banner = add_frosted_card(
        slide,
        left=sx,
        top=card_y + card_h + 0.35,
        width=sw,
        height=0.92,
        fill_alpha=20.0,
        border_color="VERIFIED_MINT",
        border_alpha=48.0,
        radius=0.14,
        name="RolloutBanner",
    )
    tf = banner.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = (
        "Rollout Rule:  Pilot on crown-jewel academic systems first  →  "
        "Verify latency & usability KPIs  →  Expand campus-wide"
    )
    r.font.name = FONT_BODY
    r.font.size = Pt(20.0)
    r.font.bold = True
    r.font.color.rgb = rgb("WHITE")

    for idx, grp in enumerate(cards):
        if idx > 0:
            anim.wipe_in(arrows[idx - 1], direction="right", trigger="after_previous", duration_ms=220)
        for j, sh in enumerate(grp):
            trig = "after_previous" if j == 0 else "with_previous"
            anim.fly_in(sh, direction="bottom", trigger=trig, duration_ms=340, delay_ms=j * 40)

    anim.fade_in(banner, trigger="after_previous", duration_ms=360)
    anim.apply()

    set_notes(
        slide,
        "[~50s] تطبيق Zero Trust في الجامعة لازم يكون تدريجي، لأننا مينفعش نغيّر البنية التحتية للشبكة بالكامل أثناء الدراسة.\n\n"
        "عشان كده بنستخدم Four-Phase Roadmap:\n\n"
        "أول Phase هي Discover and Classify، ودي بنعمل فيها Inventory للمستخدمين والأجهزة والتطبيقات، وبنحدد الـ Critical Assets.\n\n"
        "تاني Phase هي Unify Baseline Trust، ودي بنوحد فيها تسجيل الدخول باستخدام SSO، ونطبق Adaptive MFA، ونتأكد من حالة الأجهزة المُدارة.\n\n"
        "تالت Phase هي Segment and Shift Access، ودي بنعزل المعامل والأنظمة الإدارية المهمة، ونستبدل الـ Broad VPN بالـ ZTNA، بحيث كل مستخدم يوصل للتطبيق المصرح له فقط.\n\n"
        "آخر Phase هي Automate and Optimize، ودي بنربط الـ Monitoring بالـ Automated Session Cutoff، ونحسن الـ Risk Thresholds، وبعدها نوسع التطبيق لباقي الكليات.\n\n"
        "والقاعدة الأساسية إننا نبدأ بالأنظمة الأكثر أهمية، نختبر النتائج، وبعدها نتوسع.",
    )
    return slide


def _add_challenges_slide(prs):
    """Slide 4 (Custom Radial): University-Specific Challenges & Engineering Solutions."""
    slide = add_blank_slide(prs, bg="radial")
    glow = add_motif(slide, "glow", left=13.4, top=-2.6, width=8.6)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=MARGIN_Y,
        width=10.0,
        height=1.85,
        text="Campus Challenges &\nEngineering Solutions",
        font_name=FONT_TITLE,
        font_size=50.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    desc = add_textbox(
        slide,
        left=11.45,
        top=1.08,
        width=7.43,
        height=1.70,
        text=(
            "Universities balance open scholarship with tight public budgets — "
            "policies must protect critical assets without obstructing academic work."
        ),
        font_name=FONT_BODY,
        font_size=22.0,
        color="ICE_MIST",
        line_spacing=1.24,
        name="SlideIntro",
    )

    items = [
        (
            "01 · Academic Freedom",
            "Friction pushback from faculty & students",
            "Background risk scoring & SSO — trigger step-up MFA only for sensitive actions",
        ),
        (
            "02 · Student BYOD Diversity",
            "Thousands of unmanaged personal devices",
            "Isolate unmanaged endpoints; deliver browser-based clientless ZTNA to LMS",
        ),
        (
            "03 · Legacy Labs & IoT",
            "Unsupported gear cannot run EDR agents",
            "Place legacy lab equipment in micro-segmented enclaves behind gateway PEPs",
        ),
        (
            "04 · Decentralized IT",
            "Autonomous faculty & department admins",
            "Federated governance: central identity & baseline policy, local faculty operations",
        ),
        (
            "05 · Limited Budgets",
            "Strict capex & licensing constraints",
            "Reuse existing 802.1X switches, prioritize critical assets & mature open-source",
        ),
    ]

    row_y0 = 3.15
    row_h = 1.18
    gap_y = 0.18
    chal_x = MARGIN_X
    chal_w = 5.55
    sol_x = 7.62
    sol_w = 20.0 - MARGIN_X - sol_x  # 11.26"

    row_groups = []
    for idx, (ctitle, csub, sol_txt) in enumerate(items):
        y = row_y0 + idx * (row_h + gap_y)

        c_card = add_frosted_card(
            slide,
            left=chal_x,
            top=y,
            width=chal_w,
            height=row_h,
            fill_alpha=22.0,
            border_color="ALERT_CORAL",
            border_alpha=48.0,
            radius=0.08,
            name=f"ChallengeCard_{idx + 1}",
        )
        tf_c = c_card.text_frame
        tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_c.margin_left = Inches(0.28)
        tf_c.margin_right = Inches(0.22)
        p1 = tf_c.paragraphs[0]
        p1.space_after = Pt(4)
        r1 = p1.add_run()
        r1.text = ctitle
        r1.font.name = FONT_TITLE
        r1.font.size = Pt(22.0)
        r1.font.bold = True
        r1.font.color.rgb = rgb("WHITE")
        p2 = tf_c.add_paragraph()
        r2 = p2.add_run()
        r2.text = csub
        r2.font.name = FONT_BODY
        r2.font.size = Pt(18.0)
        r2.font.color.rgb = rgb("ICE_MIST")

        arr = add_arrow_connector(
            slide,
            x1=chal_x + chal_w + 0.12,
            y1=y + row_h / 2.0,
            x2=sol_x - 0.12,
            y2=y + row_h / 2.0,
            color="VERIFIED_MINT",
            width_pt=2.0,
            name=f"ChallengeArrow_{idx + 1}",
        )

        s_card = add_frosted_card(
            slide,
            left=sol_x,
            top=y,
            width=sol_w,
            height=row_h,
            fill_alpha=28.0,
            border_color="VERIFIED_MINT",
            border_alpha=48.0,
            radius=0.08,
            name=f"SolutionCard_{idx + 1}",
        )
        tf_s = s_card.text_frame
        tf_s.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_s.margin_left = Inches(0.32)
        tf_s.margin_right = Inches(0.28)
        ps = tf_s.paragraphs[0]
        rs = ps.add_run()
        rs.text = sol_txt
        rs.font.name = FONT_BODY
        rs.font.size = Pt(20.5)
        rs.font.bold = False
        rs.font.color.rgb = rgb("WHITE")

        row_groups.append((c_card, arr, s_card))

    add_footer(slide, PART_LABEL)

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(glow, trigger="on_load", duration_ms=420)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    anim.fade_in(desc, trigger="with_previous", duration_ms=340, delay_ms=80)
    for c_card, arr, s_card in row_groups:
        anim.fly_in(c_card, direction="left", trigger="after_previous", duration_ms=260)
        anim.wipe_in(arr, direction="right", trigger="with_previous", duration_ms=200, delay_ms=60)
        anim.fly_in(s_card, direction="right", trigger="with_previous", duration_ms=280, delay_ms=90)
    anim.apply()

    set_notes(
        slide,
        "[~50s] تطبيق Zero Trust في الجامعات بيواجه خمس تحديات رئيسية.\n\n"
        "أولًا، الحفاظ على Academic Freedom، وده بنعالجه باستخدام SSO وBackground Risk Scoring، بحيث نطلب Step-up MFA عند العمليات الحساسة فقط.\n\n"
        "ثانيًا، تنوع أجهزة الطلاب أو BYOD، وده بنعالجه بعزل الأجهزة غير المُدارة وتوفير Clientless ZTNA للوصول للـ LMS.\n\n"
        "ثالثًا، الأجهزة القديمة في المعامل والـ IoT، ودي ممكن متدعمش Security Agents، فبنحطها داخل Micro-segmented Enclaves.\n\n"
        "رابعًا، استقلال إدارات الكليات، والحل هو Federated Governance: سياسات أساسية مركزية مع إدارة محلية.\n\n"
        "وأخيرًا، محدودية الميزانية، وده بنواجهه بإعادة استخدام أجهزة الشبكة الموجودة، وحماية الأنظمة الأهم أولًا، والاستفادة من Open-source Tools.",
    )
    return slide


def build_part_08(prs) -> list:
    """Append 6 on-theme, animated slides for Part 08 (Adoption & Evaluation)."""
    slides = []

    # Slide 1: Section Divider (L03)
    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 08  ·  STRATEGY & METRICS",
        title="Adoption &\nEvaluation",
        summary=(
            "Turning continuous monitoring and telemetry into a practical brownfield rollout — "
            "balancing phased maturity milestones, campus realities, and measurable KPIs."
        ),
        presenter="Amr Khaled Hesseny",
        motif="crescents",
        transition="push_left",
    )
    set_notes(
        s1,
        "[~30s] السلام عليكم. أنا عمرو، والجزء بتاعي هو Adoption and Evaluation في Zero Trust Architecture for University Campus Networks.\n\n"
        "بعد ما Part 7 اتكلم عن Continuous Monitoring وAutomated Response، الجزء بتاعي بيركز على سؤالين: إزاي نطبق Zero Trust على شبكة الجامعة الحالية من غير ما نعطل العملية التعليمية؟ وإزاي نقيس نجاح التطبيق من ناحية Security وPerformance وUsability؟",
    )
    slides.append(s1)

    # Slide 2: Four-Phase Adoption Roadmap (L04 Custom Pipeline)
    slides.append(_add_roadmap_slide(prs))

    # Slide 3: CISA Zero Trust Maturity Model v2.0 (L05 Table)
    s3, _, _ = L05_table_compare(
        prs,
        title="CISA Maturity\nModel v2.0",
        intro=(
            "Tracking campus progression across five technology pillars and four maturity stages — "
            "governed by cross-cutting visibility, automation, and policy."
        ),
        headers=["Pillar", "1 · Traditional", "2 · Initial", "3 · Advanced", "4 · Optimal"],
        rows=[
            ("Identity", "Local passwords & silos", "Central SSO & basic MFA", "Risk-based adaptive MFA", "Continuous session re-auth"),
            ("Devices", "Untracked BYOD & labs", "Basic asset inventory", "Posture & compliance gates", "Automated health quarantine"),
            ("Networks", "Flat campus VLANs", "Macro zone separation", "Micro-segments & TLS", "Dynamic software-defined zones"),
            ("Applications", "Perimeter & broad VPN", "Gateway access rules", "Per-app ZTNA / PEP", "Per-session adaptive access"),
            ("Data", "Static & unclassified", "Manual tier labeling", "Policy-enforced DLP", "Automated encryption & IRM"),
        ],
        col_widths=[2.56, 3.80, 3.80, 3.80, 3.80],
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        s3,
        "[~40s] علشان نقيس مدى تطور التطبيق، بنستخدم CISA Zero Trust Maturity Model v2.0.\n\n"
        "الموديل بيقيّم خمس Pillars: Identity, Devices, Networks, Applications, and Data.\n\n"
        "وكل Pillar بيمر بأربع مراحل: Traditional, Initial, Advanced, وOptimal.\n\n"
        "وده بيساعد الجامعة تحدد مستواها الحالي، وتعرف إيه الخطوات المطلوبة للوصول لمستوى أعلى من الـ Automation والـ Security.",
    )
    slides.append(s3)

    # Slide 4: Campus Adoption Challenges & Engineering Solutions (Custom Radial)
    slides.append(_add_challenges_slide(prs))

    # Slide 5: Evaluation Framework & Key Engineering Metrics (L06 Native Chart)
    s5, _, _ = L06_chart_insight(
        prs,
        title="Evaluation Framework\n& Key Metrics",
        insight=(
            "Success balances security, speed & TCO:\n\n"
            "▸ MTTD < 60s; cutoff < 15s\n"
            "▸ Setup delay: +10–16 ms\n"
            "▸ False blocks < 1.5%\n"
            "▸ Lower helpdesk & capex TCO\n"
            "▸ Tested via iPerf3 & breach drills"
        ),
        categories=[
            "Attack Surface\nHidden",
            "Lateral Breach\nContained",
            "Throughput\nRetained",
            "Valid Access\nPass Rate",
        ],
        series_name="Target Benchmark (%)",
        values=[95.0, 92.0, 92.0, 98.5],
        chart_title="Campus ZTA Target Benchmarks (%)",
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        s5,
        "[~50s] بعد التطبيق لازم نقيس كفاءته باستخدام Engineering Metrics.\n\n"
        "من ناحية Security، بنستهدف إخفاء أكتر من 95% من الـ Internal Hosts، واحتواء أكتر من 90% من محاولات Lateral Movement، مع MTTD أقل من 60 ثانية وSession Cutoff أقل من 15 ثانية.\n\n"
        "من ناحية Network Performance، الزيادة في Session Setup Delay تكون من 10 إلى 16 milliseconds، مع الحفاظ على Throughput لا يقل عن 92% من المعدل الأساسي.\n\n"
        "أما User Experience، فبنستهدف إن نسبة False Blocks تكون أقل من 1.5%، مع تقليل Helpdesk Tickets والتكلفة الكلية TCO.\n\n"
        "وبنختبر النتائج باستخدام iPerf3 وBreach Drills.",
    )
    slides.append(s5)

    # Slide 6: Practical Campus Pilot & Handoff to Part 09 (L09 Two Columns)
    s6, _, _ = L09_two_columns(
        prs,
        title="Campus Pilot Trial\n& Next Frontier",
        subtitle=(
            "Validating containment, exam-peak load, and usability before full rollout — "
            "because every phase and metric ultimately protects core university data."
        ),
        left_badge="LOW FRICTION",
        left_col=(
            "Student BYOD  ·  LMS Portal",
            [
                "Connects via isolated BYOD onboarding segment",
                "Browser-based clientless ZTNA — zero agent install",
                "Benchmarked at registration peak for >= 92% throughput",
                "Tunes login friction and tracks < 1.5% false blocks",
            ],
        ),
        right_badge="HIGH ASSURANCE",
        right_col=(
            "Faculty  ·  Grading Control",
            [
                "Requires managed workstation & strict posture pass",
                "Verifies committee role plus adaptive step-up MFA",
                "Temporary encrypted PEP session with < 15s cutoff",
                "Anchors directly into data classification & DLP tiers",
            ],
        ),
        left_accent="SAGE_GLASS",
        right_accent="VERIFIED_MINT",
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        s6,
        "[~45s] قبل ما نعمم النظام على الجامعة، بنعمل Pilot Trial على حالتين.\n\n"
        "الحالة الأولى طالب بيستخدم Personal Laptop للوصول للـ LMS، وبيتم توجيهه من خلال Isolated BYOD Segment وClientless ZTNA، وبنختبر الأداء وقت ضغط التسجيل، مع الحفاظ على Throughput لا يقل عن 92%.\n\n"
        "الحالة التانية عضو هيئة تدريس بيحاول يوصل للـ Grading Control Server. هنا لازم يستخدم Managed Workstation، ويجتاز Device Posture Check وAdaptive MFA، وبعدها بياخد Temporary Encrypted Session من خلال الـ PEP، مع إمكانية إنهاء الاتصال في أقل من 15 ثانية.\n\n"
        "وفي النهاية، كل مراحل التطبيق والاختبارات دي هدفها حماية بيانات الجامعة والأنظمة الحساسة، وده بيمهد للجزء التالي Part 9: Data Security and Classification.\n\n"
        "Thank you for listening.",
    )
    slides.append(s6)

    return slides
