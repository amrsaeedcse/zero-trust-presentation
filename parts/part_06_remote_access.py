"""Part 06 — Remote Resource Access.

Owner: Amr Abdelazeem Saeed Mohamed (Repo Lead)
Slide Budget: 6 slides (strictly within 5–6 target, <= 7 hard cap).

Slide Arc (Clear, Intuitive Networking Story — 6 Distinct Layouts):
1. L03 Divider        — Remote Access to University Systems (LMS, Labs, SIS, Servers; VPN -> ZTNA)
2. L04 Custom Diagram — Why Legacy Campus VPNs Fail (Encapsulated Outer/Inner Packet + 4 plain flaws)
3. Custom Clean Flow  — How Campus ZTNA Works (0 Open Inbound Ports + Inside-Out App Connector)
4. Custom Protocol    — Hiding Non-Web Gateways with SPA (Why SPA for faculty apps + UDP Packet fields)
5. L09 Two Columns    — Two Remote Access Modes (Students: Browser-Only IAP vs. Faculty: Agent + SPA)
6. L05 Table + Bridge — Legacy VPN vs. Campus ZTNA + Bridge to Part 7 (Monitoring & Response)

Out of Scope (treated as already-established inputs — NOT re-explained):
- MFA/SSO mechanics (P2), MDM/EDR checks (P3), Policy Engine internals (P4),
  campus VLAN micro-segmentation (P5), SIEM/SOAR playbooks (P7), rollout KPIs (P8), data tiers (P9).
"""

from __future__ import annotations

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import (
    L03_section_divider,
    L04_wide_visual,
    L05_table_compare,
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
    add_solid_card,
    add_textbox,
    rgb,
    set_notes,
)

PART_LABEL = "Part 06  ·  Remote Resource Access"


def _add_node(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    title: str,
    subtitle: str,
    fill_color: str = "DARK_TEAL",
    fill_alpha: float = 64.0,
    border_color: str = "SAGE_GLASS",
    border_alpha: float = 55.0,
    border_pt: float = 1.25,
    title_color: str = "WHITE",
    sub_color: str = "ICE_MIST",
    title_fs: float = 20.0,
    sub_fs: float = 16.5,
    name: str = "ArchNode",
):
    """Crisp, high-contrast on-theme architecture node."""
    card = add_frosted_card(
        slide,
        left=x,
        top=y,
        width=w,
        height=h,
        fill_color=fill_color,
        fill_alpha=fill_alpha,
        border_color=border_color,
        border_alpha=border_alpha,
        border_pt=border_pt,
        radius=0.10,
        name=name,
    )
    tf = card.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.10)
    tf.margin_bottom = Inches(0.10)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(4 if subtitle else 0)
    r1 = p1.add_run()
    r1.text = title
    r1.font.name = FONT_TITLE
    r1.font.size = Pt(title_fs)
    r1.font.bold = True
    r1.font.color.rgb = rgb(title_color)

    if subtitle:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.line_spacing = 1.16
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.name = FONT_BODY
        r2.font.size = Pt(sub_fs)
        r2.font.color.rgb = rgb(sub_color)
    return card


def _build_slide_01_divider(prs):
    slide, _ = L03_section_divider(
        prs,
        part_badge="PART 06  ·  REMOTE RESOURCE ACCESS",
        title="Remote Access to\nUniversity Systems",
        summary=(
            "How students, faculty, and researchers securely access university systems "
            "(LMS, labs, grading SIS, and servers) from off campus — replacing legacy VPNs "
            "with Zero Trust Network Access (ZTNA)."
        ),
        presenter="Amr Abdelazeem Saeed Mohamed",
        motif="signal_waves",
        transition="push_left",
    )
    set_notes(
        slide,
        "Bridge from Part 5: Inside campus, micro-segmentation isolates internal zones. "
        "Now Part 6 answers how students, faculty, and researchers connect to university systems "
        "(LMS, virtual labs, SIS grades, ERP, and research servers) from home or abroad — "
        "replacing traditional campus VPNs with Zero Trust Network Access (ZTNA).",
    )
    return slide


def _build_slide_02_legacy_vpn(prs):
    slide, anim, (sx, sy, sw, sh) = L04_wide_visual(
        prs,
        title="Why Legacy\nCampus VPNs Fail",
        intro=(
            "A VPN encapsulates your home packet and assigns an internal campus IP — "
            "letting your device reach any host on the university network."
        ),
        part_label=PART_LABEL,
        stage_card=False,
        transition="fade",
    )

    diag_y = sy + 0.05
    diag_h = 3.65
    diag_bg = add_frosted_card(
        slide,
        sx,
        diag_y,
        sw,
        diag_h,
        fill_color="OBSIDIAN",
        fill_alpha=32.0,
        border_color="SAGE_GLASS",
        border_alpha=35.0,
        radius=0.03,
        name="VpnFlowStage",
    )
    diag_lbl = add_badge(
        slide,
        left=sx + 0.30,
        top=diag_y + 0.22,
        width=6.40,
        height=0.46,
        text="ENCAPSULATED VPN PACKET  →  ANY CAMPUS HOST",
        fill_color="ALERT_CORAL",
        fill_alpha=24.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        font_size=15.0,
        name="VpnExposureBadge",
    )

    # Node 1: Home Laptop
    n_home = _add_node(
        slide,
        x=sx + 0.32,
        y=diag_y + 1.02,
        w=3.15,
        h=1.98,
        title="Home Laptop",
        subtitle="Home Wi-Fi · Personal PC\n(May Have Malware)",
        fill_color="DARK_TEAL",
        fill_alpha=68.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        name="VpnNodeHome",
    )

    # Arrow 1: Outer Packet
    arr1 = add_arrow_connector(
        slide, sx + 3.55, diag_y + 2.02, sx + 5.45, diag_y + 2.02, color="ALERT_CORAL", width_pt=2.2, name="VpnArr1"
    )
    arr1_lbl = add_textbox(
        slide,
        sx + 3.50,
        diag_y + 1.10,
        2.00,
        0.82,
        "1. Outer Packet\nHome → VPN Server",
        font_size=14.5,
        color="ICE_MIST",
        align=PP_ALIGN.CENTER,
        name="VpnArr1Lbl",
    )

    # Node 2: Public VPN Server
    n_vpn = _add_node(
        slide,
        x=sx + 5.53,
        y=diag_y + 1.02,
        w=3.20,
        h=1.98,
        title="Campus VPN Server",
        subtitle="Open Public Ports\nGives Internal Campus IP",
        fill_color="DARK_TEAL",
        fill_alpha=68.0,
        border_color="ALERT_CORAL",
        border_alpha=80.0,
        name="VpnNodeServer",
    )

    # Arrow 2: Inner Packet
    arr2 = add_arrow_connector(
        slide, sx + 8.81, diag_y + 2.02, sx + 10.75, diag_y + 2.02, color="ALERT_CORAL", width_pt=2.2, name="VpnArr2"
    )
    arr2_lbl = add_textbox(
        slide,
        sx + 8.76,
        diag_y + 1.10,
        2.05,
        0.82,
        "2. Inner Packet\nCampus IP (10.20.x.x)",
        font_size=14.5,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="VpnArr2Lbl",
    )

    # Right Zone: University Internal Network
    lan_box = add_frosted_card(
        slide,
        sx + 10.85,
        diag_y + 0.35,
        sw - 11.18,
        diag_h - 0.65,
        fill_color="DARK_TEAL",
        fill_alpha=65.0,
        border_color="ALERT_CORAL",
        border_alpha=65.0,
        border_pt=1.25,
        radius=0.05,
        name="VpnCampusLan",
    )
    lan_hdr = add_textbox(
        slide,
        sx + 11.00,
        diag_y + 0.48,
        sw - 11.48,
        0.42,
        "UNIVERSITY NETWORK  —  ANY HOST REACHABLE",
        font_size=14.5,
        color="ALERT_CORAL",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="VpnLanHeader",
    )

    tx = sx + 11.15
    tw = sw - 11.78
    t_lab = _add_node(
        slide,
        tx,
        diag_y + 1.00,
        tw,
        0.64,
        "Linux Lab Server  (Needed)",
        "",
        fill_color="DARK_TEAL",
        fill_alpha=55.0,
        border_color="VERIFIED_MINT",
        border_alpha=80.0,
        title_fs=17.0,
        name="VpnTargetLab",
    )
    t_sis = _add_node(
        slide,
        tx,
        diag_y + 1.76,
        tw,
        0.64,
        "SIS Grades Server  (Exposed!)",
        "",
        fill_color="ALERT_CORAL",
        fill_alpha=28.0,
        border_color="ALERT_CORAL",
        border_alpha=90.0,
        title_fs=17.0,
        name="VpnTargetSis",
    )
    t_erp = _add_node(
        slide,
        tx,
        diag_y + 2.52,
        tw,
        0.64,
        "Finance ERP & Admin  (Exposed!)",
        "",
        fill_color="ALERT_CORAL",
        fill_alpha=28.0,
        border_color="ALERT_CORAL",
        border_alpha=90.0,
        title_fs=17.0,
        name="VpnTargetErp",
    )

    # Bottom 4 plain-English engineering flaw cards
    flaws = [
        ("1. Full Network Trust", "Assigns a campus IP at Network Layer (L3), letting malware scan any internal host."),
        ("2. Open Public Ports", "VPN server exposes public ports (UDP 500 / TCP 443) to scans, CVEs & DDoS."),
        ("3. Traffic Bottleneck", "Full-tunnel slows campus internet; split-tunnel turns your PC into an open bridge."),
        ("4. Checked Once Only", "Verifies login once at start — stays open even if device is infected later."),
    ]
    cw = (sw - 3 * 0.36) / 4.0
    cy = diag_y + diag_h + 0.36
    ch = 2.54
    flaw_cards = []
    for idx, (ftitle, fbody) in enumerate(flaws):
        fx = sx + idx * (cw + 0.36)
        card = add_frosted_card(
            slide,
            fx,
            cy,
            cw,
            ch,
            fill_color="DARK_TEAL",
            fill_alpha=66.0,
            border_color="SAGE_GLASS",
            border_alpha=45.0,
            radius=0.06,
            name=f"VpnFlaw_{idx+1}",
        )
        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.24)
        tf.margin_right = Inches(0.24)
        tf.margin_top = Inches(0.24)
        p1 = tf.paragraphs[0]
        p1.space_after = Pt(10)
        r1 = p1.add_run()
        r1.text = ftitle
        r1.font.name = FONT_TITLE
        r1.font.size = Pt(20.5)
        r1.font.bold = True
        r1.font.color.rgb = rgb("WHITE")
        p2 = tf.add_paragraph()
        p2.line_spacing = 1.22
        r2 = p2.add_run()
        r2.text = fbody
        r2.font.name = FONT_BODY
        r2.font.size = Pt(17.5)
        r2.font.color.rgb = rgb("ICE_MIST")
        flaw_cards.append(card)

    # Smooth automatic entrance choreography on slide load (1 click = next slide, zero mid-slide clicks)
    anim.fade_in(diag_bg, trigger="after_previous", duration_ms=260)
    anim.fade_in(diag_lbl, trigger="with_previous", duration_ms=260)
    anim.fade_in(n_home, trigger="after_previous", duration_ms=260)
    for sh in (arr1, arr1_lbl, n_vpn):
        anim.fade_in(sh, trigger="after_previous" if sh is arr1 else "with_previous", duration_ms=260, delay_ms=40)
    for sh in (arr2, arr2_lbl, lan_box, lan_hdr, t_lab):
        anim.fade_in(sh, trigger="after_previous" if sh is arr2 else "with_previous", duration_ms=260, delay_ms=40)
    anim.fade_in(t_sis, trigger="after_previous", duration_ms=240)
    anim.fade_in(t_erp, trigger="with_previous", duration_ms=240, delay_ms=80)
    anim.stagger(flaw_cards, effect="fly_bottom", duration_ms=320, step_ms=90, on_click_first=False, base_delay_ms=60)
    anim.apply()

    set_notes(
        slide,
        "Explain the Encapsulated Packet intuitively: "
        "1) From home Wi-Fi, your laptop sends an Outer Packet to the campus VPN server. "
        "2) Inside it, the VPN gives your laptop an internal campus IP (10.20.x.x) to send an Inner Packet to the university network. "
        "3) Because you have an internal IP on Layer 3, you (or malware on your PC) can reach any host on that network — not just the Linux Lab, "
        "also SIS Grades and Finance ERP. "
        "4) Summarize the 4 clear flaws below.",
    )
    return slide


def _build_slide_03_ztna_architecture(prs):
    """Super clean, intuitive left-to-right ZTNA diagram with zero visual clutter."""
    slide = add_blank_slide(prs, bg="diagonal")

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=MARGIN_Y,
        width=17.76,
        height=1.05,
        text="How Campus ZTNA Works",
        font_name=FONT_TITLE,
        font_size=56.0,
        color="WHITE",
        bold=True,
        name="ZtnaTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=2.08,
        width=17.76,
        height=0.58,
        text="Remote users talk to a ZTNA Proxy over HTTPS — while the university router accepts zero incoming connections.",
        font_name=FONT_BODY,
        font_size=22.5,
        color="ICE_MIST",
        name="ZtnaSubtitle",
    )

    # Clean 5-stage horizontal flow on y=3.05..8.75 (huge whitespace, super easy to trace!)
    flow_y = 3.05
    flow_h = 5.65

    # Stage 1: Remote User (Off-Campus)
    n_user = _add_node(
        slide,
        x=MARGIN_X,
        y=flow_y + 1.35,
        w=3.10,
        h=2.95,
        title="Remote User",
        subtitle="Student (Browser)\nor Faculty (App)\n\nNo Campus IP",
        fill_color="DARK_TEAL",
        fill_alpha=70.0,
        border_color="SAGE_GLASS",
        border_alpha=65.0,
        title_fs=22.0,
        sub_fs=17.0,
        name="ZtnaUser",
    )

    # Arrow 1: User -> ZTNA Proxy (HTTPS / TLS Request)
    arr1 = add_arrow_connector(
        slide, x1=4.30, y1=flow_y + 2.82, x2=5.70, y2=flow_y + 2.82, color="WHITE", width_pt=2.5, name="ArrUserToProxy"
    )
    lbl1 = add_textbox(
        slide, 4.25, flow_y + 1.85, 1.50, 0.85, "HTTPS / TLS\nRequest", font_size=15.5, color="WHITE", bold=True, align=PP_ALIGN.CENTER
    )

    # Stage 2: ZTNA Broker / Proxy
    n_proxy = _add_node(
        slide,
        x=5.78,
        y=flow_y + 0.85,
        w=3.45,
        h=3.95,
        title="ZTNA Proxy\n(Broker)",
        subtitle="Checks Identity, Device\n& Policy (Parts 2–4)\n\nAllows 1 App Only",
        fill_color="DARK_TEAL",
        fill_alpha=78.0,
        border_color="VERIFIED_MINT",
        border_alpha=90.0,
        border_pt=1.5,
        title_fs=22.5,
        sub_fs=17.0,
        name="ZtnaProxy",
    )

    # Stage 3: Campus Edge Firewall (Zero Open Inbound Ports!)
    fw_x = 10.10
    fw_w = 1.65
    fw_bar = add_solid_card(
        slide,
        left=fw_x,
        top=flow_y + 0.20,
        width=fw_w,
        height=flow_h - 0.40,
        fill_color="DARK_TEAL",
        border_color="VERIFIED_MINT",
        border_pt=1.75,
        radius=0.12,
        name="CampusFirewall",
    )
    fw_zero = add_textbox(
        slide, fw_x + 0.08, flow_y + 2.55, fw_w - 0.16, 0.90, "0", font_name=FONT_TITLE, font_size=46.0, color="VERIFIED_MINT", bold=True, align=PP_ALIGN.CENTER
    )
    fw_lbl = add_textbox(
        slide,
        fw_x + 0.08,
        flow_y + 3.48,
        fw_w - 0.16,
        1.80,
        "OPEN\nINBOUND\nPORTS\n\nCampus Router",
        font_size=14.5,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.CENTER,
        line_spacing=1.18,
    )

    # Stage 4: Internal Campus App Connector
    conn_x = 12.60
    conn_w = 2.95
    n_conn = _add_node(
        slide,
        x=conn_x,
        y=flow_y + 0.85,
        w=conn_w,
        h=3.95,
        title="App Connector",
        subtitle="Sits inside campus\n\nKeeps persistent\nOUTBOUND TCP + mTLS\nlink to Proxy",
        fill_color="DARK_TEAL",
        fill_alpha=75.0,
        border_color="VERIFIED_MINT",
        border_alpha=85.0,
        border_pt=1.5,
        title_fs=21.5,
        sub_fs=16.5,
        name="ZtnaConnector",
    )

    # SINGLE crystal-clear LEFT-pointing arrow from App Connector across Firewall to ZTNA Proxy!
    arr_out = add_arrow_connector(
        slide,
        x1=conn_x - 0.06,
        y1=flow_y + 1.50,
        x2=5.78 + 3.45 + 0.06,
        y2=flow_y + 1.50,
        color="VERIFIED_MINT",
        width_pt=2.8,
        head_end=False,
        tail_end=True,
        name="ArrInsideOut",
    )
    lbl_out = add_badge(
        slide,
        left=9.28,
        top=flow_y + 0.68,
        width=3.32,
        height=0.52,
        text="← Outbound mTLS (Inside → Out)",
        fill_color="OBSIDIAN",
        fill_alpha=90.0,
        border_color="VERIFIED_MINT",
        border_alpha=90.0,
        text_color="VERIFIED_MINT",
        font_size=13.0,
        name="LblInsideOut",
    )

    # Stage 5: Campus Apps (Allowed vs Hidden)
    app_x = 16.25
    app_w = 20.0 - MARGIN_X - app_x
    arr_app = add_arrow_connector(
        slide, x1=conn_x + conn_w + 0.06, y1=flow_y + 1.85, x2=app_x - 0.06, y2=flow_y + 1.85, color="VERIFIED_MINT", width_pt=2.4, name="ArrToApp"
    )
    n_ok = _add_node(
        slide,
        x=app_x,
        y=flow_y + 0.90,
        w=app_w,
        h=1.90,
        title="Allowed App",
        subtitle="e.g., Linux Lab\nor SIS Only",
        fill_color="VERIFIED_MINT",
        fill_alpha=24.0,
        border_color="VERIFIED_MINT",
        border_alpha=90.0,
        title_fs=19.5,
        sub_fs=16.0,
        name="ZtnaAllowedApp",
    )
    n_dark = _add_node(
        slide,
        x=app_x,
        y=flow_y + 3.20,
        w=app_w,
        h=1.65,
        title="Other Servers",
        subtitle="Hidden &\nUnreachable",
        fill_color="OBSIDIAN",
        fill_alpha=58.0,
        border_color="SAGE_GLASS",
        border_alpha=32.0,
        title_color="ICE_MIST",
        sub_color="SOFT_SAGE",
        title_fs=18.5,
        sub_fs=15.5,
        name="ZtnaHiddenApps",
    )

    # Bottom plain-English takeaway strip explaining WHY Inside-Out works!
    key_bar = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=9.12,
        width=20.0 - 2 * MARGIN_X,
        height=0.98,
        fill_color="DARK_TEAL",
        fill_alpha=72.0,
        border_color="VERIFIED_MINT",
        border_alpha=70.0,
        border_pt=1.25,
        radius=0.14,
        name="ZtnaKeyTakeaway",
    )
    tf_k = key_bar.text_frame
    tf_k.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_k.margin_left = Inches(0.30)
    tf_k.margin_right = Inches(0.30)
    pk = tf_k.paragraphs[0]
    pk.alignment = PP_ALIGN.CENTER
    rk1 = pk.add_run()
    rk1.text = "0 INBOUND PORTS:  "
    rk1.font.name = FONT_TITLE
    rk1.font.size = Pt(18.0)
    rk1.font.bold = True
    rk1.font.color.rgb = rgb("VERIFIED_MINT")
    rk2 = pk.add_run()
    rk2.text = "Nobody on the internet can reach the router — ONLY the ZTNA Proxy talks inside via the connector's persistent outbound mTLS link."
    rk2.font.name = FONT_BODY
    rk2.font.size = Pt(17.5)
    rk2.font.color.rgb = rgb("WHITE")

    add_footer(slide, PART_LABEL)

    # Smooth automatic flow choreography on slide load (1 click = next slide, zero mid-slide clicks)
    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=300)
    anim.fade_in(sub, trigger="with_previous", duration_ms=300, delay_ms=60)
    anim.stagger([fw_bar, fw_zero, fw_lbl], effect="fade", duration_ms=280, step_ms=50, on_click_first=False, base_delay_ms=40)
    for idx, sh in enumerate((n_conn, arr_out, lbl_out, n_proxy)):
        anim.fade_in(sh, trigger="after_previous" if idx == 0 else "with_previous", duration_ms=280, delay_ms=idx * 60)
    for idx, sh in enumerate((n_user, arr1, lbl1)):
        anim.fade_in(sh, trigger="after_previous" if idx == 0 else "with_previous", duration_ms=280, delay_ms=idx * 60)
    for idx, sh in enumerate((arr_app, n_ok, n_dark)):
        anim.fade_in(sh, trigger="after_previous" if idx == 0 else "with_previous", duration_ms=280, delay_ms=idx * 60)
    anim.fly_in(key_bar, direction="bottom", trigger="after_previous", duration_ms=320, delay_ms=60)
    anim.apply()

    set_notes(
        slide,
        "Narration order (super intuitive): "
        "1) First point to the Campus Router in the middle: 0 Open Inbound Ports — no packet from the internet can enter directly. "
        "2) Click 1: Instead, an App Connector inside the campus initiates an OUTBOUND Mutual-TLS (mTLS) connection out to the ZTNA Proxy (both sides verify certificates). "
        "3) Click 2: A remote student or faculty member sends a normal HTTPS/TLS request to the ZTNA Proxy, which checks Identity, Device Posture, and Policy (Parts 2–4). "
        "4) Click 3: Only the ZTNA Proxy passes the approved request through that existing tunnel to ONE allowed app — all other campus servers stay completely invisible.",
    )
    return slide


def _build_slide_04_spa_packet(prs):
    """Clearly explains WHY faculty non-web apps need SPA + simple UDP SPA Packet breakdown."""
    slide = add_blank_slide(prs, bg="radial")

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=MARGIN_Y,
        width=17.76,
        height=1.05,
        text="Hiding Non-Web Gateways with SPA",
        font_name=FONT_TITLE,
        font_size=54.0,
        color="WHITE",
        bold=True,
        name="SpaTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=2.06,
        width=17.76,
        height=0.62,
        text="Faculty & researchers use non-web apps (SSH, RDP, CAD) via a separate SDP Gateway — SPA hides it 100% from scanners.",
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        name="SpaSubtitle",
    )

    # Top Section: 4 Plain-English Steps
    steps = [
        ("01", "Drop Everything", "Drops 100% of direct TCP requests — students use the separate web proxy."),
        ("02", "1 UDP Packet First", "Before TCP, the agent sends 1 encrypted UDP packet (no session opened yet)."),
        ("03", "Verify Before Opening", "Checks HMAC signature, timestamp & random nonce (stops replay attacks)."),
        ("04", "Open 5s for 1 IP", "Opens TCP 443 for ~5s to start a stateful mTLS tunnel (re-knocks if dropped)."),
    ]
    step_y = 2.98
    step_h = 2.82
    gap = 0.56
    total_w = 20.0 - 2 * MARGIN_X
    step_w = (total_w - 3 * gap) / 4.0

    step_cards = []
    step_arrows = []
    for idx, (num, stitle, sdesc) in enumerate(steps):
        x = MARGIN_X + idx * (step_w + gap)
        border_c = "VERIFIED_MINT" if idx == 3 else "SAGE_GLASS"
        card = add_frosted_card(
            slide,
            left=x,
            top=step_y,
            width=step_w,
            height=step_h,
            fill_color="DARK_TEAL",
            fill_alpha=68.0,
            border_color=border_c,
            border_alpha=80.0 if idx == 3 else 50.0,
            border_pt=1.5 if idx == 3 else 1.0,
            radius=0.07,
            name=f"SpaStep_{num}",
        )
        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.24)
        tf.margin_right = Inches(0.24)
        tf.margin_top = Inches(0.22)

        p0 = tf.paragraphs[0]
        p0.space_after = Pt(6)
        r0 = p0.add_run()
        r0.text = f"STEP {num}"
        r0.font.name = FONT_BODY
        r0.font.size = Pt(15.0)
        r0.font.bold = True
        r0.font.color.rgb = rgb("VERIFIED_MINT" if idx == 3 else "SOFT_SAGE")

        p1 = tf.add_paragraph()
        p1.space_after = Pt(8)
        r1 = p1.add_run()
        r1.text = stitle
        r1.font.name = FONT_TITLE
        r1.font.size = Pt(21.0)
        r1.font.bold = True
        r1.font.color.rgb = rgb("WHITE")

        p2 = tf.add_paragraph()
        p2.line_spacing = 1.20
        r2 = p2.add_run()
        r2.text = sdesc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(17.5)
        r2.font.color.rgb = rgb("ICE_MIST")

        step_cards.append(card)
        if idx < 3:
            ax1 = x + step_w + 0.08
            ax2 = x + step_w + gap - 0.08
            ay = step_y + step_h / 2.0
            step_arrows.append(add_arrow_connector(slide, ax1, ay, ax2, ay, color="VERIFIED_MINT", width_pt=2.2, name=f"SpaArr_{idx+1}"))

    # Bottom Section: Inside One UDP SPA Packet (Simple, Self-Explanatory Fields!)
    pkt_y = 6.22
    pkt_h = 3.68
    pkt_frame = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=pkt_y,
        width=total_w,
        height=pkt_h,
        fill_color="OBSIDIAN",
        fill_alpha=44.0,
        border_color="VERIFIED_MINT",
        border_alpha=55.0,
        border_pt=1.25,
        radius=0.04,
        name="SpaPacketFrame",
    )
    pkt_badge = add_badge(
        slide,
        left=MARGIN_X + 0.32,
        top=pkt_y + 0.24,
        width=5.45,
        height=0.48,
        text="INSIDE ONE UDP SPA PACKET",
        fill_color="DARK_TEAL",
        fill_alpha=88.0,
        border_color="VERIFIED_MINT",
        border_alpha=75.0,
        text_color="VERIFIED_MINT",
        font_size=15.5,
        name="SpaPacketBadge",
    )
    pkt_note = add_textbox(
        slide,
        left=MARGIN_X + 6.10,
        top=pkt_y + 0.30,
        width=total_w - 6.45,
        height=0.42,
        text="Sent before TCP handshake  ·  Fake packets are dropped immediately",
        font_size=16.5,
        color="ICE_MIST",
        align=PP_ALIGN.RIGHT,
        name="SpaNote",
    )

    inner_x = MARGIN_X + 0.32
    inner_w = total_w - 0.64
    fields = [
        ("UDP Header", "Source IP & Port\n(No TCP yet)", "DARK_TEAL", "SAGE_GLASS", 2.45),
        ("User & Device ID", "Who is knocking\non the gateway", "DARK_TEAL", "VERIFIED_MINT", 3.15),
        ("Time + Nonce", "Random number\n(Anti-Replay)", "DARK_TEAL", "VERIFIED_MINT", 3.15),
        ("Target App", "Exact app requested\n(e.g., SIS :8443)", "DARK_TEAL", "VERIFIED_MINT", 3.15),
        ("HMAC-SHA256", "Digital signature\n(Checked first!)", "VERIFIED_MINT", "VERIFIED_MINT", 3.45),
    ]
    fgap = 0.22
    raw_sum = sum(f[4] for f in fields) + fgap * (len(fields) - 1)
    scale = inner_w / raw_sum

    field_y = pkt_y + 0.98
    field_h = 1.65
    cur_x = inner_x
    field_shapes = []
    enc_start_x = None
    enc_end_x = None

    for idx, (fname, fsub, ffill, fborder, fw_raw) in enumerate(fields):
        fw = fw_raw * scale
        if idx == 1:
            enc_start_x = cur_x
        if idx == 3:
            enc_end_x = cur_x + fw
        is_hmac = idx == 4
        sh = _add_node(
            slide,
            x=cur_x,
            y=field_y,
            w=fw,
            h=field_h,
            title=fname,
            subtitle=fsub,
            fill_color=ffill,
            fill_alpha=26.0 if is_hmac else 78.0,
            border_color=fborder,
            border_alpha=90.0 if idx >= 1 else 55.0,
            border_pt=1.5 if idx >= 1 else 1.0,
            title_color="WHITE",
            sub_color="WHITE" if is_hmac else "ICE_MIST",
            title_fs=18.5,
            sub_fs=15.5,
            name=f"SpaField_{idx+1}",
        )
        field_shapes.append(sh)
        cur_x += fw + fgap

    enc_bar = add_badge(
        slide,
        left=enc_start_x,
        top=field_y + field_h + 0.22,
        width=enc_end_x - enc_start_x,
        height=0.52,
        text="Encrypted with AES-256  (Hides user & target app)",
        fill_color="DARK_TEAL",
        fill_alpha=85.0,
        border_color="VERIFIED_MINT",
        border_alpha=75.0,
        text_color="VERIFIED_MINT",
        font_size=15.5,
        name="SpaAesBar",
    )
    hmac_bar = add_badge(
        slide,
        left=field_shapes[4].left.inches,
        top=field_y + field_h + 0.22,
        width=field_shapes[4].width.inches,
        height=0.52,
        text="Drops fakes before decrypting",
        fill_color="OBSIDIAN",
        fill_alpha=82.0,
        border_color="VERIFIED_MINT",
        border_alpha=75.0,
        text_color="WHITE",
        font_size=14.5,
        name="SpaHmacBar",
    )

    add_footer(slide, PART_LABEL)

    # Smooth automatic entrance choreography on slide load (1 click = next slide, zero mid-slide clicks)
    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=300)
    anim.fade_in(sub, trigger="with_previous", duration_ms=300, delay_ms=60)
    anim.fade_in(step_cards[0], trigger="after_previous", duration_ms=260)
    for i in range(3):
        anim.fade_in(step_arrows[i], trigger="after_previous", duration_ms=180)
        anim.fade_in(step_cards[i + 1], trigger="with_previous", duration_ms=260, delay_ms=50)
    anim.fade_in(pkt_frame, trigger="after_previous", duration_ms=280, delay_ms=60)
    anim.fade_in(pkt_badge, trigger="with_previous", duration_ms=260)
    anim.fade_in(pkt_note, trigger="with_previous", duration_ms=260)
    anim.stagger(field_shapes, effect="fly_bottom", duration_ms=280, step_ms=75, on_click_first=False, base_delay_ms=40)
    anim.fade_in(enc_bar, trigger="after_previous", duration_ms=260)
    anim.fade_in(hmac_bar, trigger="with_previous", duration_ms=260, delay_ms=60)
    anim.apply()

    set_notes(
        slide,
        "Why do we need SPA? Because faculty, engineers, and researchers use non-web desktop tools (SSH, RDP, MATLAB/CAD) "
        "that need a gateway — and SPA keeps that gateway invisible to internet attackers. "
        "Walk through the 4 clicks: 1) Gateway drops every normal TCP request silently. "
        "2) Before TCP, the client sends 1 encrypted UDP packet (stateless — no memory/session allocated yet). "
        "3) Gateway checks the HMAC-SHA256 signature FIRST (dropping fake floods cheaply), then checks timestamp + nonce to prevent replay attacks. "
        "4) Only then does TCP 443 open for that single IP for ~5 seconds to establish mTLS.",
    )
    return slide


def _build_slide_05_two_modes(prs):
    slide, _, _ = L09_two_columns(
        prs,
        title="Two Remote\nAccess Modes",
        subtitle=(
            "Students connect through a browser with nothing installed; "
            "faculty & researchers use an agent for native apps."
        ),
        left_col=(
            "Clientless Proxy (Browser)",
            [
                "Who: Students (personal laptops) & external guests",
                "How: Normal browser over HTTPS (port 443) — zero install",
                "Labs (SSH / RDP): HTML5 gateway displays the lab inside a browser tab",
                "Isolation: Personal laptop never touches the campus network",
            ],
        ),
        right_col=(
            "Client-Based ZTNA (Agent)",
            [
                "Who: Faculty, SIS/ERP staff, IT & HPC researchers",
                "How: Small agent on university laptops using UDP SPA + two-way mTLS",
                "Native Apps: Runs non-web tools like SSH, RDP & MATLAB/CAD",
                "Scoped Tunnel: Connects to one app only — never the whole subnet",
            ],
        ),
        left_badge="BYOD · BROWSER",
        right_badge="FACULTY · AGENT",
        left_accent="VERIFIED_MINT",
        right_accent="SAGE_GLASS",
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        slide,
        "Clear distinction between our two university remote access modes: "
        "Left (Students / Browser-Only): Clientless Identity-Aware Proxy (IAP). Students just open Chrome/Edge on HTTPS 443 — "
        "zero software installed. For SSH or RDP labs, an HTML5 gateway (Apache Guacamole) streams the screen inside the browser tab, "
        "keeping unmanaged personal laptops completely isolated from campus networks. "
        "Right (Faculty / Installed Agent): Client-Based ZTNA for university-owned laptops needing native non-web tools (SSH, RDP, MATLAB/CAD), "
        "using the UDP SPA knock + mTLS tunnel scoped strictly to one app.",
    )
    return slide


def _build_slide_06_comparison_and_bridge(prs):
    slide, anim, tbl_shape = L05_table_compare(
        prs,
        title="Legacy VPN vs.\nCampus ZTNA",
        intro=(
            "Direct comparison between traditional campus VPNs "
            "and Zero Trust Remote Access."
        ),
        headers=["Comparison Point", "Legacy Campus VPN", "Campus ZTNA (Proxy + Agent)"],
        rows=[
            [
                "Access Scope",
                "Whole campus network (internal IP)",
                "Single approved app only (no campus IP)",
            ],
            [
                "Malware Spread",
                "High — can scan any internal host",
                "Blocked — no internal network path",
            ],
            [
                "Campus Firewall",
                "Public VPN ports open to the internet",
                "0 open inbound ports (inside-out connector)",
            ],
            [
                "Internet Traffic",
                "Full-tunnel slows campus; split-tunnel risky",
                "Only app traffic routed — direct & fast",
            ],
            [
                "Student Laptops",
                "Risky — personal PC joins campus LAN",
                "Safe — isolated inside a browser tab",
            ],
        ],
        col_widths=[3.55, 7.05, 7.16],
        part_label=PART_LABEL,
        transition="fade",
    )

    tbl_shape.top = Inches(3.22)
    table = tbl_shape.table
    table.rows[0].height = Inches(0.80)
    for ri in range(1, len(table.rows)):
        table.rows[ri].height = Inches(0.90)

    bridge_bar = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=9.25,
        width=20.0 - 2 * MARGIN_X,
        height=0.92,
        fill_color="DARK_TEAL",
        fill_alpha=72.0,
        border_color="VERIFIED_MINT",
        border_alpha=75.0,
        border_pt=1.25,
        radius=0.14,
        name="BridgeToPart07",
    )
    tf = bridge_bar.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.35)
    tf.margin_right = Inches(0.35)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run()
    r1.text = "NEXT  ·  ASSUME BREACH:  "
    r1.font.name = FONT_TITLE
    r1.font.size = Pt(18.5)
    r1.font.bold = True
    r1.font.color.rgb = rgb("VERIFIED_MINT")
    r2 = p.add_run()
    r2.text = "Every remote ZTNA session sends live logs to Monitoring & Automated Response."
    r2.font.name = FONT_BODY
    r2.font.size = Pt(19.0)
    r2.font.color.rgb = rgb("WHITE")

    anim.fly_in(bridge_bar, direction="bottom", trigger="after_previous", duration_ms=340, delay_ms=100)
    anim.apply()

    set_notes(
        slide,
        "Final side-by-side scorecard wrapping up Part 6, followed by a natural bridge to Part 7 (Monitoring and Response): "
        "Even with zero open campus ports, SPA-cloaked gateways, and per-app isolation, Zero Trust assumes breach — "
        "every remote ZTNA session streams live telemetry so suspicious behavior can be detected and stopped immediately.",
    )
    return slide


def build_part_06(prs) -> list:
    """Append all 6 slides for Part 06: Remote Resource Access."""
    return [
        _build_slide_01_divider(prs),
        _build_slide_02_legacy_vpn(prs),
        _build_slide_03_ztna_architecture(prs),
        _build_slide_04_spa_packet(prs),
        _build_slide_05_two_modes(prs),
        _build_slide_06_comparison_and_bridge(prs),
    ]
