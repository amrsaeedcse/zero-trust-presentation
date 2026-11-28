"""Part 05 — Network Segmentation.

Owner: Ali Mohamed Ali Abdallah Omara
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Macro- vs. micro-segmentation for campus zones (dorms/student Wi-Fi, labs, research clusters,
  admin/ERP, IoT/BMS, guest Wi-Fi).
- Software-Defined Perimeter (SDP), SDN/VXLAN/host-based micro-segmentation.
- East-west traffic inspection & blast-radius containment inside the campus network.

Out of Scope (DO NOT re-explain):
- Re-explaining castle-and-moat (P1), PE/PA decision math (P4),
  off-campus remote access/ZTNA vs VPN (P6), SOC monitoring pipelines (P7).

Narrative Bridge:
- Builds on Part 04 (policy enforcement) -> Bridges to Remote Resource Access (Part 06).
"""

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.layouts import (
    L03_section_divider,
    L04_wide_visual,
    L05_table_compare,
    L07_flowchart_6node,
    L10_three_cards,
    L12_split_checklist,
)
from core.theme import (
    FONT_BODY,
    FONT_TITLE,
    add_arrow_connector,
    add_badge,
    add_frosted_card,
    add_textbox,
    rgb,
    set_notes,
)

PART_LABEL = "Part 05  ·  Network Segmentation"


def _add_zone_card(slide, left: float, top: float, width: float, height: float,
                   title: str, subtitle: str, accent: str = "SAGE_GLASS", name: str = "ZoneCard"):
    """Create a compact frosted-glass campus zone card with crisp hierarchy."""
    card = add_frosted_card(
        slide,
        left=left,
        top=top,
        width=width,
        height=height,
        fill_alpha=24.0,
        border_color=accent,
        border_alpha=60.0,
        border_pt=1.25,
        radius=0.08,
        name=name,
    )
    tf = card.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.26)
    tf.margin_right = Inches(0.26)
    tf.margin_top = Inches(0.14)
    tf.margin_bottom = Inches(0.14)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.LEFT
    p1.space_after = Pt(4)
    r1 = p1.add_run()
    r1.text = title
    r1.font.name = FONT_TITLE
    r1.font.size = Pt(22.0)
    r1.font.bold = True
    r1.font.color.rgb = rgb("WHITE")

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.LEFT
    p2.line_spacing = 1.15
    r2 = p2.add_run()
    r2.text = subtitle
    r2.font.name = FONT_BODY
    r2.font.size = Pt(18.0)
    r2.font.color.rgb = rgb("ICE_MIST")
    return card


def build_part_05(prs) -> list:
    slides = []

    # ------------------------------------------------------------------
    # Slide 1 — Section Divider (L03)
    # ------------------------------------------------------------------
    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 05  ·  CAMPUS NETWORK",
        title="Network\nSegmentation",
        summary=(
            "Enforcing authorization boundaries inside the campus: isolating student Wi-Fi, "
            "academic labs, research clusters, ERP, and smart IoT to inspect east-west traffic "
            "and contain lateral blast radius."
        ),
        presenter="Ali Mohamed Ali Abdallah Omara",
        motif="signal_waves",
        transition="push_left",
    )
    set_notes(
        s1,
        "Bridge from Part 4: Once the Policy Engine and Policy Administrator issue dynamic access "
        "decisions, Network Segmentation enforces those boundaries across every campus wire and Wi-Fi link.",
    )
    slides.append(s1)

    # ------------------------------------------------------------------
    # Slide 2 — Macro- vs. Micro-Segmentation Comparison (L05)
    # ------------------------------------------------------------------
    s2, _, _ = L05_table_compare(
        prs,
        title="Macro- vs. Micro-\nSegmentation",
        intro=(
            "Macro-segmentation separates broad campus zones, while micro-segmentation "
            "enforces identity-bound rules down to individual workloads and hosts."
        ),
        headers=["Dimension", "Macro-Segmentation", "Micro-Segmentation"],
        rows=[
            [
                "Granularity",
                "Subnet, VLAN, or VRF level across buildings and departments",
                "Individual host, VM, container, or application workload",
            ],
            [
                "Campus Scope",
                "Separates Dorms, Guest Wi-Fi, Labs, Research, ERP, and IoT",
                "Isolates specific lab servers, GPU nodes, and ERP tiers",
            ],
            [
                "East-West Control",
                "Coarse inter-VLAN ACLs; hosts in the same VLAN can still talk",
                "Default-deny L4–L7 inspection even within the same subnet",
            ],
            [
                "Blast Radius",
                "Limited to one campus zone or subnet upon compromise",
                "Contained to a single endpoint or workload instance",
            ],
        ],
        col_widths=[3.4, 7.18, 7.18],
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        s2,
        "Explain why universities need both: VLANs/VRFs provide macro-level separation between "
        "high-churn student/guest networks and core systems, while micro-segmentation stops lateral "
        "movement inside a single subnet.",
    )
    slides.append(s2)

    # ------------------------------------------------------------------
    # Slide 3 — Campus Micro-Segmented Topology Diagram (L04 Custom Stage)
    # ------------------------------------------------------------------
    s3, anim3, (sx, sy, sw, sh) = L04_wide_visual(
        prs,
        title="Campus Zone\nArchitecture",
        intro=(
            "Six explicit campus trust zones connected through an east-west inspection "
            "fabric—blocking direct lateral paths between endpoints and critical assets."
        ),
        part_label=PART_LABEL,
        stage_card=True,
        transition="fade",
    )

    # Column headers inside stage
    hdr_left = add_textbox(
        s3, sx + 0.43, sy + 0.20, 4.65, 0.36,
        "HIGH-CHURN & IOT ZONES",
        font_name=FONT_BODY, font_size=16.0, color="SOFT_SAGE", bold=True, align=PP_ALIGN.CENTER,
        name="ColHdrLeft",
    )
    hdr_mid = add_textbox(
        s3, sx + 6.38, sy + 0.20, 5.00, 0.36,
        "EAST-WEST ENFORCEMENT",
        font_name=FONT_BODY, font_size=16.0, color="VERIFIED_MINT", bold=True, align=PP_ALIGN.CENTER,
        name="ColHdrMid",
    )
    hdr_right = add_textbox(
        s3, sx + 12.68, sy + 0.20, 4.65, 0.36,
        "PROTECTED CAMPUS ASSETS",
        font_name=FONT_BODY, font_size=16.0, color="SOFT_SAGE", bold=True, align=PP_ALIGN.CENTER,
        name="ColHdrRight",
    )

    # Left column: Untrusted / High-churn zones
    left_zones = [
        ("Student Wi-Fi & Dorms", "BYOD laptops & phones; isolated from core", "ALERT_CORAL"),
        ("Guest Wi-Fi", "Internet-only egress; zero internal reach", "ALERT_CORAL"),
        ("Smart IoT & BMS", "HVAC, cameras & badge readers in strict VRF", "SAGE_GLASS"),
    ]
    left_cards = []
    for idx, (zt, zs, acc) in enumerate(left_zones):
        y = sy + 0.72 + idx * 1.92
        left_cards.append(
            _add_zone_card(s3, sx + 0.43, y, 4.65, 1.56, zt, zs, accent=acc, name=f"LeftZone_{idx+1}")
        )

    # Center column: SDN / East-West Policy Fabric
    fabric_card = add_frosted_card(
        s3,
        left=sx + 6.38,
        top=sy + 0.72,
        width=5.00,
        height=5.40,
        fill_color="DARK_TEAL",
        fill_alpha=68.0,
        border_color="VERIFIED_MINT",
        border_alpha=70.0,
        border_pt=1.5,
        radius=0.06,
        name="EastWestFabric",
    )
    tf_fab = fabric_card.text_frame
    tf_fab.vertical_anchor = MSO_ANCHOR.TOP
    tf_fab.margin_left = Inches(0.35)
    tf_fab.margin_right = Inches(0.35)
    tf_fab.margin_top = Inches(0.32)

    pf0 = tf_fab.paragraphs[0]
    pf0.alignment = PP_ALIGN.CENTER
    pf0.space_after = Pt(10)
    rf0 = pf0.add_run()
    rf0.text = "SDN & East-West\nInspection Fabric"
    rf0.font.name = FONT_TITLE
    rf0.font.size = Pt(25.0)
    rf0.font.bold = True
    rf0.font.color.rgb = rgb("WHITE")

    fabric_points = [
        "Default-deny between all zones",
        "Deep L4–L7 east-west inspection",
        "Identity & posture-bound flows",
        "No direct dorm-to-ERP routing",
    ]
    for pt_txt in fabric_points:
        pf = tf_fab.add_paragraph()
        pf.alignment = PP_ALIGN.LEFT
        pf.space_after = Pt(8)
        pf.line_spacing = 1.15
        rf = pf.add_run()
        rf.text = f"▸  {pt_txt}"
        rf.font.name = FONT_BODY
        rf.font.size = Pt(17.5)
        rf.font.color.rgb = rgb("ICE_MIST")

    fabric_badge = add_badge(
        s3,
        left=sx + 6.88,
        top=sy + 5.28,
        width=4.00,
        height=0.56,
        text="LATERAL MOVEMENT BLOCKED",
        fill_color="ALERT_CORAL",
        fill_alpha=24.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        text_color="WHITE",
        font_size=15.0,
        name="LateralBlockBadge",
    )

    # Right column: Protected Campus Assets
    right_zones = [
        ("Academic Labs", "Coursework VMs & departmental workstations", "VERIFIED_MINT"),
        ("Research Clusters", "High-value GPU nodes & grant datasets", "VERIFIED_MINT"),
        ("Admin & ERP Systems", "Grading, finance, HR & student records", "VERIFIED_MINT"),
    ]
    right_cards = []
    for idx, (zt, zs, acc) in enumerate(right_zones):
        y = sy + 0.72 + idx * 1.92
        right_cards.append(
            _add_zone_card(s3, sx + 12.68, y, 4.65, 1.56, zt, zs, accent=acc, name=f"RightZone_{idx+1}")
        )

    # Arrows: Left -> Fabric -> Right
    left_arrows = []
    right_arrows = []
    for idx in range(3):
        cy = sy + 0.72 + idx * 1.92 + 0.78
        left_arrows.append(
            add_arrow_connector(
                s3, sx + 5.16, cy, sx + 6.30, cy,
                color="ICE_MIST", width_pt=1.75, dashed=True, name=f"ArrL_{idx+1}",
            )
        )
        right_arrows.append(
            add_arrow_connector(
                s3, sx + 11.46, cy, sx + 12.60, cy,
                color="VERIFIED_MINT", width_pt=2.0, dashed=False, name=f"ArrR_{idx+1}",
            )
        )

    anim3.fade_in(hdr_left, trigger="with_previous", duration_ms=280)
    anim3.fade_in(hdr_mid, trigger="with_previous", duration_ms=280)
    anim3.fade_in(hdr_right, trigger="with_previous", duration_ms=280)
    anim3.click_sequence(
        [
            left_cards,
            [*left_arrows, fabric_card, fabric_badge],
            [*right_arrows, *right_cards],
        ],
        effect="fade",
        duration_ms=360,
        intra_delay_ms=70,
    )
    anim3.apply()
    set_notes(
        s3,
        "Walk through the diagram on click: (1) High-churn dorms, guest Wi-Fi, and IoT sensors sit in "
        "isolated zones; (2) Every east-west packet must traverse the SDN inspection fabric; "
        "(3) Only explicitly authorized flows reach Academic Labs, Research Clusters, or Admin/ERP.",
    )
    slides.append(s3)

    # ------------------------------------------------------------------
    # Slide 4 — Core Segmentation Technologies (L10)
    # ------------------------------------------------------------------
    s4, _, _ = L10_three_cards(
        prs,
        title="Segmentation Enablers",
        cards_data=[
            (
                "Overlay & Fabric",
                "SDN & VXLAN\nOverlays",
                "Decouples logical security zones from physical campus switches. "
                "Centralized SDN controllers push consistent segmentation policies across "
                "multi-building networks without manual VLAN sprawl.",
            ),
            (
                "Workload Level",
                "Host-Based Micro-\nSegmentation",
                "Enforces kernel and eBPF firewall rules directly on lab servers, "
                "ERP databases, and research nodes—controlling east-west traffic even "
                "when workloads share the same subnet.",
            ),
            (
                "Cloaked Access",
                "Software-Defined\nPerimeter (SDP)",
                "Makes internal campus applications invisible to unauthorized scanners. "
                "Ports remain dark by default and open dynamically only after policy "
                "authorization succeeds.",
            ),
        ],
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        s4,
        "Highlight how SDN/VXLAN handles campus-wide fabric agility, host-based agents enforce "
        "per-workload rules, and SDP hides critical servers from internal port scanning.",
    )
    slides.append(s4)

    # ------------------------------------------------------------------
    # Slide 5 — Blast-Radius Containment Flowchart (L07)
    # ------------------------------------------------------------------
    s5, _, _ = L07_flowchart_6node(
        prs,
        title="Blast-Radius Containment in Action",
        subtitle="Stopping lateral movement when a campus endpoint is compromised",
        steps=[
            ("Endpoint Breach", "Student BYOD laptop or smart HVAC sensor is infected"),
            ("Lateral Scan", "Malware probes adjacent subnets for SMB, SSH, or ERP"),
            ("East-West Drop", "Host firewall & SDN fabric block unauthorized probes"),
            ("Zone Quarantine", "Compromised device is isolated into a remediation VLAN"),
            ("Assets Protected", "Research GPUs and Admin/ERP remain completely unreachable"),
            ("Blast Radius = 1", "Incident contained to a single host, not a campus breach"),
        ],
        part_label=PART_LABEL,
        click_reveal=True,
        transition="fade",
    )
    set_notes(
        s5,
        "Walk through this 6-step containment flow on click to demonstrate the core payoff of "
        "micro-segmentation: assuming breach and shrinking the blast radius to a single host.",
    )
    slides.append(s5)

    # ------------------------------------------------------------------
    # Slide 6 — Summary & Bridge to Remote Resource Access (L12)
    # ------------------------------------------------------------------
    s6, _, _ = L12_split_checklist(
        prs,
        title="Campus Contained.\nWhat About Off-Campus?",
        subtitle=(
            "Macro-zones, micro-segmentation, and SDP eliminate implicit LAN trust inside "
            "university walls. Next, those same Zero Trust boundaries must follow students "
            "and faculty when they connect remotely."
        ),
        checklist_title="Campus Segmentation Takeaways",
        items=[
            "Separate campus into 6 explicit zones (Dorms, Guest, Labs, Research, ERP, IoT)",
            "Enforce default-deny east-west inspection between and inside subnets",
            "Use SDN/VXLAN and host-based controls instead of static IP ACLs",
            "Cloak high-value research & ERP services behind Software-Defined Perimeter",
            "Extend campus segmentation boundaries to off-campus remote users",
        ],
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        s6,
        "Wrap up Part 5 and bridge cleanly to Part 6 (Remote Resource Access): internal campus "
        "networks are now segmented, so the next challenge is replacing legacy remote access "
        "trust with application-level Zero Trust for off-campus users.",
    )
    slides.append(s6)

    return slides
