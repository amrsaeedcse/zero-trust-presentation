"""Part 04 — Authorization Policies.

Owner: Omar Ahmed Abdelaziz Mashhour
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- NIST SP 800-207 core logical components (Policy Engine, Policy Administrator, Policy Enforcement Point).
- RBAC vs. ABAC vs. context/risk-aware PBAC.
- Dynamic least-privilege, Just-In-Time (JIT) access grants, per-session continuous evaluation & revocation.

Out of Scope (DO NOT re-explain):
- How credentials/MFA work (P2), how device telemetry is gathered (P3),
  L2/L3 micro-segmentation mechanics (P5), remote ZTNA tunneling (P6), data labeling schemas (P9).

Narrative Bridge:
- Builds on Parts 02+03 (identity + device signals) -> Bridges to Network Segmentation (Part 05).
"""

from pptx.oxml.ns import qn

from core.layouts import (
    L02_grid_2x2,
    L03_section_divider,
    L05_table_compare,
    L07_flowchart_6node,
    L10_three_cards,
    L11_stepped_stack,
)
from core.theme import MARGIN_X, add_textbox, apply_bg, set_notes


PART_LABEL = "Part 04  ·  Authorization policies"


def _speaker_notes(slide, script: str) -> None:
    set_notes(slide, script)
    for paragraph in slide.notes_slide.notes_text_frame.paragraphs:
        properties = paragraph._p.get_or_add_pPr()
        properties.set("rtl", "1")
        properties.set("algn", "r")


def build_part_04(prs) -> list:
    opening, _ = L03_section_divider(
        prs,
        part_badge="PART 04  ·  ACCESS CONTROL",
        title="Authorization\nPolicies",
        summary=(
            "Dr-Ahmed has signed in from a healthy laptop. "
            "Can he change this student's grade today?"
        ),
        presenter="Omar Ahmed Abdelaziz Mashhour",
        motif="prism",
        transition="push_left",
    )
    _speaker_notes(opening, (
        "خلّونا نكمّل من عند آخر نقطة. إحنا اتأكدنا إن اللي سجّل الدخول هو الدكتور أحمد، وإن "
        "اللابتوب بتاعه سليم. دلوقتي هو فاتح نظام الدرجات وعايز يعدّل درجة طالب. هل نسمح له "
        "يدوس حفظ؟ عشان نجاوب، محتاجين نعرف الطالب ده في أنهي مادة، ومين بيدرّسها، والرصد "
        "مفتوح النهارده ولا لأ. القرار ده هو موضوع جزئي: صلاحيات الوصول."
    ))

    signals, _, _ = L11_stepped_stack(
        prs,
        title="What shapes\nthe decision?",
        subtitle="The policy engine (PE) evaluates each request using four inputs.",
        rows=[
            ("Identity", "The identity provider confirms Dr-Ahmed's role, courses and MFA level."),
            ("Device health", "His laptop reports its patch, encryption and malware status."),
            ("Resource and action", "Dr-Ahmed wants to edit a sensitive record in the grading system."),
            ("Current context", "The grading calendar and live risk alerts can change the outcome."),
        ],
        part_label=PART_LABEL,
    )
    _speaker_notes(signals, (
        "قدّامنا أربع حاجات محرك السياسات بيبصّ عليها. أول حاجة الهوية: إحنا عارفين إن الدكتور "
        "أحمد عضو هيئة تدريس، وعارفين المواد المسندة له. تاني حاجة الجهاز: هل اللابتوب لسه "
        "مستوفي شروط الجامعة، ولا ظهر عليه إنذار إصابة؟\n\n"
        "تالت حاجة هي الطلب نفسه. فيه فرق بين إنه يفتح صفحة الدرجات ويقراها، وإنه يغيّر رقم "
        "ويتحفظ باسم طالب. رابع حاجة هي الظروف وقت الطلب: هل ده ميعاد الرصد؟ وهل نظام المراقبة "
        "شايف نشاط غريب؟ المحرك بيجمع الإجابات دي عشان يقرر يسمح بالعملية دي تحديدًا."
    ))

    models, _, _ = L10_three_cards(
        prs,
        title="When can Dr-Ahmed edit a grade?",
        cards_data=[
            ("RBAC", "Faculty role", "The professor role makes Dr-Ahmed eligible to enter course grades."),
            ("ABAC", "Assigned course", "The course ID must match Dr-Ahmed's teaching assignment, and his laptop must be compliant."),
            ("Risk-aware policy", "Time-limited access", "Open write access during grading only. If risk rises, withdraw the grant."),
        ],
        part_label=PART_LABEL,
    )
    apply_bg(models, "diagonal")
    _speaker_notes(models, (
        "طب نطبّق ده على مادة الشبكات. أول كارت هو RBAC، يعني بنبدأ من دور الشخص. بما إن "
        "الدكتور أحمد عضو هيئة تدريس، فرصد الدرجات من ضمن الوظائف اللي ممكن يؤديها.\n\n"
        "الكارت التاني هو ABAC. هنا بنراجع تفاصيل تخص الطلب: هل كود المادة اللي فاتحها موجود "
        "ضمن المواد اللي بيدرّسها؟ لو فتح مادة زميله، طلب التعديل يترفض حتى لو جهازه سليم.\n\n"
        "نيجي للوقت والمخاطر. لو دي مادته وفترة الرصد مفتوحة والشروط مستوفاة، ياخد صلاحية "
        "الكتابة وقت احتياجه ليها. أول ما الفترة تخلص، صلاحية التعديل تنتهي. ولو مسموح له يقرأ "
        "بعد كده، ده بيكون بإذن قراءة منفصل."
    ))

    matrix, _, matrix_shape = L05_table_compare(
        prs,
        title="Campus access\npolicy matrix",
        intro="Each rule grants one action on one resource. Requests without a matching grant are denied.",
        headers=["Requester", "Resource", "Allowed action", "Conditions or limits"],
        col_widths=[3.1, 3.0, 4.0, 7.66],
        rows=[
            ["Student", "LMS", "Read / upload", "Enrolled · TOTP MFA · low risk"],
            ["Student", "SIS", "None", "Deny before the SIS sign-in page"],
            ["Teaching assistant", "SIS lab marks", "Enter lab scores", "Assigned lab · compliant PC · grading window"],
            ["Professor", "SIS course", "Read / edit grades", "Assigned course · FIDO2 · compliant PC · JIT"],
            ["Lab server", "Lab grades API", "POST scores only", "Service certificate (mTLS) · exam window"],
            ["Registrar", "SIS records", "Approve and lock", "Campus-managed PC · FIDO2 MFA"],
        ],
        part_label=PART_LABEL,
    )
    # DrawingML requires the four cell borders before the fill, in L/R/T/B order.
    for cell in matrix_shape.table.iter_cells():
        props = cell._tc.tcPr
        for index, edge in enumerate(("lnL", "lnR", "lnT", "lnB")):
            props.insert(index, props.find(qn(f"a:{edge}")))
    _speaker_notes(matrix, (
        "الجدول ده بنقراه من الشمال لليمين: مين بيطلب، عايز يوصل لإيه، مسموح له يعمل إيه، وبأي "
        "شروط. بصّوا على أول صفّين. نفس الطالب يقدر يرفع التكليف على منصة التعلّم لو مسجّل في "
        "المادة ومستوفي شروط الدخول. إنما لو طلب يفتح نظام الدرجات، الطلب بيقف قبل ما تظهر له "
        "صفحة الدخول.\n\n"
        "ولو تابعنا درجة واحدة، هنلاقي المعيد يقدر يدخل درجة المعمل المسند له، والدكتور أحمد "
        "يقدر يعدّل درجات مادته وقت الرصد، وموظف التسجيل عنده صلاحية الاعتماد والقفل. وحتى "
        "سيرفر المعمل له صف في الجدول: مسموح له يبعت درجات المعمل من الواجهة المخصصة، من غير "
        "ما يدخل على قاعدة البيانات مباشرة."
    ))

    flow, _, _ = L07_flowchart_6node(
        prs,
        title="How a grade edit is authorized",
        subtitle="Dr-Ahmed submits a change to a grade in his assigned course.",
        steps=[
            ("Intercept", "PEP holds the request."),
            ("Check rules", "PE verifies role, course and device."),
            ("Assess risk", "PE checks time and current risk."),
            ("Issue grant", "PA issues a short-lived credential."),
            ("Enforce", "PEP allows the scoped action."),
            ("Re-evaluate", "New risk can end the session."),
        ],
        part_label=PART_LABEL,
        click_reveal=True,
    )
    apply_bg(flow, "diagonal")
    add_textbox(
        flow, MARGIN_X, 9.67, 17.76, 0.5,
        "PEP  ·  Policy enforcement point      PE  ·  Policy engine      PA  ·  Policy administrator",
        font_size=20, color="ICE_MIST", name="PolicyComponents",
    )
    _speaker_notes(flow, (
        "دلوقتي الدكتور أحمد ضغط حفظ. نمشي مع الأسهم ونشوف الطلب ده بيمرّ على مين.\n\n"
        "أول خطوة، الطلب يوصل لنقطة تنفيذ السياسة، أو PEP. هي اللي تمسك الطلب قبل ما التعديل "
        "يوصل لنظام الدرجات. تاني خطوة، محرك السياسات، أو PE، يراجع الشروط الأساسية: هل دي "
        "مادة الدكتور فعلًا؟ وهل الجهاز سليم؟ تالت خطوة، نفس المحرك يراجع الوقت وآخر إشارات "
        "المخاطر وياخد القرار.\n\n"
        "لو القرار سماح، نيجي للخطوة الرابعة: مسؤول إدارة السياسة، أو PA، بيصدر إذن قصير المدة "
        "ومحدد للعملية دي. في الخطوة الخامسة نرجع لنقطة التنفيذ، وهي اللي تسمح للتعديل يعدّي. "
        "الخطوة السادسة بتحصل حتى بعد السماح: لو ظهر إنذار جديد، المحرك يعيد التقييم، ومسؤول "
        "السياسة يبلّغ نقطة التنفيذ إنها تنهي الجلسة."
    ))

    outcomes, _, _ = L02_grid_2x2(
        prs,
        title="What the campus policy blocks",
        subtitle="These decisions define which connections campus segments should permit.",
        cards=[
            ("Direct SIS access", "A student's request stops at the PEP before the grading application opens."),
            ("Lab server pivot", "The server can submit lab scores. It has no grant for the SIS interface or database."),
            ("Infected faculty laptop", "A malware alert triggers a fresh decision. The PEP closes Dr-Ahmed's active session."),
            ("Stolen session token", "A device-bound grant requires proof of the original device key when presented."),
        ],
        part_label=PART_LABEL,
        on_click_cards=True,
    )
    apply_bg(outcomes, "diagonal")
    _speaker_notes(outcomes, (
        "نختم بأربع مواقف من شبكة الجامعة. أولًا، لو طالب حاول يفتح نظام الدرجات مباشرة، نقطة "
        "التنفيذ توقفه. تانيًا، لو سيرفر المعمل اتصاب، صلاحيته المحدودة بإرسال الدرجات ما "
        "بتديهوش طريق مباشر لقاعدة البيانات.\n\n"
        "ثالثًا، نرجع للدكتور أحمد. لو جهازه كان سليم ساعة الدخول وبعدها ظهر عليه برنامج ضار، "
        "نقدر نسحب الإذن ونقفل الجلسة وهي شغّالة. رابعًا، لو حد نسخ رمز الجلسة وحاول يستخدمه "
        "من جهاز تاني، ربط الإذن بمفتاح الجهاز الأصلي بيخلّي الرمز لوحده غير كافي.\n\n"
        "وده يوصلنا للجزء الجاي: إحنا حددنا إن سيرفر المعمل محتاج يوصل لواجهة إرسال الدرجات "
        "بس. إزاي نخلي تصميم الشبكة نفسها يسمح بالمسار ده ويقفل المسارات التانية؟ ده دور تقسيم "
        "الشبكة."
    ))

    return [opening, signals, models, matrix, flow, outcomes]
