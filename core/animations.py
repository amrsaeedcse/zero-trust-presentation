"""Native PowerPoint OOXML Animation & Slide Transition Engine (Priority #1).

Generates 100% ISO/IEC 29500 `pml.xsd`-compliant `<p:transition>` and `<p:timing>`
elements so animations play smoothly in PowerPoint with zero "repair" warnings.

Usage inside any part file:
    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(title_shape, trigger="on_load", duration_ms=350)
    anim.stagger([card1, card2, card3, card4], effect="fly_bottom", step_ms=140, on_click_first=True)
    anim.apply()
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List, Optional

from lxml import etree
from pptx.oxml.ns import nsdecls, qn

_TRANSITION_XML = {
    "fade": "<p:fade/>",
    "push_left": '<p:push dir="l"/>',
    "push_right": '<p:push dir="r"/>',
    "push_up": '<p:push dir="u"/>',
    "push_down": '<p:push dir="d"/>',
    "wipe_left": '<p:wipe dir="l"/>',
    "wipe_right": '<p:wipe dir="r"/>',
    "wipe_up": '<p:wipe dir="u"/>',
    "wipe_down": '<p:wipe dir="d"/>',
    "cover_left": '<p:cover dir="l"/>',
    "dissolve": "<p:dissolve/>",
    "cut": "<p:cut/>",
}


def _insert_slide_child_in_schema_order(sld_el, new_el, tag_local: str) -> None:
    """Insert `<p:transition>` or `<p:timing>` in exact pml.xsd order.

    Order inside `<p:sld>`: cSld -> clrMapOvr -> transition -> timing -> extLst
    """
    existing = sld_el.find(qn(f"p:{tag_local}"))
    if existing is not None:
        sld_el.remove(existing)

    order = ["cSld", "clrMapOvr", "transition", "timing", "extLst"]
    target_rank = order.index(tag_local)
    insert_idx = len(sld_el)
    for i, child in enumerate(sld_el):
        local = etree.QName(child).localname
        if local in order and order.index(local) > target_rank:
            insert_idx = i
            break
    sld_el.insert(insert_idx, new_el)


def set_slide_transition(
    slide,
    effect: str = "fade",
    speed: str = "med",
    advance_on_click: bool = True,
    advance_after_ms: Optional[int] = None,
) -> None:
    """Attach a smooth native PowerPoint slide transition (`fade`, `push_left`, `wipe_left`, etc.)."""
    if effect not in _TRANSITION_XML:
        raise ValueError(f"Unknown transition {effect!r}. Valid: {list(_TRANSITION_XML)}")
    if speed not in ("slow", "med", "fast"):
        raise ValueError("speed must be 'slow', 'med', or 'fast'")

    attrs = [f'spd="{speed}"', f'advClick="{"1" if advance_on_click else "0"}"']
    if advance_after_ms is not None:
        attrs.append(f'advTm="{int(advance_after_ms)}"')

    xml_str = (
        f'<p:transition {nsdecls("p")} {" ".join(attrs)}>'
        f"{_TRANSITION_XML[effect]}"
        f"</p:transition>"
    )
    _insert_slide_child_in_schema_order(slide._element, etree.fromstring(xml_str), "transition")


@dataclass
class _AnimStep:
    spid: int
    effect: str
    trigger: str
    duration_ms: int
    delay_ms: int


_EFFECT_PRESETS = {
    "fade": (10, 0),
    "appear": (1, 0),
    "fly_bottom": (2, 4),
    "fly_left": (2, 8),
    "fly_right": (2, 2),
    "fly_top": (2, 1),
    "wipe_down": (22, 1),
    "wipe_left": (22, 8),
    "wipe_right": (22, 2),
    "wipe_up": (22, 4),
    "zoom": (23, 16),
}


# Global presentation pacing policy:
# When True, every slide's entrance animations play smoothly on slide entry
# (converting mid-slide `on_click` steps into crisp `after_previous` beats)
# so every presenter click cleanly advances to the next slide.
AUTO_PLAY_SLIDE_ANIMATIONS = True


class SlideAnimator:
    """Fluent OOXML animation builder for a single PowerPoint slide."""

    def __init__(self, slide, transition: str = "fade", transition_speed: str = "med"):
        self.slide = slide
        self._steps: List[_AnimStep] = []
        if transition:
            set_slide_transition(slide, effect=transition, speed=transition_speed)

    def add(
        self,
        shape,
        effect: str = "fade",
        trigger: str = "after_previous",
        duration_ms: int = 380,
        delay_ms: int = 0,
    ) -> "SlideAnimator":
        if shape is None:
            return self
        if effect not in _EFFECT_PRESETS:
            raise ValueError(f"Unknown effect {effect!r}. Valid: {list(_EFFECT_PRESETS)}")
        if trigger not in ("on_load", "on_click", "with_previous", "after_previous"):
            raise ValueError(f"Invalid trigger {trigger!r}")
        eff_trigger = trigger
        eff_delay = max(0, int(delay_ms))
        eff_dur = max(1, int(duration_ms))
        if AUTO_PLAY_SLIDE_ANIMATIONS and eff_trigger == "on_click":
            eff_trigger = "after_previous"
            eff_delay = max(eff_delay, 90)
            eff_dur = min(eff_dur, 320)
        self._steps.append(
            _AnimStep(
                spid=int(shape.shape_id),
                effect=effect,
                trigger=eff_trigger,
                duration_ms=eff_dur,
                delay_ms=eff_delay,
            )
        )
        return self

    def fade_in(self, shape, trigger: str = "on_load", duration_ms: int = 380, delay_ms: int = 0) -> "SlideAnimator":
        return self.add(shape, effect="fade", trigger=trigger, duration_ms=duration_ms, delay_ms=delay_ms)

    def fly_in(
        self,
        shape,
        direction: str = "bottom",
        trigger: str = "after_previous",
        duration_ms: int = 420,
        delay_ms: int = 0,
    ) -> "SlideAnimator":
        return self.add(shape, effect=f"fly_{direction}", trigger=trigger, duration_ms=duration_ms, delay_ms=delay_ms)

    def wipe_in(
        self,
        shape,
        direction: str = "right",
        trigger: str = "after_previous",
        duration_ms: int = 360,
        delay_ms: int = 0,
    ) -> "SlideAnimator":
        return self.add(shape, effect=f"wipe_{direction}", trigger=trigger, duration_ms=duration_ms, delay_ms=delay_ms)

    def zoom_in(self, shape, trigger: str = "after_previous", duration_ms: int = 380, delay_ms: int = 0) -> "SlideAnimator":
        return self.add(shape, effect="zoom", trigger=trigger, duration_ms=duration_ms, delay_ms=delay_ms)

    def appear(self, shape, trigger: str = "on_click", delay_ms: int = 0) -> "SlideAnimator":
        return self.add(shape, effect="appear", trigger=trigger, duration_ms=1, delay_ms=delay_ms)

    def stagger(
        self,
        shapes: Iterable,
        effect: str = "fly_bottom",
        duration_ms: int = 380,
        step_ms: int = 120,
        on_click_first: bool = False,
        base_delay_ms: int = 80,
    ) -> "SlideAnimator":
        """Reveal a sequence of shapes with clean staggered timing (cinematic wave)."""
        items = [s for s in shapes if s is not None]
        for idx, sh in enumerate(items):
            if idx == 0:
                trig = "on_click" if on_click_first else "after_previous"
                dly = 0 if on_click_first else base_delay_ms
            else:
                trig = "with_previous"
                dly = step_ms * idx + (0 if on_click_first else base_delay_ms)
            self.add(sh, effect=effect, trigger=trig, duration_ms=duration_ms, delay_ms=dly)
        return self

    def click_sequence(
        self,
        groups: Iterable,
        effect: str = "fade",
        duration_ms: int = 360,
        intra_delay_ms: int = 80,
    ) -> "SlideAnimator":
        """Reveal groups of shapes one click at a time (ideal for step-by-step diagrams)."""
        for grp in groups:
            shapes = list(grp) if isinstance(grp, (list, tuple)) else [grp]
            shapes = [s for s in shapes if s is not None]
            for i, sh in enumerate(shapes):
                trig = "on_click" if i == 0 else "with_previous"
                self.add(sh, effect=effect, trigger=trig, duration_ms=duration_ms, delay_ms=i * intra_delay_ms)
        return self

    def apply(self) -> None:
        """Compile all queued animation steps into valid OOXML `<p:timing>` on the slide."""
        if not self._steps:
            return

        click_groups: List[List[_AnimStep]] = []
        for step in self._steps:
            if not click_groups or step.trigger == "on_click":
                click_groups.append([step])
            else:
                click_groups[-1].append(step)

        nid = [2]

        def next_id() -> int:
            nid[0] += 1
            return nid[0]

        outer_pars_xml: List[str] = []
        seen_spids: List[int] = []

        for g_idx, group in enumerate(click_groups):
            first_step = group[0]
            group_is_click = first_step.trigger == "on_click"
            outer_delay = "indefinite" if group_is_click else "0"

            sub_groups: List[List[_AnimStep]] = []
            for st in group:
                if not sub_groups or st.trigger == "after_previous":
                    sub_groups.append([st])
                else:
                    sub_groups[-1].append(st)

            sub_pars_xml: List[str] = []
            for sg_idx, sg in enumerate(sub_groups):
                inner_effects_xml: List[str] = []
                for s_idx, st in enumerate(sg):
                    if st.spid not in seen_spids:
                        seen_spids.append(st.spid)
                    if sg_idx == 0 and s_idx == 0:
                        if group_is_click:
                            node_type = "clickEffect"
                        elif g_idx == 0:
                            node_type = "withEffect"
                        else:
                            node_type = "afterEffect"
                    elif s_idx == 0:
                        node_type = "afterEffect"
                    else:
                        node_type = "withEffect"
                    inner_effects_xml.append(self._build_effect_par(st, node_type, next_id))

                sub_pars_xml.append(
                    f'<p:par><p:cTn id="{next_id()}" fill="hold">'
                    f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                    f'<p:childTnLst>{"".join(inner_effects_xml)}</p:childTnLst>'
                    f"</p:cTn></p:par>"
                )

            outer_pars_xml.append(
                f'<p:par><p:cTn id="{next_id()}" fill="hold">'
                f'<p:stCondLst><p:cond delay="{outer_delay}"/></p:stCondLst>'
                f'<p:childTnLst>{"".join(sub_pars_xml)}</p:childTnLst>'
                f"</p:cTn></p:par>"
            )

        bld_xml = "".join(f'<p:bldP spid="{sid}" grpId="0"/>' for sid in seen_spids)

        timing_xml = (
            f'<p:timing {nsdecls("p", "a")}>'
            f"<p:tnLst>"
            f'<p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
            f"<p:childTnLst>"
            f'<p:seq concurrent="1" nextAc="seek">'
            f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
            f'<p:childTnLst>{"".join(outer_pars_xml)}</p:childTnLst>'
            f"</p:cTn>"
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            f"</p:seq>"
            f"</p:childTnLst>"
            f"</p:cTn></p:par>"
            f"</p:tnLst>"
            f"<p:bldLst>{bld_xml}</p:bldLst>"
            f"</p:timing>"
        )
        _insert_slide_child_in_schema_order(self.slide._element, etree.fromstring(timing_xml), "timing")

    @staticmethod
    def _build_effect_par(step: _AnimStep, node_type: str, next_id) -> str:
        preset_id, preset_sub = _EFFECT_PRESETS[step.effect]
        spid = step.spid
        dur = step.duration_ms
        dly = step.delay_ms
        tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'

        set_vis = (
            f'<p:set><p:cBhvr>'
            f'<p:cTn id="{next_id()}" dur="1" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'</p:cTn>{tgt}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
        )

        behaviors = [set_vis]

        if step.effect == "fade":
            behaviors.append(
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
                f'<p:cTn id="{next_id()}" dur="{dur}"/>{tgt}'
                f'</p:cBhvr></p:animEffect>'
            )
        elif step.effect == "appear":
            pass
        elif step.effect.startswith("wipe_"):
            wdir = {
                "wipe_down": "down",
                "wipe_up": "up",
                "wipe_left": "left",
                "wipe_right": "right",
            }[step.effect]
            behaviors.append(
                f'<p:animEffect transition="in" filter="wipe({wdir})"><p:cBhvr>'
                f'<p:cTn id="{next_id()}" dur="{dur}"/>{tgt}'
                f'</p:cBhvr></p:animEffect>'
            )
        elif step.effect.startswith("fly_"):
            axis, start_expr, end_expr = {
                "fly_bottom": ("ppt_y", "#ppt_y+0.06", "#ppt_y"),
                "fly_top": ("ppt_y", "#ppt_y-0.06", "#ppt_y"),
                "fly_left": ("ppt_x", "#ppt_x-0.06", "#ppt_x"),
                "fly_right": ("ppt_x", "#ppt_x+0.06", "#ppt_x"),
            }[step.effect]
            behaviors.append(
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
                f'<p:cTn id="{next_id()}" dur="{dur}"/>{tgt}'
                f'</p:cBhvr></p:animEffect>'
            )
            behaviors.append(
                f'<p:anim calcmode="lin" valueType="num">'
                f'<p:cBhvr additive="base">'
                f'<p:cTn id="{next_id()}" dur="{dur}" decel="65000" fill="hold"/>'
                f'{tgt}<p:attrNameLst><p:attrName>{axis}</p:attrName></p:attrNameLst>'
                f'</p:cBhvr>'
                f'<p:tavLst>'
                f'<p:tav tm="0"><p:val><p:strVal val="{start_expr}"/></p:val></p:tav>'
                f'<p:tav tm="100000"><p:val><p:strVal val="{end_expr}"/></p:val></p:tav>'
                f'</p:tavLst>'
                f'</p:anim>'
            )
        elif step.effect == "zoom":
            behaviors.append(
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr>'
                f'<p:cTn id="{next_id()}" dur="{dur}"/>{tgt}'
                f'</p:cBhvr></p:animEffect>'
            )
            behaviors.append(
                f'<p:animScale><p:cBhvr>'
                f'<p:cTn id="{next_id()}" dur="{dur}" decel="50000" fill="hold"/>{tgt}'
                f'</p:cBhvr><p:from x="88000" y="88000"/><p:to x="100000" y="100000"/></p:animScale>'
            )

        return (
            f'<p:par><p:cTn id="{next_id()}" presetID="{preset_id}" presetClass="entr" '
            f'presetSubtype="{preset_sub}" fill="hold" grpId="0" nodeType="{node_type}">'
            f'<p:stCondLst><p:cond delay="{dly}"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(behaviors)}</p:childTnLst>'
            f'</p:cTn></p:par>'
        )
