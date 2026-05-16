#!/usr/bin/env python3
"""Pass 1 numbering: assign section/figure/table/chapter numbers to AST nodes."""

import re
from ast_nodes import Figure, Table, PlantUMLBlock, Thesis, Section

_FIG_NUM_RE = re.compile(r'^图\s*\d+\.\d+')
_TBL_NUM_RE = re.compile(r'^表\s*\d+\.\d+')

_NUMBERING_RULES = {
    Figure: (lambda i: i.caption, _FIG_NUM_RE, "图", "图片", "fig"),
    PlantUMLBlock: (lambda i: i.caption or "", _FIG_NUM_RE, "图", "UML", "fig"),
    Table: (lambda i: i.caption, _TBL_NUM_RE, "表", "表格", "tbl"),
}


def collect_numbering(thesis: Thesis):
    chapter_num = 0
    fig_counters = {}
    tbl_counters = {}
    sec_counters = [0, 0, 0]

    for section in thesis.sections:
        if section.level == 1:
            chapter_num += 1
        _assign_section_numbers(section, sec_counters)
        _walk_for_numbering(section, chapter_num, fig_counters, tbl_counters, thesis)

    for section in thesis.appendix_sections:
        _assign_section_numbers(section, sec_counters)


def _assign_section_numbers(section, counters):
    level = min(section.level, 3)
    counters[level - 1] += 1
    for i in range(level, 3):
        counters[i] = 0
    parts = [str(counters[i]) for i in range(level)]
    section.auto_number = ".".join(parts)

    for sub in section.subsections:
        _assign_section_numbers(sub, counters)


def _walk_for_numbering(section, chapter, fig_counters, tbl_counters, thesis):
    counters_by_key = {"fig": fig_counters, "tbl": tbl_counters}

    for item in section.items:
        rule = _NUMBERING_RULES.get(type(item))
        if rule is None:
            continue
        get_caption, num_re, prefix, label, ck = rule
        counters = counters_by_key[ck]
        caption = get_caption(item)
        if caption and num_re.match(caption):
            thesis.numbering_warnings.append(
                f"行 {item.line_no}: {label}标题含手写编号 \"{caption[:20]}\"，"
                f"自动编号已跳过（建议删除手写编号）"
            )
        else:
            count = counters.get(chapter, 0) + 1
            counters[chapter] = count
            item.assigned_number = f"{prefix}{chapter}.{count}"

    for sub in section.subsections:
        _walk_for_numbering(sub, chapter, fig_counters, tbl_counters, thesis)


def caption_with_number(assigned_number: str, caption: str) -> str:
    if not assigned_number:
        return caption
    if not caption:
        return assigned_number
    return f"{assigned_number} {caption}"
