#!/usr/bin/env python3

import os
import re

from ast_nodes import (
    Metadata, Figure, CodeBlock, Table, PlantUMLBlock, Reference,
    Section, Thesis,
)

_CODE_START_RE = re.compile(r'@code\{([^,]+),\s*([^}]+)\}')
_DIRECTIVE_RE = re.compile(r'@(figure|table|plantuml)\{(.*)\}')
_KW_CN_RE = re.compile(r'^关键(?:词|字)[：:]*\s*')
_KW_EN_RE = re.compile(r'^Key\s*words[:\s]', re.IGNORECASE)
_REF_RE = re.compile(r'^\[\d+\]')
_HEADING_RE = re.compile(r'^#{1,6}\s+(.*)')


_META_FIELDS = {"title", "english_title", "student_id", "student_name",
                "english_name", "advisor", "english_advisor",
                "co_advisor", "english_co_advisor",
                "college", "major", "date"}


def _consume_block(lines, pos):
    content_lines = []
    while pos < len(lines):
        stripped = lines[pos].strip()
        if stripped == "@end":
            return pos + 1, "\n".join(content_lines)
        content_lines.append(lines[pos])
        pos += 1
    return pos, "\n".join(content_lines)


def _consume_table_block(lines, pos):
    headers = []
    rows = []
    while pos < len(lines):
        stripped = lines[pos].strip()
        if stripped == "@end":
            return pos + 1, headers, rows
        if not stripped or '|' not in stripped:
            pos += 1
            continue
        row = _parse_pipe_row(lines[pos])
        if not row:
            pos += 1
            continue
        if not headers:
            headers = row
        elif all(re.match(r'^[-:]+$', cell) for cell in row):
            pass
        else:
            rows.append(row)
        pos += 1
    return pos, headers, rows


def _parse_pipe_row(line):
    stripped = line.strip()
    if not stripped or '|' not in stripped:
        return []
    if stripped.startswith('|'):
        stripped = stripped[1:]
    if stripped.endswith('|'):
        stripped = stripped[:-1]
    return [cell.strip() for cell in stripped.split('|')]


def _parse_directive(line):
    match = _DIRECTIVE_RE.match(line.strip())
    if not match:
        return None
    inner = match.group(2).strip()
    if not inner:
        return None
    comma = inner.find(',')
    if comma == -1:
        if '=' in inner:
            return [], _parse_options(inner)
        return [inner.strip()], {}
    positional = [inner[:comma].strip()]
    options = _parse_options(inner[comma + 1:])
    return positional, options


def _parse_options(raw):
    options = {}
    if not raw:
        return options
    for token in raw.split(","):
        token = token.strip()
        if not token or "=" not in token:
            continue
        k, v = token.split("=", 1)
        options[k.strip().lower()] = v.strip().strip('"').strip("'")
    return options


def _match_heading(line):
    m = _HEADING_RE.match(line)
    if m:
        level = len(m.group(0)) - len(m.group(1)) - 1
        return level, m.group(1).strip()
    return None


def _is_special_heading(title):
    return (title in ('参考文献', 'References')
            or title in ('致谢', 'Acknowledgments')
            or title.startswith('附录')
            or 'Appendix' in title)


def _skip_heading(lines, pos, match_fn):
    while pos < len(lines):
        h = _match_heading(lines[pos])
        if h:
            _, title = h
            if match_fn(title):
                return pos
        pos += 1
    return len(lines)


def _parse_metadata(lines, pos):
    meta = Metadata()
    if pos >= len(lines) or lines[pos].strip() != "---":
        return pos, meta
    pos += 1
    while pos < len(lines) and lines[pos].strip() != "---":
        stripped = lines[pos].strip()
        if ":" in stripped:
            k, v = stripped.split(":", 1)
            k, v = k.strip(), v.strip()
            if k in _META_FIELDS:
                setattr(meta, k, v)
        pos += 1
    if pos < len(lines):
        pos += 1
    return pos, meta


def _parse_abstract(lines, pos, lang):
    abstract = []
    keywords = []
    target = "摘要" if lang == "cn" else "ABSTRACT"
    match = (lambda t: t.strip() == target) if lang == "cn" else (lambda t: t.strip().upper() == "ABSTRACT")
    pos = _skip_heading(lines, pos, match)
    if pos >= len(lines):
        return pos, abstract, keywords
    pos += 1
    while pos < len(lines):
        stripped = lines[pos].strip()
        if not stripped:
            pos += 1
            continue
        h = _match_heading(stripped)
        if h:
            break
        if lang == "cn":
            m = _KW_CN_RE.match(stripped)
            if m:
                kw_text = stripped[m.end():]
                keywords = [k.strip() for k in kw_text.split('；') if k.strip()]
                pos += 1
                break
        if lang == "en":
            m = _KW_EN_RE.match(stripped)
            if m:
                kw_text = stripped[m.end():]
                keywords = [k.strip() for k in kw_text.split(';') if k.strip()]
                pos += 1
                break
        abstract.append(stripped)
        pos += 1
    return pos, abstract, keywords


def _parse_sections(lines, pos, thesis):
    section_stack = []
    current_chapter = None

    while pos < len(lines):
        stripped = lines[pos].strip()
        line_no = pos + 1

        if not stripped:
            pos += 1
            continue

        h = _match_heading(stripped)
        if h:
            level, title = h
            if _is_special_heading(title):
                break

            if '本章小结' in title:
                if current_chapter:
                    current_chapter.has_summary = True

            section = Section(level=level, title=title, title_line_no=line_no)
            if level == 1:
                thesis.sections.append(section)
                section_stack = [section]
                current_chapter = section
            else:
                while section_stack and section_stack[-1].level >= level:
                    section_stack.pop()
                if section_stack:
                    section_stack[-1].subsections.append(section)
                section_stack.append(section)
            pos += 1
            continue

        if section_stack:
            pos = _handle_directive(lines, pos, line_no, section_stack[-1], thesis)
        else:
            pos += 1

    return pos


def _handle_directive(lines, pos, line_no, section, thesis):
    stripped = lines[pos].strip()

    if stripped.startswith("@figure{"):
        result = _parse_directive(stripped)
        if result:
            parts, options = result
            section.items.append(Figure(
                filename=parts[0],
                caption=options.get("caption", ""),
                scale=float(options.get("scale", "1.0")),
                line_no=line_no,
            ))
        else:
            thesis.parse_errors.append(f"行 {line_no}: 无法解析 @figure 指令")
        return pos + 1

    if stripped.startswith("@table{"):
        result = _parse_directive(stripped)
        if result:
            _, options = result
            pos += 1
            pos, headers, rows = _consume_table_block(lines, pos)
            section.items.append(Table(
                caption=options.get("caption", ""),
                headers=headers,
                rows=rows,
                line_no=line_no,
            ))
        else:
            thesis.parse_errors.append(f"行 {line_no}: 无法解析 @table 指令")
        return pos

    code_match = _CODE_START_RE.match(stripped)
    if code_match:
        language = code_match.group(1).strip()
        filename = code_match.group(2).strip()
        pos += 1
        pos, content = _consume_block(lines, pos)
        section.items.append(CodeBlock(
            language=language,
            filename=filename,
            content=content,
            line_no=line_no,
        ))
        return pos

    if stripped.startswith("@plantuml{"):
        result = _parse_directive(stripped)
        if result:
            positional, options = result
            if positional:
                options.update(_parse_options(positional[0]))
            pos += 1
            pos, content = _consume_block(lines, pos)
            section.items.append(PlantUMLBlock(
                content=content,
                caption=options.get("caption", ""),
                scale=float(options.get("scale", "1.0")),
                line_no=line_no,
            ))
        else:
            thesis.parse_errors.append(f"行 {line_no}: 无法解析 @plantuml 指令")
        return pos

    section.items.append(stripped)
    return pos + 1


def _parse_references(lines, pos):
    refs = []
    pos = _skip_heading(lines, pos, lambda t: t.strip() in ('参考文献', 'References'))
    if pos >= len(lines):
        return pos, refs
    pos += 1
    while pos < len(lines):
        stripped = lines[pos].strip()
        if _match_heading(stripped):
            break
        if _REF_RE.match(stripped):
            refs.append(Reference(index=len(refs) + 1, content=stripped))
        pos += 1
    return pos, refs


def _parse_appendix(lines, pos, thesis):
    sections = []
    pos = _skip_heading(lines, pos, lambda t: '附录' in t or 'Appendix' in t)
    if pos >= len(lines):
        return sections

    while pos < len(lines):
        stripped = lines[pos].strip()
        line_no = pos + 1
        h = _match_heading(stripped)
        if h:
            _, title = h
            if title in ('致谢', 'Acknowledgments'):
                break
            section = Section(level=h[0], title=title, title_line_no=line_no)
            sections.append(section)
            pos += 1
            continue
        if sections and stripped:
            pos = _handle_directive(lines, pos, line_no, sections[-1], thesis)
        else:
            pos += 1
    return sections


def _parse_acknowledgments(lines, pos):
    acks = []
    pos = _skip_heading(lines, pos, lambda t: t.strip() in ('致谢', 'Acknowledgments'))
    if pos >= len(lines):
        return pos, acks
    pos += 1
    while pos < len(lines):
        stripped = lines[pos].strip()
        if _match_heading(stripped):
            break
        if stripped:
            acks.append(stripped)
        pos += 1
    return pos, acks


def parse_thesis(filepath: str) -> Thesis:
    thesis = Thesis()
    thesis.basedir = os.path.dirname(filepath)

    with open(filepath, 'r', encoding='utf-8') as f:
        raw = f.read()

    lines = raw.split('\n')

    pos = 0
    pos, thesis.metadata = _parse_metadata(lines, pos)
    pos, thesis.abstract, thesis.abstract_keywords = _parse_abstract(lines, pos, 'cn')
    pos, thesis.english_abstract, thesis.english_keywords = _parse_abstract(lines, pos, 'en')
    pos = _parse_sections(lines, pos, thesis)
    pos, thesis.references = _parse_references(lines, pos)
    thesis.appendix_sections = _parse_appendix(lines, pos, thesis)
    _, thesis.acknowledgments = _parse_acknowledgments(lines, pos)

    return thesis
