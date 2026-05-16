#!/usr/bin/env python3

import os
import re
from typing import List, Dict
from dataclasses import dataclass
from ast_nodes import (
    Thesis, Section, Metadata,
    normalize_title, CITATION_BRACKET_RE,
    Figure, Table, PlantUMLBlock, CodeBlock,
)
from builder.styles import StyleConfig


def _preview(items, max_n=5):
    preview = ", ".join(str(x) for x in items[:max_n])
    suffix = "..." if len(items) > max_n else ""
    return preview + suffix


@dataclass
class CheckResult:
    passed: bool
    category: str
    message: str
    severity: str = "info"


class ThesisChecker:

    def __init__(self, styles: StyleConfig):
        self.styles = styles
        self._init_config()

    def _init_config(self):
        content_cfg = self.styles.content
        abstract_cfg = content_cfg["abstract"]
        references_cfg = content_cfg["references"]
        metadata_cfg = content_cfg.get("metadata", {})
        chapter_summary_cfg = content_cfg.get("chapter_summary", {})

        required_sections = self.styles.required_sections
        required_sections = [s for s in required_sections if s not in ("参考文献", "致谢")]

        chapter_rules = []
        for _, chapter in content_cfg.get("chapters", {}).items():
            if not isinstance(chapter, dict) or not chapter.get("name"):
                continue
            chapter_rules.append({
                "name": chapter["name"],
                "alternatives": chapter.get("alternatives", []),
                "min_ratio": chapter.get("min_ratio"),
                "max_ratio": chapter.get("max_ratio"),
                "ratio": chapter.get("ratio"),
                "forbid_summary": chapter.get("forbid_summary", False),
            })

        self.title_max_chars = metadata_cfg.get("title_max_chars", 24)
        self.abstract_cfg = {
            "min_chars": abstract_cfg["min_chars"],
            "max_chars": abstract_cfg["max_chars"],
            "keywords_min": abstract_cfg["keywords_min"],
            "keywords_max": abstract_cfg["keywords_max"],
        }
        self.references_cfg = {
            "min_count": references_cfg["min_count"],
        }
        self.required_sections = required_sections
        self.chapter_summary_cfg = {
            "min_length": chapter_summary_cfg.get("min_length", 80),
        }
        self.chapter_rules = chapter_rules
        self.figures_path = self.styles.figures_path
        self.code_cfg = {
            "max_lines": content_cfg.get("code_max_lines", 60),
        }
        self.chart_cfg = {
            "max_consecutive": content_cfg.get("chart_density", {}).get("max_consecutive", 3),
        }

    def check_all(self, thesis: Thesis) -> List[CheckResult]:
        results = []

        results.extend(self._check_metadata(thesis.metadata))
        results.extend(self._check_abstract(thesis))
        results.extend(self._check_english_abstract(thesis))
        results.extend(self._check_structure(thesis))
        results.extend(self._check_chapter_ratios(thesis))
        results.extend(self._check_summaries(thesis))
        results.extend(self._check_references(thesis))
        results.extend(self._check_reference_citations(thesis))
        results.extend(self._check_figures(thesis))
        results.extend(self._check_code_blocks(thesis))
        results.extend(self._check_chart_density(thesis))

        return results

    def _check_keyword_count(self, count: int, category: str, label: str,
                             kw_min: int, kw_max: int,
                             empty_severity: str = "error") -> List[CheckResult]:
        results = []
        if count == 0:
            results.append(CheckResult(False, category, f"未找到{label}", empty_severity))
        elif count < kw_min:
            results.append(CheckResult(
                False, category,
                f"{label}数量不足（{count}个，要求≥{kw_min}个）", empty_severity,
            ))
        elif count > kw_max:
            results.append(CheckResult(
                False, category,
                f"{label}数量过多（{count}个，要求≤{kw_max}个）", "warning",
            ))
        else:
            results.append(CheckResult(True, category, f"{label}{count}个", "info"))
        return results

    @staticmethod
    def _fuzzy_match_title(title: str, candidate: str) -> bool:
        return candidate == title or candidate in title or title in candidate

    def _section_char_count(self, section: Section) -> int:
        total = sum(len(line) for line in section.content if line.strip())
        if section.summary_content:
            total += len(section.summary_content)
        for sub in section.iter_subsections():
            total += sum(len(line) for line in sub.content if line.strip())
            if sub.summary_content:
                total += len(sub.summary_content)
        return total

    def _find_section_for_rule(self, sections: List[Section], rule: Dict) -> Section:
        candidates = [rule["name"], *rule.get("alternatives", [])]
        normalized_candidates = [c.strip() for c in candidates if c and c.strip()]

        for section in sections:
            title = normalize_title(section.title)
            if any(self._fuzzy_match_title(title, c) for c in normalized_candidates):
                return section
        return None

    def _check_metadata(self, metadata: Metadata) -> List[CheckResult]:
        results = []

        if not metadata.title:
            results.append(CheckResult(False, "元数据", "论文题目未填写", "error"))
        elif len(metadata.title) > self.title_max_chars:
            results.append(CheckResult(False, "元数据", f"论文题目过长（{len(metadata.title)}字，要求≤{self.title_max_chars}字）", "error"))

        if not metadata.student_id:
            results.append(CheckResult(False, "元数据", "学号未填写", "error"))

        if not metadata.student_name:
            results.append(CheckResult(False, "元数据", "姓名未填写", "error"))

        if not metadata.advisor:
            results.append(CheckResult(False, "元数据", "指导教师未填写", "error"))

        if not metadata.college:
            results.append(CheckResult(False, "元数据", "学院未填写", "error"))

        if not metadata.major:
            results.append(CheckResult(False, "元数据", "专业未填写", "error"))

        if not metadata.year or not metadata.month:
            results.append(CheckResult(False, "元数据", "年月未填写", "error"))

        return results

    def _check_abstract(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        abstract_text = "".join(thesis.abstract)
        char_count = len(abstract_text)

        if char_count == 0:
            results.append(CheckResult(False, "摘要", "中文摘要为空", "error"))
        elif char_count < self.abstract_cfg["min_chars"]:
            results.append(CheckResult(
                False, "摘要",
                f"摘要字数不足（{char_count}字，要求≥{self.abstract_cfg['min_chars']}字）",
                "error",
            ))
        elif char_count > self.abstract_cfg["max_chars"]:
            results.append(CheckResult(
                False, "摘要",
                f"摘要字数过多（{char_count}字，要求≤{self.abstract_cfg['max_chars']}字）",
                "warning",
            ))
        else:
            results.append(CheckResult(True, "摘要", f"摘要字数符合要求（{char_count}字）", "info"))

        keyword_count = len(thesis.abstract_keywords)
        results.extend(self._check_keyword_count(
            keyword_count, "关键词", "关键词",
            self.abstract_cfg["keywords_min"], self.abstract_cfg["keywords_max"],
            empty_severity="error",
        ))

        para_count = sum(1 for p in thesis.abstract if p.strip())
        if para_count < 3:
            results.append(CheckResult(False, "摘要结构", f"摘要只有{para_count}段，建议采用三段式结构", "info"))

        return results

    def _check_english_abstract(self, thesis: Thesis) -> List[CheckResult]:
        results = []

        if not thesis.english_abstract:
            results.append(CheckResult(False, "英文摘要", "英文摘要为空", "error"))
            return results

        keyword_count = len(thesis.english_keywords)
        results.extend(self._check_keyword_count(
            keyword_count, "英文关键词", "英文关键词",
            self.abstract_cfg["keywords_min"], self.abstract_cfg["keywords_max"],
            empty_severity="warning",
        ))

        return results

    def _check_structure(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        section_titles = [normalize_title(s.title) for s in thesis.sections]

        for required in self.required_sections:
            found = any(
                self._fuzzy_match_title(title, required)
                for title in section_titles
            )
            if found:
                results.append(CheckResult(True, "章节结构", f"包含: {required}", "info"))
            else:
                results.append(CheckResult(False, "章节结构", f"缺少章节: {required}", "error"))

        return results

    def _check_chapter_ratios(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        chapter_rules = self.chapter_rules
        if not chapter_rules:
            return results

        section_chars = {id(s): self._section_char_count(s) for s in thesis.sections}
        total_chars = sum(section_chars.values())
        if total_chars <= 0:
            return [CheckResult(False, "章节比例", "正文内容为空，无法计算章节比例", "warning")]

        tolerance = 0.03
        for rule in chapter_rules:
            section = self._find_section_for_rule(thesis.sections, rule)
            if not section:
                continue

            ratio = section_chars[id(section)] / total_chars
            ratio_text = f"{ratio * 100:.1f}%"

            min_ratio = rule.get("min_ratio")
            max_ratio = rule.get("max_ratio")
            target_ratio = rule.get("ratio")

            if min_ratio is not None and ratio < min_ratio:
                results.append(CheckResult(
                    False, "章节比例",
                    f"{rule['name']}篇幅占比偏低（{ratio_text}，要求≥{min_ratio * 100:.0f}%）",
                    "error",
                ))
                continue

            if max_ratio is not None and ratio > max_ratio:
                results.append(CheckResult(
                    False, "章节比例",
                    f"{rule['name']}篇幅占比偏高（{ratio_text}，要求≤{max_ratio * 100:.0f}%）",
                    "error",
                ))
                continue

            if target_ratio is not None:
                lower = max(0.0, target_ratio - tolerance)
                upper = target_ratio + tolerance
                if ratio < lower or ratio > upper:
                    results.append(CheckResult(
                        False, "章节比例",
                        f"{rule['name']}篇幅占比偏离建议值（{ratio_text}，建议约{target_ratio * 100:.0f}%）",
                        "warning",
                    ))

        return results

    def _check_summaries(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        min_length = self.chapter_summary_cfg["min_length"]

        for rule in self.chapter_rules:
            section = self._find_section_for_rule(thesis.sections, rule)
            if not section:
                continue

            normalized_title = normalize_title(section.title)

            if rule.get("forbid_summary"):
                if section.has_summary:
                    results.append(CheckResult(
                        False, f"{normalized_title}小结",
                        f"{normalized_title}不应包含本章小结",
                        "error",
                    ))
                continue

            if not section.has_summary:
                results.append(CheckResult(False, f"{normalized_title}小结", "缺少本章小结", "warning"))
                continue

            if len(section.summary_content.strip()) < min_length:
                results.append(CheckResult(
                    False, f"{normalized_title}小结",
                    f"本章小结过短，建议不少于{min_length}字",
                    "warning",
                ))
            else:
                results.append(CheckResult(True, f"{normalized_title}小结", "已包含本章小结", "info"))

        return results

    def _check_references(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        ref_count = len(thesis.references)
        min_count = self.references_cfg["min_count"]

        if ref_count == 0:
            results.append(CheckResult(False, "参考文献", "未找到参考文献", "error"))
        elif ref_count < min_count:
            results.append(CheckResult(False, "参考文献", f"参考文献数量不足（{ref_count}条，要求≥{min_count}条）", "error"))
        else:
            results.append(CheckResult(True, "参考文献", f"参考文献数量符合要求（{ref_count}条）", "info"))

        return results

    def _extract_citation_numbers_ordered(self, text: str) -> List[int]:
        seen = set()
        result = []
        for match in CITATION_BRACKET_RE.finditer(text):
            for part in re.split(r"[,，]", match.group(1)):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    start_text, end_text = part.split("-", 1)
                    if start_text.strip().isdigit() and end_text.strip().isdigit():
                        start, end = int(start_text.strip()), int(end_text.strip())
                        if start <= end:
                            for n in range(start, end + 1):
                                if n not in seen:
                                    seen.add(n)
                                    result.append(n)
                elif part.isdigit():
                    n = int(part)
                    if n not in seen:
                        seen.add(n)
                        result.append(n)
        return result

    def _check_reference_citations(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        ref_count = len(thesis.references)
        if ref_count == 0:
            return results

        body_text = "\n".join(
            line
            for section in thesis.iter_sections()
            for line in section.content
        )
        ordered = self._extract_citation_numbers_ordered(body_text)
        cited_numbers = set(ordered)

        if not cited_numbers:
            results.append(CheckResult(False, "参考文献引用", "正文中未检测到参考文献引用标注", "error"))
            return results

        invalid = sorted(n for n in cited_numbers if n < 1 or n > ref_count)
        if invalid:
            results.append(CheckResult(
                False, "参考文献引用",
                f"存在越界引用编号（{_preview(invalid)}），参考文献共{ref_count}条",
                "error",
            ))

        valid_cited = {n for n in cited_numbers if 1 <= n <= ref_count}
        uncited = sorted(set(range(1, ref_count + 1)) - valid_cited)
        if uncited:
            results.append(CheckResult(
                False, "参考文献引用",
                f"存在未在正文引用的参考文献（共{len(uncited)}条，如{_preview(uncited)}）",
                "error",
            ))
        else:
            results.append(CheckResult(True, "参考文献引用", "所有参考文献均在正文中被引用", "info"))

        valid_ordered = [n for n in ordered if 1 <= n <= ref_count]
        if all(valid_ordered[i] == i + 1 for i in range(len(valid_ordered))):
            results.append(CheckResult(True, "参考文献引用顺序", "参考文献按首次引用顺序编号", "info"))
        else:
            first_mismatch = next(
                (i + 1 for i, n in enumerate(valid_ordered) if n != i + 1),
                None,
            )
            if first_mismatch:
                results.append(CheckResult(
                    False, "参考文献引用顺序",
                    f"参考文献应按正文中首次引用的顺序编号，"
                    f"第{first_mismatch}处首次出现的引用为[{valid_ordered[first_mismatch - 1]}]，期望为[{first_mismatch}]",
                    "warning",
                ))

        return results

    def _check_figures(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        figures = []
        tables = []
        uml_blocks = []
        for section in thesis.iter_sections():
            figures.extend(section.figures)
            tables.extend(section.tables)
            uml_blocks.extend(section.uml_blocks)

        total = len(figures) + len(tables) + len(uml_blocks)
        if total == 0:
            results.append(CheckResult(False, "图表", "未找到任何图表引用", "warning"))
            return results

        results.append(CheckResult(
            True,
            "图表",
            f"包含{len(figures)}个图片引用，{len(tables)}个表格引用，{len(uml_blocks)}个PlantUML图",
            "info",
        ))

        if not thesis.basedir:
            return results

        figures_dir = os.path.join(thesis.basedir, self.figures_path)
        missing_assets = []

        for figure in figures:
            if not os.path.exists(os.path.join(figures_dir, figure.filename)):
                missing_assets.append(f"figure:{figure.filename}")

        if missing_assets:
            results.append(CheckResult(
                False, "图表资源",
                f"存在缺失图表资源（共{len(missing_assets)}项，如{_preview(missing_assets)}）",
                "warning",
            ))

        return results

    def _check_code_blocks(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        max_lines = self.code_cfg["max_lines"]

        long_blocks = []
        for section in thesis.iter_sections():
            for item in section.items:
                if isinstance(item, CodeBlock):
                    lines = item.content.count('\n') + 1
                    if lines > max_lines:
                        long_blocks.append(
                            f"{item.filename}（{lines}行）"
                        )

        if long_blocks:
            results.append(CheckResult(
                False, "代码块",
                f"存在过长代码块（建议≤{max_lines}行/页）：{'; '.join(long_blocks)}",
                "warning",
            ))
        else:
            results.append(CheckResult(True, "代码块", "所有代码块长度合适", "info"))
        return results

    def _check_chart_density(self, thesis: Thesis) -> List[CheckResult]:
        results = []
        max_consecutive = self.chart_cfg["max_consecutive"]
        chart_types = (Figure, Table, PlantUMLBlock)

        warnings_found = 0
        for section in thesis.iter_sections():
            consecutive = 0
            for item in section.items:
                if isinstance(item, chart_types):
                    consecutive += 1
                elif isinstance(item, (str, CodeBlock)):
                    if consecutive > max_consecutive:
                        results.append(CheckResult(
                            False, "图表密度",
                            f"\"{section.title[:15]}\" 连续{consecutive}个图表无文字间隔",
                            "warning",
                        ))
                        warnings_found += 1
                    consecutive = 0

            if consecutive > max_consecutive:
                results.append(CheckResult(
                    False, "图表密度",
                    f"\"{section.title[:15]}\" 结尾连续{consecutive}个图表无文字间隔",
                    "warning",
                ))
                warnings_found += 1

        if warnings_found == 0:
            results.append(CheckResult(True, "图表密度", "图表与文字交替排列良好", "info"))
        return results

