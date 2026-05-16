#!/usr/bin/env python3

import re
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Metadata:
    title: str = ""
    english_title: str = ""
    student_id: str = ""
    student_name: str = ""
    advisor: str = ""
    co_advisor: str = ""
    college: str = ""
    major: str = ""
    year: str = ""
    month: str = ""


@dataclass
class Figure:
    filename: str
    caption: str = ""
    scale: float = 1.0
    width: Optional[float] = None
    line_no: int = 0
    assigned_number: str = ""


@dataclass
class CodeBlock:
    language: str
    filename: str
    content: str
    line_no: int = 0


@dataclass
class Table:
    caption: str = ""
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    line_no: int = 0
    assigned_number: str = ""


@dataclass
class PlantUMLBlock:
    content: str
    caption: str = ""
    scale: float = 1.0
    line_no: int = 0
    assigned_number: str = ""


@dataclass
class Reference:
    index: int
    content: str


LEADING_DIGITS_RE = re.compile(r"^[\d.\s]+")

CITATION_BRACKET_RE = re.compile(r"\[(\d+(?:\s*[-,，]\s*\d+)*)\]")
CITATION_DISPLAY_RE = re.compile(r'(\[\d+(?:\s*[-,，]\s*\d+)*\])')


def normalize_title(title: str) -> str:
    return LEADING_DIGITS_RE.sub("", title).strip()


@dataclass
class Section:
    level: int
    title: str
    items: list = field(default_factory=list)
    subsections: List['Section'] = field(default_factory=list)
    has_summary: bool = False
    summary_content: str = ""
    title_line_no: int = 0
    auto_number: str = ""

    @property
    def content(self) -> List[str]:
        return [item for item in self.items if isinstance(item, str)]

    @property
    def figures(self) -> List[Figure]:
        return [item for item in self.items if isinstance(item, Figure)]

    @property
    def tables(self) -> List[Table]:
        return [item for item in self.items if isinstance(item, Table)]

    @property
    def uml_blocks(self) -> List[PlantUMLBlock]:
        return [item for item in self.items if isinstance(item, PlantUMLBlock)]

    @property
    def codes(self) -> List[CodeBlock]:
        return [item for item in self.items if isinstance(item, CodeBlock)]

    def iter_subsections(self):
        for sub in self.subsections:
            yield sub
            yield from sub.iter_subsections()


@dataclass
class Thesis:
    metadata: Metadata = field(default_factory=Metadata)
    basedir: Optional[str] = None
    abstract: List[str] = field(default_factory=list)
    abstract_keywords: List[str] = field(default_factory=list)
    english_abstract: List[str] = field(default_factory=list)
    english_keywords: List[str] = field(default_factory=list)
    sections: List[Section] = field(default_factory=list)
    references: List[Reference] = field(default_factory=list)
    appendix_sections: List[Section] = field(default_factory=list)
    acknowledgments: List[str] = field(default_factory=list)
    parse_errors: List[str] = field(default_factory=list)
    numbering_warnings: List[str] = field(default_factory=list)

    def iter_sections(self):
        for s in self.sections:
            yield s
            yield from s.iter_subsections()
