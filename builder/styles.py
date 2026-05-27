#!/usr/bin/env python3

from dataclasses import dataclass, field
from typing import Dict, Optional
import yaml


@dataclass
class LayoutMetrics:
    first_line_indent_pt: float = 24.0
    cover_line_spacing_pt: float = 32.0
    chapter_before_lines: int = 80
    chapter_after_lines: int = 50
    section_before_lines: int = 50
    section_after_lines: int = 50
    subsection_before_lines: int = 0
    subsection_after_lines: int = 0
    caption_before_lines: int = 50
    caption_after_lines: int = 50

    line_spacing_pt: float = 23.0


class StyleConfig:

    def __init__(self, config: Optional[dict] = None):
        if config is None:
            config = {}
        self.fonts: Dict = config.get("fonts", {})
        self.page: Dict = config.get("page", {})
        self.layout = LayoutMetrics(**config.get("layout", {}))
        self.figures_path: str = config.get("figures_path", "figures")
        self.caption: Dict = config.get("caption", {})
        self.content: Dict = config.get("content", {})
        self.required_sections = config.get("required_sections", [])

    def font(self, key: str, fallback: str = "body") -> Dict:
        return self.fonts.get(key, self.fonts.get(fallback, {"name": "宋体", "size": 12}))

    @property
    def use_native_caption(self) -> bool:
        return self.caption.get("native", False)


def load_style_config(yaml_path: str) -> StyleConfig:
    with open(yaml_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}
    return StyleConfig(config)
