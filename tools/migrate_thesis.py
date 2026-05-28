#!/usr/bin/env python3
import re
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from ast_nodes import Section, Reference, Thesis


def is_heading(para):
    if para.style.name.startswith('Heading'):
        return int(para.style.name.replace('Heading ', ''))

    text = para.text.strip()
    if re.match(r'^(\d+|第\d+章)\s+\S', text):
        return 1
    # 节标题: "1.1 研究背景"
    if re.match(r'^\d+\.\d+\s+\S', text):
        return 2
    # 小节标题: "1.1.1 研究"
    if re.match(r'^\d+\.\d+\.\d+\s+\S', text):
        return 3
    return 0


def extract_text_from_docx(doc_path: str) -> Thesis:
    """从Word文档中提取内容"""
    doc = Document(doc_path)
    thesis = Thesis()

    state = None  # 'abstract', 'english_abstract', 'references', 'acknowledgments'
    current_section = None
    section_stack = []

    abstract_lines = []
    english_abstract_lines = []
    acknowledgments_lines = []

    _SECTION_MAP = {
        '摘  要': 'abstract', '摘要': 'abstract',
        'ABSTRACT': 'english_abstract', 'Abstract': 'english_abstract', '英文摘要': 'english_abstract',
        '参 考 文 献': 'references', '参考文献': 'references',
        '致  谢': 'acknowledgments', '致谢': 'acknowledgments', 'Acknowledgments': 'acknowledgments',
    }

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        heading_level = is_heading(para)

        if text in _SECTION_MAP:
            state = _SECTION_MAP[text]
            continue

        # 收集内容
        if state == 'abstract':
            if text.startswith('关键词'):
                kw_text = text.replace('关键词', '').replace('：', '').replace(':', '').strip()
                thesis.abstract_keywords = [k.strip() for k in kw_text.split('；') if k.strip()]
            else:
                abstract_lines.append(text)
        elif state == 'english_abstract':
            if text.startswith('Key words') or text.startswith('Keywords'):
                kw_text = text.replace('Key words', '').replace('Keywords', '').replace(':', '').strip()
                thesis.english_keywords = [k.strip() for k in kw_text.split(';') if k.strip()]
            else:
                english_abstract_lines.append(text)
        elif state == 'references':
            if text.startswith('[') and ']' in text:
                thesis.references.append(Reference(index=len(thesis.references)+1, content=text))
        elif state == 'acknowledgments':
            acknowledgments_lines.append(text)
        else:
            if heading_level > 0:
                section = Section(level=heading_level, title=text)

                if heading_level == 1:
                    thesis.sections.append(section)
                    section_stack = [section]
                else:
                    while len(section_stack) > heading_level - 1:
                        section_stack.pop()
                    if section_stack:
                        section_stack[-1].subsections.append(section)
                    section_stack.append(section)

                current_section = section
            elif current_section:
                current_section.items.append(text)

    thesis.abstract = abstract_lines
    thesis.english_abstract = english_abstract_lines
    thesis.acknowledgments = acknowledgments_lines

    return thesis


def thesis_to_markdown(thesis: Thesis, output_path: str):
    """将Thesis对象转换为Markdown DSL格式"""
    lines = []

    # YAML frontmatter
    lines.append('---')
    lines.append(f'title: {thesis.metadata.title or "论文题目"}')
    lines.append(f'english_title: {thesis.metadata.english_title or "Thesis Title"}')
    lines.append(f'student_id: {thesis.metadata.student_id or "XXXXXXXXXX"}')
    lines.append(f'student_name: {thesis.metadata.student_name or "姓名"}')
    lines.append(f'advisor: {thesis.metadata.advisor or "导师姓名 教授"}')
    lines.append(f'college: {thesis.metadata.college or "软件学院"}')
    lines.append(f'major: {thesis.metadata.major or "软件工程"}')
    lines.append(f'date: {thesis.metadata.date or "2025-06"}')
    lines.append('---')
    lines.append('')

    # 中文摘要
    lines.append('# 摘要')
    lines.append('')
    for line in thesis.abstract:
        lines.append(line)
    lines.append('')
    if thesis.abstract_keywords:
        lines.append(f"关键词：{'；'.join(thesis.abstract_keywords)}")
    lines.append('')

    # 英文摘要
    lines.append('# ABSTRACT')
    lines.append('')
    for line in thesis.english_abstract:
        lines.append(line)
    lines.append('')
    if thesis.english_keywords:
        lines.append(f"Key words: {'; '.join(thesis.english_keywords)}")
    lines.append('')

    # 正文
    def add_section(section, level=1):
        prefix = '#' * level
        lines.append(f'{prefix} {section.title}')
        lines.append('')
        for content in section.content:
            lines.append(content)
        lines.append('')
        for subsection in section.subsections:
            add_section(subsection, level + 1)

    for section in thesis.sections:
        add_section(section)

    # 参考文献
    if thesis.references:
        lines.append('# 参考文献')
        lines.append('')
        for ref in thesis.references:
            lines.append(ref.content)
        lines.append('')

    # 致谢
    if thesis.acknowledgments:
        lines.append('# 致谢')
        lines.append('')
        for line in thesis.acknowledgments:
            lines.append(line)
        lines.append('')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    print(f"已生成: {output_path}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} input.docx [output.md]")
        sys.exit(1)
    input_doc = sys.argv[1]
    output_md = sys.argv[2] if len(sys.argv) > 2 else 'output.md'

    print(f"读取: {input_doc}")
    thesis = extract_text_from_docx(input_doc)

    print(f"  - 摘要: {len(thesis.abstract)} 行")
    print(f"  - 英文摘要: {len(thesis.english_abstract)} 行")
    print(f"  - 章节: {len(thesis.sections)} 章")
    print(f"  - 参考文献: {len(thesis.references)} 条")
    print(f"  - 致谢: {len(thesis.acknowledgments)} 行")

    print(f"\n写入: {output_md}")
    thesis_to_markdown(thesis, output_md)

    print("\n完成!")
