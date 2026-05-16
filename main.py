#!/usr/bin/env python3
"""NEU thesis builder — Markdown DSL → NEU-compliant .docx"""

import sys
import os
import argparse
import logging

sys.path.insert(0, os.path.dirname(__file__))

from parser.markdown import parse_thesis
from checker.content import ThesisChecker
from builder.document import DocumentBuilder
from builder.styles import load_style_config

__version__ = "0.1.0"

log = logging.getLogger("thesis-builder")


class _Formatter(logging.Formatter):
    def format(self, record):
        if record.levelno >= logging.ERROR:
            return f"thesis-builder: error: {record.getMessage()}"
        if record.levelno >= logging.WARNING:
            return f"thesis-builder: warning: {record.getMessage()}"
        return record.getMessage()


def _setup_logging(quiet: bool, verbose: bool):
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(_Formatter())
    root = logging.getLogger()
    root.addHandler(handler)
    if quiet:
        root.setLevel(logging.WARNING)
    elif verbose:
        root.setLevel(logging.DEBUG)
    else:
        root.setLevel(logging.INFO)


def main():
    parser = argparse.ArgumentParser(
        prog="thesis-builder",
        description="NEU thesis builder — Markdown DSL → NEU-compliant .docx",
    )
    parser.add_argument("input", help="Thesis source file (Markdown DSL)")
    parser.add_argument("-o", "--output", help="Output .docx path", default="thesis.docx")
    parser.add_argument("--check-only", action="store_true", help="Check only, do not build")
    parser.add_argument("-q", "--quiet", action="store_true", help="Suppress non-error output")
    parser.add_argument("-y", "--yes", action="store_true", help="Auto-confirm prompts")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    parser.add_argument("--version", action="version", version=f"thesis-builder {__version__}")

    args = parser.parse_args()
    _setup_logging(args.quiet, args.verbose)

    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output) if os.path.isabs(args.output) \
                  else os.path.abspath(os.path.join(os.getcwd(), args.output))

    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "format.yaml")
    style_config = load_style_config(config_path)

    log.info("Parsing %s ...", input_path)
    try:
        thesis = parse_thesis(input_path)
    except Exception as e:
        log.error("Parse failed: %s", e)
        return 1

    parts = [f"{len(thesis.sections)} chapters"]
    if thesis.references:
        parts.append(f"{len(thesis.references)} references")
    if thesis.parse_errors:
        parts.append(f"{len(thesis.parse_errors)} parse warnings")
    log.info("  done (%s)", ", ".join(parts))

    if args.verbose:
        log.debug("  title: %s", thesis.metadata.title)
        log.debug("  abstract: %d chars", len("".join(thesis.abstract)))
        for err in thesis.parse_errors[:5]:
            log.warning("  %s", err)

    checker = ThesisChecker(styles=style_config)
    results = checker.check_all(thesis)
    errors = sum(1 for r in results if r.severity == "error")
    warnings = sum(1 for r in results if r.severity == "warning")

    if errors or warnings:
        log.info("Checking content ...")
        for r in results:
            if r.severity == "error":
                log.error("[%s] %s", r.category, r.message)
            elif r.severity == "warning":
                log.warning("[%s] %s", r.category, r.message)
            else:
                log.info("  ok [%s] %s", r.category, r.message)
    else:
        log.info("Checking content ... OK")

    if thesis.numbering_warnings:
        for w in thesis.numbering_warnings[:5]:
            log.warning("[%s]", w)

    summary_parts = []
    if errors:
        summary_parts.append(f"{errors} error{'s' if errors > 1 else ''}")
    if warnings:
        summary_parts.append(f"{warnings} warning{'s' if warnings > 1 else ''}")
    if summary_parts:
        log.info("  %s", ", ".join(summary_parts))

    if args.check_only:
        return 1 if errors else 0

    if errors > 0 and not args.yes:
        try:
            response = input("Continue anyway? [y/N] ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            return 1
        if response != "y":
            return 1

    log.info("Building %s ...", output_path)
    try:
        builder = DocumentBuilder(config=style_config)
        builder.build(thesis, output_path)
    except Exception as e:
        log.error("Build failed: %s", e)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1
    log.info("  done")

    return 0


if __name__ == "__main__":
    sys.exit(main())
