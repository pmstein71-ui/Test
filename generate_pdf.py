#!/usr/bin/env python3
"""Generate Repository Structure & Ozon Checklist Analysis PDF and text report."""

from __future__ import annotations

import glob
import os
import subprocess
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


def nfc(text: object) -> str:
    """Normalize text to NFC Unicode form for clean PDF rendering."""
    if text is None:
        return ""
    return unicodedata.normalize("NFC", str(text)).strip()


def register_cyrillic_fonts() -> tuple[str, str]:
    """Register DejaVuSans regular and bold TTF fonts supporting Cyrillic."""
    candidates = [
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ),
        (
            "/usr/share/fonts/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
        ),
    ]
    for reg_path, bold_path in candidates:
        if os.path.exists(reg_path) and os.path.exists(bold_path):
            pdfmetrics.registerFont(TTFont("DejaVuSans", reg_path))
            pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", bold_path))
            return "DejaVuSans", "DejaVuSans-Bold"
    raise FileNotFoundError(
        "DejaVuSans TTF fonts not found. Please install fonts-dejavu-core."
    )


def get_git_info() -> dict[str, str]:
    """Collect current repository metadata."""
    info = {
        "branch": os.environ.get("GITHUB_REF_NAME", ""),
        "commit": os.environ.get("GITHUB_SHA", ""),
        "repo": os.environ.get("GITHUB_REPOSITORY", "pmstein71-ui/Test"),
    }
    try:
        if not info["branch"]:
            info["branch"] = subprocess.check_output(
                ["git", "rev-parse", "--abbrev-ref", "HEAD"], text=True
            ).strip()
    except Exception:
        info["branch"] = "unknown"

    try:
        commit_short = subprocess.check_output(
            ["git", "log", "-1", "--format=%h %s"], text=True
        ).strip()
    except Exception:
        commit_short = info["commit"][:7] if info["commit"] else "unknown"
    info["commit_summary"] = nfc(commit_short)
    return info


def collect_repo_files(root: Path) -> tuple[list[tuple[str, int]], list[str]]:
    """Return list of workspace files with sizes and any .exe files."""
    files_info: list[tuple[str, int]] = []
    exe_files: list[str] = []
    for p in sorted(root.rglob("*")):
        if ".git" in p.parts or "__pycache__" in p.parts:
            continue
        if p.is_file():
            rel = nfc(str(p.relative_to(root)))
            size = p.stat().st_size
            files_info.append((rel, size))
            if p.suffix.lower() == ".exe":
                exe_files.append(rel)
    return files_info, exe_files


def load_excel_data(xlsx_path: str) -> dict[str, list[list[str]]]:
    """Read sheets and non-empty rows from the Excel workbook."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    data: dict[str, list[list[str]]] = {}
    for name in wb.sheetnames:
        ws = wb[name]
        rows: list[list[str]] = []
        for row in ws.iter_rows(values_only=True):
            cleaned = [nfc(cell) for cell in row]
            if any(cleaned):
                rows.append(cleaned)
        data[nfc(name)] = rows
    return data


def escape_xml(text: str, allow_slash_wrap: bool = False) -> str:
    """Escape XML special characters and preserve line breaks for ReportLab Paragraph."""
    escaped = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    if allow_slash_wrap:
        escaped = escaped.replace("/", "/\u200b")
    return escaped.replace("\n", "<br/>")


def build_pdf_and_report(
    pdf_path: str = "Structure_Analysis.pdf",
    txt_path: str = "analysis_report.txt",
) -> None:
    font_reg, font_bold = register_cyrillic_fonts()
    root = Path(".").resolve()
    git_info = get_git_info()
    files_info, exe_files = collect_repo_files(root)

    xlsx_files = sorted(glob.glob("*.xlsx"))
    excel_data: dict[str, list[list[str]]] = {}
    if xlsx_files:
        excel_data = load_excel_data(xlsx_files[0])

    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    # Prepare text report
    report_lines = [
        "============================================================",
        "Repository Structure Analysis & Ozon Checklist Report",
        "============================================================",
        f"Generated at : {now_str}",
        f"Repository   : {git_info['repo']}",
        f"Branch       : {git_info['branch']}",
        f"Commit       : {git_info['commit_summary']}",
        f".exe files   : {len(exe_files)} (No executable files found)",
        "",
        "Files in repository:",
    ]
    for rel, size in files_info:
        report_lines.append(f"  - {rel} ({size} bytes)")

    if excel_data:
        report_lines.append("")
        report_lines.append(f"Excel Workbook: {nfc(xlsx_files[0])}")
        for sheet_name, rows in excel_data.items():
            report_lines.append(f"  * Sheet '{sheet_name}': {len(rows)} non-empty rows")

    Path(txt_path).write_text("\n".join(report_lines) + "\n", encoding="utf-8")

    # Build PDF using ReportLab Platypus (Landscape A4 for wide tables)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Title"],
        fontName=font_bold,
        fontSize=15,
        leading=19,
        alignment=0,
        spaceAfter=6,
        textColor=colors.HexColor("#1A365D"),
    )
    h2_style = ParagraphStyle(
        "SectionHeader",
        parent=styles["Heading2"],
        fontName=font_bold,
        fontSize=12,
        leading=15,
        spaceBefore=10,
        spaceAfter=6,
        textColor=colors.HexColor("#2B6CB0"),
    )
    body_style = ParagraphStyle(
        "BodyCyr",
        parent=styles["Normal"],
        fontName=font_reg,
        fontSize=9,
        leading=12,
        spaceAfter=4,
    )
    cell_style = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName=font_reg,
        fontSize=7.5,
        leading=9.5,
    )
    cell_header_style = ParagraphStyle(
        "TableHeaderCell",
        parent=styles["Normal"],
        fontName=font_bold,
        fontSize=7.5,
        leading=9.5,
        textColor=colors.white,
    )

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Repository Structure Analysis & Ozon Operational Readiness Checklist",
        author="GitHub Actions PDF Builder",
    )

    story = []

    # Section 1: Repository Structure Analysis
    story.append(
        Paragraph(
            escape_xml(
                "Repository Structure Analysis / Анализ структуры репозитория"
            ),
            title_style,
        )
    )
    meta_lines = [
        f"<b>Repository:</b> {escape_xml(git_info['repo'])}",
        f"<b>Branch:</b> {escape_xml(git_info['branch'])}",
        f"<b>Commit:</b> {escape_xml(git_info['commit_summary'])}",
        f"<b>Generated:</b> {escape_xml(now_str)}",
        f"<b>.exe files found:</b> {len(exe_files)} (Analysis: No executable files found)",
    ]
    for line in meta_lines:
        story.append(Paragraph(line, body_style))

    story.append(Spacer(1, 4 * mm))
    story.append(
        Paragraph(escape_xml("Files in repository / Файлы в репозитории:"), h2_style)
    )

    file_table_data = [
        [
            Paragraph("<b>№</b>", cell_header_style),
            Paragraph("<b>Файл / Путь (File Path)</b>", cell_header_style),
            Paragraph("<b>Размер (Bytes)</b>", cell_header_style),
        ]
    ]
    for idx, (rel, size) in enumerate(files_info, 1):
        file_table_data.append(
            [
                Paragraph(str(idx), cell_style),
                Paragraph(escape_xml(rel), cell_style),
                Paragraph(f"{size:,}", cell_style),
            ]
        )

    file_table = Table(file_table_data, colWidths=[15 * mm, 210 * mm, 40 * mm])
    file_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2B6CB0")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [colors.white, colors.HexColor("#F7FAFC")],
                ),
                ("PADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(file_table)

    # Section 2: Excel Workbook Data
    if excel_data:
        for sheet_name, rows in excel_data.items():
            story.append(PageBreak())
            story.append(
                Paragraph(
                    escape_xml(f"Лист: {sheet_name} ({nfc(xlsx_files[0])})"),
                    title_style,
                )
            )

            # Separate leading/trailing single-cell note rows from table rows
            table_rows: list[list[str]] = []
            footer_notes: list[str] = []
            for r in rows:
                non_empty = [c for c in r if c]
                if len(non_empty) == 1 and not table_rows:
                    story.append(Paragraph(escape_xml(non_empty[0]), body_style))
                elif len(non_empty) == 1 and table_rows:
                    footer_notes.append(non_empty[0])
                else:
                    table_rows.append(r)

            if not table_rows:
                continue

            story.append(Spacer(1, 3 * mm))
            num_cols = max(len(r) for r in table_rows)
            usable_width = 273 * mm

            if num_cols == 8:
                col_widths = [
                    27 * mm,
                    47 * mm,
                    34 * mm,
                    31 * mm,
                    36 * mm,
                    28 * mm,
                    50 * mm,
                    20 * mm,
                ]
            elif num_cols == 4:
                col_widths = [45 * mm, 120 * mm, 53 * mm, 55 * mm]
            else:
                col_widths = [usable_width / num_cols] * num_cols

            formatted_rows = []
            for r_idx, r in enumerate(table_rows):
                padded = (r + [""] * num_cols)[:num_cols]
                style_to_use = cell_header_style if r_idx == 0 else cell_style
                formatted_rows.append(
                    [
                        Paragraph(
                            escape_xml(cell, allow_slash_wrap=True), style_to_use
                        )
                        for cell in padded
                    ]
                )

            sheet_table = Table(
                formatted_rows, colWidths=col_widths, repeatRows=1
            )
            sheet_table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E0")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        (
                            "ROWBACKGROUNDS",
                            (0, 1),
                            (-1, -1),
                            [colors.white, colors.HexColor("#F7FAFC")],
                        ),
                        ("LEFTPADDING", (0, 0), (-1, -1), 3),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ]
                )
            )
            story.append(sheet_table)

            if footer_notes:
                story.append(Spacer(1, 3 * mm))
                for note in footer_notes:
                    story.append(Paragraph(escape_xml(note), body_style))

    doc.build(story)
    print(f"Successfully generated {pdf_path} ({os.path.getsize(pdf_path)} bytes)")
    print(f"Successfully generated {txt_path} ({os.path.getsize(txt_path)} bytes)")


if __name__ == "__main__":
    build_pdf_and_report()
