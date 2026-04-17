"""
Report Utilities
================
Shared helpers for persisting pipeline artefacts to disk.

Functions
---------
  save_pipeline_log(log_rows, out_dir, label)
      Write a semicolon-delimited CSV with one row per pipeline step:
      timestamp | step | status | message | duration_s

  save_summary_csv(summary, cfg, out_dir, label)
      Write run metadata followed by the LLM summary text line-by-line.

  save_kpi_csv(numerics, out_dir, label)
      Write a tidy CSV with one row per (signal_type, object, metric):
      signal_type | object | metric | value | unit

  save_final_report_txt(improvements, final_report, out_dir, label)
      Write the reviewer corrections and final report as plain text.

  save_report_docx(final_report, out_dir, label, plot_paths=None)
      Build a .docx file using only the Python standard library (zipfile +
      raw OOXML strings).  Embeds all PNG plots as inline images.
      No dependency on python-docx.
"""

import csv as csv_mod
import os
import re
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape

import numpy as np


def save_pipeline_log(log_rows: list[dict], out_dir: str, label: str) -> str:
    path = os.path.join(out_dir, f"{label}_pipeline_log.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv_mod.DictWriter(
            f,
            fieldnames=["timestamp", "step", "status", "message", "duration_s"],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerows(log_rows)
    return path


def save_summary_csv(summary: str, cfg, out_dir: str, label: str) -> str:
    path = os.path.join(out_dir, f"{label}_llm_summary.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv_mod.writer(f, delimiter=";")
        writer.writerow(["section", "key", "value"])
        writer.writerow(["metadata", "run_label", cfg.run_label])
        writer.writerow(["metadata", "study_case", cfg.study_case])
        writer.writerow(["metadata", "fault_element", cfg.fault_element])
        writer.writerow(["metadata", "t_fault", cfg.t_fault])
        writer.writerow(["metadata", "t_clear", cfg.t_clear])
        writer.writerow(["metadata", "t_end", cfg.t_end])
        writer.writerow(["metadata", "timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        writer.writerow([])
        writer.writerow(["section", "line_no", "text"])
        for i, line in enumerate(summary.splitlines(), start=1):
            writer.writerow(["llm_summary", i, line])
    return path


def save_kpi_csv(numerics: dict, out_dir: str, label: str) -> str:
    rows = []

    unit_map = {
        "voltages": "p.u.",
        "speeds": "p.u.",
        "angles": "rad",
    }
    deg_metrics = {"delta_max"}

    for sig_type, obj_dict in [
        ("voltages", numerics.get("voltages", {})),
        ("speeds", numerics.get("speeds", {})),
        ("angles", numerics.get("angles", {})),
    ]:
        unit = unit_map[sig_type]
        for obj, kpis in sorted(obj_dict.items()):
            for metric, value in kpis.items():
                if sig_type == "angles" and metric in deg_metrics:
                    value = float(np.degrees(value))
                    metric_unit = "deg"
                else:
                    metric_unit = unit
                rows.append(
                    {
                        "signal_type": sig_type,
                        "object": obj,
                        "metric": metric,
                        "value": round(float(value), 6),
                        "unit": metric_unit,
                    }
                )

    path = os.path.join(out_dir, f"{label}_kpi_summary.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv_mod.DictWriter(
            f,
            fieldnames=["signal_type", "object", "metric", "value", "unit"],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerows(rows)
    return path


def save_final_report_txt(improvements: str, final_report: str,
                          out_dir: str, label: str) -> str:
    path = os.path.join(out_dir, f"{label}_final_report.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write("═" * 60 + "\n")
        f.write("  REVIEW IMPROVEMENTS\n")
        f.write("═" * 60 + "\n")
        f.write(improvements + "\n\n")
        f.write("═" * 60 + "\n")
        f.write("  FINAL REPORT\n")
        f.write("═" * 60 + "\n")
        f.write(final_report + "\n")
    return path


def save_report_docx(final_report: str, out_dir: str, label: str,
                     plot_paths: list[str] | None = None) -> str:
    path = os.path.join(out_dir, f"{label}_full_report.docx")
    plot_paths = [p for p in (plot_paths or []) if os.path.isfile(p)]

    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    ns_pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    rel_img = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    ct_png = '  <Default Extension="png" ContentType="image/png"/>\n' if plot_paths else ""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        + ct_png
        + '  <Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument'
        '.wordprocessingml.document.main+xml"/>\n'
        '</Types>\n'
    )

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>\n'
        '</Relationships>\n'
    )

    img_rel_lines = []
    for i, _ in enumerate(plot_paths, start=1):
        img_rel_lines.append(
            f'  <Relationship Id="rImg{i}" Type="{rel_img}" Target="media/image{i}.png"/>'
        )
    word_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        + "\n".join(img_rel_lines)
        + "\n"
        + '</Relationships>\n'
    )

    def para(text: str, bold: bool = False, heading: bool = False) -> str:
        safe = escape(text)
        run_props = "<w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr>" if bold or heading else ""
        para_props = (
            "<w:pPr><w:pStyle w:val=\"Heading1\"/><w:spacing w:before=\"240\" w:after=\"60\"/></w:pPr>"
            if heading
            else ""
        )
        return (
            f"<w:p>{para_props}<w:r>{run_props}"
            f"<w:t xml:space=\"preserve\">{safe}</w:t></w:r></w:p>"
        )

    def drawing(rel_id: str, img_name: str, cx: int, cy: int, draw_id: int) -> str:
        return (
            f'<w:p><w:r><w:drawing>'
            f'<wp:inline xmlns:wp="{ns_wp}" distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{draw_id}" name="{img_name}"/>'
            f'<wp:cNvGraphicFramePr>'
            f'<a:graphicFrameLocks xmlns:a="{ns_a}" noChangeAspect="1"/>'
            f'</wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="{ns_a}">'
            f'<a:graphicData uri="{ns_pic}">'
            f'<pic:pic xmlns:pic="{ns_pic}">'
            f'<pic:nvPicPr>'
            f'<pic:cNvPr id="{draw_id}" name="{img_name}"/>'
            f'<pic:cNvPicPr/>'
            f'</pic:nvPicPr>'
            f'<pic:blipFill>'
            f'<a:blip xmlns:r="{ns_r}" r:embed="{rel_id}"/>'
            f'<a:stretch><a:fillRect/></a:stretch>'
            f'</pic:blipFill>'
            f'<pic:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'</pic:spPr>'
            f'</pic:pic>'
            f'</a:graphicData>'
            f'</a:graphic>'
            f'</wp:inline>'
            f'</w:drawing></w:r></w:p>'
        )

    heading_re = re.compile(
        r"^(\d+\.\s+[A-Z]|VOLTAGE STABILITY|ROTOR ANGLE|SPEED STABILITY|OVERALL VERDICT|FINAL REPORT)"
    )

    cx_full = 5943600

    body_parts = []
    body_parts.append(para("FINAL REPORT", heading=True))
    body_parts.append(para(""))
    for line in final_report.splitlines() or [""]:
        is_heading = bool(heading_re.match(line.strip()))
        body_parts.append(para(line, bold=is_heading, heading=is_heading))

    if plot_paths:
        body_parts.append(para(""))
        body_parts.append(para("SIMULATION PLOTS", heading=True))
        for i, img_path in enumerate(plot_paths, start=1):
            stem = os.path.splitext(os.path.basename(img_path))[0]
            caption = stem.replace("_", " ").title()
            body_parts.append(para(caption, bold=True))

            if img_path.endswith("_dashboard.png"):
                cy = int(cx_full * 8 / 14)
            else:
                cy = int(cx_full * 5 / 12)

            body_parts.append(drawing(f"rImg{i}", f"image{i}.png", cx_full, cy, draw_id=i))
            body_parts.append(para(""))

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{ns_w}">'
        "<w:body>"
        + "".join(body_parts)
        + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
        "</w:body></w:document>"
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", word_rels)
        zf.writestr("word/document.xml", document_xml)
        for i, img_path in enumerate(plot_paths, start=1):
            with open(img_path, "rb") as f:
                zf.writestr(f"word/media/image{i}.png", f.read())

    return path
