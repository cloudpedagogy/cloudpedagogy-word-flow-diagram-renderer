#!/usr/bin/env python3
"""Render a robust interactive flow diagram from tables in a Word document."""

from __future__ import annotations

import argparse
import html
import json
import re
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from docx import Document

VERSION = "1.0.0"
ORIENTATIONS = {"LR", "RL", "TB", "BT"}
NODE_ALIASES = {
    "id": {"id", "node id", "node_id", "key"},
    "label": {"label", "name", "title", "node"},
    "type": {"type", "kind", "shape", "node type"},
    "description": {"description", "details", "summary", "notes"},
    "status": {"status", "state"},
    "color": {"color", "colour", "fill", "fill colour", "fill color"},
    "link": {"link", "url", "web link"},
    "group": {"group", "phase", "section", "lane", "swimlane"},
}
LINK_ALIASES = {
    "source": {"source", "from", "source id", "source_id", "start"},
    "target": {"target", "to", "target id", "target_id", "end"},
    "label": {"label", "name", "text", "condition"},
    "style": {"style", "line style", "line_style"},
    "color": {"color", "colour", "line color", "line colour"},
    "description": {"description", "details", "notes"},
}
SETTING_ALIASES = {
    "title": {"title", "visualisation title", "visualization title"},
    "subtitle": {"subtitle", "description"},
    "orientation": {"orientation", "direction", "layout"},
    "theme": {"theme"},
    "show_legend": {"show legend", "legend", "show_legend"},
    "allow_orientation_switching": {
        "allow orientation switching", "orientation switching", "switch orientation"
    },
}
TYPE_MAP = {
    "process": "process", "step": "process", "activity": "process", "task": "process",
    "decision": "decision", "choice": "decision", "question": "decision",
    "start": "terminal", "end": "terminal", "terminal": "terminal",
    "input": "data", "output": "data", "data": "data", "document": "document",
    "database": "database", "store": "database", "delay": "delay",
}
DEFAULT_COLORS = {
    "process": "#DBEAFE", "decision": "#FEF3C7", "terminal": "#DCFCE7",
    "data": "#EDE9FE", "document": "#FCE7F3", "database": "#E0F2FE",
    "delay": "#FEE2E2",
}


@dataclass
class Issue:
    level: str
    message: str
    table: str | None = None
    row: int | None = None


@dataclass
class ParseResult:
    settings: dict[str, str] = field(default_factory=dict)
    nodes: list[dict[str, Any]] = field(default_factory=list)
    links: list[dict[str, Any]] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)


def normalise(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("_", " ").strip().lower())


def clean_cell(cell) -> str:
    return "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip()).strip()


def canonical(raw: str, aliases: dict[str, set[str]]) -> str | None:
    value = normalise(raw)
    for key, variants in aliases.items():
        if value == normalise(key) or value in {normalise(v) for v in variants}:
            return key
    return None


def valid_link(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def valid_color(value: str) -> bool:
    return bool(re.fullmatch(r"#[0-9a-fA-F]{6}", value) or re.fullmatch(r"[a-zA-Z]+", value))


def as_bool(value: str, default: bool = True) -> bool:
    text = normalise(value)
    if text in {"true", "yes", "1", "on"}:
        return True
    if text in {"false", "no", "0", "off"}:
        return False
    return default


def iter_blocks(document):
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, document)
        elif child.tag.endswith("}tbl"):
            yield Table(child, document)


def table_rows(table) -> list[list[str]]:
    return [[clean_cell(cell) for cell in row.cells] for row in table.rows]


def detect_table(rows: list[list[str]], heading: str) -> str | None:
    title = normalise(heading)
    if title in {"settings", "flow settings", "configuration"}:
        return "settings"
    if title in {"nodes", "flow nodes", "steps", "activities"}:
        return "nodes"
    if title in {"links", "edges", "connections", "transitions"}:
        return "links"
    if not rows:
        return None
    headers = {normalise(v) for v in rows[0]}
    if headers & NODE_ALIASES["id"] and headers & NODE_ALIASES["label"]:
        return "nodes"
    if headers & LINK_ALIASES["source"] and headers & LINK_ALIASES["target"]:
        return "links"
    if ({"setting", "value"} <= headers) or ({"key", "value"} <= headers):
        return "settings"
    return None


def parse_settings(rows: list[list[str]], result: ParseResult) -> None:
    headers = [normalise(x) for x in rows[0]] if rows else []
    try:
        key_i = next(i for i, h in enumerate(headers) if h in {"setting", "key", "option", "name"})
        value_i = next(i for i, h in enumerate(headers) if h in {"value", "setting value"})
    except StopIteration:
        result.issues.append(Issue("warning", "SETTINGS table ignored: Setting and Value columns are required.", "SETTINGS"))
        return
    for row_num, row in enumerate(rows[1:], 2):
        if max(key_i, value_i) >= len(row) or not row[key_i].strip():
            continue
        key = canonical(row[key_i], SETTING_ALIASES)
        if key:
            result.settings[key] = row[value_i].strip()
        else:
            result.issues.append(Issue("warning", f"Unknown setting '{row[key_i]}' ignored.", "SETTINGS", row_num))


def mapped_headers(raw_headers: list[str], aliases: dict[str, set[str]], table: str,
                   result: ParseResult) -> tuple[list[str | None], set[str]]:
    mapped: list[str | None] = []
    seen: set[str] = set()
    for raw in raw_headers:
        key = canonical(raw, aliases)
        if key and key in seen:
            result.issues.append(Issue("warning", f"Duplicate '{key}' column; later column ignored.", table))
            key = None
        if key:
            seen.add(key)
        elif raw.strip():
            result.issues.append(Issue("warning", f"Unknown {table} column '{raw}' ignored.", table))
        mapped.append(key)
    return mapped, seen


def row_record(row: list[str], headers: list[str | None]) -> dict[str, str]:
    return {headers[i]: value.strip() for i, value in enumerate(row)
            if i < len(headers) and headers[i]}


def parse_nodes(rows: list[list[str]], result: ParseResult) -> None:
    if len(rows) < 2:
        result.issues.append(Issue("warning", "NODES table has no data rows.", "NODES"))
        return
    headers, seen = mapped_headers(rows[0], NODE_ALIASES, "NODES", result)
    if not {"id", "label"} <= seen:
        result.issues.append(Issue("error", "NODES requires ID and Label/Name columns.", "NODES"))
        return
    known: set[str] = set()
    for row_num, row in enumerate(rows[1:], 2):
        record = row_record(row, headers)
        if not any(record.values()):
            continue
        node_id, label = record.get("id", ""), record.get("label", "")
        if not node_id or not label:
            result.issues.append(Issue("error", "Missing node ID or label. Row skipped.", "NODES", row_num))
            continue
        if node_id in known:
            result.issues.append(Issue("error", f"Duplicate node ID '{node_id}'. Row skipped.", "NODES", row_num))
            continue
        known.add(node_id)
        raw_type = normalise(record.get("type", "process")) or "process"
        node_type = TYPE_MAP.get(raw_type, "process")
        if raw_type not in TYPE_MAP:
            result.issues.append(Issue("warning", f"Unknown node type '{record.get('type')}' treated as process.", "NODES", row_num))
        color = record.get("color", "")
        if color and not valid_color(color):
            result.issues.append(Issue("warning", f"Invalid colour '{color}' ignored.", "NODES", row_num))
            color = ""
        link = record.get("link", "")
        if link and not valid_link(link):
            result.issues.append(Issue("warning", f"Unsafe or invalid link '{link}' ignored.", "NODES", row_num))
            link = ""
        result.nodes.append({
            "id": node_id, "label": label, "type": node_type,
            "description": record.get("description", ""), "status": record.get("status", ""),
            "color": color or DEFAULT_COLORS[node_type], "link": link,
            "group": record.get("group", ""), "_row": row_num,
        })


def parse_links(rows: list[list[str]], result: ParseResult) -> None:
    if len(rows) < 2:
        result.issues.append(Issue("warning", "LINKS table has no data rows.", "LINKS"))
        return
    headers, seen = mapped_headers(rows[0], LINK_ALIASES, "LINKS", result)
    if not {"source", "target"} <= seen:
        result.issues.append(Issue("error", "LINKS requires Source/From and Target/To columns.", "LINKS"))
        return
    for row_num, row in enumerate(rows[1:], 2):
        record = row_record(row, headers)
        if not any(record.values()):
            continue
        source, target = record.get("source", ""), record.get("target", "")
        if not source or not target:
            result.issues.append(Issue("error", "Missing source or target. Link skipped.", "LINKS", row_num))
            continue
        style = normalise(record.get("style", "solid")) or "solid"
        if style not in {"solid", "dashed", "dotted", "thick"}:
            result.issues.append(Issue("warning", f"Unknown link style '{style}' treated as solid.", "LINKS", row_num))
            style = "solid"
        color = record.get("color", "")
        if color and not valid_color(color):
            result.issues.append(Issue("warning", f"Invalid link colour '{color}' ignored.", "LINKS", row_num))
            color = ""
        result.links.append({
            "source": source, "target": target, "label": record.get("label", ""),
            "style": style, "color": color, "description": record.get("description", ""),
            "_row": row_num,
        })


def validate(result: ParseResult) -> None:
    ids = {n["id"] for n in result.nodes}
    valid_links = []
    for link in result.links:
        missing = [x for x in (link["source"], link["target"]) if x not in ids]
        if missing:
            result.issues.append(Issue("error", f"Link references missing node(s): {', '.join(missing)}. Link removed.", "LINKS", link["_row"]))
        elif link["source"] == link["target"]:
            result.issues.append(Issue("warning", f"Self-link on '{link['source']}' retained.", "LINKS", link["_row"]))
            valid_links.append(link)
        else:
            valid_links.append(link)
    result.links = valid_links
    orientation = result.settings.get("orientation", "LR").upper().replace("-", "")
    orientation = {"LEFTTORIGHT": "LR", "RIGHTTOLEFT": "RL", "TOPTOBOTTOM": "TB",
                   "TOPDOWN": "TB", "BOTTOMTOTOP": "BT"}.get(orientation, orientation)
    if orientation not in ORIENTATIONS:
        result.issues.append(Issue("warning", f"Unknown orientation '{orientation}'; using LR.", "SETTINGS"))
        orientation = "LR"
    result.settings["orientation"] = orientation
    theme = normalise(result.settings.get("theme", "light"))
    if theme not in {"light", "dark", "neutral"}:
        result.issues.append(Issue("warning", f"Unknown theme '{theme}'; using light.", "SETTINGS"))
        theme = "light"
    result.settings["theme"] = theme
    if not result.nodes:
        result.issues.append(Issue("error", "No valid nodes were found."))
    if result.nodes and not result.links:
        result.issues.append(Issue("warning", "No valid links were found; nodes will be unconnected."))


def read_docx(path: Path) -> ParseResult:
    result = ParseResult()
    document = Document(path)
    heading = ""
    counts = {"settings": 0, "nodes": 0, "links": 0}
    for block in iter_blocks(document):
        if hasattr(block, "style") and block.style and block.style.name.startswith("Heading"):
            heading = block.text.strip()
            continue
        if not hasattr(block, "rows"):
            continue
        rows = table_rows(block)
        kind = detect_table(rows, heading)
        if not kind:
            continue
        counts[kind] += 1
        if kind == "settings":
            parse_settings(rows, result)
        elif kind == "nodes":
            parse_nodes(rows, result)
        else:
            parse_links(rows, result)
    if not counts["nodes"]:
        result.issues.append(Issue("error", "No NODES table was found."))
    if not counts["links"]:
        result.issues.append(Issue("warning", "No LINKS table was found."))
    validate(result)
    return result


def safe_id(raw: str, used: set[str]) -> str:
    base = "n_" + re.sub(r"[^A-Za-z0-9_]", "_", raw)
    candidate, number = base, 2
    while candidate in used:
        candidate, number = f"{base}_{number}", number + 1
    used.add(candidate)
    return candidate


def mermaid_label(text: str) -> str:
    return text.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")


def diagram_text(result: ParseResult, orientation: str) -> tuple[str, dict[str, str]]:
    used: set[str] = set()
    ids = {n["id"]: safe_id(n["id"], used) for n in result.nodes}
    lines = [f"flowchart {orientation}"]
    shapes = {
        "process": ('["', '"]'),
        "decision": ('{"', '"}'),
        "terminal": ('(["', '"])'),
        "data": ('[/"', '"/]'),
        "document": ('[["', '"]]'),
        "database": ('[("', '")]'),
        "delay": ('(["', '"])'),
    }
    for node in result.nodes:
        start, end = shapes[node["type"]]
        lines.append(f"  {ids[node['id']]}{start}{mermaid_label(node['label'])}{end}")
    for index, link in enumerate(result.links):
        source, target = ids[link["source"]], ids[link["target"]]
        label = mermaid_label(link["label"])
        connector = "-.->" if link["style"] in {"dashed", "dotted"} else "==>" if link["style"] == "thick" else "-->"
        if label:
            connector = f"-. {label} .->" if link["style"] in {"dashed", "dotted"} else f"== {label} ==>" if link["style"] == "thick" else f"-- {label} -->"
        lines.append(f"  {source} {connector} {target}")
        if link["color"]:
            width = "3px" if link["style"] == "thick" else "2px"
            dash = ",stroke-dasharray:5 4" if link["style"] in {"dashed", "dotted"} else ""
            lines.append(f"  linkStyle {index} stroke:{link['color']},stroke-width:{width}{dash}")
    for node in result.nodes:
        lines.append(f"  style {ids[node['id']]} fill:{node['color']},stroke:#334155,stroke-width:1.5px,color:#0f172a")
    return "\n".join(lines), ids


def payload(result: ParseResult) -> dict[str, Any]:
    clean_nodes = [{k: v for k, v in n.items() if not k.startswith("_")} for n in result.nodes]
    clean_links = [{k: v for k, v in e.items() if not k.startswith("_")} for e in result.links]
    return {"schema_version": "1.0", "generator_version": VERSION,
            "settings": result.settings, "nodes": clean_nodes, "links": clean_links}


def build_html(result: ParseResult, mermaid_js: str) -> str:
    data = payload(result)
    diagrams, id_maps = {}, {}
    for orientation in sorted(ORIENTATIONS):
        diagrams[orientation], id_maps[orientation] = diagram_text(result, orientation)
    title = html.escape(result.settings.get("title", "Interactive flow diagram"))
    subtitle = html.escape(result.settings.get("subtitle", "Generated from Word tables"))
    initial = result.settings["orientation"]
    switch = as_bool(result.settings.get("allow_orientation_switching", "true"))
    show_legend = as_bool(result.settings.get("show_legend", "true"))
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><style>
:root{{--ink:#0f172a;--muted:#64748b;--line:#cbd5e1;--panel:#fff;--bg:#f8fafc;--accent:#1d4ed8}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 system-ui,-apple-system,Segoe UI,sans-serif}}
header{{padding:22px 28px 14px;background:#fff;border-bottom:1px solid var(--line)}}h1{{margin:0;font-size:1.55rem}}header p{{margin:.3rem 0 0;color:var(--muted)}}
.toolbar{{display:flex;gap:10px;flex-wrap:wrap;align-items:end;padding:12px 28px;background:#fff;border-bottom:1px solid var(--line)}}
label{{font-size:.78rem;font-weight:700;color:#475569}}input,select,button{{font:inherit;border:1px solid #94a3b8;border-radius:6px;background:#fff;padding:7px 9px}}
button{{cursor:pointer}}button:hover{{background:#eff6ff}}.field{{display:grid;gap:3px}}#search{{width:230px}}
main{{display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:14px;padding:14px;height:calc(100vh - 169px);min-height:520px}}
.canvas,.details{{background:var(--panel);border:1px solid var(--line);border-radius:10px;overflow:hidden}}.viewport{{height:100%;overflow:auto;position:relative}}
#diagram{{min-width:100%;min-height:100%;padding:34px;transform-origin:top left}}#diagram svg{{max-width:none!important;height:auto}}
.details{{padding:18px;overflow:auto}}.details h2{{font-size:1.05rem;margin:0 0 10px}}.details dl{{margin:0}}.details dt{{font-size:.72rem;text-transform:uppercase;color:var(--muted);font-weight:700;margin-top:12px}}.details dd{{margin:2px 0;white-space:pre-wrap}}
.hint{{color:var(--muted)}}.legend{{display:flex;gap:8px;flex-wrap:wrap;padding-top:12px}}.key{{display:flex;gap:5px;align-items:center;font-size:.78rem}}.swatch{{width:13px;height:13px;border:1px solid #64748b;border-radius:3px}}
.match rect,.match polygon,.match path,.match circle,.match ellipse{{stroke:#dc2626!important;stroke-width:4px!important}}.muted-node{{opacity:.2}}
.fallback{{position:absolute;left:-10000px;width:1px;height:1px;overflow:hidden}}a{{color:var(--accent)}}@media(max-width:850px){{main{{grid-template-columns:1fr;height:auto}}.canvas{{height:65vh}}}}
</style></head><body>
<header><h1>{title}</h1><p>{subtitle}</p></header>
<div class="toolbar">
<div class="field"><label for="search">Find a node</label><input id="search" type="search" placeholder="Search labels and descriptions"></div>
<div class="field" {'hidden' if not switch else ''}><label for="orientation">Orientation</label><select id="orientation"><option value="LR">Left → right</option><option value="TB">Top → bottom</option><option value="RL">Right → left</option><option value="BT">Bottom → top</option></select></div>
<button id="zoomIn" type="button" aria-label="Zoom in">Zoom +</button><button id="zoomOut" type="button" aria-label="Zoom out">Zoom −</button><button id="reset" type="button">Reset view</button>
</div>
<main><section class="canvas" aria-label="Flow diagram"><div class="viewport"><div id="diagram" class="mermaid"></div></div></section>
<aside class="details" aria-live="polite"><h2>Node details</h2><div id="detail"><p class="hint">Select a node in the diagram to see its details.</p></div>
<div class="legend" {'hidden' if not show_legend else ''}>{''.join(f'<span class="key"><i class="swatch" style="background:{c}"></i>{html.escape(t.title())}</span>' for t,c in DEFAULT_COLORS.items())}</div></aside></main>
<table class="fallback"><caption>Flow diagram data</caption><thead><tr><th>Node</th><th>Type</th><th>Description</th></tr></thead><tbody>{''.join(f"<tr><td>{html.escape(n['label'])}</td><td>{html.escape(n['type'])}</td><td>{html.escape(n['description'])}</td></tr>" for n in result.nodes)}</tbody></table>
<script>{mermaid_js}</script><script>
const DATA={json.dumps(data, ensure_ascii=False).replace("</", "<\\/")};
const DIAGRAMS={json.dumps(diagrams, ensure_ascii=False).replace("</", "<\\/")};
const MAPS={json.dumps(id_maps, ensure_ascii=False).replace("</", "<\\/")};
let orientation={json.dumps(initial)}, scale=1;
const box=document.getElementById("diagram"), viewport=document.querySelector(".viewport"), detail=document.getElementById("detail");
document.getElementById("orientation").value=orientation;
mermaid.initialize({{startOnLoad:false,securityLevel:"strict",theme:{json.dumps(result.settings["theme"])},flowchart:{{htmlLabels:true,useMaxWidth:false,curve:"basis"}}}});
function esc(s){{return String(s??"").replace(/[&<>"']/g,c=>({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}}[c]))}}
async function draw(){{
 box.removeAttribute("data-processed");box.innerHTML=DIAGRAMS[orientation];
 await mermaid.run({{nodes:[box]}});
 const map=MAPS[orientation];
 DATA.nodes.forEach(n=>{{const el=box.querySelector(`[id*="${{CSS.escape(map[n.id])}}"]`);if(el){{el.style.cursor="pointer";el.setAttribute("tabindex","0");el.setAttribute("role","button");el.setAttribute("aria-label",n.label);el.dataset.nodeId=n.id;el.addEventListener("click",()=>show(n));el.addEventListener("keydown",e=>{{if(e.key==="Enter"||e.key===" ")show(n)}})}}}});
 applySearch();applyZoom();
}}
function show(n){{detail.innerHTML=`<h3>${{esc(n.label)}}</h3><dl><dt>Type</dt><dd>${{esc(n.type)}}</dd>${{n.group?`<dt>Group</dt><dd>${{esc(n.group)}}</dd>`:""}}${{n.status?`<dt>Status</dt><dd>${{esc(n.status)}}</dd>`:""}}${{n.description?`<dt>Description</dt><dd>${{esc(n.description)}}</dd>`:""}}${{n.link?`<dt>Link</dt><dd><a href="${{esc(n.link)}}" target="_blank" rel="noopener noreferrer">Open related page</a></dd>`:""}}</dl>`}}
function applySearch(){{const q=document.getElementById("search").value.trim().toLowerCase();document.querySelectorAll("[data-node-id]").forEach(el=>{{const n=DATA.nodes.find(x=>x.id===el.dataset.nodeId);const hit=!q||[n.label,n.description,n.group,n.status,n.type].join(" ").toLowerCase().includes(q);el.classList.toggle("match",!!q&&hit);el.classList.toggle("muted-node",!!q&&!hit)}})}}
function applyZoom(){{box.style.transform=`scale(${{scale}})`;box.style.width=`${{100/scale}}%`;box.style.height=`${{100/scale}}%`}}
document.getElementById("orientation").addEventListener("change",e=>{{orientation=e.target.value;draw()}});
document.getElementById("search").addEventListener("input",applySearch);
document.getElementById("zoomIn").onclick=()=>{{scale=Math.min(2.5,scale+.15);applyZoom()}};
document.getElementById("zoomOut").onclick=()=>{{scale=Math.max(.4,scale-.15);applyZoom()}};
document.getElementById("reset").onclick=()=>{{scale=1;viewport.scrollTo(0,0);document.getElementById("search").value="";applySearch();applyZoom()}};
draw().catch(err=>{{box.innerHTML=`<p>Unable to render diagram: ${{esc(err.message)}}</p>`;console.error(err)}});
</script></body></html>"""


def write_qa(result: ParseResult, path: Path) -> None:
    errors = sum(i.level == "error" for i in result.issues)
    warnings = sum(i.level == "warning" for i in result.issues)
    lines = ["# Flow diagram QA report", "", f"- Nodes accepted: {len(result.nodes)}",
             f"- Links accepted: {len(result.links)}", f"- Errors: {errors}",
             f"- Warnings: {warnings}", "", "## Findings", ""]
    if not result.issues:
        lines.append("No issues found.")
    else:
        for issue in result.issues:
            where = " ".join(x for x in [issue.table, f"row {issue.row}" if issue.row else ""] if x)
            lines.append(f"- **{issue.level.upper()}**{f' ({where})' if where else ''}: {issue.message}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def render(source: Path, output: Path, vendor: Path, strict: bool = False) -> ParseResult:
    result = read_docx(source)
    errors = [i for i in result.issues if i.level == "error"]
    if strict and errors:
        raise ValueError(f"{len(errors)} validation error(s); see QA report.")
    if not result.nodes:
        raise ValueError("No valid nodes are available to render.")
    if not vendor.exists():
        raise FileNotFoundError(f"Mermaid asset not found: {vendor}")
    output.mkdir(parents=True, exist_ok=True)
    (output / "data.json").write_text(json.dumps(payload(result), indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    write_qa(result, output / "qa_report.md")
    (output / "index.html").write_text(build_html(result, vendor.read_text(encoding="utf-8")), encoding="utf-8")
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Word .docx containing SETTINGS, NODES and LINKS tables")
    parser.add_argument("-o", "--output", type=Path, help="Output directory (default: output/<input-name>)")
    parser.add_argument("--vendor", type=Path, default=Path(__file__).parent / "vendor/mermaid.min.js")
    parser.add_argument("--strict", action="store_true", help="Stop if validation errors are found")
    parser.add_argument("--overwrite", action="store_true", help="Replace an existing output directory")
    parser.add_argument("--version", action="version", version=VERSION)
    args = parser.parse_args(argv)
    if args.input.suffix.lower() != ".docx" or not args.input.is_file():
        parser.error("input must be an existing .docx file")
    output = args.output or Path("output") / args.input.stem
    if output.exists() and args.overwrite:
        shutil.rmtree(output)
    try:
        result = render(args.input, output, args.vendor, args.strict)
    except (ValueError, FileNotFoundError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    errors = sum(i.level == "error" for i in result.issues)
    warnings = sum(i.level == "warning" for i in result.issues)
    print(f"Rendered {len(result.nodes)} nodes and {len(result.links)} links to {output / 'index.html'}")
    print(f"QA: {errors} error(s), {warnings} warning(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
