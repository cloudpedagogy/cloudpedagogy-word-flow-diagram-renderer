import ast
import json
import tempfile
import unittest
from pathlib import Path

from docx import Document

import render_flow_diagram as flow

ROOT_DIR = Path(__file__).resolve().parents[1]
VENDOR_PATH = ROOT_DIR / "vendor" / "mermaid.min.js"
SAMPLE_DOCX = ROOT_DIR / "examples" / "flow_diagram_example.docx"


def make_doc(path: Path, nodes, links, node_headers=None, link_headers=None,
             settings=None, settings_headers=None):
    doc = Document()
    if settings is not None:
        doc.add_heading("SETTINGS", 1)
        table = doc.add_table(rows=1, cols=len(settings_headers or ["Setting", "Value"]))
        for cell, value in zip(table.rows[0].cells, settings_headers or ["Setting", "Value"]):
            cell.text = value
        for values in settings:
            for cell, value in zip(table.add_row().cells, values):
                cell.text = value
    doc.add_heading("NODES", 1)
    table = doc.add_table(rows=1, cols=len(node_headers or ["ID", "Label", "Type"]))
    for cell, value in zip(table.rows[0].cells, node_headers or ["ID", "Label", "Type"]):
        cell.text = value
    for values in nodes:
        for cell, value in zip(table.add_row().cells, values):
            cell.text = value
    doc.add_heading("LINKS", 1)
    table = doc.add_table(rows=1, cols=len(link_headers or ["Source", "Target", "Label"]))
    for cell, value in zip(table.rows[0].cells, link_headers or ["Source", "Target", "Label"]):
        cell.text = value
    for values in links:
        for cell, value in zip(table.add_row().cells, values):
            cell.text = value
    doc.save(path)


class FlowTests(unittest.TestCase):
    def parse(self, nodes, links, node_headers=None, link_headers=None,
              settings=None, settings_headers=None):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.docx"
            make_doc(path, nodes, links, node_headers, link_headers, settings, settings_headers)
            return flow.read_docx(path)

    def test_valid_flow_and_aliases(self):
        result = self.parse(
            [["a", "Start", "Start"], ["b", "Review", "Decision"]],
            [["a", "b", "Continue"]],
            ["Key", "Name", "Kind"], ["From", "To", "Condition"],
        )
        self.assertEqual(2, len(result.nodes))
        self.assertEqual("decision", result.nodes[1]["type"])
        self.assertEqual(1, len(result.links))
        self.assertFalse([i for i in result.issues if i.level == "error"])

    def test_duplicate_node_removed(self):
        result = self.parse([["a", "A", "Process"], ["a", "Again", "Process"]], [])
        self.assertEqual(1, len(result.nodes))
        self.assertTrue(any("Duplicate node" in i.message for i in result.issues))

    def test_missing_link_node_removed(self):
        result = self.parse([["a", "A", "Process"]], [["a", "missing", "Next"]])
        self.assertFalse(result.links)
        self.assertTrue(any("missing node" in i.message for i in result.issues))

    def test_unsafe_link_and_bad_colour_cleaned(self):
        result = self.parse(
            [["a", "A", "Process", "javascript:alert(1)", "#bad-color"]],
            [],
            ["ID", "Label", "Type", "Link", "Colour"],
        )
        self.assertEqual("", result.nodes[0]["link"])
        self.assertEqual(flow.DEFAULT_COLORS["process"], result.nodes[0]["color"])

    def test_unknown_type_becomes_process(self):
        result = self.parse([["a", "A", "Mystery"]], [])
        self.assertEqual("process", result.nodes[0]["type"])
        self.assertTrue(any("Unknown node type" in i.message for i in result.issues))

    def test_mermaid_escapes_labels(self):
        result = self.parse([["a", 'A "quoted" <node>', "Process"]], [])
        diagram, _ = flow.diagram_text(result, "LR")
        self.assertIn("&quot;quoted&quot;", diagram)
        self.assertIn("&lt;node&gt;", diagram)

    def test_build_html_and_script_escaping(self):
        result = self.parse(
            [["a", 'Payload with </script><script>alert("xss")</script>', "Process"]],
            [],
        )
        html_out = flow.build_html(result, "/* mock-mermaid-js */")
        self.assertTrue(html_out.startswith("<!doctype html>"))
        self.assertIn("/* mock-mermaid-js */", html_out)
        self.assertIn("const DATA=", html_out)
        self.assertIn("const DIAGRAMS=", html_out)
        self.assertIn("const MAPS=", html_out)
        self.assertIn("<\\/script>", html_out)
        self.assertNotIn("</script><script>alert", html_out)

    def test_end_to_end_render(self):
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "render_output"
            result = flow.render(SAMPLE_DOCX, out_dir, VENDOR_PATH, strict=False)
            self.assertEqual(10, len(result.nodes))
            self.assertEqual(11, len(result.links))

            html_file = out_dir / "index.html"
            data_file = out_dir / "data.json"
            qa_file = out_dir / "qa_report.md"

            self.assertTrue(html_file.is_file() and html_file.stat().st_size > 0)
            self.assertTrue(data_file.is_file() and data_file.stat().st_size > 0)
            self.assertTrue(qa_file.is_file() and qa_file.stat().st_size > 0)

            data = json.loads(data_file.read_text(encoding="utf-8"))
            self.assertEqual("1.0", data["schema_version"])
            self.assertEqual(flow.VERSION, data["generator_version"])
            self.assertEqual(10, len(data["nodes"]))
            self.assertEqual(11, len(data["links"]))
            self.assertTrue(all(not k.startswith("_") for n in data["nodes"] for k in n))
            self.assertTrue(all(not k.startswith("_") for l in data["links"] for k in l))

            qa_content = qa_file.read_text(encoding="utf-8")
            self.assertIn("# Flow diagram QA report", qa_content)
            self.assertIn("Errors: 0", qa_content)

    def test_canonical_sample_fidelity(self):
        self.assertTrue(SAMPLE_DOCX.is_file(), f"Sample docx not found at {SAMPLE_DOCX}")
        result = flow.read_docx(SAMPLE_DOCX)
        self.assertEqual(10, len(result.nodes))
        self.assertEqual(11, len(result.links))
        errors = [i for i in result.issues if i.level == "error"]
        warnings = [i for i in result.issues if i.level == "warning"]
        self.assertEqual([], errors)
        self.assertEqual([], warnings)
        self.assertEqual("LR", result.settings.get("orientation"))
        self.assertEqual("Online Course Development Workflow", result.settings.get("title"))

    def test_strict_mode_rejection(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad_path = Path(tmp) / "bad.docx"
            out_dir = Path(tmp) / "out"
            make_doc(bad_path, [["n1", "Node 1", "Process"]], [["n1", "missing_target", "Link"]])

            # strict=True must reject validation errors
            with self.assertRaises(ValueError) as ctx:
                flow.render(bad_path, out_dir, VENDOR_PATH, strict=True)
            self.assertIn("validation error", str(ctx.exception))

            # strict=False must proceed and drop invalid link
            result = flow.render(bad_path, out_dir, VENDOR_PATH, strict=False)
            self.assertEqual(1, len(result.nodes))
            self.assertEqual(0, len(result.links))
            self.assertTrue(any(i.level == "error" for i in result.issues))

    def test_python_syntax_and_fstring_guard(self):
        src_path = ROOT_DIR / "render_flow_diagram.py"
        source = src_path.read_text(encoding="utf-8")

        # Must compile cleanly
        compile(source, str(src_path), "exec")

        # Static guard: No f-string formatted value expression may contain backslashes
        tree = ast.parse(source, filename=str(src_path))
        for node in ast.walk(tree):
            if isinstance(node, ast.JoinedStr):
                for part in node.values:
                    if isinstance(part, ast.FormattedValue):
                        segment = ast.unparse(part.value)
                        self.assertNotIn(
                            "\\", segment,
                            f"Disallowed backslash in f-string expression at line {part.lineno}: {segment}"
                        )

    def test_orientation_aliases_and_theme_fallbacks(self):
        settings = [
            ["orientation", "top-to-bottom"],
            ["theme", "custom_theme"],
            ["allow_orientation_switching", "false"],
            ["show_legend", "no"],
        ]
        result = self.parse(
            [["n1", "Node 1", "Process"]],
            [],
            settings=settings,
        )
        self.assertEqual("TB", result.settings["orientation"])
        self.assertEqual("light", result.settings["theme"])
        self.assertTrue(any("Unknown theme" in i.message for i in result.issues))

        html_out = flow.build_html(result, "/* js */")
        self.assertIn('id="orientation"', html_out)
        self.assertIn("hidden", html_out)

    def test_link_styles_and_colors(self):
        nodes = [["a", "A", "Process"], ["b", "B", "Process"], ["c", "C", "Process"]]
        links = [
            ["a", "b", "Dashed Link", "dashed", "#15803D", "Desc 1"],
            ["b", "c", "Thick Link", "thick", "#DC2626", "Desc 2"],
            ["c", "a", "Unknown Style Link", "wobbly", "", "Desc 3"],
        ]
        result = self.parse(
            nodes, links,
            link_headers=["Source", "Target", "Label", "Style", "Colour", "Description"],
        )
        self.assertEqual(3, len(result.links))
        self.assertEqual("solid", result.links[2]["style"])
        self.assertTrue(any("Unknown link style" in i.message for i in result.issues))

        diagram, _ = flow.diagram_text(result, "LR")
        self.assertIn("-. Dashed Link .->", diagram)
        self.assertIn("== Thick Link ==>", diagram)
        self.assertIn("-- Unknown Style Link -->", diagram)
        self.assertIn("stroke:#15803D", diagram)
        self.assertIn("stroke:#DC2626", diagram)

    def test_missing_or_invalid_required_tables(self):
        with tempfile.TemporaryDirectory() as tmp:
            empty_doc = Path(tmp) / "empty.docx"
            Document().save(empty_doc)

            result = flow.read_docx(empty_doc)
            self.assertTrue(any("No NODES table was found" in i.message for i in result.issues))
            self.assertTrue(any("No LINKS table was found" in i.message for i in result.issues))

            out_dir = Path(tmp) / "empty_out"
            with self.assertRaises(ValueError) as ctx:
                flow.render(empty_doc, out_dir, VENDOR_PATH, strict=False)
            self.assertIn("No valid nodes are available", str(ctx.exception))

            # Table with invalid headers missing ID
            bad_headers_doc = Path(tmp) / "no_id.docx"
            doc = Document()
            doc.add_heading("NODES", 1)
            tbl = doc.add_table(rows=2, cols=2)
            tbl.rows[0].cells[0].text = "WrongCol1"
            tbl.rows[0].cells[1].text = "WrongCol2"
            tbl.rows[1].cells[0].text = "Val1"
            tbl.rows[1].cells[1].text = "Val2"
            doc.save(bad_headers_doc)

            result_bad = flow.read_docx(bad_headers_doc)
            self.assertTrue(any("NODES requires ID and Label" in i.message for i in result_bad.issues))


if __name__ == "__main__":
    unittest.main()
