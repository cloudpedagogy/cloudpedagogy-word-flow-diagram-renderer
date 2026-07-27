# CloudPedagogy Word Flow Diagram Renderer

Convert editable Word tables into a standalone interactive process or decision-flow diagram. Edit the supplied example and run one Python command; no Mermaid or JavaScript editing is required.

## Files and demonstration

- [Editable Word example](examples/flow_diagram_example.docx)
- [Renderer script](render_flow_diagram.py)
- [Generated HTML example](output/flow_diagram_example/index.html)
- [Normalised example data](output/flow_diagram_example/data.json)
- [Example QA report](output/flow_diagram_example/qa_report.md)
- [Automated tests](tests/test_render_flow_diagram.py)

After enabling GitHub Pages, the live demonstration will be:

https://cloudpedagogy.github.io/cloudpedagogy-word-flow-diagram-renderer/output/flow_diagram_example/

## Quick start

```bash
git clone https://github.com/cloudpedagogy/cloudpedagogy-word-flow-diagram-renderer.git
cd cloudpedagogy-word-flow-diagram-renderer

python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt

python3 render_flow_diagram.py examples/flow_diagram_example.docx \
  --output output/flow_diagram_example --overwrite
```

Windows PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py render_flow_diagram.py examples/flow_diagram_example.docx --output output/flow_diagram_example --overwrite
```

Open `output/flow_diagram_example/index.html`.

## Create your own diagram

Copy [the Word example](examples/flow_diagram_example.docx) and edit these tables:

- `SETTINGS` — title, subtitle, orientation, theme and interface options
- `NODES` — required ID and Label; optional Type, Description, Status, Colour, Link and Group
- `LINKS` — required Source and Target; optional Label, Style, Colour and Description

Common heading aliases such as `Key`, `Name`, `Kind`, `From`, `To`, `Condition`, `Phase`, and British or US spelling of colour are accepted. Supported node types include start, end, process, decision, input, output, data, document, database and delay. Link styles include solid, dashed, dotted and thick.

## Customisation and limits

The generated page supports four orientations, search, zoom, keyboard selection, node details, safe external links, legends and an accessible tabular fallback. Mermaid is embedded for offline use.

The input is flexible within the documented schema, but every source and target must match a unique node ID. It does not infer a flow from an arbitrary Word document.

## Output and validation

- `index.html` — interactive offline diagram
- `data.json` — parsed and normalised data
- `qa_report.md` — findings with table and row references

```bash
python3 render_flow_diagram.py --help
python3 render_flow_diagram.py INPUT.docx --strict
python3 -m unittest discover -s tests -v
```

Without `--strict`, invalid rows are excluded where possible and reported. With it, validation errors stop generation.

## Licence

MIT. See [LICENSE](LICENSE).
