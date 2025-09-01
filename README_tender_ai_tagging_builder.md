# Tender AI Tagging Builder

This tool builds **AI-tagging Excel workbooks** from tender PDFs using your **Lakhwar-style** template.

## Features
- Clone your template’s sheet structure
- Parse **main RfS/ITB PDFs** and **corrigendum/addendum PDFs**
- Extract clause-level parameters using **regex rules** in YAML
- Populate `AI_Tagging_Master` and `BID_INFO`
- Update `AmendmentTracker` with basic metadata
- Add build/version info in `TenderMeta`

## Requirements
- Python 3.9+
- `pip install pandas PyYAML PyPDF2 openpyxl xlsxwriter`

## Quick Start
```bash
python tender_ai_tagging_builder.py \
  --template "8.1 Lakhwar_Master_v8_MLready_v2.xlsx" \
  --pdf "SECI000196-...500MWhfinal.pdf" \
  --amendment "Corrigendum1.pdf" --amendment "AddendumA.pdf" \
  --rules "rules_example.yaml" \
  --out "SECI_BESS_AI_Tagging.xlsx"
```

## Extending to New Tender Types
- Copy `rules_example.yaml` to a new file, e.g. `rules_ntpc_boiler.yaml`
- Add/modify `extractors` patterns for the new tender
- Adjust `bid_info_map` keys and regex
- Run the script with `--rules rules_ntpc_boiler.yaml`

### YAML Tips
- Use `mode: all` to capture multiples (e.g., multi-year tables).
- Add `flags: ["IGNORECASE","DOTALL","MULTILINE"]` for complex paragraphs.
- In `value_expr`, you can use `{0}`, `{1}` for capture groups, or named groups like `(?P<cap_mw>\d+)` and `{cap_mw}`.

## Sheets Produced
- `AI_Tagging_Master` – rows with: Section, Clause/Ref, Parameter, Value, Unit, Notes, SourcePage
- `BID_INFO` – field/value summary of headline items
- `AmendmentTracker` – each corrigendum/addendum file with type/date (heuristic), filename
- `TenderMeta` – version string and counts

## Notes
- PDF text extraction is best-effort; complex tables may need rule tweaking.
- For annexure drawings or SLD IDs, add specialized patterns or parse those PDFs separately.
- You can merge/normalize parameters into `Parameters_Master` with a post-step if needed.
