# Example Data

Redistributable fixtures for OfficeMD examples and smoke tests.

## Included Files

- `showcase.docx`
- `showcase.xlsx`
- `showcase.csv`
- `showcase.pptx`
- `OpenXML_WhitePaper.pdf`
- OCR classification fixtures for PDF inspection examples

## Regenerate

```bash
uv run --with python-docx --with openpyxl --with python-pptx python examples/generate_data.py
```
