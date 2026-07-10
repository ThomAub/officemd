# YC Presentation Walkthrough

This is a two-minute live demo of the same Pi agent operating with and without the OfficeMD skill.

## 0:00 - Problem

Open `PROMPT.md` and the three GDPval inputs.

Say:

> Agents can write spreadsheet files, but they usually lack reliable artifact controls. This task requires Pi to understand three workbooks, preserve a template, calculate shipping costs, write formulas, and return a valid XLSX.

Show:

- `Pick Tickets 062525.xlsx`
- `Shipping parameters.xlsx`
- `Blank Daily Shipment Manifest.xlsx`

## 0:20 - Controlled comparison

Open the two generated `command.txt` files in the run directory.

Say:

> Both runs use the same Pi model, thinking level, prompt, files, built-in tools, and runtime flags. The only difference is the explicit OfficeMD skill.

Point to:

```text
baseline: --no-skills
skill:    --no-skills --skill .pi/skills/officemd
```

The process runs both variants in separate temporary directories containing only the task inputs.

## 0:40 - Baseline

Open the baseline score and rendered workbook from `report.html`.

Say:

> This is what Pi produced without artifact-specific guidance.

Discuss only measured failures. The scorer checks for missing native output, incorrect weights or methods, hard-coded savings, lost formatting, changed source files, and absent verification. Do not claim a failure that the current run did not produce.

## 1:00 - Skilled workflow

Open the skilled `events.jsonl` and show the lifecycle:

```text
probe -> bounded inspect -> render inputs -> typed patch
     -> apply to new output -> inspect -> verify -> render evidence
```

Say:

> The skill does not contain the benchmark answers. It teaches Pi how to operate on Office artifacts safely.

Highlight the skill rules:

- Stable sheet and cell locators
- Source SHA-256 precondition
- Typed cell and formula operations
- No in-place mutation
- Atomic output and provenance sidecar
- Structural, formula, and visual verification

## 1:25 - Result

Open `report.html` and show the two score columns.

Say:

> The score is computed from the workbook itself, not from the agent's explanation.

The 100-point rubric checks:

- Exact ticket set
- Customers and weights
- Shipping methods
- Actual and industry costs
- Formula-based savings
- Currency formatting
- Unchanged source hashes
- OfficeMD verification
- Rendered PNG evidence

Open the skilled workbook PNG and point out that it remains a native, formatted manifest.

## 1:45 - Product message

Say:

> OfficeMD is not another model. It is an artifact-control layer for existing agents: stable locators, typed mutations, provenance, rendering, verification, and semantic or visual diffs.

The model stays the same. The artifact workflow becomes dependable.

## 1:55 - Honest boundary

Finish with:

> OfficeMD also exposes what it cannot prove. An image-only GDP.pdf document is flagged as requiring OCR. It can render the pages, but it does not pretend that empty extraction means understanding.

This demonstrates a product boundary rather than hiding it: OCR and vision are explicit capabilities that can be added above the native artifact layer.

## Before presenting

Authenticate Pi first:

```bash
pi
/login
```

Then run:

```bash
./demos/yc-pi-officemd/run-demo.sh \
  --provider openai \
  --model gpt-5.5 \
  --thinking high
```

Use the scores from the current run. The repository includes `calibrate-scorer.sh` to validate the scorer against an empty workbook and the held-out GDPval reference without leaving the reference file in the repository.
