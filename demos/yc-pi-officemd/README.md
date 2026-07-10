# YC demo: Pi with and without the OfficeMD skill

This demo runs the same Pi model on the same [GDPval](https://huggingface.co/datasets/openai/gdpval) spreadsheet task in two isolated directories:

- `baseline`: Pi with skill, extension, prompt-template, and context-file discovery disabled.
- `skill`: the same command plus the repository's explicit `.pi/skills/officemd` skill.

The task is GDPval `76418a2c-a3c0-4894-b89d-2493369135d9`. Pi must read three XLSX files, preserve a blank shipment-manifest template, derive three shipments, write formulas, and visually verify a native workbook. Each Pi process runs in an external temporary directory containing only the three task inputs. The GDPval reference deliverable is never stored in the repository or copied into an agent work directory.

## Run

Prerequisites:

- `pi` authenticated with the selected provider (`pi`, then `/login` if needed)
- `uv`, `cargo`, `curl`, and `jq`
- LibreOffice and Poppler for XLSX-to-PNG evidence

Prepare fixtures and validate the command without invoking a model:

```bash
./demos/yc-pi-officemd/run-demo.sh --dry-run
```

Calibrate the objective scorer against an empty run and the held-out GDPval reference:

```bash
./demos/yc-pi-officemd/calibrate-scorer.sh
```

The calibration downloads the reference into a temporary directory and deletes it on exit.

Run the controlled comparison:

```bash
./demos/yc-pi-officemd/run-demo.sh \
  --provider openai-codex \
  --model gpt-5.5 \
  --thinking high
```

The command prints the path to `runs/<timestamp>/report.html`. Each variant also retains:

- `command.txt`: exact shell command
- `config.txt` and `run.json`: model and execution metadata
- `events.jsonl` and `stderr.log`: raw Pi output
- `work/`: isolated inputs and the produced workbook
- `score.json`: the 100-point objective rubric
- `evidence/`: OfficeMD inspect, verify, and render reports plus PNGs

Override `--run-id` to make paths stable for a recorded demo.

## Two-minute talk track

1. **0:00 - 0:20, establish control.** Show `PROMPT.md`, then the two `command.txt` files. Call out that model, thinking level, prompt, files, built-in tools, and runtime flags are identical. The only difference is `--skill .pi/skills/officemd`.
2. **0:20 - 0:45, show the baseline.** Open its rendered workbook and score. Focus on concrete failure modes: missing native output, changed source files, wrong weight band, hard-coded savings, lost currency formatting, or no visual verification. Do not claim a failure that the measured run did not show.
3. **0:45 - 1:25, show the skilled run.** Open `events.jsonl` and point to the lifecycle: probe, bounded inspect, source hash, one typed patch plan, verification, then PNG evidence. Open the output render and its score.
4. **1:25 - 1:45, explain the product.** OfficeMD is not another model. It gives an existing agent stable locators, typed edits, atomic outputs, provenance, renderer discovery, formula checks, and semantic/visual evidence.
5. **1:45 - 2:00, show the boundary.** A GDP.pdf image-only scan is correctly reported as `pdf_may_require_ocr`. OfficeMD renders it but does not pretend that empty extraction is understanding. OCR and vision remain explicit capabilities.

## Claims discipline

- Use the scores from the current run, not a rehearsed number.
- A passed `formula-references` check means no stored formula errors were found. It does not prove that every spreadsheet engine recalculates identically.
- A passed visual-render check means PNG evidence was produced. The agent or presenter must still inspect it for clipping and layout defects.
- The benchmark fixtures remain subject to the source dataset's terms. This repository stores only the fetch URLs and checksums.
