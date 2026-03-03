# Steel Cycle Anomaly Detection

This project analyzes EAF cycle Excel files and generates anomaly plots for:

- `Energie / FEM (4)`
- `(Energie / FEM (4)) / Tap to Tap`

It supports:

- single-file processing
- folder processing (recursive), where all `.xlsx` files are merged and treated as one dataset

## Requirements

- Python 3.11+
- Packages:
  - `pandas`
  - `matplotlib`
  - `openpyxl`

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install pandas matplotlib openpyxl
```

## Usage

### Process a folder (recommended)

```bash
python script.py --folder 2024
python script.py --folder 2025
python script.py --folder 2026
```

### Process default files

If `--folder` is not provided, the script processes:

- `1.xlsx`
- `2.xlsx`

```bash
python script.py
```

### Optional anomaly threshold factor

Default threshold is:

`mean + 1.0 * std`

You can change it:

```bash
python script.py --folder 2025 --factor 1.5
```

## Output Files

For each processed folder, the script writes:

- `*_fem4_anomalies.png`
- `*_fem4_counts.png`
- `*_fem4_per_tap_to_tap_anomalies.png`
- `*_fem4_per_tap_to_tap_counts.png`

## Notes

- Rows with invalid denominators (`<= 0`) are excluded per metric.
- Count plot x-axis is forced to start at `0` to avoid visual negative values.
