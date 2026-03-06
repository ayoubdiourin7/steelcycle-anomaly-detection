import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

ENERGY_COLUMN_CANDIDATES = ["Energie Elec.", "Energie Elec. EAF"]
METRICS = {
    "energie_tappe_power_on": {
        "denominator_candidates": [["Poids Tappé", "Poids Tapp", "Poids Tappe"]],
        "multiplier_candidates": [["Power On"]],
        "energy_multiplier": 1.0,
        "ratio_label": "Energie_Tappe_PowerOn = (Energie Elec. / Poids Tappé) * Power On",
        "y_label": "Energie_Tappe * Power On",
    },
    "energie_tappe": {
        "denominator_candidates": [["Poids Tappé", "Poids Tapp", "Poids Tappe"]],
        "energy_multiplier": 1.0,
        "ratio_label": "Energie_Tappe = Energie Elec. / Poids Tappé",
        "y_label": "Energie_Tappe",
    },
}


def _normalize_column_name(name: str) -> str:
    return " ".join(str(name).replace("\xa0", " ").strip().lower().split())


def _resolve_column_name(columns: pd.Index, candidates: list[str]) -> str | None:
    normalized_lookup = {_normalize_column_name(col): str(col) for col in columns}
    for candidate in candidates:
        resolved = normalized_lookup.get(_normalize_column_name(candidate))
        if resolved is not None:
            return resolved
    return None


def _read_excel(path: Path) -> pd.DataFrame | None:
    try:
        df = pd.read_excel(path, sheet_name="Sheet1", header=1).copy()
    except Exception as exc:
        print(f"Skipping {path}: {exc}")
        return None

    if _resolve_column_name(df.columns, ENERGY_COLUMN_CANDIDATES) is None:
        print(f"Skipping {path}: missing energy column in {ENERGY_COLUMN_CANDIDATES}")
        return None
    return df


def _prepare_metric_dataframe(
    df: pd.DataFrame,
    energy_column: str,
    denominator_columns: list[str],
    multiplier_columns: list[str] | None = None,
    energy_multiplier: float = 1.0,
) -> tuple[pd.DataFrame, int]:
    multiplier_columns = multiplier_columns or []
    df_work = df.copy()
    if "Cycle" not in df_work.columns:
        df_work["Cycle"] = range(1, len(df_work) + 1)

    df_work[energy_column] = pd.to_numeric(df_work[energy_column], errors="coerce")
    for denominator_column in denominator_columns:
        df_work[denominator_column] = pd.to_numeric(df_work[denominator_column], errors="coerce")
    for multiplier_column in multiplier_columns:
        df_work[multiplier_column] = pd.to_numeric(df_work[multiplier_column], errors="coerce")

    valid_denominator_mask = pd.Series(True, index=df_work.index)
    for denominator_column in denominator_columns:
        valid_denominator_mask &= df_work[denominator_column] > 0
    for multiplier_column in multiplier_columns:
        valid_denominator_mask &= df_work[multiplier_column] > 0
    ignored_rows = int((~valid_denominator_mask).sum())
    df_metric = df_work.loc[valid_denominator_mask].copy()
    if df_metric.empty:
        return df_metric, ignored_rows

    ratio = df_metric[energy_column] * energy_multiplier
    for denominator_column in denominator_columns:
        ratio = ratio / df_metric[denominator_column]
    for multiplier_column in multiplier_columns:
        ratio = ratio * df_metric[multiplier_column]
    df_metric["Ratio"] = ratio
    df_metric = df_metric[df_metric["Ratio"].notna()].copy()
    return df_metric, ignored_rows


def _save_metric_plots(
    df_metric: pd.DataFrame,
    output_dir: Path,
    output_prefix: str,
    title_label: str,
    metric_name: str,
    ratio_label: str,
    y_label: str,
    anomaly_factor: float,
) -> None:
    ratio_values = df_metric["Ratio"]
    q1 = ratio_values.quantile(0.25)
    q3 = ratio_values.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 if pd.isna(iqr) else q1 - anomaly_factor * iqr
    upper_bound = q3 if pd.isna(iqr) else q3 + anomaly_factor * iqr
    df_metric["Anomalie"] = (df_metric["Ratio"] < lower_bound) | (df_metric["Ratio"] > upper_bound)

    anomaly_output_file = output_dir / f"{output_prefix}_{metric_name}_anomalies.png"
    ratio_count_output_file = output_dir / f"{output_prefix}_{metric_name}_counts.png"

    # Plot 1: ratio by cycle with IQR bounds.
    y_top = max(float(ratio_values.max()) * 1.05, float(upper_bound) * 1.05)
    plt.figure(figsize=(12, 6))
    plt.plot(df_metric["Cycle"], df_metric["Ratio"], label=ratio_label, color="#1d4ed8", linewidth=1.7)
    plt.scatter(
        df_metric.loc[~df_metric["Anomalie"], "Cycle"],
        df_metric.loc[~df_metric["Anomalie"], "Ratio"],
        color="#64748b",
        s=14,
        alpha=0.55,
        label="Cycle normal",
    )
    plt.scatter(
        df_metric.loc[df_metric["Anomalie"], "Cycle"],
        df_metric.loc[df_metric["Anomalie"], "Ratio"],
        color="#dc2626",
        edgecolors="black",
        linewidths=0.4,
        s=52,
        label="Anomalie",
    )
    plt.axhline(lower_bound, color="#dc2626", linestyle="--", linewidth=1.6, label="Seuil Bas")
    plt.axhline(upper_bound, color="#dc2626", linestyle="--", linewidth=1.8, label="Seuil Haut")
    plt.text(
        float(df_metric["Cycle"].min()),
        float(upper_bound) * 1.01,
        f"Bas={lower_bound:.3f} | Haut={upper_bound:.3f}",
        color="#9a3412",
        fontsize=10,
        va="bottom",
    )
    plt.title(f"{title_label}: {ratio_label}")
    plt.xlabel("Cycle")
    plt.ylabel(y_label)
    plt.grid(alpha=0.25)
    plt.legend()
    plt.tight_layout()
    plt.savefig(anomaly_output_file, dpi=300, bbox_inches="tight")
    print(f"Plot saved to {anomaly_output_file}")
    plt.close()

    # Plot 2: count plot with anomaly bins in red.
    ratio_counts = ratio_values.round(2).value_counts().sort_index()
    bar_width = 0.008
    if len(ratio_counts) > 1:
        step = (float(ratio_counts.index.max()) - float(ratio_counts.index.min())) / (len(ratio_counts) - 1)
        bar_width = max(step * 0.8, 0.002)

    bar_colors = ["#dc2626" if (value < lower_bound or value > upper_bound) else "#0f766e" for value in ratio_counts.index]
    x_left = 0.0

    plt.figure(figsize=(12, 6))
    plt.bar(ratio_counts.index, ratio_counts.values, width=bar_width, color=bar_colors, edgecolor="white", linewidth=0.3)
    plt.axvline(lower_bound, color="#dc2626", linestyle="--", linewidth=1.8, label="Seuil Bas")
    plt.axvline(upper_bound, color="#dc2626", linestyle="--", linewidth=2, label="Seuil Haut")
    plt.text(
        upper_bound + (bar_width * 0.8),
        float(ratio_counts.max()) * 0.95,
        f"Anomalie si ratio < {lower_bound:.3f} ou > {upper_bound:.3f}",
        color="#9a3412",
        fontsize=10,
        ha="left",
        va="top",
    )
    plt.title(f"{title_label}: Distribution {ratio_label}")
    plt.xlabel(f"{ratio_label} (arrondi a 0.01)")
    plt.ylabel("Nombre de cycles")
    plt.xlim(left=x_left)
    plt.grid(axis="y", alpha=0.25)
    plt.legend()
    plt.tight_layout()
    plt.savefig(ratio_count_output_file, dpi=300, bbox_inches="tight")
    print(f"Plot saved to {ratio_count_output_file}")
    plt.close()


def _process_dataframe(
    df: pd.DataFrame,
    output_dir: Path,
    output_prefix: str,
    title_label: str,
    anomaly_factor: float,
) -> None:
    df_base = df.copy()
    df_base["Cycle"] = range(1, len(df_base) + 1)
    energy_column = _resolve_column_name(df_base.columns, ENERGY_COLUMN_CANDIDATES)
    if energy_column is None:
        print(f"No energy column found in {ENERGY_COLUMN_CANDIDATES}, skipping.")
        return
    print(f"Using energy column: {energy_column}")

    for metric_name, metric in METRICS.items():
        resolved_denominators: list[str] = []
        missing_candidates: list[list[str]] = []
        for candidates in metric["denominator_candidates"]:
            denominator = _resolve_column_name(df_base.columns, candidates)
            if denominator is None:
                missing_candidates.append(candidates)
            else:
                resolved_denominators.append(denominator)
        resolved_multipliers: list[str] = []
        missing_multiplier_candidates: list[list[str]] = []
        for candidates in metric.get("multiplier_candidates", []):
            multiplier = _resolve_column_name(df_base.columns, candidates)
            if multiplier is None:
                missing_multiplier_candidates.append(candidates)
            else:
                resolved_multipliers.append(multiplier)

        if missing_candidates or missing_multiplier_candidates:
            missing_text: list[str] = []
            if missing_candidates:
                missing_text.append(f"denominator column in {missing_candidates}")
            if missing_multiplier_candidates:
                missing_text.append(f"multiplier column in {missing_multiplier_candidates}")
            print(f"Skipping metric '{metric_name}': missing {' and '.join(missing_text)}")
            continue

        ratio_label = metric["ratio_label"]
        y_label = metric["y_label"]
        df_metric, ignored_rows = _prepare_metric_dataframe(
            df_base,
            energy_column,
            resolved_denominators,
            multiplier_columns=resolved_multipliers,
            energy_multiplier=float(metric.get("energy_multiplier", 1.0)),
        )
        valid_columns_text = " and ".join(resolved_denominators + resolved_multipliers)
        print(f"[{metric_name}] Ignored rows with invalid metric columns ({valid_columns_text} <= 0): {ignored_rows}")
        if df_metric.empty:
            print(f"[{metric_name}] No valid rows left after filtering, skipping plots.")
            continue

        _save_metric_plots(
            df_metric=df_metric,
            output_dir=output_dir,
            output_prefix=output_prefix,
            title_label=title_label,
            metric_name=metric_name,
            ratio_label=ratio_label,
            y_label=y_label,
            anomaly_factor=anomaly_factor,
        )


def process_file(input_file: str, anomaly_factor: float = 1.5) -> None:
    path = Path(input_file)
    if not path.exists():
        print(f"File not found: {input_file}")
        return

    print(f"\nProcessing file: {path}")
    df = _read_excel(path)
    if df is None:
        return

    _process_dataframe(
        df=df,
        output_dir=path.parent,
        output_prefix=path.stem,
        title_label=path.stem,
        anomaly_factor=anomaly_factor,
    )


def process_folder(input_folder: str, anomaly_factor: float = 1.5) -> None:
    folder = Path(input_folder)
    if not folder.exists() or not folder.is_dir():
        print(f"Invalid folder: {input_folder}")
        return

    xlsx_files = sorted(folder.rglob("*.xlsx"))
    if not xlsx_files:
        print(f"No .xlsx files found in {folder}")
        return

    print(f"\nProcessing folder: {folder}")
    print(f"Found {len(xlsx_files)} .xlsx files")

    frames: list[pd.DataFrame] = []
    skipped_files = 0
    for xlsx_file in xlsx_files:
        df = _read_excel(xlsx_file)
        if df is None:
            skipped_files += 1
            continue
        frames.append(df)

    if not frames:
        print("No readable .xlsx files to process.")
        return

    folder_name = folder.name if folder.name else folder.resolve().name
    combined_df = pd.concat(frames, ignore_index=True)
    print(f"Rows loaded from all files: {len(combined_df)}")
    if skipped_files:
        print(f"Skipped files: {skipped_files}")

    _process_dataframe(
        df=combined_df,
        output_dir=folder,
        output_prefix=f"{folder_name}_all_xlsx",
        title_label=f"{folder_name} (all xlsx)",
        anomaly_factor=anomaly_factor,
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Plot Energy ratios and anomalies from Excel files.")
    parser.add_argument(
        "--file",
        type=str,
        help="Input .xlsx file to process.",
    )
    parser.add_argument(
        "--folder",
        type=str,
        help="Input folder containing .xlsx files (recursive). All files are merged and treated as one dataset.",
    )
    parser.add_argument(
        "--factor",
        type=float,
        default=1.5,
        help="Outlier factor for IQR bounds: lower = Q1 - factor * IQR, upper = Q3 + factor * IQR (default: 1.5).",
    )
    args = parser.parse_args()

    if args.file and args.folder:
        parser.error("Use only one input: --file or --folder.")
    if args.folder:
        process_folder(args.folder, anomaly_factor=args.factor)
    elif args.file:
        process_file(args.file, anomaly_factor=args.factor)
    else:
        parser.error("Missing input. Provide --folder <path> or --file <path>.")
