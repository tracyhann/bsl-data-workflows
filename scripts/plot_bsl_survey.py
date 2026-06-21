#!/usr/bin/env python3
from __future__ import annotations

import csv
import os
import sys
import textwrap
from collections import Counter
from pathlib import Path

_CACHE_ROOT = Path.cwd() / ".plot_cache"
os.environ.setdefault("MPLCONFIGDIR", str(_CACHE_ROOT / "matplotlib"))
os.environ.setdefault("XDG_CACHE_HOME", str(_CACHE_ROOT))

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns


ROLE_COL = "What is your primary role in the lab?"
TENURE_COL = "How long have you been in the Brain Stimulation Lab?"
MODALITIES_COL = "What modalities did you use in the last 4 weeks?"
ANALYSIS_TIME_COL = (
    "Approximately what percentage of your time in the last 4 weeks involved "
    "data analysis-related tasks?"
)
TIME_TO_FIND_COL = "Typical time to locate a needed file/dataset:"

DOMAIN_COLUMNS = [
    "Discoverability & Access",
    "Organization & Versioning",
    "Metadata & Interpretability",
    "Documentation & Training",
    "QC & Data Quality",
    "Overall",
]
TENURE_ORDER = ["6 months to a year", "1 to 3 years", "3 to 5 years", "More than 5 years"]
ANALYSIS_TIME_ORDER = ["0–10%", "11–25%", "26–50%", "51–75%", "76–100%"]
TIME_TO_FIND_ORDER = ["< 20 min", "11-20 min", "21 min - 1 hr", "> 1 hr but within a day", "days"]
ROLE_NORMALIZATION = {
    "Research Associate": "Clinical Researcher",
}


def wrap_label(text: str, width: int = 42) -> str:
    return "\n".join(textwrap.wrap(text, width=width, break_long_words=False))


def respondent_count(counter: Counter[str], order: list[str] | None = None) -> pd.DataFrame:
    if order is None:
        items = counter.most_common()
    else:
        items = [(key, counter[key]) for key in order if key in counter]
        items.extend((key, counter[key]) for key in counter if key not in {item[0] for item in items})
    return pd.DataFrame(items, columns=["label", "count"])


def save_figure(fig: plt.Figure, path: Path) -> None:
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)


def style_plotting() -> None:
    sns.set_theme(style="whitegrid", context="talk")
    plt.rcParams.update(
        {
            "figure.facecolor": "#f7f4ed",
            "axes.facecolor": "#fffdf8",
            "axes.edgecolor": "#d8cfbf",
            "grid.color": "#e5dccd",
            "axes.labelcolor": "#3b342b",
            "xtick.color": "#4b4339",
            "ytick.color": "#4b4339",
            "text.color": "#2d281f",
            "axes.titleweight": "bold",
            "font.size": 12,
        }
    )


def read_raw_survey(csv_path: Path) -> list[dict[str, str]]:
    with csv_path.open(newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def normalize_role(value: str | None) -> str:
    cleaned = str(value or "").strip()
    return ROLE_NORMALIZATION.get(cleaned, cleaned)


def plot_domain_scores(question_summary: pd.DataFrame, out_dir: Path) -> None:
    domain_scores = (
        question_summary.groupby("domain", as_index=False)["mean_score_1_to_5_higher_is_better"]
        .mean()
        .sort_values("mean_score_1_to_5_higher_is_better", ascending=True)
    )
    fig, ax = plt.subplots(figsize=(10, 6))
    palette = sns.color_palette("YlGnBu", n_colors=len(domain_scores))
    ax.barh(
        domain_scores["domain"],
        domain_scores["mean_score_1_to_5_higher_is_better"],
        color=palette,
        edgecolor="#6d6251",
    )
    ax.set_xlim(0, 5)
    ax.set_xlabel("Mean score (1-5, higher is healthier)")
    ax.set_ylabel("")
    ax.set_title("Overall Data Workflow Health by Domain")
    for idx, value in enumerate(domain_scores["mean_score_1_to_5_higher_is_better"]):
        ax.text(value + 0.06, idx, f"{value:.2f}", va="center", fontsize=11)
    save_figure(fig, out_dir / "01_domain_scores_overall.png")


def plot_item_strengths(question_summary: pd.DataFrame, out_dir: Path) -> None:
    ordered = question_summary.sort_values("mean_score_1_to_5_higher_is_better")
    focus = pd.concat([ordered.head(5), ordered.tail(5)]).copy()
    focus["wrapped_question"] = focus["question"].map(lambda text: wrap_label(text, width=44))
    focus["segment"] = ["Lowest friction score"] * 5 + ["Highest friction score"] * 5
    fig, ax = plt.subplots(figsize=(12, 8))
    colors = ["#cb6d51"] * 5 + ["#3f7f75"] * 5
    ax.barh(
        focus["wrapped_question"],
        focus["mean_score_1_to_5_higher_is_better"],
        color=colors,
        edgecolor="#5f5648",
    )
    ax.set_xlim(0, 5)
    ax.set_xlabel("Mean score (1-5)")
    ax.set_ylabel("")
    ax.set_title("Strongest and Weakest Survey Items")
    for idx, (_, row) in enumerate(focus.iterrows()):
        ax.text(
            row["mean_score_1_to_5_higher_is_better"] + 0.05,
            idx,
            f"{row['mean_score_1_to_5_higher_is_better']:.2f}",
            va="center",
            fontsize=10,
        )
    ax.axhline(4.5, color="#bfb5a5", linewidth=2)
    ax.text(4.95, 2.1, "Lowest-scoring items", ha="right", va="center", fontsize=11, color="#7a2f1d")
    ax.text(4.95, 7.1, "Highest-scoring items", ha="right", va="center", fontsize=11, color="#1e5e56")
    save_figure(fig, out_dir / "02_key_item_strengths_and_gaps.png")


def _heatmap_labels(df: pd.DataFrame, label_col: str) -> list[str]:
    return [f"{row[label_col]}\n(n={int(row['n'])})" for _, row in df.iterrows()]


def plot_role_heatmap(role_summary: pd.DataFrame, out_dir: Path) -> None:
    role_summary = role_summary.sort_values("Overall", ascending=False).reset_index(drop=True)
    display = role_summary.copy()
    display.index = _heatmap_labels(display, "group")
    fig, ax = plt.subplots(figsize=(11, 6.8))
    sns.heatmap(
        display[DOMAIN_COLUMNS],
        annot=True,
        fmt=".2f",
        cmap="RdYlGn",
        vmin=1,
        vmax=5,
        linewidths=0.6,
        linecolor="#f7f4ed",
        cbar_kws={"label": "Mean score"},
        ax=ax,
    )
    ax.set_title("Role Differences Across Survey Domains")
    ax.set_xlabel("")
    ax.set_ylabel("")
    save_figure(fig, out_dir / "03_role_domain_heatmap.png")


def plot_tenure_heatmap(tenure_summary: pd.DataFrame, out_dir: Path) -> None:
    tenure_summary = tenure_summary.copy()
    tenure_summary["group"] = pd.Categorical(tenure_summary["group"], categories=TENURE_ORDER, ordered=True)
    tenure_summary = tenure_summary.sort_values("group").reset_index(drop=True)
    display = tenure_summary.copy()
    display.index = _heatmap_labels(display, "group")
    fig, ax = plt.subplots(figsize=(11, 5.5))
    sns.heatmap(
        display[DOMAIN_COLUMNS],
        annot=True,
        fmt=".2f",
        cmap="RdYlGn",
        vmin=1,
        vmax=5,
        linewidths=0.6,
        linecolor="#f7f4ed",
        cbar_kws={"label": "Mean score"},
        ax=ax,
    )
    ax.set_title("Tenure Differences Across Survey Domains")
    ax.set_xlabel("")
    ax.set_ylabel("")
    save_figure(fig, out_dir / "04_tenure_domain_heatmap.png")


def plot_friction_distributions(raw_rows: list[dict[str, str]], out_dir: Path) -> None:
    time_counter = Counter(row[TIME_TO_FIND_COL].strip() for row in raw_rows if row.get(TIME_TO_FIND_COL, "").strip())
    analysis_counter = Counter(
        row[ANALYSIS_TIME_COL].strip() for row in raw_rows if row.get(ANALYSIS_TIME_COL, "").strip()
    )
    time_df = respondent_count(time_counter, TIME_TO_FIND_ORDER)
    analysis_df = respondent_count(analysis_counter, ANALYSIS_TIME_ORDER)

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    axes[0].barh(
        time_df["label"],
        time_df["count"],
        color=sns.color_palette("crest", n_colors=len(time_df)),
        edgecolor="#5a5246",
    )
    axes[0].set_title("How Long It Takes to Find a Needed File")
    axes[0].set_xlabel("Respondents")
    axes[0].set_ylabel("")
    for idx, value in enumerate(time_df["count"]):
        axes[0].text(value + 0.08, idx, str(value), va="center", fontsize=11)

    axes[1].bar(
        analysis_df["label"],
        analysis_df["count"],
        color=sns.color_palette("flare", n_colors=len(analysis_df)),
        edgecolor="#5a5246",
    )
    axes[1].set_title("Analysis Time Share in the Last 4 Weeks")
    axes[1].set_xlabel("")
    axes[1].set_ylabel("Respondents")
    axes[1].tick_params(axis="x", rotation=25)
    for idx, value in enumerate(analysis_df["count"]):
        axes[1].text(idx, value + 0.1, str(value), ha="center", va="bottom", fontsize=11)

    save_figure(fig, out_dir / "05_workflow_friction_distributions.png")


def plot_roadblock_themes(free_responses: pd.DataFrame, out_dir: Path) -> None:
    theme_counter: Counter[str] = Counter()
    blank_count = 0
    for themes in free_responses["themes"].fillna(""):
        parts = [part.strip() for part in themes.split(";") if part.strip()]
        if not parts:
            blank_count += 1
            continue
        theme_counter.update(parts)
    if blank_count:
        theme_counter["No substantive roadblock"] = blank_count

    theme_df = pd.DataFrame(theme_counter.most_common(), columns=["theme", "count"])
    fig, ax = plt.subplots(figsize=(11, 6))
    colors = ["#788b5a" if theme != "No substantive roadblock" else "#b5afa1" for theme in theme_df["theme"]]
    ax.barh(theme_df["theme"], theme_df["count"], color=colors, edgecolor="#645a4d")
    ax.invert_yaxis()
    ax.set_title("Roadblock Themes Mentioned in Free Responses")
    ax.set_xlabel("Mentions")
    ax.set_ylabel("")
    for idx, value in enumerate(theme_df["count"]):
        ax.text(value + 0.06, idx, str(value), va="center", fontsize=11)
    save_figure(fig, out_dir / "06_roadblock_themes.png")


def expand_theme_mentions(free_responses: pd.DataFrame, group_col: str) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for _, response in free_responses.iterrows():
        group_value = str(response.get(group_col, "") or "").strip()
        raw_themes = str(response.get("themes", "") or "").strip()
        if raw_themes.lower() == "nan":
            raw_themes = ""
        themes = [part.strip() for part in raw_themes.split(";") if part.strip()]
        if not themes:
            themes = ["No substantive roadblock"]
        for theme in themes:
            rows.append({"group": group_value, "theme": theme})
    return pd.DataFrame(rows)


def plot_roadblock_themes_by_group(
    free_responses: pd.DataFrame,
    group_col: str,
    title: str,
    out_path: Path,
    group_order: list[str] | None = None,
) -> None:
    theme_mentions = expand_theme_mentions(free_responses, group_col)
    heatmap_df = (
        theme_mentions.groupby(["group", "theme"])
        .size()
        .reset_index(name="count")
        .pivot(index="group", columns="theme", values="count")
        .fillna(0)
    )

    column_totals = heatmap_df.sum(axis=0).sort_values(ascending=False)
    ordered_columns = [column for column in column_totals.index if column != "No substantive roadblock"]
    if "No substantive roadblock" in heatmap_df.columns:
        ordered_columns.append("No substantive roadblock")
    heatmap_df = heatmap_df[ordered_columns]

    if group_order is not None:
        ordered_rows = [group for group in group_order if group in heatmap_df.index]
        heatmap_df = heatmap_df.reindex(ordered_rows)
    else:
        heatmap_df = heatmap_df.loc[heatmap_df.sum(axis=1).sort_values(ascending=False).index]

    heatmap_df = heatmap_df.astype(int)
    fig_width = max(14, 1.9 * len(heatmap_df.columns) + 4)
    fig_height = max(5.2, 1.0 * len(heatmap_df.index) + 2.4)
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    sns.heatmap(
        heatmap_df,
        annot=True,
        fmt="d",
        cmap="YlOrBr",
        linewidths=0.6,
        linecolor="#f7f4ed",
        cbar_kws={"label": "Theme mentions"},
        ax=ax,
    )
    ax.set_title(title)
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_xticklabels(
        [wrap_label(str(label), width=16) for label in heatmap_df.columns],
        rotation=28,
        ha="right",
        rotation_mode="anchor",
    )
    ax.set_yticklabels(ax.get_yticklabels(), rotation=0)
    save_figure(fig, out_path)


def plot_modality_usage(raw_rows: list[dict[str, str]], out_dir: Path) -> None:
    counter: Counter[str] = Counter()
    for row in raw_rows:
        entries = [item.strip() for item in row.get(MODALITIES_COL, "").splitlines() if item.strip()]
        counter.update(entries)
    modality_df = pd.DataFrame(counter.most_common(10), columns=["modality", "count"])
    modality_df["wrapped_modality"] = modality_df["modality"].map(lambda text: wrap_label(text, width=28))

    fig, ax = plt.subplots(figsize=(12, 7))
    ax.barh(
        modality_df["wrapped_modality"],
        modality_df["count"],
        color=sns.color_palette("mako", n_colors=len(modality_df)),
        edgecolor="#5a5246",
    )
    ax.invert_yaxis()
    ax.set_title("Most Common Data Modalities Used Recently")
    ax.set_xlabel("Respondents")
    ax.set_ylabel("")
    for idx, value in enumerate(modality_df["count"]):
        ax.text(value + 0.06, idx, str(value), va="center", fontsize=11)
    save_figure(fig, out_dir / "07_modality_usage.png")


def _time_to_find_group_table(
    raw_rows: list[dict[str, str]],
    group_col: str,
    group_order: list[str] | None = None,
) -> pd.DataFrame:
    rows = []
    for row in raw_rows:
        time_bucket = str(row.get(TIME_TO_FIND_COL, "") or "").strip()
        if not time_bucket:
            continue
        group_value = str(row.get(group_col, "") or "").strip()
        if group_col == ROLE_COL:
            group_value = normalize_role(group_value)
        rows.append({"group": group_value, "time_bucket": time_bucket})

    df = pd.DataFrame(rows)
    pivot = (
        df.groupby(["group", "time_bucket"])
        .size()
        .reset_index(name="count")
        .pivot(index="group", columns="time_bucket", values="count")
        .fillna(0)
    )
    ordered_cols = [bucket for bucket in TIME_TO_FIND_ORDER if bucket in pivot.columns]
    pivot = pivot[ordered_cols]

    if group_order is not None:
        ordered_rows = [group for group in group_order if group in pivot.index]
        pivot = pivot.reindex(ordered_rows)
    else:
        pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]

    return pivot.astype(int)


def plot_time_to_find_by_group(raw_rows: list[dict[str, str]], out_dir: Path) -> None:
    role_order = ["Clinical Researcher", "Data Analyst", "Faculty", "Regulatory", "Postdoc", "Scholar"]
    role_table = _time_to_find_group_table(raw_rows, ROLE_COL, group_order=role_order)
    tenure_table = _time_to_find_group_table(raw_rows, TENURE_COL, group_order=TENURE_ORDER)

    fig, axes = plt.subplots(2, 1, figsize=(13, 9))
    sns.heatmap(
        role_table,
        annot=True,
        fmt="d",
        cmap="PuBuGn",
        linewidths=0.6,
        linecolor="#f7f4ed",
        cbar_kws={"label": "Respondents"},
        ax=axes[0],
    )
    axes[0].set_title("Typical File-Finding Time by Role")
    axes[0].set_xlabel("")
    axes[0].set_ylabel("")
    axes[0].set_xticklabels([wrap_label(label, width=14) for label in role_table.columns], rotation=0)
    axes[0].set_yticklabels(axes[0].get_yticklabels(), rotation=0)

    sns.heatmap(
        tenure_table,
        annot=True,
        fmt="d",
        cmap="YlGnBu",
        linewidths=0.6,
        linecolor="#f7f4ed",
        cbar_kws={"label": "Respondents"},
        ax=axes[1],
    )
    axes[1].set_title("Typical File-Finding Time by Tenure")
    axes[1].set_xlabel("")
    axes[1].set_ylabel("")
    axes[1].set_xticklabels([wrap_label(label, width=14) for label in tenure_table.columns], rotation=0)
    axes[1].set_yticklabels(axes[1].get_yticklabels(), rotation=0)

    fig.suptitle("How Long It Takes to Find a Needed File Across Groups", fontsize=18, fontweight="bold")
    save_figure(fig, out_dir / "10_time_to_find_by_role_and_tenure.png")


def plot_summary_dashboard(
    question_summary: pd.DataFrame,
    role_summary: pd.DataFrame,
    tenure_summary: pd.DataFrame,
    free_responses: pd.DataFrame,
    raw_rows: list[dict[str, str]],
    out_dir: Path,
) -> None:
    domain_scores = (
        question_summary.groupby("domain", as_index=False)["mean_score_1_to_5_higher_is_better"]
        .mean()
        .sort_values("mean_score_1_to_5_higher_is_better", ascending=True)
    )
    role_sorted = role_summary.sort_values("Overall", ascending=False)
    tenure_sorted = tenure_summary.copy()
    tenure_sorted["group"] = pd.Categorical(tenure_sorted["group"], categories=TENURE_ORDER, ordered=True)
    tenure_sorted = tenure_sorted.sort_values("group")

    theme_counter: Counter[str] = Counter()
    for themes in free_responses["themes"].fillna(""):
        parts = [part.strip() for part in themes.split(";") if part.strip()]
        theme_counter.update(parts)
    theme_df = pd.DataFrame(theme_counter.most_common(5), columns=["theme", "count"])

    time_counter = Counter(row[TIME_TO_FIND_COL].strip() for row in raw_rows if row.get(TIME_TO_FIND_COL, "").strip())
    time_df = respondent_count(time_counter, TIME_TO_FIND_ORDER)

    fig, axes = plt.subplots(2, 2, figsize=(15, 11))

    axes[0, 0].barh(
        domain_scores["domain"],
        domain_scores["mean_score_1_to_5_higher_is_better"],
        color=sns.color_palette("YlGnBu", n_colors=len(domain_scores)),
        edgecolor="#6d6251",
    )
    axes[0, 0].set_xlim(0, 5)
    axes[0, 0].set_title("Domain scores")
    axes[0, 0].set_xlabel("Mean score")
    axes[0, 0].set_ylabel("")

    axes[0, 1].bar(
        role_sorted["group"],
        role_sorted["Overall"],
        color="#6c9a8b",
        edgecolor="#5d5548",
    )
    axes[0, 1].set_ylim(0, 5)
    axes[0, 1].set_title("Overall score by role")
    axes[0, 1].set_ylabel("Mean score")
    axes[0, 1].tick_params(axis="x", rotation=35)

    axes[1, 0].plot(
        tenure_sorted["group"].astype(str),
        tenure_sorted["Overall"],
        marker="o",
        color="#bf6f50",
        linewidth=2.5,
    )
    axes[1, 0].set_ylim(0, 5)
    axes[1, 0].set_title("Overall score by tenure")
    axes[1, 0].set_ylabel("Mean score")
    axes[1, 0].tick_params(axis="x", rotation=20)

    axes[1, 1].barh(theme_df["theme"], theme_df["count"], color="#8a9d5f", edgecolor="#5d5548")
    axes[1, 1].invert_yaxis()
    axes[1, 1].set_title("Top roadblock themes")
    axes[1, 1].set_xlabel("Mentions")

    fig.suptitle("BSL Data Workflow Survey Visual Summary", fontsize=20, fontweight="bold")
    save_figure(fig, out_dir / "00_survey_visual_summary.png")


def main() -> int:
    if len(sys.argv) < 2:
        print(
            "Usage: plot_bsl_survey.py <survey.csv> [analysis_dir] [plot_dir]",
            file=sys.stderr,
        )
        return 1

    csv_path = Path(sys.argv[1]).resolve()
    analysis_dir = (
        Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else csv_path.parent / "reports" / "survey_analysis"
    )
    plot_dir = Path(sys.argv[3]).resolve() if len(sys.argv) > 3 else csv_path.parent / "survey_plots"
    plot_dir.mkdir(parents=True, exist_ok=True)

    style_plotting()

    question_summary = pd.read_csv(analysis_dir / "bsl_survey_question_summary.csv")
    role_summary = pd.read_csv(analysis_dir / "bsl_survey_by_role.csv")
    tenure_summary = pd.read_csv(analysis_dir / "bsl_survey_by_tenure.csv")
    free_responses = pd.read_csv(analysis_dir / "bsl_survey_free_responses.csv")
    raw_rows = read_raw_survey(csv_path)

    numeric_cols = ["mean_score_1_to_5_higher_is_better"]
    question_summary[numeric_cols] = question_summary[numeric_cols].apply(pd.to_numeric)
    for df in (role_summary, tenure_summary):
        for col in DOMAIN_COLUMNS:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    plot_summary_dashboard(question_summary, role_summary, tenure_summary, free_responses, raw_rows, plot_dir)
    plot_domain_scores(question_summary, plot_dir)
    plot_item_strengths(question_summary, plot_dir)
    plot_role_heatmap(role_summary, plot_dir)
    plot_tenure_heatmap(tenure_summary, plot_dir)
    plot_friction_distributions(raw_rows, plot_dir)
    plot_roadblock_themes(free_responses, plot_dir)
    plot_roadblock_themes_by_group(
        free_responses,
        "role",
        "Roadblock Themes by Role",
        plot_dir / "08_roadblock_themes_by_role.png",
    )
    plot_roadblock_themes_by_group(
        free_responses,
        "tenure",
        "Roadblock Themes by Tenure",
        plot_dir / "09_roadblock_themes_by_tenure.png",
        group_order=TENURE_ORDER,
    )
    plot_time_to_find_by_group(raw_rows, plot_dir)
    plot_modality_usage(raw_rows, plot_dir)

    print(f"Wrote plots to {plot_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
