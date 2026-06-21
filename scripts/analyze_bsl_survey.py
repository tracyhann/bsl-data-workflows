#!/usr/bin/env python3
from __future__ import annotations

import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from statistics import mean


ROLE_COL = "What is your primary role in the lab?"
TENURE_COL = "How long have you been in the Brain Stimulation Lab?"
MODALITIES_COL = "What modalities did you use in the last 4 weeks?"
ANALYSIS_TIME_COL = (
    "Approximately what percentage of your time in the last 4 weeks involved "
    "data analysis-related tasks?"
)
TIME_TO_FIND_COL = "Typical time to locate a needed file/dataset:"
ROADBLOCKS_COL = (
    "What roadblocks do you often experience when accessing, analyzing, and/or "
    "performing any other data operations? (feel free to vent)"
)

ROLE_NORMALIZATION = {
    "Research Associate": "Clinical Researcher",
}


DISCOVERABILITY_ITEMS = [
    "I can reliably locate the correct dataset/file I need without trial-and-error.",
    "I know the single best location to look first for each data type I use.",
    "I can retrieve needed data fast enough to keep work moving (finding is not a bottleneck).",
    "I need to ask another person where something lives.",
    "I discover that the “right file” exists but I cannot access it due to permissions.",
    TIME_TO_FIND_COL,
]

ORGANIZATION_ITEMS = [
    "Folder structure and naming conventions are consistent enough that I can predict where things should be.",
    'For the datasets I use, there is a clear "source of truth" (not multiple competing versions).',
    "Versioning is handled well (I can tell what is current, what is archived, and what changed).",
    "Data is stored in a way that supports cross-study reuse (not trapped inside one person's organization).",
    "Archiving practices prevent loss and make older data retrievable.",
    "Data is organized in a way that is intuitive and supports downstream analysis (not just storage).",
    "I create a local copy/personal folder because the shared structure is not practical or hard to use.",
]

METADATA_ITEMS = [
    "Files/datasets consistently include the minimum metadata needed to interpret them later (IDs, dates, modality, version).",
    "Subject identifiers are consistent across systems (REDCap imaging physiology files).",
    "Timepoints/visit labels are consistent across systems.",
    "Metadata are sufficiently standardized that merging datasets could be automated (at least partially).",
    "I can understand a dataset without needing the original creator to explain it.",
    'I trust that "what the variable means" is documented (data dictionary / README / codebook exists and is findable).',
    "During analysis, how often do you have to stop to clarify what a variable/file meant (missing/unclear metadata)?",
]

DOCUMENTATION_ITEMS = [
    "SOPs and how-to docs exist for the workflows I use.",
    "Documentation is easy to find when I need it.",
    "Documentation is specific enough to execute without improvising or asking the team/study lead.",
    "Documentation is up to date (matches current practice).",
    "For certain tasks, new staff could follow documentation and perform key data tasks without heavy 1:1 training.",
    "How often do you have to troubleshoot due to missing/unclear documentation?",
]

QUALITY_ITEMS = [
    "QC/validation steps are clearly defined for the data types I use.",
    "QC is performed early enough to prevent downstream rework.",
    "QC outcomes are documented.",
    "When corrections occur, there is a traceable record of what changed and why.",
    "I trust the accuracy of the data I retrieve from lab systems.",
    "Data issues are handled in a consistent way rather than depending on the individual.",
    "In the last 4 weeks, how often did you encounter missing/corrupted/incomplete data?",
    "In the last 4 weeks, how often did QC catch issues only at the analysis stage (late discovery)?",
    "How often do you spend time fixing data quality issues (not counting planned QC - actual work)?",
]

DOMAIN_ITEMS = {
    "Discoverability & Access": DISCOVERABILITY_ITEMS,
    "Organization & Versioning": ORGANIZATION_ITEMS,
    "Metadata & Interpretability": METADATA_ITEMS,
    "Documentation & Training": DOCUMENTATION_ITEMS,
    "QC & Data Quality": QUALITY_ITEMS,
}

NEGATIVE_AGREEMENT_ITEMS = {
    "I need to ask another person where something lives.",
    "I discover that the “right file” exists but I cannot access it due to permissions.",
}

NEGATIVE_FREQUENCY_ITEMS = {
    "I create a local copy/personal folder because the shared structure is not practical or hard to use.",
    "During analysis, how often do you have to stop to clarify what a variable/file meant (missing/unclear metadata)?",
    "How often do you have to troubleshoot due to missing/unclear documentation?",
    "In the last 4 weeks, how often did you encounter missing/corrupted/incomplete data?",
    "In the last 4 weeks, how often did QC catch issues only at the analysis stage (late discovery)?",
    "How often do you spend time fixing data quality issues (not counting planned QC - actual work)?",
}

ALL_SCORABLE_ITEMS = {
    item for items in DOMAIN_ITEMS.values() for item in items
}

LIKERT_ITEMS = (
    set(DISCOVERABILITY_ITEMS[:-1])
    | set(ORGANIZATION_ITEMS[:-1])
    | set(METADATA_ITEMS[:-1])
    | set(DOCUMENTATION_ITEMS[:-1])
    | set(QUALITY_ITEMS[:-3])
)

FREQUENCY_ITEMS = NEGATIVE_FREQUENCY_ITEMS.copy()

TIME_TO_FIND_SCORES = {
    "< 20 min": 5,
    "11-20 min": 4,
    "21 min - 1 hr": 3,
    "> 1 hr but within a day": 2,
    "days": 1,
}

TIME_TO_FIND_ORDER = [
    "< 20 min",
    "11-20 min",
    "21 min - 1 hr",
    "> 1 hr but within a day",
    "days",
]

ANALYSIS_TIME_ORDER = ["0–10%", "11–25%", "26–50%", "51–75%", "76–100%"]

FREE_RESPONSE_THEME_RULES = {
    "Findability / fragmented storage": [
        r"where data is located",
        r"place where all clinical assessments data",
        r"lot of data storage platforms",
        r"platform to platform",
        r"finding documentation",
        r"file locations",
        r"find lists of projects",
        r"where data is located",
        r"figuring out where data is located",
    ],
    "Documentation / metadata gaps": [
        r"headers mean",
        r"read\.?me",
        r"documentation",
        r"scanner specifications",
        r"regressor",
        r"not recorded anywhere",
        r"lost knowledge",
    ],
    "Versioning / source-of-truth confusion": [
        r"most up-to-date",
        r"different versions",
        r"who has already worked with the data",
        r"what has been done",
    ],
    "Naming / data consistency": [
        r"file name",
        r"data number",
    ],
    "Legacy system / historical gaps": [
        r"cni",
        r"older studies",
        r"last 3 years",
    ],
    "Cross-dataset linkage / cohort logic": [
        r"inclusion/exclusion",
        r"open label retreatment",
        r"matching derived data across modalities",
    ],
}


def clean_text(value: str | None) -> str:
    return (value or "").strip()


def normalize_group_value(group_col: str, value: str | None) -> str:
    cleaned = clean_text(value)
    if group_col == ROLE_COL:
        return ROLE_NORMALIZATION.get(cleaned, cleaned)
    return cleaned


def is_missing(value: str | None) -> bool:
    cleaned = clean_text(value)
    return not cleaned or cleaned.lower() in {"n/a", "na", "none"}


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip())


def parse_likert(value: str | None) -> int | None:
    if is_missing(value):
        return None
    cleaned = clean_text(value)
    if cleaned in {"1", "2", "3", "4", "5"}:
        return int(cleaned)
    match = re.match(r"^\s*([1-5])\s*[- ]", cleaned)
    if match:
        return int(match.group(1))
    return None


def parse_frequency(value: str | None) -> int | None:
    if is_missing(value):
        return None
    cleaned = normalize_whitespace(clean_text(value)).lower()
    mapping = {
        "never": 1,
        "rarely": 2,
        "sometimes": 3,
        "often": 4,
        "very often": 5,
    }
    if cleaned in mapping:
        return mapping[cleaned]
    match = re.match(r"^\s*([1-5])\s*[- ]", cleaned)
    if match:
        return int(match.group(1))
    return None


def parse_time_to_find(value: str | None) -> int | None:
    if is_missing(value):
        return None
    cleaned = normalize_whitespace(clean_text(value))
    return TIME_TO_FIND_SCORES.get(cleaned)


def question_score(question: str, value: str | None) -> float | None:
    if question == TIME_TO_FIND_COL:
        return parse_time_to_find(value)
    if question in LIKERT_ITEMS:
        parsed = parse_likert(value)
        if parsed is None:
            return None
        if question in NEGATIVE_AGREEMENT_ITEMS:
            return 6 - parsed
        return float(parsed)
    if question in FREQUENCY_ITEMS:
        parsed = parse_frequency(value)
        if parsed is None:
            return None
        if question in NEGATIVE_FREQUENCY_ITEMS:
            return float(6 - parsed)
        return float(parsed)
    return None


def favorable_bucket(score: float) -> str:
    if score >= 4:
        return "favorable"
    if score <= 2:
        return "unfavorable"
    return "neutral"


def split_multiline_list(value: str | None) -> list[str]:
    cleaned = clean_text(value)
    if not cleaned:
        return []
    items = [part.strip() for part in cleaned.splitlines()]
    return [item for item in items if item]


def domain_score(row: dict[str, str], questions: list[str]) -> float | None:
    scores = [question_score(question, row.get(question)) for question in questions]
    scores = [score for score in scores if score is not None]
    return mean(scores) if scores else None


def overall_score(row: dict[str, str]) -> float | None:
    scores = [domain_score(row, questions) for questions in DOMAIN_ITEMS.values()]
    scores = [score for score in scores if score is not None]
    return mean(scores) if scores else None


def percent(part: int, total: int) -> str:
    if total == 0:
        return "0.0%"
    return f"{(part / total) * 100:.1f}%"


def format_score(value: float | None) -> str:
    if value is None:
        return "NA"
    return f"{value:.2f}"


def summarize_distribution(counter: Counter[str], order: list[str] | None = None) -> str:
    total = sum(counter.values())
    if total == 0:
        return "No responses"
    items: list[tuple[str, int]]
    if order:
        ordered_keys = [key for key in order if key in counter]
        remainder = [key for key in counter if key not in ordered_keys]
        items = [(key, counter[key]) for key in ordered_keys + sorted(remainder)]
    else:
        items = counter.most_common()
    return "; ".join(f"{label}: {count} ({percent(count, total)})" for label, count in items)


def top_items(counter: Counter[str], limit: int = 8) -> str:
    total = sum(counter.values())
    if total == 0:
        return "No responses"
    return "; ".join(
        f"{label}: {count} ({percent(count, total)})" for label, count in counter.most_common(limit)
    )


def respondent_share(counter: Counter[str], denominator: int, limit: int = 8) -> str:
    if denominator == 0 or not counter:
        return "No responses"
    return "; ".join(
        f"{label}: {count} ({percent(count, denominator)})"
        for label, count in counter.most_common(limit)
    )


def question_domain(question: str) -> str:
    for domain, questions in DOMAIN_ITEMS.items():
        if question in questions:
            return domain
    return "Other"


def classify_response_themes(response: str) -> list[str]:
    cleaned = clean_text(response)
    if is_missing(cleaned):
        return []
    lowered = cleaned.lower()
    themes: list[str] = []
    for theme, patterns in FREE_RESPONSE_THEME_RULES.items():
        if any(re.search(pattern, lowered) for pattern in patterns):
            themes.append(theme)
    if not themes:
        themes.append("Other / uncategorized")
    return themes


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def build_question_summary(rows: list[dict[str, str]]) -> list[dict[str, object]]:
    summaries: list[dict[str, object]] = []
    for question in ALL_SCORABLE_ITEMS:
        scored = [question_score(question, row.get(question)) for row in rows]
        scored = [score for score in scored if score is not None]
        buckets = Counter(favorable_bucket(score) for score in scored)
        summaries.append(
            {
                "domain": question_domain(question),
                "question": question,
                "n": len(scored),
                "mean_score_1_to_5_higher_is_better": f"{mean(scored):.2f}" if scored else "",
                "favorable_count": buckets["favorable"],
                "favorable_pct": f"{(buckets['favorable'] / len(scored)) * 100:.1f}" if scored else "",
                "neutral_count": buckets["neutral"],
                "unfavorable_count": buckets["unfavorable"],
                "reverse_coded": "yes"
                if question in NEGATIVE_AGREEMENT_ITEMS or question in NEGATIVE_FREQUENCY_ITEMS
                else "no",
            }
        )
    return sorted(
        summaries,
        key=lambda item: (item["domain"], float(item["mean_score_1_to_5_higher_is_better"] or 0)),
    )


def build_group_summary(
    rows: list[dict[str, str]], group_col: str
) -> list[dict[str, object]]:
    groups: defaultdict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        groups[normalize_group_value(group_col, row.get(group_col)) or "Missing"].append(row)

    summaries: list[dict[str, object]] = []
    for group, group_rows in sorted(groups.items()):
        item: dict[str, object] = {
            "group_type": group_col,
            "group": group,
            "n": len(group_rows),
        }
        domain_values: list[float] = []
        for domain, questions in DOMAIN_ITEMS.items():
            row_scores = [domain_score(row, questions) for row in group_rows]
            row_scores = [score for score in row_scores if score is not None]
            group_domain_score = mean(row_scores) if row_scores else None
            item[domain] = f"{group_domain_score:.2f}" if group_domain_score is not None else ""
            if group_domain_score is not None:
                domain_values.append(group_domain_score)
        item["Overall"] = f"{mean(domain_values):.2f}" if domain_values else ""
        summaries.append(item)
    return summaries


def build_group_question_summary(
    rows: list[dict[str, str]], group_col: str
) -> list[dict[str, object]]:
    groups: defaultdict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        groups[normalize_group_value(group_col, row.get(group_col)) or "Missing"].append(row)

    summaries: list[dict[str, object]] = []
    for group, group_rows in sorted(groups.items()):
        for question in sorted(ALL_SCORABLE_ITEMS):
            scored = [question_score(question, row.get(question)) for row in group_rows]
            scored = [score for score in scored if score is not None]
            if not scored:
                continue
            buckets = Counter(favorable_bucket(score) for score in scored)
            summaries.append(
                {
                    "group_type": group_col,
                    "group": group,
                    "domain": question_domain(question),
                    "question": question,
                    "n": len(scored),
                    "mean_score_1_to_5_higher_is_better": f"{mean(scored):.2f}",
                    "favorable_pct": f"{(buckets['favorable'] / len(scored)) * 100:.1f}",
                    "reverse_coded": "yes"
                    if question in NEGATIVE_AGREEMENT_ITEMS or question in NEGATIVE_FREQUENCY_ITEMS
                    else "no",
                }
            )
    return summaries


def strongest_and_weakest_questions(
    question_summary: list[dict[str, object]], limit: int = 5
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    ordered = sorted(
        question_summary,
        key=lambda item: float(item["mean_score_1_to_5_higher_is_better"] or 0),
    )
    return ordered[-limit:][::-1], ordered[:limit]


def notable_group_gaps(group_rows: list[dict[str, object]], domain: str) -> str | None:
    eligible = []
    for row in group_rows:
        try:
            n = int(row["n"])
        except (KeyError, TypeError, ValueError):
            continue
        value = row.get(domain)
        if n >= 2 and value not in {"", None}:
            eligible.append((row["group"], float(value), n))
    if len(eligible) < 2:
        return None
    low_group, low_score, low_n = min(eligible, key=lambda item: item[1])
    high_group, high_score, high_n = max(eligible, key=lambda item: item[1])
    return (
        f"{domain}: {high_group} {high_score:.2f} (n={high_n}) vs "
        f"{low_group} {low_score:.2f} (n={low_n}), gap {high_score - low_score:.2f}"
    )


def render_markdown_report(
    csv_path: Path,
    out_dir: Path,
    rows: list[dict[str, str]],
    question_summary: list[dict[str, object]],
    role_summary: list[dict[str, object]],
    tenure_summary: list[dict[str, object]],
    free_response_rows: list[dict[str, object]],
) -> str:
    role_counts = Counter(normalize_group_value(ROLE_COL, row.get(ROLE_COL)) for row in rows)
    tenure_counts = Counter(clean_text(row.get(TENURE_COL)) for row in rows)
    analysis_time_counts = Counter(clean_text(row.get(ANALYSIS_TIME_COL)) for row in rows)

    modality_counts: Counter[str] = Counter()
    for row in rows:
        modality_counts.update(split_multiline_list(row.get(MODALITIES_COL)))

    domain_averages = []
    for domain, questions in DOMAIN_ITEMS.items():
        scores = [domain_score(row, questions) for row in rows]
        scores = [score for score in scores if score is not None]
        domain_averages.append((domain, mean(scores), len(scores)))
    domain_averages.sort(key=lambda item: item[1], reverse=True)

    strongest, weakest = strongest_and_weakest_questions(question_summary)

    roadblock_theme_counts: Counter[str] = Counter()
    for item in free_response_rows:
        for theme in item["themes"].split("; "):
            if theme:
                roadblock_theme_counts[theme] += 1

    quick_find = Counter(clean_text(row.get(TIME_TO_FIND_COL)) for row in rows if not is_missing(row.get(TIME_TO_FIND_COL)))

    role_gaps = [
        gap for domain in DOMAIN_ITEMS for gap in [notable_group_gaps(role_summary, domain)] if gap
    ]
    tenure_gaps = [
        gap for domain in DOMAIN_ITEMS for gap in [notable_group_gaps(tenure_summary, domain)] if gap
    ]

    report_lines = [
        "# BSL Data Workflow Survey Report",
        "",
        f"- Source file: `{csv_path.name}`",
        f"- Responses analyzed: **{len(rows)}**",
        f"- Output directory: `{out_dir}`",
        "- Scoring note: all scored items are normalized to a 1-5 health scale where higher is better; negatively worded and high-friction frequency items are reverse-coded.",
        "",
        "## Executive Summary",
        "",
        f"- The strongest domain was **{domain_averages[0][0]}** at **{domain_averages[0][1]:.2f}/5**; the weakest was **{domain_averages[-1][0]}** at **{domain_averages[-1][1]:.2f}/5**.",
        f"- File-finding speed was mixed: {summarize_distribution(quick_find, TIME_TO_FIND_ORDER)}.",
        f"- The most common analysis-time band was **{analysis_time_counts.most_common(1)[0][0]}** ({analysis_time_counts.most_common(1)[0][1]} respondents).",
        f"- The most common roadblock themes were: {respondent_share(roadblock_theme_counts, len(rows), limit=5)}.",
        "",
        "## Respondent Profile",
        "",
        "### Roles",
        "",
        f"- {top_items(role_counts)}",
        "",
        "### Tenure in Lab",
        "",
        f"- {summarize_distribution(tenure_counts, ['6 months to a year', '1 to 3 years', '3 to 5 years', 'More than 5 years'])}",
        "",
        "### Analysis Time",
        "",
        f"- {summarize_distribution(analysis_time_counts, ANALYSIS_TIME_ORDER)}",
        "",
        "### Modalities Used in Last 4 Weeks",
        "",
        f"- {respondent_share(modality_counts, len(rows), limit=12)}",
        "",
        "## Domain Scores",
        "",
        "| Domain | Mean Score | Respondents with Score |",
        "| --- | ---: | ---: |",
    ]

    for domain, score, n_scored in domain_averages:
        report_lines.append(f"| {domain} | {score:.2f} | {n_scored} |")

    report_lines.extend(
        [
            "",
            "## Highest-Scoring Items",
            "",
            "| Question | Domain | Mean Score | Favorable % | n |",
            "| --- | --- | ---: | ---: | ---: |",
        ]
    )
    for item in strongest:
        report_lines.append(
            "| {question} | {domain} | {score} | {favorable}% | {n} |".format(
                question=item["question"],
                domain=item["domain"],
                score=item["mean_score_1_to_5_higher_is_better"],
                favorable=item["favorable_pct"],
                n=item["n"],
            )
        )

    report_lines.extend(
        [
            "",
            "## Lowest-Scoring Items",
            "",
            "| Question | Domain | Mean Score | Favorable % | n |",
            "| --- | --- | ---: | ---: | ---: |",
        ]
    )
    for item in weakest:
        report_lines.append(
            "| {question} | {domain} | {score} | {favorable}% | {n} |".format(
                question=item["question"],
                domain=item["domain"],
                score=item["mean_score_1_to_5_higher_is_better"],
                favorable=item["favorable_pct"],
                n=item["n"],
            )
        )

    report_lines.extend(
        [
            "",
            "## Group Differences",
            "",
            "### By Role",
            "",
            "| Role | n | Discoverability & Access | Organization & Versioning | Metadata & Interpretability | Documentation & Training | QC & Data Quality | Overall |",
            "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
    )
    for row in role_summary:
        report_lines.append(
            "| {group} | {n} | {d1} | {d2} | {d3} | {d4} | {d5} | {overall} |".format(
                group=row["group"],
                n=row["n"],
                d1=row.get("Discoverability & Access", ""),
                d2=row.get("Organization & Versioning", ""),
                d3=row.get("Metadata & Interpretability", ""),
                d4=row.get("Documentation & Training", ""),
                d5=row.get("QC & Data Quality", ""),
                overall=row.get("Overall", ""),
            )
        )

    report_lines.extend(["", "Notable role gaps:", ""])
    for gap in role_gaps:
        report_lines.append(f"- {gap}")

    report_lines.extend(
        [
            "",
            "### By Tenure",
            "",
            "| Tenure | n | Discoverability & Access | Organization & Versioning | Metadata & Interpretability | Documentation & Training | QC & Data Quality | Overall |",
            "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
    )
    for row in tenure_summary:
        report_lines.append(
            "| {group} | {n} | {d1} | {d2} | {d3} | {d4} | {d5} | {overall} |".format(
                group=row["group"],
                n=row["n"],
                d1=row.get("Discoverability & Access", ""),
                d2=row.get("Organization & Versioning", ""),
                d3=row.get("Metadata & Interpretability", ""),
                d4=row.get("Documentation & Training", ""),
                d5=row.get("QC & Data Quality", ""),
                overall=row.get("Overall", ""),
            )
        )

    report_lines.extend(["", "Notable tenure gaps:", ""])
    for gap in tenure_gaps:
        report_lines.append(f"- {gap}")

    report_lines.extend(
        [
            "",
            "## Free-Response Analysis",
            "",
            "### Thematic Summary",
            "",
        ]
    )
    non_substantive_count = sum(1 for item in free_response_rows if not item["themes"])
    if non_substantive_count:
        report_lines.append(f"- **No substantive roadblock provided**: {non_substantive_count} responses")
    for theme, count in roadblock_theme_counts.most_common():
        report_lines.append(
            f"- **{theme}**: {count} responses ({percent(count, len(rows))} of all respondents)"
        )

    report_lines.extend(
        [
            "",
            "### Verbatim Responses",
            "",
            "| ID | Role | Tenure | Themes | Response |",
            "| --- | --- | --- | --- | --- |",
        ]
    )
    for item in free_response_rows:
        response = item["response"].replace("\n", " ").replace("|", "\\|")
        report_lines.append(
            f"| {item['id']} | {item['role']} | {item['tenure']} | {item['themes']} | {response} |"
        )

    report_lines.extend(
        [
            "",
            "## Question-Level Detail",
            "",
            "| Domain | Question | Mean Score | Favorable % | Reverse Coded | n |",
            "| --- | --- | ---: | ---: | --- | ---: |",
        ]
    )
    for item in sorted(
        question_summary,
        key=lambda row: (
            row["domain"],
            float(row["mean_score_1_to_5_higher_is_better"] or 0),
            row["question"],
        ),
    ):
        report_lines.append(
            "| {domain} | {question} | {score} | {favorable}% | {reverse} | {n} |".format(
                domain=item["domain"],
                question=item["question"],
                score=item["mean_score_1_to_5_higher_is_better"],
                favorable=item["favorable_pct"],
                reverse=item["reverse_coded"],
                n=item["n"],
            )
        )

    return "\n".join(report_lines) + "\n"


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: analyze_bsl_survey.py <survey.csv> [output_dir]", file=sys.stderr)
        return 1

    csv_path = Path(sys.argv[1]).resolve()
    out_dir = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else csv_path.parent / "reports" / "survey_analysis"
    out_dir.mkdir(parents=True, exist_ok=True)

    with csv_path.open(newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.DictReader(handle))

    question_summary = build_question_summary(rows)
    role_summary = build_group_summary(rows, ROLE_COL)
    tenure_summary = build_group_summary(rows, TENURE_COL)
    role_question_summary = build_group_question_summary(rows, ROLE_COL)
    tenure_question_summary = build_group_question_summary(rows, TENURE_COL)

    free_response_rows: list[dict[str, object]] = []
    for index, row in enumerate(rows, start=1):
        response = clean_text(row.get(ROADBLOCKS_COL))
        if not response:
            continue
        themes = classify_response_themes(response)
        free_response_rows.append(
            {
                "id": index,
                "role": normalize_group_value(ROLE_COL, row.get(ROLE_COL)),
                "tenure": clean_text(row.get(TENURE_COL)),
                "themes": "; ".join(themes),
                "response": response,
            }
        )

    report = render_markdown_report(
        csv_path=csv_path,
        out_dir=out_dir,
        rows=rows,
        question_summary=question_summary,
        role_summary=role_summary,
        tenure_summary=tenure_summary,
        free_response_rows=free_response_rows,
    )

    report_path = out_dir / "bsl_survey_report.md"
    report_path.write_text(report, encoding="utf-8")

    write_csv(
        out_dir / "bsl_survey_question_summary.csv",
        [
            "domain",
            "question",
            "n",
            "mean_score_1_to_5_higher_is_better",
            "favorable_count",
            "favorable_pct",
            "neutral_count",
            "unfavorable_count",
            "reverse_coded",
        ],
        question_summary,
    )

    group_fieldnames = [
        "group_type",
        "group",
        "n",
        "Discoverability & Access",
        "Organization & Versioning",
        "Metadata & Interpretability",
        "Documentation & Training",
        "QC & Data Quality",
        "Overall",
    ]
    write_csv(out_dir / "bsl_survey_by_role.csv", group_fieldnames, role_summary)
    write_csv(out_dir / "bsl_survey_by_tenure.csv", group_fieldnames, tenure_summary)
    write_csv(
        out_dir / "bsl_survey_questions_by_role.csv",
        [
            "group_type",
            "group",
            "domain",
            "question",
            "n",
            "mean_score_1_to_5_higher_is_better",
            "favorable_pct",
            "reverse_coded",
        ],
        role_question_summary,
    )
    write_csv(
        out_dir / "bsl_survey_questions_by_tenure.csv",
        [
            "group_type",
            "group",
            "domain",
            "question",
            "n",
            "mean_score_1_to_5_higher_is_better",
            "favorable_pct",
            "reverse_coded",
        ],
        tenure_question_summary,
    )
    write_csv(
        out_dir / "bsl_survey_free_responses.csv",
        ["id", "role", "tenure", "themes", "response"],
        free_response_rows,
    )

    print(f"Wrote {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
