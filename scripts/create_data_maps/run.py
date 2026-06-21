#!/usr/bin/env python3
"""Create lightweight data-map workbooks for a cleaned study folder."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook


if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[2]))

from scripts.create_study_overview.run import (  # noqa: E402
    DEFAULT_SEMANTICS,
    autosize_columns,
    instrument_key_from_workbook,
    load_dictionary_labels,
    load_semantic_examples,
    match_description,
    overview_label_for_workbook,
)


PLATFORMS_COLUMNS = ["stage", "privacy", "description", "location"]
PLATFORM_PLACEHOLDER_ROWS = 10
CLASS_MAP_COLUMNS = ["stage", "description", "location"]
RAW_PLACEHOLDER_ROW = ["raw", None, None]


@dataclass(frozen=True)
class DataMapsResult:
    study_folder: Path
    output_dir: Path
    platforms_path: Path
    class_paths: list[Path]
    map_count: int
    item_count: int


def cleaned_dir_for_study(study_folder: Path) -> Path:
    return study_folder / "data" / "cleaned"


def class_dirs(cleaned_dir: Path) -> list[Path]:
    if not cleaned_dir.exists():
        return []
    return sorted(path for path in cleaned_dir.iterdir() if path.is_dir() and not path.name.startswith("."))


def cleaned_items(class_dir: Path) -> list[Path]:
    return sorted(path for path in class_dir.iterdir() if path.is_file() and not path.name.startswith("~$"))


def relative_location(study_folder: Path, path: Path) -> str:
    return f"./{path.relative_to(study_folder).as_posix()}"


def write_platforms(output_dir: Path) -> Path:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "platforms"
    worksheet.append(PLATFORMS_COLUMNS)
    for _ in range(PLATFORM_PLACEHOLDER_ROWS):
        worksheet.append(["", "", "", ""])
    autosize_columns(worksheet)
    output_path = output_dir / "platforms-data-map.xlsx"
    workbook.save(output_path)
    return output_path


def description_for_item(
    path: Path,
    data_class: str,
    dictionary_labels: dict[str, str],
    semantic_examples: dict[str, list[str]],
) -> str:
    instrument_key = instrument_key_from_workbook(path)
    data_label = overview_label_for_workbook(path, dictionary_labels)
    if data_label:
        return match_description(data_class, data_label, instrument_key, semantic_examples) or data_label
    return path.stem


def write_class_map(
    study_folder: Path,
    output_dir: Path,
    class_dir: Path,
    dictionary_labels: dict[str, str],
    semantic_examples: dict[str, list[str]],
) -> tuple[Path, int]:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "data_map"
    worksheet.append(CLASS_MAP_COLUMNS)
    worksheet.append(RAW_PLACEHOLDER_ROW)

    item_count = 0
    for item in cleaned_items(class_dir):
        worksheet.append(
            [
                "cleaned/processed",
                description_for_item(item, class_dir.name, dictionary_labels, semantic_examples),
                relative_location(study_folder, item),
            ]
        )
        item_count += 1

    autosize_columns(worksheet)
    output_path = output_dir / f"{class_dir.name}-data-map.xlsx"
    workbook.save(output_path)
    return output_path, item_count


def create_data_maps(
    study_folder: str | Path,
    output_dir: str | Path | None = None,
    semantics_path: str | Path = DEFAULT_SEMANTICS,
) -> DataMapsResult:
    study_folder = Path(study_folder)
    cleaned_dir = cleaned_dir_for_study(study_folder)
    output_dir = Path(output_dir) if output_dir is not None else study_folder / "data-map"
    output_dir.mkdir(parents=True, exist_ok=True)

    dictionary_labels = load_dictionary_labels(study_folder)
    semantic_examples = load_semantic_examples(semantics_path)

    platforms_path = write_platforms(output_dir)
    class_paths: list[Path] = []
    item_count = 0
    for class_dir in class_dirs(cleaned_dir):
        class_path, class_item_count = write_class_map(
            study_folder,
            output_dir,
            class_dir,
            dictionary_labels,
            semantic_examples,
        )
        class_paths.append(class_path)
        item_count += class_item_count

    return DataMapsResult(
        study_folder=study_folder,
        output_dir=output_dir,
        platforms_path=platforms_path,
        class_paths=class_paths,
        map_count=len(class_paths),
        item_count=item_count,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--study-folder", required=True, type=Path, help="Cleaned study folder.")
    parser.add_argument("--out-dir", type=Path, help="Optional data-map output directory.")
    parser.add_argument(
        "--semantics",
        type=Path,
        default=DEFAULT_SEMANTICS,
        help="Instrument semantics JSON used to fill description values.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    result = create_data_maps(args.study_folder, output_dir=args.out_dir, semantics_path=args.semantics)
    print(result.output_dir)
    print(f"maps={result.map_count}")
    print(f"items={result.item_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
