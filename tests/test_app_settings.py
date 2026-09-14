from datetime import date

import pytest

from app import load_settings, normalize_excel_paths, save_settings, validate_run_parameters
from extractor import MODE_DEPARTMENT_TOP10, MODE_RANDOM


def test_date_range_is_remembered(tmp_path):
    path = tmp_path / "settings.json"
    save_settings(date(2024, 1, 2), date(2026, 12, 30), "C:/input", "D:/output", path)
    assert load_settings(path) == {
        "start_date": "2024-01-02",
        "end_date": "2026-12-30",
        "input_dir": "C:/input",
        "output_dir": "D:/output",
        "extraction_mode": MODE_RANDOM,
    }


def test_invalid_settings_are_ignored(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text('{"start_date":"bad"}', encoding="utf-8")
    assert load_settings(path) == {}


def test_invalid_extraction_mode_settings_are_ignored(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text('{"start_date":"2025-01-01","end_date":"2025-12-31","extraction_mode":"bad"}', encoding="utf-8")
    assert load_settings(path) == {}


def test_reversed_date_settings_are_ignored(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text('{"start_date":"2026-01-01","end_date":"2025-01-01"}', encoding="utf-8")
    assert load_settings(path) == {}


def test_excel_paths_are_filtered_and_deduplicated(tmp_path):
    first = tmp_path / "sample.xlsx"
    result = normalize_excel_paths([first, str(first), tmp_path / "notes.txt"])
    assert result == [str(first.resolve())]


def test_invalid_run_parameters_are_rejected_before_saving():
    with pytest.raises(ValueError, match="开始日期"):
        validate_run_parameters(date(2026, 1, 1), date(2025, 1, 1), 5, "out")
    with pytest.raises(ValueError, match="抽取条数"):
        validate_run_parameters(date(2025, 1, 1), date(2026, 1, 1), 0, "out")
    with pytest.raises(ValueError, match="输出目录"):
        validate_run_parameters(date(2025, 1, 1), date(2026, 1, 1), 5, "  ")


def test_department_mode_does_not_require_random_count():
    validate_run_parameters(date(2025, 1, 1), date(2026, 1, 1), 0, "out", MODE_DEPARTMENT_TOP10)
