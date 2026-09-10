from datetime import date

import pytest

from app import load_settings, save_settings, validate_run_parameters


def test_date_range_is_remembered(tmp_path):
    path = tmp_path / "settings.json"
    save_settings(date(2024, 1, 2), date(2026, 12, 30), path)
    assert load_settings(path) == {"start_date": "2024-01-02", "end_date": "2026-12-30"}


def test_invalid_settings_are_ignored(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text('{"start_date":"bad"}', encoding="utf-8")
    assert load_settings(path) == {}


def test_invalid_run_parameters_are_rejected_before_saving():
    with pytest.raises(ValueError, match="开始日期"):
        validate_run_parameters(date(2026, 1, 1), date(2025, 1, 1), 5, "out")
    with pytest.raises(ValueError, match="抽取条数"):
        validate_run_parameters(date(2025, 1, 1), date(2026, 1, 1), 0, "out")
    with pytest.raises(ValueError, match="输出目录"):
        validate_run_parameters(date(2025, 1, 1), date(2026, 1, 1), 5, "  ")
