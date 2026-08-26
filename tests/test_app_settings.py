from datetime import date

from app import load_settings, save_settings


def test_date_range_is_remembered(tmp_path):
    path = tmp_path / "settings.json"
    save_settings(date(2024, 1, 2), date(2026, 12, 30), path)
    assert load_settings(path) == {"start_date": "2024-01-02", "end_date": "2026-12-30"}


def test_invalid_settings_are_ignored(tmp_path):
    path = tmp_path / "settings.json"
    path.write_text('{"start_date":"bad"}', encoding="utf-8")
    assert load_settings(path) == {}
