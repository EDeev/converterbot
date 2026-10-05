import io
import os
import zipfile

import pytest

os.environ.setdefault("BOT_TOKEN", "123456:TEST")

import bot  # noqa: E402
from rep_to_txt import generate_complete_project_structure  # noqa: E402


def test_zip_bomb_is_rejected(tmp_path, monkeypatch):
    monkeypatch.setattr(bot, "MAX_UNPACKED_SIZE", 1000)
    archive = tmp_path / "a.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("big.txt", "0" * 10_000)
    with pytest.raises(bot.ArchiveTooLarge):
        bot.analyze_archive(str(archive), str(tmp_path))


def test_archive_structure(tmp_path):
    archive = tmp_path / "p.zip"
    with zipfile.ZipFile(archive, "w") as z:
        z.writestr("proj/main.py", "print(1)\n")
        z.writestr("proj/img.png", b"\x89PNG\x00")
        z.writestr("proj/node_modules/x.js", "ignored")
    text = open(bot.analyze_archive(str(archive), str(tmp_path)), encoding="utf-8").read()
    assert "main.py" in text and "   1 | print(1)" in text
    assert "[Image - content not displayed]" in text and "node_modules" not in text


def test_structure_of_missing_path():
    assert generate_complete_project_structure("/no/such/dir").startswith("Error")


def test_docx_conversion_in_bot(tmp_path):
    md = tmp_path / "a.md"
    md.write_text("# Тест\n", encoding="utf-8")
    out = bot.convert_md_to_docx(str(md), str(tmp_path))
    assert zipfile.is_zipfile(io.BytesIO(open(out, "rb").read()))
