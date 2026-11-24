"""文字翻译和视频字幕翻译。"""

from __future__ import annotations

import subprocess
from pathlib import Path

from fastapi.testclient import TestClient

from shop.shopapp.api import app
from shop.shopapp.translate.engine import read_google_payload, translate_text
from shop.shopapp.translate.video import Segment, translate_upload

client = TestClient(app)


def test_google_payload_joins_chunks_and_reads_detected_language():
    text, detected = read_google_payload(
        [[["This is a Bluetooth headset", "这是一款蓝牙耳机"]], None, "zh-CN"]
    )
    assert text == "This is a Bluetooth headset"
    assert detected == "zh"


def test_translate_text_reads_the_curl_response(monkeypatch):
    def fake_run(cmd, **kwargs):
        class Result:
            returncode = 0
            stdout = '[[["This is a Bluetooth headset","这是一款蓝牙耳机"]],null,"zh-CN"]'

        return Result()

    monkeypatch.setattr("shop.shopapp.translate.engine.subprocess.run", fake_run)
    text, detected = translate_text("这是一款蓝牙耳机", "zh", "en")
    assert text == "This is a Bluetooth headset"
    assert detected == "zh"


def test_text_translation_uses_selected_languages(monkeypatch):
    def fake(text, source, target):
        assert text == "这是一款蓝牙耳机"
        assert source == "zh"
        assert target == "en"
        return "This is a Bluetooth headset", "zh"

    monkeypatch.setattr("shop.shopapp.api.translate_text", fake)
    response = client.post(
        "/api/translate/text",
        data={"source": "zh", "target": "en", "text": "这是一款蓝牙耳机"},
    )
    assert response.status_code == 200
    body = response.json()
    assert body["translation"] == "This is a Bluetooth headset"
    assert body["source"] == "zh"
    assert body["target"] == "en"


def test_text_translation_rejects_empty_and_unknown_language():
    empty = client.post("/api/translate/text", data={"source": "zh", "target": "en", "text": "  "})
    unknown = client.post("/api/translate/text", data={"source": "xx", "target": "en", "text": "hello"})
    assert empty.status_code == 400
    assert empty.json()["error"] == "请输入要翻译的文字"
    assert unknown.status_code == 400
    assert unknown.json()["error"] == "没有这个输入语言"


def test_video_translation_burns_subtitles(monkeypatch, tmp_path):
    def fake_transcribe(path, language):
        assert language == "zh"
        return [Segment(0.2, 1.4, "这是一款蓝牙耳机")], "zh"

    def fake_translate(text, source, target):
        assert text == "这是一款蓝牙耳机"
        assert source == "zh"
        assert target == "en"
        return "This is a Bluetooth headset", "zh"

    monkeypatch.setattr("shop.shopapp.translate.video.transcribe", fake_transcribe)
    monkeypatch.setattr("shop.shopapp.translate.video.translate_text", fake_translate)
    video = tmp_path / "clip.mp4"
    _silent_clip(video)
    response = client.post(
        "/api/translate/video",
        data={"source": "zh", "target": "en"},
        files={"file": ("clip.mp4", video.read_bytes(), "video/mp4")},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["transcript"] == "这是一款蓝牙耳机"
    assert body["translation"] == "This is a Bluetooth headset"
    assert body["source_name"] == "中文"
    assert body["target_name"] == "英语"
    saved = client.get(f"/api/translate/video/{body['video_id']}")
    assert saved.status_code == 200
    assert saved.headers["content-type"].startswith("video/mp4")
    assert saved.content[:8] == b"\x00\x00\x00\x18ftyp" or b"ftyp" in saved.content[:32]


def test_video_without_speech_is_rejected(monkeypatch, tmp_path):
    monkeypatch.setattr("shop.shopapp.translate.video.transcribe", lambda path, language: ([], "zh"))
    video = tmp_path / "quiet.mp4"
    _silent_clip(video)
    with client:
        try:
            translate_upload(video.read_bytes(), "quiet.mp4", "auto", "en")
        except ValueError as exc:
            assert str(exc) == "这段视频里没有听出说话内容"
        else:
            raise AssertionError("silent video should be rejected")


def _silent_clip(path: Path) -> None:
    result = subprocess.run(
        [
            "ffmpeg",
            "-y",
            "-f",
            "lavfi",
            "-i",
            "color=c=blue:s=320x180:d=1",
            "-f",
            "lavfi",
            "-i",
            "anullsrc=r=16000:cl=mono",
            "-shortest",
            "-c:v",
            "libx264",
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            str(path),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr[-400:])
