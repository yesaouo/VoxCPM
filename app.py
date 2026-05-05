import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import gradio as gr
import numpy as np
import soundfile as sf
import torch
from funasr import AutoModel

import voxcpm

os.environ["TOKENIZERS_PARALLELISM"] = "false"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


MAX_REFERENCE_AUDIO_SECONDS = 45.0
RECOMMENDED_REFERENCE_AUDIO_SECONDS = 30
PPTX_EXPORT_WIDTH = 1920
PPTX_EXPORT_HEIGHT = 1080
PPTX_VIDEO_FPS = 30
PPTX_DEFAULT_PAUSE_SECONDS = 0.35
PPTX_EMPTY_SLIDE_SECONDS = 1.5
PPTX_DEFAULT_AUDIO_SR = 24000
PPTX_DEFAULT_SUBTITLE_LANGUAGE = "zh"
PPTX_SUBTITLE_LANGUAGE_CHOICES = [
    ("中文（zh）", "zh"),
    ("英文（en）", "en"),
    ("日文（ja）", "ja"),
    ("韓文（ko）", "ko"),
    ("法文（fr）", "fr"),
    ("德文（de）", "de"),
    ("西班牙文（es）", "es"),
    ("義大利文（it）", "it"),
    ("葡萄牙文（pt）", "pt"),
    ("俄文（ru）", "ru"),
]
PPTX_SUBTITLE_MAX_SECONDS = 5.0
PPTX_SUBTITLE_MAX_CHARS = 24

DEFAULT_TARGET_TEXT = "請在這裡輸入要合成的文字。"

INTRO_TEXT = """
# VoxCPM2 Voice & Slide Narration Studio

<div style="font-size: 0.8rem; color: #666; margin-top: -0.5rem;">
Made by Yu-Jie Li
</div>

參考音訊建議約 30 秒；超過 45 秒會自動截斷，只使用前 45 秒。
"""

_CUSTOM_CSS = """
.logo-container {
    text-align: center;
    margin: 0.5rem 0 1rem 0;
}

.logo-container img {
    height: 72px;
    width: auto;
    max-width: 200px;
    display: inline-block;
}

/* Toggle switch style */
.switch-toggle {
    padding: 8px 12px;
    border-radius: 8px;
    background: var(--block-background-fill);
}

.switch-toggle input[type="checkbox"] {
    appearance: none;
    -webkit-appearance: none;
    width: 44px;
    height: 24px;
    background: #ccc;
    border-radius: 12px;
    position: relative;
    cursor: pointer;
    transition: background 0.3s ease;
    flex-shrink: 0;
}

.switch-toggle input[type="checkbox"]::after {
    content: "";
    position: absolute;
    top: 2px;
    left: 2px;
    width: 20px;
    height: 20px;
    background: white;
    border-radius: 50%;
    transition: transform 0.3s ease;
    box-shadow: 0 1px 3px rgba(0,0,0,0.2);
}

.switch-toggle input[type="checkbox"]:checked {
    background: var(--color-accent);
}

.switch-toggle input[type="checkbox"]:checked::after {
    transform: translateX(20px);
}
"""

_APP_THEME = gr.themes.Soft(
    primary_hue="blue",
    secondary_hue="gray",
    neutral_hue="slate",
    font=[gr.themes.GoogleFont("Inter"), "Arial", "sans-serif"],
)


def _trim_reference_audio_if_needed(audio_path: str) -> Tuple[str, bool]:
    try:
        info = sf.info(audio_path)
    except Exception as exc:
        logger.warning("Could not inspect reference audio duration: %s", exc)
        return audio_path, False

    if not info.samplerate or not info.frames:
        return audio_path, False

    duration = info.frames / info.samplerate
    if duration <= MAX_REFERENCE_AUDIO_SECONDS:
        return audio_path, False

    max_frames = int(MAX_REFERENCE_AUDIO_SECONDS * info.samplerate)
    logger.info(
        "Reference audio is %.2fs; truncating to %.0fs.",
        duration,
        MAX_REFERENCE_AUDIO_SECONDS,
    )

    with sf.SoundFile(audio_path) as source:
        audio_data = source.read(frames=max_frames, always_2d=True)
        samplerate = source.samplerate

    with tempfile.NamedTemporaryFile(prefix="voxcpm_ref_", suffix=".wav", delete=False) as temp_file:
        trimmed_path = temp_file.name

    sf.write(trimmed_path, audio_data, samplerate)
    return trimmed_path, True


def _cleanup_temp_audio(audio_path: str, should_delete: bool) -> None:
    if not should_delete:
        return
    try:
        os.unlink(audio_path)
    except OSError as exc:
        logger.warning("Failed to remove temporary audio file %s: %s", audio_path, exc)


def _find_ffmpeg() -> str:
    ffmpeg_path = shutil.which("ffmpeg")
    if ffmpeg_path:
        return ffmpeg_path

    common_path = Path("C:/ffmpeg/bin/ffmpeg.exe")
    if common_path.exists():
        return str(common_path)

    raise gr.Error("找不到 ffmpeg，請先安裝 ffmpeg，或把 ffmpeg.exe 加到 PATH。")


def _run_ffmpeg(args: List[str]) -> None:
    result = subprocess.run(
        args,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.returncode != 0:
        logger.error("ffmpeg failed: %s", result.stderr[-2000:])
        raise gr.Error("影片製作失敗，請確認 ffmpeg 可正常執行。")


def _valid_notes_text(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    placeholders = {
        "click to add notes",
        "click to add text",
        "按一下以新增備忘稿",
        "按一下以新增文字",
        "新增備忘稿",
    }
    return value.lower() not in placeholders


def _shape_text(shape) -> str:
    try:
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return ""
        return str(shape.TextFrame.TextRange.Text).strip()
    except Exception:
        return ""


def _placeholder_type(shape) -> Optional[int]:
    try:
        return int(shape.PlaceholderFormat.Type)
    except Exception:
        return None


def _extract_slide_notes(slide) -> str:
    notes_page = slide.NotesPage
    shapes = notes_page.Shapes

    try:
        body_text = _shape_text(shapes.Placeholders(2))
        if _valid_notes_text(body_text):
            return body_text
    except Exception:
        pass

    skip_placeholder_types = {13, 15, 16}
    parts: List[str] = []
    seen = set()
    for index in range(1, shapes.Count + 1):
        shape = shapes.Item(index)
        placeholder = _placeholder_type(shape)
        if placeholder in skip_placeholder_types:
            continue

        text = _shape_text(shape)
        if not _valid_notes_text(text) or text in seen:
            continue
        parts.append(text)
        seen.add(text)

    return "\n".join(parts).strip()


def _export_pptx_slides_and_notes(pptx_path: str, output_dir: Path) -> List[Tuple[Path, str]]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise gr.Error("PPTX 模式需要 pywin32，請先執行：pip install pywin32") from exc

    pythoncom.CoInitialize()
    powerpoint = None
    presentation = None
    slides: List[Tuple[Path, str]] = []

    try:
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        pptx_file = str(Path(pptx_path).resolve())
        try:
            presentation = powerpoint.Presentations.Open(
                pptx_file,
                ReadOnly=1,
                Untitled=0,
                WithWindow=0,
            )
        except Exception:
            powerpoint.Visible = 1
            presentation = powerpoint.Presentations.Open(
                pptx_file,
                ReadOnly=1,
                Untitled=0,
                WithWindow=1,
            )

        for index in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides.Item(index)
            image_path = output_dir / f"slide_{index:03d}.png"
            slide.Export(str(image_path), "PNG", PPTX_EXPORT_WIDTH, PPTX_EXPORT_HEIGHT)
            slides.append((image_path, _extract_slide_notes(slide)))
    except Exception as exc:
        logger.exception("Failed to export pptx slides and notes.")
        raise gr.Error("無法讀取 PPTX。請確認 PowerPoint 已安裝，且檔案未被其他程式鎖定。") from exc
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return slides


def _audio_to_mono(audio: np.ndarray) -> np.ndarray:
    audio = np.asarray(audio, dtype=np.float32).squeeze()
    if audio.ndim == 0:
        return np.zeros(0, dtype=np.float32)
    if audio.ndim == 1:
        return audio
    if audio.shape[0] <= 2 and audio.shape[1] > audio.shape[0]:
        return audio.mean(axis=0)
    return audio.mean(axis=1)


def _write_audio_with_pauses(
    output_path: Path,
    sample_rate: int,
    audio: Optional[np.ndarray],
    pause_before: float,
    pause_after: float,
    empty_seconds: float = PPTX_EMPTY_SLIDE_SECONDS,
) -> float:
    if audio is None:
        body = np.zeros(int(sample_rate * empty_seconds), dtype=np.float32)
    else:
        body = _audio_to_mono(audio)
        if body.size == 0:
            body = np.zeros(int(sample_rate * empty_seconds), dtype=np.float32)

    before = np.zeros(max(0, int(sample_rate * pause_before)), dtype=np.float32)
    after = np.zeros(max(0, int(sample_rate * pause_after)), dtype=np.float32)
    padded = np.concatenate([before, body, after])
    sf.write(str(output_path), padded, sample_rate)
    return len(padded) / sample_rate


def _render_slide_clip(ffmpeg_path: str, image_path: Path, audio_path: Path, output_path: Path, duration: float) -> None:
    _run_ffmpeg(
        [
            ffmpeg_path,
            "-y",
            "-loop",
            "1",
            "-framerate",
            str(PPTX_VIDEO_FPS),
            "-i",
            str(image_path),
            "-i",
            str(audio_path),
            "-t",
            f"{duration:.3f}",
            "-vf",
            "scale=trunc(iw/2)*2:trunc(ih/2)*2,format=yuv420p",
            "-r",
            str(PPTX_VIDEO_FPS),
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-tune",
            "stillimage",
            "-c:a",
            "aac",
            "-b:a",
            "192k",
            "-shortest",
            str(output_path),
        ]
    )


def _concat_video_clips(ffmpeg_path: str, clip_paths: List[Path], output_path: Path, list_path: Path) -> None:
    lines = []
    for clip_path in clip_paths:
        safe_path = clip_path.resolve().as_posix().replace("'", "'\\''")
        lines.append(f"file '{safe_path}'")
    list_path.write_text("\n".join(lines), encoding="utf-8")

    _run_ffmpeg(
        [
            ffmpeg_path,
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(list_path),
            "-c",
            "copy",
            str(output_path),
        ]
    )


def _load_whisperx():
    try:
        import whisperx
    except ImportError as exc:
        raise gr.Error("要產生字幕請先安裝 WhisperX：pip install whisperx") from exc
    return whisperx


def _get_whisperx_align_model(
    language_code: str,
    device: str,
    align_model_cache: Dict[str, Tuple[Any, Any, Dict[str, Any]]],
) -> Tuple[Any, Any, Dict[str, Any]]:
    language = (language_code or PPTX_DEFAULT_SUBTITLE_LANGUAGE).strip().lower()
    cache_key = f"{language}:{device}"
    if cache_key not in align_model_cache:
        whisperx = _load_whisperx()
        logger.info("Loading WhisperX alignment model for language=%s on %s", language, device)
        try:
            align_model, metadata = whisperx.load_align_model(language_code=language, device=device)
        except Exception as exc:
            raise gr.Error(f"WhisperX 載入 {language} 對齊模型失敗，請確認語言碼或模型下載狀態。") from exc
        align_model_cache[cache_key] = (whisperx, align_model, metadata)
    return align_model_cache[cache_key]


def _align_transcript_with_whisperx(
    audio_path: Path,
    transcript: str,
    duration: float,
    language_code: str,
    device: str,
    align_model_cache: Dict[str, Tuple[Any, Any, Dict[str, Any]]],
) -> Dict[str, Any]:
    whisperx, align_model, metadata = _get_whisperx_align_model(language_code, device, align_model_cache)
    segments = [
        {
            "start": 0.0,
            "end": max(0.1, float(duration)),
            "text": transcript.strip(),
        }
    ]
    return whisperx.align(
        segments,
        align_model,
        metadata,
        str(audio_path),
        device,
        return_char_alignments=False,
    )


def _language_without_word_spaces(language_code: str) -> bool:
    return (language_code or "").lower().split("-")[0] in {"zh", "ja", "th", "lo", "my", "km"}


def _join_subtitle_tokens(tokens: List[str], language_code: str) -> str:
    clean_tokens = [token.strip() for token in tokens if token and token.strip()]
    if _language_without_word_spaces(language_code):
        parts: List[str] = []
        for token in clean_tokens:
            if parts and _subtitle_tokens_need_space(parts[-1], token):
                parts.append(" ")
            parts.append(token)
        text = "".join(parts)
    else:
        text = " ".join(clean_tokens)

    text = re.sub(r"\s+([,.;:!?，。！？、；：）\]\}])", r"\1", text)
    text = re.sub(r"([(（\[\{])\s+", r"\1", text)
    return text.strip()


def _subtitle_tokens_need_space(previous: str, current: str) -> bool:
    if not previous or not current:
        return False
    return _ascii_word_edge(previous[-1]) and _ascii_word_edge(current[0])


def _ascii_word_edge(char: str) -> bool:
    return char.isascii() and char.isalnum()


def _ascii_word_char(char: str) -> bool:
    return char.isascii() and (char.isalnum() or char in {"_", "-", "'"})


def _subtitle_boundary_inside_ascii_word(
    text: str,
    left_end: Optional[int],
    right_start: int,
) -> bool:
    if left_end is None or left_end <= 0 or right_start >= len(text):
        return False

    # 如果中間有空白或標點，就不是同一個英文單字
    if text[left_end:right_start]:
        return False

    return _ascii_word_char(text[left_end - 1]) and _ascii_word_char(text[right_start])


def _split_text_preserving_ascii_words(text: str, max_chars: int) -> List[str]:
    units: List[str] = []
    index = 0

    while index < len(text):
        char = text[index]

        # 連續英文、數字、底線、連字號、撇號視為同一個 token
        if _ascii_word_char(char):
            start = index
            index += 1
            while index < len(text) and _ascii_word_char(text[index]):
                index += 1
            units.append(text[start:index])
            continue

        units.append(char)
        index += 1

    chunks: List[str] = []
    current = ""

    for unit in units:
        candidate = current + unit

        if current and len(candidate) > max_chars:
            chunk = current.strip()
            if chunk:
                chunks.append(chunk)
            current = unit.lstrip()
        else:
            current = candidate

    final_chunk = current.strip()
    if final_chunk:
        chunks.append(final_chunk)

    return chunks


def _subtitle_sentence_ended(text: str) -> bool:
    return bool(text) and text.rstrip()[-1] in ".!?。！？；;"


def _clean_subtitle_display_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def _build_timed_transcript_units(
    word_segments: List[Dict[str, Any]],
    transcript: str,
) -> List[Dict[str, float]]:
    nonspace_positions = [index for index, char in enumerate(transcript) if not char.isspace()]
    if not nonspace_positions:
        return []

    units: List[Dict[str, float]] = []
    cursor = 0
    for word in word_segments:
        token = str(word.get("word", "")).strip()
        token_length = sum(1 for char in token if not char.isspace())
        if token_length <= 0 or cursor >= len(nonspace_positions):
            continue
        try:
            word_start = float(word.get("start"))
        except (TypeError, ValueError):
            continue
        try:
            word_end = float(word.get("end"))
        except (TypeError, ValueError):
            word_end = word_start + 0.25
        word_end = max(word_end, word_start + 0.05)

        end_cursor = min(len(nonspace_positions), cursor + token_length)
        units.append(
            {
                "start": word_start,
                "end": word_end,
                "text_start": float(nonspace_positions[cursor]),
                "text_end": float(nonspace_positions[end_cursor - 1] + 1),
            }
        )
        cursor = end_cursor

    return units


def _build_subtitle_cues_from_transcript_units(
    units: List[Dict[str, float]],
    transcript: str,
    offset_seconds: float,
    language_code: str,
) -> List[Tuple[float, float, str]]:
    max_chars = PPTX_SUBTITLE_MAX_CHARS if _language_without_word_spaces(language_code) else 42
    cues: List[Tuple[float, float, str]] = []
    current_text_start: Optional[int] = None
    current_text_end: Optional[int] = None
    current_audio_start: Optional[float] = None
    current_audio_end: Optional[float] = None

    for unit in units:
        unit_text_start = int(unit["text_start"])
        unit_text_end = int(unit["text_end"])
        unit_audio_start = float(unit["start"])
        unit_audio_end = float(unit["end"])

        candidate_text_start = current_text_start if current_text_start is not None else unit_text_start
        candidate_audio_start = current_audio_start if current_audio_start is not None else unit_audio_start
        candidate_text = _clean_subtitle_display_text(transcript[candidate_text_start:unit_text_end])
        candidate_duration = unit_audio_end - candidate_audio_start

        should_split = current_text_start is not None and (
            len(candidate_text) > max_chars or candidate_duration > PPTX_SUBTITLE_MAX_SECONDS
        )

        split_would_cut_ascii_word = (
            should_split
            and current_text_end is not None
            and _subtitle_boundary_inside_ascii_word(
                transcript,
                current_text_end,
                unit_text_start,
            )
        )

        if should_split and not split_would_cut_ascii_word:
            cue_text = _clean_subtitle_display_text(transcript[current_text_start:current_text_end])
            if cue_text and current_audio_start is not None and current_audio_end is not None:
                cues.append((offset_seconds + current_audio_start, offset_seconds + current_audio_end, cue_text))
            current_text_start = None
            current_text_end = None
            current_audio_start = None
            current_audio_end = None

        if current_text_start is None:
            current_text_start = unit_text_start
            current_audio_start = unit_audio_start
        current_text_end = unit_text_end
        current_audio_end = unit_audio_end

        cue_text = _clean_subtitle_display_text(transcript[current_text_start:current_text_end])
        if _subtitle_sentence_ended(cue_text):
            if cue_text and current_audio_start is not None and current_audio_end is not None:
                cues.append((offset_seconds + current_audio_start, offset_seconds + current_audio_end, cue_text))
            current_text_start = None
            current_text_end = None
            current_audio_start = None
            current_audio_end = None

    if current_text_start is not None and current_text_end is not None:
        cue_text = _clean_subtitle_display_text(transcript[current_text_start:current_text_end])
        if cue_text and current_audio_start is not None and current_audio_end is not None:
            cues.append((offset_seconds + current_audio_start, offset_seconds + current_audio_end, cue_text))

    return cues


def _build_subtitle_cues_from_words(
    word_segments: List[Dict[str, Any]],
    offset_seconds: float,
    language_code: str,
    transcript: str = "",
) -> List[Tuple[float, float, str]]:
    if transcript:
        units = _build_timed_transcript_units(word_segments, transcript)
        if units:
            return _build_subtitle_cues_from_transcript_units(
                units,
                transcript,
                offset_seconds,
                language_code,
            )

    max_chars = PPTX_SUBTITLE_MAX_CHARS if _language_without_word_spaces(language_code) else 42
    cues: List[Tuple[float, float, str]] = []
    current_tokens: List[str] = []
    current_start: Optional[float] = None
    current_end: Optional[float] = None

    for word in word_segments:
        token = str(word.get("word", "")).strip()
        if not token:
            continue
        try:
            word_start = float(word.get("start"))
        except (TypeError, ValueError):
            continue
        try:
            word_end = float(word.get("end"))
        except (TypeError, ValueError):
            word_end = word_start + 0.25
        word_end = max(word_end, word_start + 0.05)

        candidate_tokens = current_tokens + [token]
        candidate_text = _join_subtitle_tokens(candidate_tokens, language_code)
        candidate_start = current_start if current_start is not None else word_start
        candidate_duration = word_end - candidate_start

        if current_tokens and (
            len(candidate_text) > max_chars or candidate_duration > PPTX_SUBTITLE_MAX_SECONDS
        ):
            cue_text = _join_subtitle_tokens(current_tokens, language_code)
            if cue_text and current_start is not None and current_end is not None:
                cues.append((offset_seconds + current_start, offset_seconds + current_end, cue_text))
            current_tokens = []
            current_start = None
            current_end = None

        if current_start is None:
            current_start = word_start
        current_tokens.append(token)
        current_end = word_end

        if _subtitle_sentence_ended(token):
            cue_text = _join_subtitle_tokens(current_tokens, language_code)
            if cue_text and current_start is not None and current_end is not None:
                cues.append((offset_seconds + current_start, offset_seconds + current_end, cue_text))
            current_tokens = []
            current_start = None
            current_end = None

    cue_text = _join_subtitle_tokens(current_tokens, language_code)
    if cue_text and current_start is not None and current_end is not None:
        cues.append((offset_seconds + current_start, offset_seconds + current_end, cue_text))

    return cues


def _split_subtitle_text(text: str, language_code: str) -> List[str]:
    normalized = re.sub(r"\s+", " ", (text or "").replace("\r", " ").replace("\n", " ")).strip()
    if not normalized:
        return []

    max_chars = PPTX_SUBTITLE_MAX_CHARS if _language_without_word_spaces(language_code) else 42
    sentence_parts = [
        part.strip()
        for part in re.split(r"(?<=[。！？.!?；;])\s*", normalized)
        if part and part.strip()
    ]
    chunks: List[str] = []

    for part in sentence_parts:
        if _language_without_word_spaces(language_code):
            chunks.extend(_split_text_preserving_ascii_words(part, max_chars))
            continue

        current_words: List[str] = []
        for word in part.split():
            candidate = " ".join(current_words + [word])
            if current_words and len(candidate) > max_chars:
                chunks.append(" ".join(current_words))
                current_words = [word]
            else:
                current_words.append(word)
        if current_words:
            chunks.append(" ".join(current_words))

    return chunks


def _build_fallback_subtitle_cues(
    text: str,
    start_seconds: float,
    end_seconds: float,
    language_code: str,
) -> List[Tuple[float, float, str]]:
    chunks = _split_subtitle_text(text, language_code)
    if not chunks:
        return []

    body_start = max(0.0, float(start_seconds))
    body_end = max(body_start + 0.5, float(end_seconds))
    available = body_end - body_start
    cues: List[Tuple[float, float, str]] = []
    for index, chunk in enumerate(chunks):
        cue_start = body_start + available * index / len(chunks)
        cue_end = body_start + available * (index + 1) / len(chunks)
        cues.append((cue_start, cue_end, chunk))
    return cues


def _format_srt_timestamp(seconds: float) -> str:
    total_ms = max(0, int(round(seconds * 1000)))
    hours, remainder = divmod(total_ms, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    secs, millis = divmod(remainder, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"


def _write_srt_file(cues: List[Tuple[float, float, str]], output_path: Path) -> None:
    lines: List[str] = []
    last_end = 0.0
    for index, (start, end, text) in enumerate(cues, start=1):
        cue_start = max(0.0, float(start))
        cue_end = max(cue_start + 0.1, float(end))
        if cue_start < last_end:
            cue_start = last_end + 0.001
            cue_end = max(cue_end, cue_start + 0.1)
        last_end = cue_end

        lines.extend(
            [
                str(index),
                f"{_format_srt_timestamp(cue_start)} --> {_format_srt_timestamp(cue_end)}",
                text.strip(),
                "",
            ]
        )

    output_path.write_text("\n".join(lines), encoding="utf-8")


def _escape_ffmpeg_filter_path(path: Path) -> str:
    value = path.resolve().as_posix()
    value = value.replace("\\", "/")
    value = value.replace(":", "\\:")
    value = value.replace("'", "\\'")
    value = value.replace(",", "\\,")
    return value


def _burn_subtitles_into_video(
    ffmpeg_path: str,
    input_video_path: Path,
    subtitle_path: Path,
    output_video_path: Path,
) -> None:
    subtitle_file = _escape_ffmpeg_filter_path(subtitle_path)
    style = (
        "FontName=Microsoft JhengHei,"
        "FontSize=18,"
        "PrimaryColour=&H00FFFFFF,"
        "OutlineColour=&H00000000,"
        "BorderStyle=1,"
        "Outline=2,"
        "Shadow=0,"
        "MarginV=36"
    )
    _run_ffmpeg(
        [
            ffmpeg_path,
            "-y",
            "-i",
            str(input_video_path),
            "-vf",
            f"subtitles='{subtitle_file}':charenc=UTF-8:force_style='{style}'",
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-c:a",
            "copy",
            str(output_video_path),
        ]
    )


class VoxCPMDemo:
    def __init__(self, model_id: str = "openbmb/VoxCPM2") -> None:
        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        logger.info("Running on device: %s", self.device)

        self.asr_model_id = "iic/SenseVoiceSmall"
        self.asr_model: Optional[AutoModel] = AutoModel(
            model=self.asr_model_id,
            disable_update=True,
            log_level="DEBUG",
            device="cuda:0" if self.device == "cuda" else "cpu",
        )

        self.voxcpm_model: Optional[voxcpm.VoxCPM] = None
        self._model_id = model_id

    def get_or_load_voxcpm(self) -> voxcpm.VoxCPM:
        if self.voxcpm_model is not None:
            return self.voxcpm_model
        logger.info("Loading model: %s", self._model_id)
        self.voxcpm_model = voxcpm.VoxCPM.from_pretrained(self._model_id, optimize=True)
        logger.info("Model loaded successfully.")
        return self.voxcpm_model

    def prompt_wav_recognition(self, prompt_wav: Optional[str]) -> str:
        if prompt_wav is None:
            return ""
        audio_path, should_cleanup = _trim_reference_audio_if_needed(prompt_wav)
        try:
            res = self.asr_model.generate(input=audio_path, language="auto", use_itn=True)
            return res[0]["text"].split("|>")[-1]
        finally:
            _cleanup_temp_audio(audio_path, should_cleanup)

    def _generate_tts_with_prepared_reference(
        self,
        *,
        text: str,
        audio_path: str,
        prompt_text_clean: str,
        cfg_value_input: float,
        do_normalize: bool,
        denoise: bool,
        inference_timesteps: int,
    ) -> Tuple[int, np.ndarray]:
        current_model = self.get_or_load_voxcpm()
        logger.info("[Ultimate Cloning] prompt_wav + prompt_text + reference_wav")
        logger.info("Generating audio for text: '%s...'", text[:80])

        wav = current_model.generate(
            text=text,
            reference_wav_path=audio_path,
            prompt_wav_path=audio_path,
            prompt_text=prompt_text_clean,
            cfg_value=float(cfg_value_input),
            inference_timesteps=int(inference_timesteps),
            normalize=do_normalize,
            denoise=denoise,
        )
        return (current_model.tts_model.sample_rate, wav)

    def generate_tts_audio(
        self,
        text_input: str,
        reference_wav_path_input: Optional[str],
        prompt_text: str,
        cfg_value_input: float = 2.0,
        do_normalize: bool = False,
        denoise: bool = False,
        inference_timesteps: int = 10,
    ) -> Tuple[int, np.ndarray]:
        text = (text_input or "").strip()
        if not text:
            raise gr.Error("請輸入要合成的文字。")

        audio_path = reference_wav_path_input if reference_wav_path_input else None
        if not audio_path:
            raise gr.Error("請上傳參考音訊。")

        prompt_text_clean = (prompt_text or "").strip()
        if not prompt_text_clean:
            raise gr.Error("請提供參考音訊逐字稿。")

        audio_path, should_cleanup = _trim_reference_audio_if_needed(audio_path)
        try:
            return self._generate_tts_with_prepared_reference(
                text=text,
                audio_path=audio_path,
                prompt_text_clean=prompt_text_clean,
                cfg_value_input=cfg_value_input,
                do_normalize=do_normalize,
                denoise=denoise,
                inference_timesteps=inference_timesteps,
            )
        finally:
            _cleanup_temp_audio(audio_path, should_cleanup)

    def generate_pptx_video(
        self,
        pptx_path_input: Optional[str],
        reference_wav_path_input: Optional[str],
        prompt_text: str,
        cfg_value_input: float = 2.0,
        do_normalize: bool = False,
        denoise: bool = False,
        inference_timesteps: int = 10,
        pause_before: float = PPTX_DEFAULT_PAUSE_SECONDS,
        pause_after: float = PPTX_DEFAULT_PAUSE_SECONDS,
        enable_subtitles: bool = True,
        subtitle_language: str = PPTX_DEFAULT_SUBTITLE_LANGUAGE,
        burn_subtitles: bool = True,
        progress=None,
    ) -> Tuple[str, Optional[str], str]:
        if not pptx_path_input:
            raise gr.Error("請上傳 PPTX 檔案。")

        audio_path = reference_wav_path_input if reference_wav_path_input else None
        if not audio_path:
            raise gr.Error("請上傳參考音訊。")

        prompt_text_clean = (prompt_text or "").strip()
        if not prompt_text_clean:
            raise gr.Error("請提供參考音訊逐字稿。")

        output_file = tempfile.NamedTemporaryFile(prefix="voxcpm_pptx_", suffix=".mp4", delete=False)
        output_path = Path(output_file.name)
        output_file.close()
        output_video_path = output_path
        subtitle_path: Optional[Path] = None

        prepared_audio_path, should_cleanup = _trim_reference_audio_if_needed(audio_path)
        subtitle_language_clean = (subtitle_language or PPTX_DEFAULT_SUBTITLE_LANGUAGE).strip().lower()
        subtitle_device = "cuda" if self.device == "cuda" else "cpu"
        align_model_cache: Dict[str, Tuple[Any, Any, Dict[str, Any]]] = {}
        subtitle_cues: List[Tuple[float, float, str]] = []
        subtitle_fallback_count = 0
        timeline_offset = 0.0
        burn_status = ""

        try:
            ffmpeg_path = _find_ffmpeg()
            with tempfile.TemporaryDirectory(prefix="voxcpm_pptx_work_") as temp_dir:
                work_dir = Path(temp_dir)
                if progress:
                    progress(0, desc="讀取 PPTX 並擷取投影片")
                slides = _export_pptx_slides_and_notes(pptx_path_input, work_dir)
                if not slides:
                    raise gr.Error("PPTX 沒有可用的投影片。")

                clip_paths: List[Path] = []
                slides_with_notes = 0
                total = len(slides)

                for index, (slide_image, notes) in enumerate(slides, start=1):
                    if progress:
                        progress((index - 1) / total, desc=f"處理第 {index} / {total} 頁")

                    audio_file = work_dir / f"slide_{index:03d}.wav"
                    clip_file = work_dir / f"clip_{index:03d}.mp4"
                    notes_clean = notes.strip()

                    if notes_clean:
                        slides_with_notes += 1
                        sample_rate, wav = self._generate_tts_with_prepared_reference(
                            text=notes_clean,
                            audio_path=prepared_audio_path,
                            prompt_text_clean=prompt_text_clean,
                            cfg_value_input=cfg_value_input,
                            do_normalize=do_normalize,
                            denoise=denoise,
                            inference_timesteps=inference_timesteps,
                        )
                        duration = _write_audio_with_pauses(
                            audio_file,
                            sample_rate,
                            wav,
                            pause_before,
                            pause_after,
                        )
                    else:
                        duration = _write_audio_with_pauses(
                            audio_file,
                            PPTX_DEFAULT_AUDIO_SR,
                            None,
                            pause_before,
                            pause_after,
                        )

                    if enable_subtitles and notes_clean:
                        slide_cues: List[Tuple[float, float, str]] = []
                        try:
                            if progress:
                                progress((index - 0.35) / total, desc=f"對齊第 {index} / {total} 頁字幕")
                            aligned = _align_transcript_with_whisperx(
                                audio_file,
                                notes_clean,
                                duration,
                                subtitle_language_clean,
                                subtitle_device,
                                align_model_cache,
                            )
                            slide_cues = _build_subtitle_cues_from_words(
                                aligned.get("word_segments", []),
                                timeline_offset,
                                subtitle_language_clean,
                                notes_clean,
                            )
                        except gr.Error:
                            raise
                        except Exception as exc:
                            logger.warning("WhisperX alignment failed for slide %s: %s", index, exc)

                        if not slide_cues:
                            subtitle_fallback_count += 1
                            slide_cues = _build_fallback_subtitle_cues(
                                notes_clean,
                                timeline_offset + max(0.0, float(pause_before)),
                                timeline_offset + max(0.0, duration - max(0.0, float(pause_after))),
                                subtitle_language_clean,
                            )
                        subtitle_cues.extend(slide_cues)

                    _render_slide_clip(ffmpeg_path, slide_image, audio_file, clip_file, duration)
                    clip_paths.append(clip_file)
                    timeline_offset += duration

                if progress:
                    progress(0.95, desc="合成影片")
                _concat_video_clips(ffmpeg_path, clip_paths, output_path, work_dir / "clips.txt")

            if enable_subtitles and subtitle_cues:
                subtitle_file = tempfile.NamedTemporaryFile(prefix="voxcpm_pptx_subtitles_", suffix=".srt", delete=False)
                subtitle_path = Path(subtitle_file.name)
                subtitle_file.close()
                _write_srt_file(subtitle_cues, subtitle_path)

                if burn_subtitles:
                    subtitled_file = tempfile.NamedTemporaryFile(
                        prefix="voxcpm_pptx_subtitled_",
                        suffix=".mp4",
                        delete=False,
                    )
                    subtitled_path = Path(subtitled_file.name)
                    subtitled_file.close()
                    try:
                        if progress:
                            progress(0.98, desc="燒錄字幕到影片")
                        _burn_subtitles_into_video(ffmpeg_path, output_path, subtitle_path, subtitled_path)
                        output_video_path = subtitled_path
                    except gr.Error as exc:
                        burn_status = "字幕 SRT 已產生，但 ffmpeg 燒錄字幕失敗；影片先回傳未燒錄版本。"
                        logger.warning("Failed to burn subtitles into video: %s", exc)

            if progress:
                progress(1.0, desc="完成")

            status_parts = [f"完成：共 {len(slides)} 頁，{slides_with_notes} 頁有備忘稿旁白。"]
            if enable_subtitles:
                if subtitle_path:
                    status_parts.append(f"字幕：{len(subtitle_cues)} 段，已依每頁換頁停頓對齊。")
                    if subtitle_fallback_count:
                        status_parts.append(f"{subtitle_fallback_count} 頁未取得 WhisperX 字詞時間，已用講稿平均分段。")
                    if burn_subtitles and output_video_path != output_path:
                        status_parts.append("已燒錄字幕到影片，並另提供 SRT。")
                    elif burn_status:
                        status_parts.append(burn_status)
                else:
                    status_parts.append("字幕：沒有可對齊的備忘稿，因此未產生 SRT。")

            return str(output_video_path), str(subtitle_path) if subtitle_path else None, "\n\n".join(status_parts)
        finally:
            _cleanup_temp_audio(prepared_audio_path, should_cleanup)


def create_demo_interface(demo: VoxCPMDemo):
    gr.set_static_paths(paths=[Path.cwd().absolute() / "assets"])

    def _generate_single(
        ref_wav: Optional[str],
        transcript_text: str,
        text: str,
        cfg_value: float,
        do_normalize: bool,
        denoise: bool,
        dit_steps: int,
    ):
        return demo.generate_tts_audio(
            text_input=text,
            reference_wav_path_input=ref_wav,
            prompt_text=transcript_text,
            cfg_value_input=cfg_value,
            do_normalize=do_normalize,
            denoise=denoise,
            inference_timesteps=int(dit_steps),
        )

    def _generate_pptx(
        pptx_path: Optional[str],
        ref_wav: Optional[str],
        transcript_text: str,
        cfg_value: float,
        do_normalize: bool,
        denoise: bool,
        dit_steps: int,
        pause_before: float,
        pause_after: float,
        enable_subtitles: bool,
        subtitle_language: str,
        burn_subtitles: bool,
        progress=gr.Progress(),
    ):
        return demo.generate_pptx_video(
            pptx_path_input=pptx_path,
            reference_wav_path_input=ref_wav,
            prompt_text=transcript_text,
            cfg_value_input=cfg_value,
            do_normalize=do_normalize,
            denoise=denoise,
            inference_timesteps=int(dit_steps),
            pause_before=float(pause_before),
            pause_after=float(pause_after),
            enable_subtitles=bool(enable_subtitles),
            subtitle_language=subtitle_language,
            burn_subtitles=bool(burn_subtitles),
            progress=progress,
        )

    def _run_asr(audio_path: Optional[str]):
        if not audio_path:
            raise gr.Error("請先上傳參考音訊。")
        try:
            logger.info("Running ASR on reference audio...")
            asr_text = demo.prompt_wav_recognition(audio_path)
            logger.info("ASR result: %s...", asr_text[:60])
            return gr.update(value=asr_text)
        except Exception as exc:
            logger.warning("ASR recognition failed: %s", exc)
            raise gr.Error("自動辨識失敗，請手動貼上逐字稿。") from exc

    with gr.Blocks(
        title="VoxCPM2 Voice & Slide Narration Studio",
        theme=_APP_THEME,
        css=_CUSTOM_CSS,
    ) as interface:
        gr.Markdown(INTRO_TEXT)

        with gr.Tabs():
            with gr.Tab("單句語音"):
                with gr.Row():
                    with gr.Column(scale=1):
                        reference_wav = gr.Audio(
                            sources=["upload", "microphone"],
                            type="filepath",
                            label="參考音訊",
                        )
                        transcript_text = gr.Textbox(
                            value="",
                            label="參考音訊逐字稿",
                            placeholder="貼上參考音訊的逐字稿，或用下方按鈕自動辨識。若音訊超過 45 秒，請填寫前 45 秒逐字稿。",
                            lines=5,
                        )
                        transcribe_btn = gr.Button("自動辨識參考音訊", variant="secondary")
                        text = gr.Textbox(
                            value=DEFAULT_TARGET_TEXT,
                            label="要合成的文字",
                            lines=5,
                        )

                        with gr.Accordion("進階設定", open=False):
                            denoise_prompt_audio = gr.Checkbox(
                                value=False,
                                label="參考音訊增強",
                                elem_classes=["switch-toggle"],
                                info="生成前先對參考音訊套用 ZipEnhancer 降噪。",
                            )
                            normalize_text = gr.Checkbox(
                                value=False,
                                label="文字正規化",
                                elem_classes=["switch-toggle"],
                                info="使用 wetext 正規化數字、日期與縮寫。",
                            )
                            cfg_value = gr.Slider(
                                minimum=1.0,
                                maximum=3.0,
                                value=2.0,
                                step=0.1,
                                label="CFG",
                                info="數值越高越貼近參考音訊與逐字稿，越低變化較多。",
                            )
                            dit_steps = gr.Slider(
                                minimum=1,
                                maximum=50,
                                value=10,
                                step=1,
                                label="LocDiT steps",
                                info="步數越高可能提升品質，但生成速度較慢。",
                            )

                        run_btn = gr.Button("生成語音", variant="primary", size="lg")

                    with gr.Column(scale=1):
                        audio_output = gr.Audio(label="生成結果")

                transcribe_btn.click(
                    fn=_run_asr,
                    inputs=[reference_wav],
                    outputs=[transcript_text],
                    show_progress=True,
                )

                run_btn.click(
                    fn=_generate_single,
                    inputs=[
                        reference_wav,
                        transcript_text,
                        text,
                        cfg_value,
                        normalize_text,
                        denoise_prompt_audio,
                        dit_steps,
                    ],
                    outputs=[audio_output],
                    show_progress=True,
                    api_name="generate",
                )

            with gr.Tab("PPTX 影片"):
                with gr.Row():
                    with gr.Column(scale=1):
                        gr.Markdown("會自動擷取每頁投影片截圖與備忘稿；備忘稿會轉成旁白。換頁前後會各停頓一下。")
                        pptx_file = gr.File(
                            label="PPTX 檔案",
                            file_types=[".pptx"],
                            type="filepath",
                        )
                        subtitle_language = gr.Dropdown(
                            choices=PPTX_SUBTITLE_LANGUAGE_CHOICES,
                            value=PPTX_DEFAULT_SUBTITLE_LANGUAGE,
                            label="備忘稿語言",
                            info="請選擇備忘稿的語言。",
                            allow_custom_value=False,
                        )
                        pptx_reference_wav = gr.Audio(
                            sources=["upload", "microphone"],
                            type="filepath",
                            label="參考音訊",
                        )
                        pptx_transcript_text = gr.Textbox(
                            value="",
                            label="參考音訊逐字稿",
                            placeholder="貼上參考音訊的逐字稿，或用下方按鈕自動辨識。",
                            lines=5,
                        )
                        pptx_transcribe_btn = gr.Button("自動辨識參考音訊", variant="secondary")

                        with gr.Accordion("進階設定", open=False):
                            pptx_denoise_prompt_audio = gr.Checkbox(
                                value=False,
                                label="參考音訊增強",
                                elem_classes=["switch-toggle"],
                                info="生成前先對參考音訊套用 ZipEnhancer 降噪。",
                            )
                            pptx_normalize_text = gr.Checkbox(
                                value=False,
                                label="文字正規化",
                                elem_classes=["switch-toggle"],
                                info="使用 wetext 正規化數字、日期與縮寫。",
                            )
                            pptx_cfg_value = gr.Slider(
                                minimum=1.0,
                                maximum=3.0,
                                value=2.0,
                                step=0.1,
                                label="CFG",
                                info="數值越高越貼近參考音訊與逐字稿，越低變化較多。",
                            )
                            pptx_dit_steps = gr.Slider(
                                minimum=1,
                                maximum=50,
                                value=10,
                                step=1,
                                label="LocDiT steps",
                                info="步數越高可能提升品質，但生成速度較慢。",
                            )
                            pause_before = gr.Slider(
                                minimum=0.0,
                                maximum=2.0,
                                value=PPTX_DEFAULT_PAUSE_SECONDS,
                                step=0.05,
                                label="換頁後停頓秒數",
                            )
                            pause_after = gr.Slider(
                                minimum=0.0,
                                maximum=2.0,
                                value=PPTX_DEFAULT_PAUSE_SECONDS,
                                step=0.05,
                                label="換頁前停頓秒數",
                            )
                            enable_subtitles = gr.Checkbox(
                                value=True,
                                label="產生字幕檔",
                                elem_classes=["switch-toggle"],
                                info="開啟後，會根據每頁備忘稿產生 SRT 字幕檔。",
                            )
                            burn_subtitles = gr.Checkbox(
                                value=True,
                                label="在影片中顯示字幕",
                                elem_classes=["switch-toggle"],
                                info="開啟後，字幕會直接顯示在影片畫面上。",
                            )

                        pptx_run_btn = gr.Button("製作影片", variant="primary", size="lg")

                    with gr.Column(scale=1):
                        pptx_video_output = gr.Video(label="生成影片")
                        pptx_subtitle_output = gr.File(label="字幕 SRT")
                        pptx_status = gr.Markdown()

                pptx_transcribe_btn.click(
                    fn=_run_asr,
                    inputs=[pptx_reference_wav],
                    outputs=[pptx_transcript_text],
                    show_progress=True,
                )

                pptx_run_btn.click(
                    fn=_generate_pptx,
                    inputs=[
                        pptx_file,
                        pptx_reference_wav,
                        pptx_transcript_text,
                        pptx_cfg_value,
                        pptx_normalize_text,
                        pptx_denoise_prompt_audio,
                        pptx_dit_steps,
                        pause_before,
                        pause_after,
                        enable_subtitles,
                        subtitle_language,
                        burn_subtitles,
                    ],
                    outputs=[pptx_video_output, pptx_subtitle_output, pptx_status],
                    show_progress=True,
                    api_name="generate_pptx_video",
                )

    return interface


def run_demo(
    server_name: str = "0.0.0.0",
    server_port: int = 8808,
    show_error: bool = True,
    model_id: str = "openbmb/VoxCPM2",
):
    demo = VoxCPMDemo(model_id=model_id)
    interface = create_demo_interface(demo)
    interface.queue(max_size=10, default_concurrency_limit=1).launch(
        server_name=server_name,
        server_port=server_port,
        show_error=show_error,
    )


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--model-id",
        type=str,
        default="openbmb/VoxCPM2",
        help="Local path or HuggingFace repo ID (default: openbmb/VoxCPM2)",
    )
    parser.add_argument("--port", type=int, default=8808, help="Server port")
    args = parser.parse_args()
    run_demo(model_id=args.model_id, server_port=args.port)
