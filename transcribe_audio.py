#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""音频转写最小示例（faster-whisper / CPU + int8）

用途：单文件快速转写，把带时间戳的结果写成 txt。
需要批量转写、多种输出格式或断点续跑，请用 transcribe_hq.py。

用法：
    python transcribe_audio.py <音频文件> [输出txt] [-m 模型]
    python transcribe_audio.py "录音.mp3"
    python transcribe_audio.py "录音.mp3" out.txt -m small
"""
import argparse
import sys
from pathlib import Path

from faster_whisper import WhisperModel


def build_parser():
    ap = argparse.ArgumentParser(description="音频转写最小示例")
    ap.add_argument("audio", help="音频文件路径")
    ap.add_argument("out_txt", nargs="?", default=None,
                    help="输出 txt 路径；不填则输出到音频同目录下的 <音频名>_转写.txt")
    ap.add_argument("-m", "--model", default="small",
                    help="模型名或本地模型目录，默认 small")
    return ap


def main():
    args = build_parser().parse_args()

    audio = Path(args.audio).expanduser()
    if not audio.is_file():
        print(f"[错误] 找不到音频文件：{audio}", file=sys.stderr)
        return 1

    out_txt = (Path(args.out_txt).expanduser() if args.out_txt
               else audio.with_name(audio.stem + "_转写.txt"))

    model = WhisperModel(args.model, device="cpu", compute_type="int8")
    segments, info = model.transcribe(str(audio), beam_size=5)

    print(f"Detected language: {info.language} (probability {info.language_probability:.2f})")
    print(f"Audio duration: {info.duration:.1f}s")
    print("-" * 60)

    lines = []
    for seg in segments:
        line = f"[{seg.start:07.2f} - {seg.end:07.2f}] {seg.text.strip()}"
        print(line)
        lines.append(line)

    out_txt.parent.mkdir(parents=True, exist_ok=True)
    out_txt.write_text("\n".join(lines), encoding="utf-8")
    print("-" * 60)
    print("Saved to:", out_txt)
    return 0


if __name__ == "__main__":
    sys.exit(main())
