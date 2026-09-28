# -*- coding: utf-8 -*-
"""音频批量转写脚本（Whisper / faster-whisper，CPU + int8）

支持一次传入多个文件或文件夹，文件夹会自动递归扫描其中的音频；
模型只加载一次，整批复用。

用法示例：
    python transcribe_hq.py "D:\\音频"                       # 批量转写文件夹（递归）
    python transcribe_hq.py "a.mp3" "b.mp3"                  # 批量转写多个文件
    python transcribe_hq.py "D:\\音频" -m small              # 换模型
    python transcribe_hq.py "D:\\音频" -l en -p "人名,术语"   # 指定语言 + 领域提示词
    python transcribe_hq.py "D:\\音频" -o "D:\\输出"          # 统一输出到指定目录
    python transcribe_hq.py "D:\\音频" -f plain               # 只输出纯文本
    python transcribe_hq.py "D:\\音频" --skip-existing        # 已有输出则跳过（可断点续跑）
    python transcribe_hq.py "D:\\音频" --dry-run              # 只列出待处理文件，不转写
"""
import argparse
import os
import sys
import time
from datetime import datetime
from pathlib import Path

os.environ.setdefault("HF_ENDPOINT", "https://hf-mirror.com")
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

SCRIPT_DIR = Path(__file__).resolve().parent

AUDIO_EXTS = {
    ".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg", ".opus", ".wma",
    ".mp4", ".mkv", ".mov", ".avi", ".webm", ".ts", ".m4v",
}

FORMAT_CHOICES = ("txt", "plain", "srt")
SUFFIX = {"txt": "_转写.txt", "plain": "_转写_纯文本.txt", "srt": "_转写.srt"}

USAGE_TEXT = """
音频转写 - 拖入即用
------------------------------------------------------------
拖入用法：把「音频文件」或「整个文件夹」拖到桌面快捷方式
          「音频转写（拖入文件夹）」上，松开鼠标即自动开始转写。

命令行用法：
    python transcribe_hq.py <输入...> [-m 模型] [-l 语言] [-p 提示词]
                            [-o 输出目录] [-f 格式] [--no-recursive]
                            [--skip-existing] [--dry-run] [-v]

示例：
    python transcribe_hq.py "D:\\英语听力" -m large-v3 -o ".\\out"
    python transcribe_hq.py "a.mp3" "b.mp3" -f plain
    python transcribe_hq.py "D:\\英语听力" --dry-run

输入可写多个文件或文件夹；文件夹默认递归扫描其中的音频。
结果按音频名输出 3 种格式：
    xxx_转写.txt  /  xxx_转写_纯文本.txt  /  xxx_转写.srt
"""


def now() -> str:
    return time.strftime("%H:%M:%S")


def fmt_ts(t: float) -> str:
    h = int(t // 3600)
    m = int(t % 3600 // 60)
    s = t % 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


def fmt_srt_ts(t: float) -> str:
    return fmt_ts(t).replace(".", ",")


def fmt_hms(seconds: float) -> str:
    seconds = int(round(seconds))
    h, m, s = seconds // 3600, seconds % 3600 // 60, seconds % 60
    if h:
        return f"{h} 小时 {m} 分 {s} 秒"
    if m:
        return f"{m} 分 {s} 秒"
    return f"{s} 秒"


def collect_inputs(paths, recursive=True):
    """展开文件夹、过滤音频文件，排序去重。返回绝对路径。"""
    files, seen = [], set()
    for raw in paths:
        p = Path(raw).expanduser()
        if not p.exists():
            print(f"[警告] 路径不存在，已忽略：{p}")
            continue
        if p.is_dir():
            it = p.rglob("*") if recursive else p.glob("*")
            for f in sorted(it):
                if f.is_file() and f.suffix.lower() in AUDIO_EXTS:
                    f = f.resolve()
                    key = str(f).lower()
                    if key not in seen:
                        seen.add(key)
                        files.append(f)
        elif p.is_file():
            if p.suffix.lower() not in AUDIO_EXTS:
                print(f"[警告] 不是支持的音频格式，已忽略：{p.name}")
                continue
            f = p.resolve()
            key = str(f).lower()
            if key not in seen:
                seen.add(key)
                files.append(f)
    return sorted(files, key=lambda x: str(x).lower())


def resolve_model(spec: str):
    """spec 可以是本地模型目录，也可以是模型名（在脚本同级 models/ 下查找）。"""
    p = Path(spec).expanduser()
    if p.is_dir():
        return str(p.resolve()), p.resolve().name
    cand = SCRIPT_DIR / "models" / spec
    if cand.is_dir():
        return str(cand.resolve()), spec
    print(f"[警告] 未找到本地模型 {spec}，将按 HuggingFace 仓库名尝试在线加载。")
    return spec, spec


def build_output_bases(files, out_dir):
    """为每个音频确定输出目录。

    不指定 -o 时输出到各音频所在目录；指定 -o 时统一输出，
    若不同子目录下存在同名音频，则用其父目录名建子目录消歧，避免互相覆盖。
    """
    if not out_dir:
        return {f: f.parent for f in files}

    base = Path(out_dir).expanduser().resolve()
    by_stem = {}
    for f in files:
        by_stem.setdefault(f.stem.lower(), []).append(f)

    bases = {}
    for group in by_stem.values():
        if len(group) == 1:
            bases[group[0]] = base
            continue
        used = set()
        for f in group:
            tag = f.parent.name or "root"
            if tag.lower() in used:
                tag = f"{tag}_{abs(hash(str(f.parent))) % 10000:04d}"
            used.add(tag.lower())
            bases[f] = base / tag
    return bases


def output_paths(audio: Path, base_dir: Path, formats):
    return {fmt: base_dir / f"{audio.stem}{SUFFIX[fmt]}" for fmt in formats}


def render(fmt, header, seg_lines, plain_lines, segs):
    if fmt == "txt":
        return "\n".join(header) + "\n" + "-" * 60 + "\n" + "\n".join(seg_lines) + "\n"
    if fmt == "plain":
        return "\n".join(plain_lines) + "\n"
    srt = []
    for i, (s, e, t) in enumerate(segs, 1):
        srt += [str(i), f"{fmt_srt_ts(s)} --> {fmt_srt_ts(e)}", t, ""]
    return "\n".join(srt)


def write_outputs(audio, model_label, segs, info, paths):
    """逐个格式写出；单个格式失败不影响其余格式。返回 (已写出, 失败清单)。"""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dur = fmt_ts(info.duration) if info.duration else "未知"
    lang = (f"{info.language}（置信度 {info.language_probability:.2f}）"
            if info.language else "未知")
    header = [
        "# 音频转写结果",
        f"源文件：{audio.name}",
        f"模型：{model_label}",
        f"识别语言：{lang}",
        f"音频时长：{dur}",
        f"生成时间：{stamp}",
    ]
    seg_lines = [f"[{fmt_ts(s)} - {fmt_ts(e)}] {t}" for s, e, t in segs]
    plain_lines = [t for _, _, t in segs]

    written, failed = [], []
    for fmt, path in paths.items():
        try:
            path.write_text(render(fmt, header, seg_lines, plain_lines, segs),
                            encoding="utf-8")
            written.append(path)
        except OSError as exc:
            failed.append((fmt, path, exc))
    return written, failed


def transcribe_one(model, audio: Path, model_label, lang, prompt, base_dir,
                   formats, skip_existing, verbose):
    """转写单个文件，返回结果字典。"""
    paths = output_paths(audio, base_dir, formats)
    res = {"audio": audio, "outputs": [], "write_failed": [], "segments": 0,
           "duration": 0.0, "language": None, "elapsed": 0.0, "error": None,
           "skipped": False}

    if skip_existing and paths and all(p.exists() for p in paths.values()):
        res["skipped"] = True
        return res

    for p in paths.values():
        p.parent.mkdir(parents=True, exist_ok=True)

    t0 = time.time()
    segments, info = model.transcribe(
        str(audio),
        language=lang,
        initial_prompt=prompt or None,
        beam_size=5,
        patience=1.0,
        vad_filter=True,
        vad_parameters=dict(min_silence_duration_ms=300),
        word_timestamps=True,
        condition_on_previous_text=False,
        temperature=[0.0, 0.2, 0.4, 0.6, 0.8, 1.0],
    )

    segs = []
    duration = info.duration or 0.0
    marks = [0.25, 0.5, 0.75]
    hit = 0
    for seg in segments:
        segs.append((seg.start, seg.end, seg.text.strip()))
        if verbose:
            print(f"    [{fmt_ts(seg.start)} - {fmt_ts(seg.end)}] {seg.text.strip()}",
                  flush=True)
        if duration > 0:
            frac = seg.end / duration
            while hit < len(marks) and frac >= marks[hit]:
                print(f"    进度 {int(marks[hit] * 100)}%", flush=True)
                hit += 1

    written, failed = write_outputs(audio, model_label, segs, info, paths)

    res.update({
        "outputs": [str(p) for p in written],
        "write_failed": failed,
        "segments": len(segs),
        "duration": duration,
        "language": info.language,
        "elapsed": time.time() - t0,
    })
    return res


def build_parser():
    ap = argparse.ArgumentParser(
        description="音频批量转写（支持文件与文件夹）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("inputs", nargs="*", help="音频文件或文件夹，可多个")
    ap.add_argument("-m", "--model", default="large-v3",
                    help="模型目录或模型名（在脚本同级 models/ 下查找），默认 large-v3")
    ap.add_argument("-l", "--lang", default=None,
                    help="语言代码，如 en / zh；不填则自动识别")
    ap.add_argument("-p", "--prompt", default=None,
                    help="领域提示词（人名、术语），可提升专有名词准确率")
    ap.add_argument("-o", "--outdir", default=None,
                    help="输出目录；不填则输出到各音频所在目录")
    ap.add_argument("-f", "--formats", default="txt,plain,srt",
                    help="输出格式，逗号分隔，可选 txt,plain,srt，默认全出")
    ap.add_argument("--no-recursive", action="store_true",
                    help="不递归子文件夹")
    ap.add_argument("--skip-existing", action="store_true",
                    help="已有所需输出则跳过（便于断点续跑）")
    ap.add_argument("--dry-run", action="store_true",
                    help="只列出将要处理的文件，不实际转写")
    ap.add_argument("-v", "--verbose", action="store_true",
                    help="打印逐段文本")
    return ap


def main():
    args = build_parser().parse_args()

    if not args.inputs:
        print(USAGE_TEXT)
        return 2

    formats = [f.strip() for f in args.formats.split(",") if f.strip()]
    bad = [f for f in formats if f not in FORMAT_CHOICES]
    if bad:
        print(f"[错误] 不支持的输出格式：{bad}；可选 {list(FORMAT_CHOICES)}")
        return 2
    if not formats:
        print("[错误] 至少选择一种输出格式。")
        return 2

    files = collect_inputs(args.inputs, recursive=not args.no_recursive)
    if not files:
        print("[错误] 没有找到任何音频文件。")
        return 1

    model_path, model_label = resolve_model(args.model)
    bases = build_output_bases(files, args.outdir)

    print("=" * 66)
    print(f"模型      ：{model_label}  ({model_path})")
    print(f"语言      ：{args.lang or '自动识别'}")
    print(f"提示词    ：{args.prompt or '（无）'}")
    print(f"输出格式  ：{', '.join(formats)}")
    print(f"输出目录  ：{args.outdir or '与各音频同目录'}")
    print(f"待处理    ：{len(files)} 个文件")
    print("=" * 66)
    for i, f in enumerate(files, 1):
        print(f"  {i:>3}. {f}")
    print("=" * 66)

    if args.dry_run:
        print(f"[dry-run] 共 {len(files)} 个文件，未执行转写。")
        return 0

    print(f"[{now()}] 加载模型…（首次较慢，整批只加载一次）", flush=True)
    from faster_whisper import WhisperModel
    model = WhisperModel(model_path, device="cpu", compute_type="int8",
                         cpu_threads=os.cpu_count() or 4)
    print(f"[{now()}] 模型就绪，开始转写。", flush=True)

    ok, skipped, failed, partial = [], [], [], []
    audio_total = 0.0
    wall0 = time.time()

    for i, audio in enumerate(files, 1):
        print(f"\n[{now()}] ({i}/{len(files)}) {audio.name}", flush=True)
        try:
            r = transcribe_one(model, audio, model_label, args.lang, args.prompt,
                               bases[audio], formats, args.skip_existing, args.verbose)
        except Exception as exc:
            failed.append((audio, f"{type(exc).__name__}: {exc}"))
            print(f"    [失败] {type(exc).__name__}: {exc}", flush=True)
            continue

        if r["skipped"]:
            skipped.append(audio)
            print("    [跳过] 输出已存在", flush=True)
            continue

        if not r["outputs"]:
            failed.append((audio, "所有输出格式均写出失败"))
            print("    [失败] 所有输出格式均写出失败", flush=True)
            continue

        ok.append(r)
        audio_total += r["duration"]
        rtf = (r["elapsed"] / r["duration"]) if r["duration"] else 0.0
        print(f"    [完成] {r['segments']} 段 · 时长 {fmt_ts(r['duration'])} · "
              f"语言 {r['language']} · 用时 {r['elapsed']:.1f}s · 实时率 {rtf:.2f}×",
              flush=True)
        for p in r["outputs"]:
            print(f"           → {p}", flush=True)
        if r["write_failed"]:
            partial.append(r)
            for fmt, path, exc in r["write_failed"]:
                print(f"    [写出失败] {fmt}：{path} — {exc}", flush=True)

    wall = time.time() - wall0
    print("\n" + "=" * 66)
    print("汇总")
    print(f"  成功 {len(ok)} 个 ｜ 跳过 {len(skipped)} 个 ｜ 失败 {len(failed)} 个 "
          f"（共 {len(files)} 个）")
    if partial:
        print(f"  另有 {len(partial)} 个文件的部分格式写出失败（转写内容本身已成功，见上）")
    if ok:
        line = f"  音频总时长 {fmt_hms(audio_total)} ｜ 总用时 {fmt_hms(wall)}"
        if audio_total:
            line += f" ｜ 平均实时率 {wall / audio_total:.2f}×"
        print(line)
    if failed:
        print("  失败清单：")
        for audio, msg in failed:
            print(f"    - {audio.name}：{msg}")
    print("=" * 66)
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())