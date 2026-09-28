# -*- coding: utf-8 -*-
"""音频转写桌面端 —— 拖入即用，自动产出 TXT / Word

依赖：PySide6、faster-whisper、python-docx
用法：
    python main.py                        # 启动图形界面
    python main.py --selftest <音频> [模型]  # 无界面自检（跑通转写与输出）
"""
import os
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

os.environ.setdefault("HF_ENDPOINT", "https://hf-mirror.com")
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

APP_DIR = Path(__file__).resolve().parent
APP_NAME = "音频转写"
ORG_NAME = "HarnessTools"

AUDIO_EXTS = {
    ".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg", ".opus", ".wma",
    ".mp4", ".mkv", ".mov", ".avi", ".webm", ".ts", ".m4v",
}

MODEL_DEFS = [
    ("large-v3", "large-v3 · 最准（约 2.9 GB）", "large-v3"),
    ("small", "small · 最快（约 460 MB）", "small"),
]

LANG_DEFS = [
    ("自动检测", None), ("中文", "zh"), ("英语", "en"), ("日语", "ja"),
    ("韩语", "ko"), ("法语", "fr"), ("德语", "de"), ("西班牙语", "es"), ("俄语", "ru"),
]


def find_models_root() -> Path:
    candidates = [
        APP_DIR / "models",
        APP_DIR.parent / "models",
        Path(r"E:\harness\13-音频转写工具\models"),
    ]
    for c in candidates:
        if c.is_dir():
            return c
    return candidates[1]


MODELS_ROOT = find_models_root()
_MODEL_CACHE = {}


def resolve_model_dir(key: str) -> Path:
    return MODELS_ROOT / key


def fmt_ts(t: float) -> str:
    h = int(t // 3600)
    m = int(t % 3600 // 60)
    s = t % 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


def collect_audio_files(paths) -> list:
    """展开文件夹、过滤出音频文件，保持顺序并去重。"""
    found, seen = [], set()
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            for f in sorted(p.rglob("*")):
                if f.is_file() and f.suffix.lower() in AUDIO_EXTS:
                    if str(f) not in seen:
                        seen.add(str(f))
                        found.append(str(f))
        elif p.is_file() and p.suffix.lower() in AUDIO_EXTS:
            if str(p) not in seen:
                seen.add(str(p))
                found.append(str(p))
    return found


def write_txt(path: str, header_lines, body_lines) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(header_lines))
        f.write("\n" + "-" * 60 + "\n")
        f.write("\n".join(body_lines))
        f.write("\n")


def write_docx(path: str, title: str, meta_lines, segs, timestamps: bool) -> None:
    from docx import Document
    from docx.shared import Pt
    from docx.oxml.ns import qn

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    doc.add_heading(title, level=0)
    for line in meta_lines:
        doc.add_paragraph(line)

    if timestamps:
        table = doc.add_table(rows=1, cols=2)
        table.style = "Light Grid Accent 1"
        head = table.rows[0].cells
        head[0].text = "时间"
        head[1].text = "文本"
        for s, e, text in segs:
            cells = table.add_row().cells
            cells[0].text = f"{fmt_ts(s)} → {fmt_ts(e)}"
            cells[1].text = text
    else:
        for _, _, text in segs:
            doc.add_paragraph(text)

    doc.save(path)


def build_outputs(audio: str, model_key: str, segs, info, out_dir: str,
                  want_txt: bool, want_docx: bool, timestamps: bool) -> list:
    """写出 TXT / Word，返回生成的文件路径列表。"""
    stem = Path(audio).stem
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dur = fmt_ts(info.duration) if info.duration else "未知"
    lang = f"{info.language}（置信度 {info.language_probability:.2f}）" if info.language else "未知"

    meta = [
        f"源文件：{Path(audio).name}",
        f"模型：{model_key}",
        f"识别语言：{lang}",
        f"音频时长：{dur}",
        f"生成时间：{stamp}",
    ]

    body = [f"[{fmt_ts(s)} - {fmt_ts(e)}] {t}" for s, e, t in segs] if timestamps \
        else [t for _, _, t in segs]

    outputs = []
    if want_txt:
        p = os.path.join(out_dir, f"{stem}_转写.txt")
        write_txt(p, [f"# 音频转写结果"] + meta, body)
        outputs.append(p)
    if want_docx:
        p = os.path.join(out_dir, f"{stem}_转写.docx")
        write_docx(p, "音频转写结果", meta, segs, timestamps)
        outputs.append(p)
    return outputs


# ---------------------------------------------------------------- 后台线程
try:
    from PySide6.QtCore import QThread, Signal
except ImportError:  # 自检模式下无 GUI 依赖也能跑
    QThread, Signal = object, None


class TranscribeWorker(QThread):
    if Signal is not None:
        sig_item_status = Signal(int, str)
        sig_item_progress = Signal(int, float)
        sig_item_done = Signal(int, dict)
        sig_item_failed = Signal(int, str)
        sig_model_loading = Signal(str)
        sig_all_done = Signal()

    def __init__(self, jobs, model_key, model_dir, lang, prompt,
                 want_txt, want_docx, timestamps, out_dir, parent=None):
        super().__init__(parent)
        self.jobs = jobs
        self.model_key = model_key
        self.model_dir = model_dir
        self.lang = lang
        self.prompt = prompt
        self.want_txt = want_txt
        self.want_docx = want_docx
        self.timestamps = timestamps
        self.out_dir = out_dir
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def _load_model(self):
        if self.model_key not in _MODEL_CACHE:
            self.sig_model_loading.emit(self.model_key)
            from faster_whisper import WhisperModel
            _MODEL_CACHE[self.model_key] = WhisperModel(
                self.model_dir, device="cpu", compute_type="int8",
                cpu_threads=os.cpu_count() or 4,
            )
        return _MODEL_CACHE[self.model_key]

    def _run_one(self, idx, audio):
        self.sig_item_status.emit(idx, "加载模型…")
        model = self._load_model()
        if self._cancel:
            self.sig_item_status.emit(idx, "已取消")
            return

        self.sig_item_status.emit(idx, "转写中…")
        segments, info = model.transcribe(
            audio,
            language=self.lang,
            initial_prompt=self.prompt or None,
            beam_size=5,
            patience=1.0,
            vad_filter=True,
            vad_parameters=dict(min_silence_duration_ms=300),
            word_timestamps=True,
            condition_on_previous_text=False,
            temperature=[0.0, 0.2, 0.4, 0.6, 0.8, 1.0],
        )

        segs = []
        duration = info.duration or 0
        for seg in segments:
            if self._cancel:
                self.sig_item_status.emit(idx, "已取消")
                return
            segs.append((seg.start, seg.end, seg.text.strip()))
            if duration > 0:
                self.sig_item_progress.emit(idx, min(seg.end / duration, 1.0))

        out_dir = self.out_dir or str(Path(audio).parent)
        os.makedirs(out_dir, exist_ok=True)
        self.sig_item_status.emit(idx, "写出文档…")
        outputs = build_outputs(audio, self.model_key, segs, info, out_dir,
                                self.want_txt, self.want_docx, self.timestamps)
        self.sig_item_progress.emit(idx, 1.0)
        self.sig_item_done.emit(idx, {
            "outputs": outputs,
            "segments": len(segs),
            "duration": info.duration,
            "language": info.language,
            "text": " ".join(t for _, _, t in segs),
        })

    def run(self):
        try:
            for idx, audio in enumerate(self.jobs):
                if self._cancel:
                    break
                try:
                    self._run_one(idx, audio)
                except Exception as exc:
                    traceback.print_exc()
                    self.sig_item_failed.emit(idx, str(exc))
        finally:
            self.sig_all_done.emit()


# ---------------------------------------------------------------- 界面
def main_gui():
    from PySide6.QtCore import Qt, QSettings, QTimer
    from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
        QListWidget, QListWidgetItem, QComboBox, QCheckBox, QPushButton,
        QProgressBar, QLineEdit, QFileDialog, QFrame,
        QMessageBox, QRadioButton, QButtonGroup,
    )

    def make_icon():
        pm = QPixmap(64, 64)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#2563eb"))
        p.drawRoundedRect(2, 2, 60, 60, 15, 15)
        p.setPen(QPen(QColor("white"), 3, Qt.SolidLine, Qt.RoundCap))
        for i, h in enumerate([12, 26, 38, 22, 14]):
            x = 15 + i * 8.5
            p.drawLine(int(x), 32 - h // 2, int(x), 32 + h // 2)
        p.end()
        return QIcon(pm)

    class FileRow(QWidget):
        def __init__(self, audio):
            super().__init__()
            self.audio = audio
            lay = QVBoxLayout(self)
            lay.setContentsMargins(12, 9, 12, 9)
            lay.setSpacing(7)
            top = QHBoxLayout()
            top.setSpacing(8)
            name = QLabel(Path(audio).name)
            name.setStyleSheet("font-weight:600;color:#111827;")
            name.setToolTip(audio)
            self.status = QLabel("等待中")
            self.status.setStyleSheet("color:#6b7280;")
            top.addWidget(name, 1)
            top.addWidget(self.status, 0)
            lay.addLayout(top)
            self.bar = QProgressBar()
            self.bar.setRange(0, 100)
            self.bar.setValue(0)
            self.bar.setTextVisible(False)
            self.bar.setFixedHeight(5)
            lay.addWidget(self.bar)

        def set_status(self, text, color="#6b7280"):
            self.status.setText(text)
            self.status.setStyleSheet(f"color:{color};")

        def set_progress(self, frac):
            self.bar.setValue(int(max(0.0, min(1.0, frac)) * 100))

    class DropBar(QLabel):
        def __init__(self):
            super().__init__("把音频文件拖到这里\n支持多选，也可直接拖入文件夹")
            self.setObjectName("DropBar")
            self.setAlignment(Qt.AlignCenter)
            self.setFixedHeight(84)
            self.setAttribute(Qt.WA_TransparentForMouseEvents, False)

    class MainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle(f"{APP_NAME} · 拖入即用")
            self.setWindowIcon(make_icon())
            self.resize(940, 600)
            self.setAcceptDrops(True)

            self.rows = []
            self.job_rows = []
            self.worker = None
            self.settings = QSettings(ORG_NAME, APP_NAME)

            root = QWidget()
            root.setObjectName("Root")
            self.setCentralWidget(root)
            outer = QVBoxLayout(root)
            outer.setContentsMargins(20, 18, 20, 18)
            outer.setSpacing(14)

            head = QVBoxLayout()
            head.setSpacing(2)
            t = QLabel(APP_NAME)
            t.setObjectName("Title")
            s = QLabel("拖入音频 → 自动转写 → 输出 TXT / Word 文档")
            s.setObjectName("Sub")
            head.addWidget(t)
            head.addWidget(s)
            outer.addLayout(head)

            body = QHBoxLayout()
            body.setSpacing(14)
            outer.addLayout(body, 1)

            # 左：拖放区 + 文件列表
            left = QVBoxLayout()
            left.setSpacing(10)
            self.dropbar = DropBar()
            left.addWidget(self.dropbar)

            self.list = QListWidget()
            self.list.setObjectName("FileList")
            self.list.setAcceptDrops(False)
            self.list.setDragDropMode(QListWidget.NoDragDrop)
            self.list.setSelectionMode(QListWidget.ExtendedSelection)
            left.addWidget(self.list, 1)

            self.empty_hint = QLabel("还没有文件。拖入音频后会自动排队等待转写。")
            self.empty_hint.setObjectName("Sub")
            self.empty_hint.setAlignment(Qt.AlignCenter)
            left.addWidget(self.empty_hint)
            body.addLayout(left, 1)

            # 右：设置卡片
            card = QFrame()
            card.setObjectName("Card")
            card.setFixedWidth(300)
            cl = QVBoxLayout(card)
            cl.setContentsMargins(16, 16, 16, 16)
            cl.setSpacing(11)

            def section(text):
                lb = QLabel(text)
                lb.setObjectName("Section")
                cl.addWidget(lb)

            section("识别模型")
            self.model_combo = QComboBox()
            for key, label, sub in MODEL_DEFS:
                d = resolve_model_dir(sub)
                item = label if d.is_dir() else f"{label}（本地未找到）"
                self.model_combo.addItem(item, key)
            idx = self.model_combo.findData(self.settings.value("model", "large-v3"))
            self.model_combo.setCurrentIndex(max(0, idx))
            cl.addWidget(self.model_combo)

            section("音频语言")
            self.lang_combo = QComboBox()
            for label, code in LANG_DEFS:
                self.lang_combo.addItem(label, code)
            cl.addWidget(self.lang_combo)

            section("领域提示词（可选）")
            self.prompt_edit = QLineEdit()
            self.prompt_edit.setPlaceholderText("人名、术语，可提升准确率")
            self.prompt_edit.setText(self.settings.value("prompt", ""))
            cl.addWidget(self.prompt_edit)

            section("输出格式")
            self.chk_txt = QCheckBox("TXT 文本")
            self.chk_txt.setChecked(self.settings.value("txt", True, type=bool))
            self.chk_docx = QCheckBox("Word 文档")
            self.chk_docx.setChecked(self.settings.value("docx", True, type=bool))
            self.chk_ts = QCheckBox("附带时间戳")
            self.chk_ts.setChecked(self.settings.value("ts", False, type=bool))
            cl.addWidget(self.chk_txt)
            cl.addWidget(self.chk_docx)
            cl.addWidget(self.chk_ts)

            section("输出位置")
            self.rb_same = QRadioButton("与音频同目录")
            self.rb_custom = QRadioButton("指定文件夹")
            self.rb_same.setChecked(self.settings.value("same_dir", True, type=bool))
            self.rb_custom.setChecked(not self.rb_same.isChecked())
            grp = QButtonGroup(self)
            grp.addButton(self.rb_same)
            grp.addButton(self.rb_custom)
            cl.addWidget(self.rb_same)
            rowd = QHBoxLayout()
            rowd.setSpacing(6)
            rowd.addWidget(self.rb_custom)
            self.dir_edit = QLineEdit(self.settings.value("out_dir", ""))
            self.dir_edit.setPlaceholderText("默认与音频同目录")
            btn_dir = QPushButton("…")
            btn_dir.setObjectName("Ghost")
            btn_dir.setFixedWidth(34)
            btn_dir.clicked.connect(self.pick_dir)
            rowd.addWidget(self.dir_edit, 1)
            rowd.addWidget(btn_dir)
            cl.addLayout(rowd)

            self.chk_auto = QCheckBox("拖入后自动开始")
            self.chk_auto.setChecked(self.settings.value("auto", True, type=bool))
            cl.addSpacing(4)
            cl.addWidget(self.chk_auto)

            cl.addStretch(1)
            body.addWidget(card)

            # 底部：进度 + 按钮
            foot = QHBoxLayout()
            foot.setSpacing(12)
            pcol = QVBoxLayout()
            pcol.setSpacing(6)
            self.overall = QProgressBar()
            self.overall.setRange(0, 100)
            self.overall.setValue(0)
            self.overall.setTextVisible(False)
            self.overall.setFixedHeight(8)
            self.status = QLabel("就绪")
            self.status.setObjectName("Sub")
            pcol.addWidget(self.overall)
            pcol.addWidget(self.status)
            foot.addLayout(pcol, 1)

            self.btn_start = QPushButton("开始转写")
            self.btn_start.setObjectName("Primary")
            self.btn_start.clicked.connect(self.start_jobs)
            self.btn_stop = QPushButton("停止")
            self.btn_stop.setObjectName("Ghost")
            self.btn_stop.setEnabled(False)
            self.btn_stop.clicked.connect(self.stop_jobs)
            self.btn_clear = QPushButton("清空")
            self.btn_clear.setObjectName("Ghost")
            self.btn_clear.clicked.connect(self.clear_all)
            self.btn_open = QPushButton("打开输出文件夹")
            self.btn_open.setObjectName("Ghost")
            self.btn_open.clicked.connect(self.open_out_dir)
            for b in (self.btn_start, self.btn_stop, self.btn_clear, self.btn_open):
                foot.addWidget(b)
            outer.addLayout(foot)

            self.last_out_dir = self.settings.value("out_dir", "") or str(APP_DIR)
            self.refresh_empty()
            self.apply_style()

        # ---------------- 样式
        def apply_style(self):
            self.setStyleSheet("""
            #Root { background:#f3f4f6; }
            QLabel { color:#1f2937; }
            #Title { font-size:21px; font-weight:700; color:#0f172a; }
            #Sub { color:#6b7280; font-size:12px; }
            #Section { color:#374151; font-weight:600; font-size:12px; margin-top:4px; }
            #Card { background:#ffffff; border:1px solid #e5e7eb; border-radius:12px; }
            #DropBar { background:#f8fafc; border:2px dashed #cbd5e1; border-radius:12px;
                       color:#64748b; font-size:13px; }
            #FileList { background:transparent; border:none; outline:none; }
            #FileList::item { background:#ffffff; border:1px solid #e5e7eb;
                              border-radius:10px; margin:4px 2px; }
            #FileList::item:selected { border:1px solid #2563eb; }
            QComboBox, QLineEdit { background:#ffffff; border:1px solid #d1d5db;
                                    border-radius:8px; padding:6px 8px; }
            QComboBox:focus, QLineEdit:focus { border:1px solid #2563eb; }
            QComboBox::drop-down { border:none; width:22px; }
            QCheckBox, QRadioButton { spacing:7px; }
            QPushButton#Primary { background:#2563eb; color:#ffffff; border:none;
                                  border-radius:9px; padding:10px 20px; font-weight:600; }
            QPushButton#Primary:hover { background:#1d4ed8; }
            QPushButton#Primary:disabled { background:#bfdbfe; color:#eff6ff; }
            QPushButton#Ghost { background:#ffffff; color:#374151; border:1px solid #d1d5db;
                                border-radius:9px; padding:9px 14px; }
            QPushButton#Ghost:hover { background:#f9fafb; border:1px solid #9ca3af; }
            QPushButton#Ghost:disabled { color:#9ca3af; }
            QProgressBar { background:#e5e7eb; border:none; border-radius:4px; }
            QProgressBar::chunk { background:#2563eb; border-radius:4px; }
            """)

        # ---------------- 拖放
        def dragEnterEvent(self, e):
            if e.mimeData().hasUrls():
                e.acceptProposedAction()
                self.dropbar.setStyleSheet(
                    "background:#eff6ff;border:2px dashed #2563eb;border-radius:12px;"
                    "color:#1d4ed8;font-size:13px;")

        def dragLeaveEvent(self, e):
            self.apply_style()

        def dropEvent(self, e):
            paths = [u.toLocalFile() for u in e.mimeData().urls()]
            self.apply_style()
            files = collect_audio_files(paths)
            if not files:
                self.status.setText("没有识别到音频文件（支持 mp3 / wav / m4a / flac / mp4 等）")
                return
            self.add_files(files)
            e.acceptProposedAction()
            if self.chk_auto.isChecked() and not self.is_running():
                self.start_jobs()

        # ---------------- 列表
        def refresh_empty(self):
            has = self.list.count() > 0
            self.empty_hint.setVisible(not has)

        def add_files(self, files):
            added = 0
            existing = {r.audio for r in self.rows}
            for f in files:
                if f in existing:
                    continue
                item = QListWidgetItem(self.list)
                row = FileRow(f)
                item.setSizeHint(row.sizeHint())
                self.list.addItem(item)
                self.list.setItemWidget(item, row)
                self.rows.append(row)
                added += 1
            self.refresh_empty()
            if added:
                self.status.setText(f"已加入 {added} 个文件，共 {len(self.rows)} 个待转写")

        def clear_all(self):
            if self.is_running():
                return
            self.list.clear()
            self.rows = []
            self.job_rows = []
            self.overall.setValue(0)
            self.status.setText("就绪")
            self.refresh_empty()

        # ---------------- 设置
        def pick_dir(self):
            d = QFileDialog.getExistingDirectory(self, "选择输出文件夹", self.dir_edit.text() or str(APP_DIR))
            if d:
                self.dir_edit.setText(d)
                self.rb_custom.setChecked(True)

        def open_out_dir(self):
            d = self.last_out_dir or self.dir_edit.text() or str(APP_DIR)
            if os.path.isdir(d):
                os.startfile(d)
            else:
                QMessageBox.information(self, "提示", "还没有可打开的输出文件夹。")

        def is_running(self):
            return self.worker is not None and self.worker.isRunning()

        def save_settings(self):
            self.settings.setValue("model", self.model_combo.currentData())
            self.settings.setValue("prompt", self.prompt_edit.text())
            self.settings.setValue("txt", self.chk_txt.isChecked())
            self.settings.setValue("docx", self.chk_docx.isChecked())
            self.settings.setValue("ts", self.chk_ts.isChecked())
            self.settings.setValue("same_dir", self.rb_same.isChecked())
            self.settings.setValue("out_dir", self.dir_edit.text())
            self.settings.setValue("auto", self.chk_auto.isChecked())

        # ---------------- 任务
        def start_jobs(self):
            if self.is_running():
                return
            jobs = [r.audio for r in self.rows if r.status.text() in ("等待中", "失败", "已取消")]
            if not jobs:
                QMessageBox.information(self, "提示", "请先拖入音频文件。")
                return
            if not (self.chk_txt.isChecked() or self.chk_docx.isChecked()):
                QMessageBox.warning(self, "提示", "请至少勾选一种输出格式（TXT 或 Word）。")
                return

            key = self.model_combo.currentData()
            sub = dict((k, s) for k, _, s in MODEL_DEFS)[key]
            model_dir = resolve_model_dir(sub)
            if not model_dir.is_dir():
                QMessageBox.warning(self, "模型缺失",
                                    f"本地未找到模型目录：\n{model_dir}\n\n"
                                    "请先用 download_model.ps1 下载该模型。")
                return

            out_dir = "" if self.rb_same.isChecked() else self.dir_edit.text().strip()
            if out_dir and not os.path.isdir(out_dir):
                QMessageBox.warning(self, "路径无效", f"输出文件夹不存在：\n{out_dir}")
                return

            self.save_settings()
            self.last_out_dir = out_dir or str(Path(jobs[0]).parent)

            self.job_rows = [r for r in self.rows if r.audio in set(jobs)]
            for r in self.job_rows:
                r.set_status("等待中")
                r.set_progress(0)

            self.worker = TranscribeWorker(
                jobs, key, str(model_dir), self.lang_combo.currentData(),
                self.prompt_edit.text().strip(), self.chk_txt.isChecked(),
                self.chk_docx.isChecked(), self.chk_ts.isChecked(), out_dir,
            )
            self.worker.sig_model_loading.connect(self.on_model_loading)
            self.worker.sig_item_status.connect(self.on_item_status)
            self.worker.sig_item_progress.connect(self.on_item_progress)
            self.worker.sig_item_done.connect(self.on_item_done)
            self.worker.sig_item_failed.connect(self.on_item_failed)
            self.worker.sig_all_done.connect(self.on_all_done)

            self.btn_start.setEnabled(False)
            self.btn_stop.setEnabled(True)
            self.btn_clear.setEnabled(False)
            self.overall.setValue(0)
            self.worker.start()

        def stop_jobs(self):
            if self.is_running():
                self.worker.cancel()
                self.status.setText("正在停止…")
                self.btn_stop.setEnabled(False)

        def on_model_loading(self, key):
            self.status.setText(f"正在加载模型 {key}（首次较慢，请稍候）…")
            self.overall.setRange(0, 0)

        def on_item_status(self, idx, text):
            rows, total = self.job_rows, len(self.job_rows)
            if idx < total:
                color = {"转写中…": "#2563eb", "写出文档…": "#2563eb",
                         "已取消": "#b45309"}.get(text, "#6b7280")
                rows[idx].set_status(text, color)
                self.status.setText(f"[{idx + 1}/{total}] {Path(rows[idx].audio).name} · {text}")
            else:
                self.status.setText(text)

        def on_item_progress(self, idx, frac):
            rows, total = self.job_rows, len(self.job_rows)
            if idx < total:
                rows[idx].set_progress(frac)
            if total:
                if self.overall.maximum() == 0:
                    self.overall.setRange(0, 100)
                self.overall.setValue(int((idx + frac) / total * 100))

        def on_item_done(self, idx, res):
            if idx < len(self.job_rows):
                self.job_rows[idx].set_status("已完成", "#16a34a")
                self.job_rows[idx].set_progress(1.0)
                self.job_rows[idx].outputs = res["outputs"]
            if res["outputs"]:
                self.last_out_dir = str(Path(res["outputs"][0]).parent)

        def on_item_failed(self, idx, msg):
            if idx < len(self.job_rows):
                self.job_rows[idx].set_status("失败", "#dc2626")
                self.job_rows[idx].setToolTip(msg)
            self.status.setText(f"失败：{msg}")

        def on_all_done(self):
            self.btn_start.setEnabled(True)
            self.btn_stop.setEnabled(False)
            self.btn_clear.setEnabled(True)
            if self.overall.maximum() == 0:
                self.overall.setRange(0, 100)
            done = sum(1 for r in self.rows if r.status.text() == "已完成")
            failed = sum(1 for r in self.rows if r.status.text() == "失败")
            self.status.setText(f"全部结束：成功 {done} 个，失败 {failed} 个")
            if done:
                self.overall.setValue(100)

        def closeEvent(self, e):
            if self.is_running():
                self.worker.cancel()
                self.worker.wait(8000)
            self.save_settings()
            e.accept()

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    win = MainWindow()
    win.show()
    if "--guitest" in sys.argv:
        def check():
            print("GUI OK: 窗口构建成功", win.width(), "x", win.height())
            print("模型下拉:", [win.model_combo.itemText(i) for i in range(win.model_combo.count())])
            print("默认模型:", win.model_combo.currentData(), "| 语言:", win.lang_combo.currentText())
            print("输出: TXT=", win.chk_txt.isChecked(), "Word=", win.chk_docx.isChecked(),
                  "时间戳=", win.chk_ts.isChecked(), "自动开始=", win.chk_auto.isChecked())
            app.quit()
        QTimer.singleShot(1200, check)
    if "--guitest-job" in sys.argv:
        audio = sys.argv[sys.argv.index("--guitest-job") + 1]
        win.model_combo.setCurrentIndex(win.model_combo.findData("small"))
        win.chk_docx.setChecked(False)
        win.add_files([audio])
        QTimer.singleShot(300, win.start_jobs)

        def watch():
            if win.worker is not None and not win.is_running():
                print("GUI JOB DONE:", [(Path(r.audio).name, r.status.text()) for r in win.job_rows])
                print("状态栏:", win.status.text())
                app.quit()
            else:
                QTimer.singleShot(1000, watch)

        QTimer.singleShot(1500, watch)
    sys.exit(app.exec())


# ---------------------------------------------------------------- 自检
def run_selftest():
    from PySide6.QtCore import QCoreApplication

    args = [a for a in sys.argv[1:] if a != "--selftest"]
    if not args:
        print("用法: python main.py --selftest <音频文件> [模型key]")
        sys.exit(2)
    audio = args[0]
    key = args[1] if len(args) > 1 else "small"
    sub = dict((k, s) for k, _, s in MODEL_DEFS).get(key, key)
    model_dir = resolve_model_dir(sub)
    if not model_dir.is_dir():
        print(f"模型目录不存在: {model_dir}")
        sys.exit(3)

    app = QCoreApplication(sys.argv)
    worker = TranscribeWorker([audio], key, str(model_dir), None, "", True, True, False, "")
    result = {}

    def done(idx, res):
        result.update(res)
        print("DONE outputs:", res["outputs"], "segments:", res["segments"],
              "lang:", res["language"])

    worker.sig_item_status.connect(lambda i, t: print("  status:", t, flush=True))
    worker.sig_item_progress.connect(lambda i, f: None)
    worker.sig_item_done.connect(done)
    worker.sig_item_failed.connect(lambda i, m: print("FAILED:", m))
    worker.sig_all_done.connect(app.quit)
    worker.start()
    app.exec()
    worker.wait()
    for p in result.get("outputs", []):
        print("存在:", os.path.exists(p), p)
    sys.exit(0 if result.get("outputs") else 1)


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        run_selftest()
    else:
        main_gui()