# 音频转写工具

把音频丢进去，拿回来文字。**全程本地推理**，音频不出本机 —— 用 Whisper（faster-whisper 实现）在 CPU 上跑，适合会议录音、课程音频、访谈、播客这类需要文字稿的场景。

![桌面端](desktop-app/README.md)

```
音频文件 / 整个文件夹
        ↓  faster-whisper（本地，CPU + int8）
   带时间戳的转写结果
        ↓
  TXT  /  纯文本  /  SRT 字幕  /  Word 文档
```

## 三种用法

### 1. 桌面端（拖入即用，推荐）

```bash
cd desktop-app
安装依赖.bat      # 只需一次：PySide6 + faster-whisper + python-docx
启动.bat
```

把音频文件或整个文件夹拖进窗口即自动开始转写，实时显示每个文件的进度与状态，完成后在音频所在目录产出 **TXT 与 Word**。详见 [`desktop-app/README.md`](desktop-app/README.md)。

### 2. 命令行批量转写

```bash
pip install faster-whisper

python transcribe_hq.py "D:\录音"                     # 递归转写整个文件夹
python transcribe_hq.py a.mp3 b.mp3                   # 多个文件
python transcribe_hq.py "D:\录音" -m small            # 换更快的模型
python transcribe_hq.py "D:\录音" -l en -p "人名,术语" # 指定语言 + 领域提示词提准
python transcribe_hq.py "D:\录音" -o "D:\输出"        # 统一输出目录
python transcribe_hq.py "D:\录音" --skip-existing     # 断点续跑
python transcribe_hq.py "D:\录音" --dry-run           # 只列待处理文件
```

结果按音频名输出三种格式：`xxx_转写.txt`（带时间戳）/ `xxx_转写_纯文本.txt` / `xxx_转写.srt`。

模型只加载一次、整批复用；单个文件失败不会中断整批，最后打印成功 / 跳过 / 失败清单。

### 3. 最小示例

```bash
python transcribe_audio.py "录音.mp3"            # 输出到同目录 录音_转写.txt
python transcribe_audio.py "录音.mp3" out.txt -m small
```

## 模型

模型**不入库**（`large-v3` 单文件 3 GB 以上），用脚本下载到 `models/` 下：

```powershell
.\download_model.ps1 -Repo Systran/faster-whisper-large-v3 -Dest .\models\large-v3
.\download_model.ps1 -Repo Systran/faster-whisper-small   -Dest .\models\small
```

脚本默认走 `hf-mirror.com` 镜像。也可以用模型名（如 `-m small`）让 faster-whisper 自行下载。

| 模型 | 速度 | 准确率 | 适用 |
|---|---|---|---|
| `small` | 快 | 一般 | 快速过一遍、口音标准的录音 |
| `large-v3` | 慢（CPU 上约 0.3–0.6× 实时率） | 好 | 正式文字稿、多人对话、专业术语 |

## 实现细节

- **VAD 过滤**：`vad_filter=True`，静音段不送模型，明显提速
- **退化兜底**：`temperature` 从 0.0 递增到 1.0，配合 `condition_on_previous_text=False` 抑制重复循环
- **输出稳健**：多种格式逐个写出，某一格式失败不影响其余；`--skip-existing` 便于长音频断点续跑
- **文件去重**：递归扫描时按绝对路径小写去重，排序稳定

## 目录结构

```
├── desktop-app/                  # 桌面端（PySide6 图形界面）
├── transcribe_hq.py              # 批量转写 CLI（完整参数）
├── transcribe_audio.py           # 单文件最小示例
├── make_lesson_docx.py           # 由转写稿生成「逐句翻译」Word 文档（教学材料）
├── make_icon.py / 转写.ico       # 图标生成
├── download_model.ps1            # 模型下载（hf-mirror 镜像）
├── 拖入转写.bat                  # 拖放启动器（可做成桌面快捷方式）
├── TRANSCRIPTION_RUNBOOK.md      # 使用手册
└── TRANSCRIPTION_EXPERIENCE_GUIDE.md
```

## 说明

- 转写产物、音频源文件与模型均已加入 `.gitignore`，不会随仓库分发
- 长音频在纯 CPU 上较慢，可用 `-m small` 或改用 GPU 版 faster-whisper
