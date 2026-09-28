# 音频转写桌面端（拖入即用）

把音频文件拖进窗口，自动用本地 Whisper 模型转写，并在音频所在目录产出 **TXT** 和 **Word** 文档。

---

## 1. 三步开始

| 步骤 | 操作 |
| --- | --- |
| ① 装依赖（只需一次） | 双击 `安装依赖.bat` |
| ② 启动 | 双击 `启动.bat` |
| ③ 用 | 把音频文件拖进窗口 —— 自动开始转写 |

> 首次转写要先加载模型，large-v3 约需 10–30 秒，之后同一模型会缓存复用，不再重复加载。

---

## 2. 界面说明

**左侧**：拖放区 + 文件队列。每个文件显示实时状态（等待中 / 加载模型 / 转写中 / 写出文档 / 已完成 / 失败）和进度条。可一次拖入多个文件，也可直接拖入文件夹（自动扫描其中的音频）。

**右侧设置**：

| 项 | 说明 | 默认 |
| --- | --- | --- |
| 识别模型 | `large-v3` 最准；`small` 最快（约快 5–8 倍） | large-v3 |
| 音频语言 | 自动检测，或手动指定中/英/日/韩/法/德/西/俄 | 自动检测 |
| 领域提示词 | 填人名、专业术语可提升准确率 | 空 |
| 输出格式 | TXT / Word，可同时勾选 | 两个都勾 |
| 附带时间戳 | 勾选后每段前面加 `[起 - 止]` 时间 | 不勾 |
| 输出位置 | 与音频同目录，或指定文件夹 | 与音频同目录 |
| 拖入后自动开始 | 勾选即为"拖入即用" | 勾选 |

**底部**：总进度条 + 状态栏；`开始转写` / `停止` / `清空` / `打开输出文件夹`。

所有设置会自动记住，下次启动沿用上次选择。

---

## 3. 输出文件

命名规则：`<音频文件名>_转写.txt` 和 `<音频文件名>_转写.docx`。

- **TXT**：开头是元信息（源文件、模型、识别语言与置信度、音频时长、生成时间），分隔线后是正文。
- **Word**：标题 + 元信息段落 + 正文；勾选"附带时间戳"时正文改为两列表格（时间 | 文本）。

示例：

```
Thought groups.mp3  →  Thought groups_转写.txt
                        Thought groups_转写.docx
```

---

## 4. 支持的音频 / 视频格式

`mp3` `wav` `m4a` `aac` `flac` `ogg` `opus` `wma` `mp4` `mkv` `mov` `avi` `webm` `ts` `m4v`

解码由 PyAV 内置的 FFmpeg 完成，**不需要**另外安装 FFmpeg。

---

## 5. 模型文件

程序自动按以下顺序查找模型目录，命中即用：

1. `desktop-app\models\<模型名>`
2. `13-音频转写工具\models\<模型名>` ← 当前使用
3. `E:\harness\13-音频转写工具\models\<模型名>`

若下拉框显示"（本地未找到）"，说明该模型还没下载，用项目根目录的 `download_model.ps1` 下载即可：

```powershell
# large-v3（约 2.9 GB）
powershell -ExecutionPolicy Bypass -File ".\download_model.ps1" -Repo "Systran/faster-whisper-large-v3" -Dest ".\models\large-v3" -Files "config.json,model.bin,preprocessor_config.json,tokenizer.json,vocabulary.json"

# small（约 460 MB）
powershell -ExecutionPolicy Bypass -File ".\download_model.ps1" -Repo "Systran/faster-whisper-small" -Dest ".\models\small" -Files "config.json,model.bin,tokenizer.json,vocabulary.txt"
```

---

## 6. 常见问题

**转写很慢？**
全程 CPU 推理，速度取决于音频长度和模型大小。约 5 分钟音频：large-v3 需 5–8 分钟，small 约 1 分钟。想快速出稿就换 `small`。

**人名、术语识别错？**
在"领域提示词"里写正确的人名/术语。注意：只写**确定正确**的拼写，写错反而会污染输出。

**双击 `启动.bat` 没反应？**
改用 `调试启动.bat`，会保留控制台输出，能直接看到报错。

**想重跑某个文件？**
先"清空"，再重新拖入即可（已完成的文件不会被重复转写，除非清空）。

---

## 7. 命令行自检（可选）

无需界面即可验证转写链路是否正常：

```powershell
# 跑通转写与输出
python main.py --selftest "<音频文件>" small

# 只验证界面能正常构建
python main.py --guitest

# 界面 + 转写全链路
python main.py --guitest-job "<音频文件>"
```

---

## 8. 文件清单

```
desktop-app\
├── main.py             # 主程序（界面 + 转写 + 输出）
├── requirements.txt    # 依赖清单
├── 启动.bat            # 一键启动（无控制台）
├── 调试启动.bat        # 带控制台，便于排查
├── 安装依赖.bat        # 安装依赖
└── README.md           # 本说明
```

后续如需打成可安装的 exe，可在本目录用 PyInstaller + 安装程序打包（faster-whisper / ctranslate2 体积较大，需一并带上模型或用下载脚本首次运行再拉取）。