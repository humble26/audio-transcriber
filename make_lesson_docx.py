# -*- coding: utf-8 -*-
"""生成《Thought Groups》逐句翻译 Word 文档，并把原音频一并放到目标文件夹。

英文取自 Whisper large-v3 转写原文（未改写）；中文为逐句译文。
时间范围 = 该句所在转写分段的时间区间（据 transcription_large-v3.txt）。
"""
import argparse
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# 由命令行指定（见文件末尾的 main）；本模块的 build() 通过全局变量读取。
AUDIO = Path()      # --audio 源音频，会被复制到输出目录
OUT_DIR = Path()    # -o 输出目录
DOCX_NAME = "Thought Groups 逐句翻译.docx"

# (时间, 英文原句, 中文翻译)
SENTENCES = [
    ("00:00–00:12",
     "I'm Avi Arditi with Rosanne Skirbel and this week on WordMaster, English teacher Lita Baker in Los Angeles talks about improving English pronunciation by understanding the idea of thought groups.",
     "我是 Avi Arditi，和 Rosanne Skirbel 一起主持。本周的《WordMaster》节目中，洛杉矶的英语教师 Lita Baker 要谈的是：如何通过理解「意群」（thought groups）这一概念来改善英语发音。"),
    ("00:12–00:25",
     "Thought groups are something we don't even think about as native speakers of English.",
     "意群这种东西，我们作为英语母语者根本不会去留意它。"),
    ("00:12–00:25",
     "It's a way to break long sentences into shorter pieces separated by slight pauses to help listeners organize the meaning.",
     "它是一种把长句切分成若干较短片段、并用轻微停顿隔开的方法，目的是帮助听者理清意思。"),
    ("00:25–00:29",
     "But English learners need help to develop this skill when they study pronunciation.",
     "但英语学习者在学发音时，需要有人帮他们培养这项技能。"),
    ("00:30–00:40",
     "Lita says over the last 20 years, many teachers of English have come to focus not just on vowels and consonants, but also on stress and intonation.",
     "Lita 说，过去 20 年里，许多英语教师已经不再只关注元音和辅音，也开始关注重音和语调。"),
    ("00:40–00:46",
     "So we're talking about the way that the voice moves up and down and where we pause and things of that sort.",
     "所以我们谈的是声音起伏的方式、我们在哪里停顿，以及诸如此类的事情。"),
    ("00:46–00:51",
     "Take a sentence like, I took the milk from the table and I put it in the refrigerator.",
     "比如这样一个句子：“I took the milk from the table and I put it in the refrigerator.”"),
    ("00:51–00:59",
     "This is not right.",
     "这样念是不对的。"),
    ("00:51–00:59",
     "I took the milk from the table and I put it in the refrigerator.",
     "I took the milk from the table and I put it in the refrigerator.（每个词都读得一样重、一样清楚，像机器人念稿）"),
    ("00:51–00:59",
     "Nobody talks like that, right?",
     "没人这么说话，对吧？"),
    ("00:59–01:00",
     "You sound like a robot.",
     "你听起来像个机器人。"),
    ("01:00–01:19",
     "That's right, but that's not how we speak English.",
     "没错，可英语不是这么说的。"),
    ("01:00–01:19",
     "What we do is that the voice moves up and down, and there is also an alternation between syllables that are stressed and pronounced clearly and syllables that are unstressed and therefore are reduced and are spoken very quickly.",
     "我们实际的做法是：声音有高有低；而且重读、读得清楚的音节，与不重读、因而被弱化、说得很快的音节，会交替出现。"),
    ("01:19–01:25",
     "So, I took the milk becomes, I took the milk, pa-pa-pa-pa, okay?",
     "于是 “I took the milk” 就变成了 “I took the milk, pa-pa-pa-pa”，明白吗？"),
    ("01:25–01:39",
     "So, within each thought group, you will also find that there are these variations in pitch with the voice moving up and down, and then syllables that are pronounced more clearly, syllables that are reduced and pronounced unclearly.",
     "所以，在每一个意群内部，你还会发现音高有这样的变化——声音上下起伏；接着是读得比较清楚的音节，以及被弱化、读得含糊的音节。"),
    ("01:39–01:48",
     "So, you get this effect of, I took the milk, pa-pa-pa-pa, from the table, pa-pa-pa-pa, and I put it, da-da-da-da, in the refrigerator.",
     "这样就产生了这样的效果：I took the milk, pa-pa-pa-pa, from the table, pa-pa-pa-pa, and I put it, da-da-da-da, in the refrigerator."),
    ("01:49–01:50",
     "Ba-ba-ba-ba-ba-ba.",
     "Ba-ba-ba-ba-ba-ba.（拟声，模拟快速弱读的效果）"),
    ("01:50–01:51",
     "You've got a hit there.",
     "（此句语义不通，疑为听误，见文末「存疑片段」第 1 条。）"),
    ("01:52–01:53",
     "Right. Yeah, yeah.",
     "对。嗯，嗯。"),
    ("01:53–02:01",
     "Well, it's funny you should say that because one of the easiest ways to learn about thought groups is to listen to popular music.",
     "你这么说挺有意思，因为学习意群最简单的方法之一，就是听流行音乐。"),
    ("02:01–02:06",
     "And it happens that my daughter is absolutely crazy about the Beatles.",
     "恰好我女儿特别迷披头士乐队。"),
    ("02:07–02:09",
     "And she plays the guitar.",
     "她还弹吉他。"),
    ("02:09–02:11",
     "So yesterday she was singing Can't Buy Me Love.",
     "所以昨天她在唱《Can't Buy Me Love》。"),
    ("02:11–02:13",
     "Can't buy me love",
     "Can't buy me love（节目播放的歌曲片段，识别不可靠，见「存疑片段」第 2 条。）"),
    ("02:13–02:35",
     "Money can't buy",
     "Money can't buy（同上，歌曲片段，识别不可靠。）"),
    ("02:36–02:39",
     "First of all, can't buy me love.",
     "首先，“can't buy me love”。"),
    ("02:36–02:39",
     "That's a thought group right there.",
     "这里本身就是一个意群。"),
    ("02:40–02:42",
     "I'll buy you a diamond ring, my friend.",
     "I'll buy you a diamond ring, my friend.（朋友，我会给你买一枚钻戒。）"),
    ("02:42–02:46",
     "So, I'll buy you a diamond ring, my friend.",
     "那么，“I'll buy you a diamond ring, my friend.”"),
    ("02:46–02:48",
     "That's three thought groups right there.",
     "这里正好是三个意群。"),
    ("02:48–02:51",
     "What about for those that speak English as a foreign language?",
     "那么，把英语当作外语来学的人呢？"),
    ("02:52–02:55",
     "Are there some rules, or do they have to learn by doing?",
     "有没有什么规则，还是只能边做边学？"),
    ("02:56–02:57",
     "Well, I can't give you any rules.",
     "嗯，我给不出什么规则。"),
    ("02:57–02:59",
     "There are no rules, but there are guidelines.",
     "没有规则，但有一些指导原则。"),
    ("03:00–03:02",
     "Generally speaking, the pauses occur.",
     "一般来说，停顿会出现。"),
    ("03:02–03:18",
     "occur, they sort of correspond to grammatical units, such as phrases and clauses, and things like the complete subject of a sentence.",
     "（原文此处疑有重复与漏词，见「存疑片段」第 3 条。）它们大体上与语法单位相对应，比如短语、从句，以及句子的完整主语这类成分。"),
    ("03:12–03:18",
     "So if you have a sentence like, a big black cat sat on a tall white fence.",
     "所以，如果你有这样一个句子：“a big black cat sat on a tall white fence.”"),
    ("03:18–03:25",
     "So the subject there is a big black cat, and that's a thought group.",
     "那么其中的主语是 “a big black cat”，这就是一个意群。"),
    ("03:25–03:31",
     "A big black cat sat on a tall white fence.",
     "A big black cat sat on a tall white fence."),
    ("03:31–03:38",
     "On a tall white fence is also a thought group, and that's a prepositional phrase.",
     "而 “On a tall white fence” 也是一个意群，它是一个介词短语。"),
    ("03:31–03:38",
     "Now, pop music isn't the only way to learn this.",
     "不过，流行音乐并不是学习它的唯一途径。"),
    ("03:38–03:43",
     "A great way to learn this, I'm going to put in a plug here for the Voice of America, is to go to the special English broadcasts and look at the transcripts and then listen to the announcers.",
     "有个很好的方法——这里我要为美国之音（Voice of America）做个宣传——就是去听「特别英语」（Special English）广播，先看文字稿，再听播音员是怎么读的。"),
    ("03:43–03:56",
     "Because on special English, the language is slowed down, it's a wonderful way for learners to pick up on the way sentences are broken down into thought groups.",
     "因为在特别英语里，语速被放慢了；对学习者来说，这是体会句子如何被切分成意群的绝佳方式。"),
    ("03:56–04:02",
     "Another way is to use a videocassette recorder and tape any television program and do something called tracking.",
     "另一个办法是用录像机录下任何电视节目，做一种叫做「跟读」（tracking）的练习。"),
    ("04:02–04:10",
     "You tape a segment of a show and then you play it back.",
     "你录下一段节目，然后把它回放。"),
    ("04:10–04:16",
     "And what you try to do is to imitate what they're saying, just one beat behind them.",
     "你要做的就是模仿他们说的话，只比他们慢一拍。"),
    ("04:16–04:22",
     "And incidentally, it doesn't have to be done with television.",
     "顺便说一句，这不一定非得用电视来做。"),
    ("04:22–04:23",
     "It can be done with radio as well.",
     "用广播也可以。"),
    ("04:23–04:25",
     "Anywhere there's sound going on in English.",
     "只要有英语声音的地方都行。"),
    ("04:26",
     "That's right.",
     "没错。"),
    ("04:27–04:32",
     "Lita Baker teaches at the American Language Center at the University of California at Los Angeles.",
     "Lita Baker 在加州大学洛杉矶分校（UCLA）的美国语言中心任教。"),
    ("04:33–04:36",
     "She also writes and edits textbooks for English learners.",
     "她还为英语学习者编写和编辑教材。"),
    ("04:36–04:43",
     "And by the way, those special English programs she mentioned are all available online at voaspecialenglish.com.",
     "顺便提一下，她提到的那些「特别英语」节目，都可以在 voaspecialenglish.com 上在线收听。"),
    ("04:44–04:49",
     "You can also find a link from our website, voanews.com slash wordmaster.",
     "你也可以从我们的网站 voanews.com/wordmaster 找到链接。"),
    ("04:49–04:54",
     "And our email address is word at voanews.com.",
     "我们的邮箱地址是 word@voanews.com。"),
    ("04:54–04:57",
     "With Avi Arditi, I'm Roseanne Skirble.",
     "与 Avi Arditi 一起，我是 Roseanne Skirble。（节目结束语）"),
]

NOTES = [
    ("01:50", "语义可疑", "“You've got a hit there.” 语义不通，两个模型都未能给出合理解读，疑为听误，置信度低。"),
    ("02:11–02:35", "背景音乐 / 歌词", "此处为节目播放的披头士《Can't Buy Me Love》片段，音乐与人声重叠，识别结果不可靠，不宜直接引用。"),
    ("03:00–03:12", "分段重复 / 疑似漏词", "large-v3 输出 “…the pauses occur. occur, they sort of correspond to…”，small 模型则识别为 “However, they sort of correspond to…”。原文很可能含 however。"),
    ("03:12–03:31", "断句位置", "“a big black cat sat on a tall white fence.” 被切成两段，文字无误，仅为断句位置问题。"),
    ("全篇", "人名拼写", "转写原文中人名拼写不统一：large-v3 作 “Avi Arditi” / “Roseanne Skirble”，开头处又作 “Rosanne Skirbel”。官方拼写为 Avi Arditti、Rosanne Skirble。"),
]

META = [
    "源音频：Thought groups.mp3（2.27 MB）",
    "音频时长：4:58（298.0 秒）",
    "节目来源：VOA《WordMaster》——通过「意群」（thought groups）改善英语发音",
    "说话人：Avi Arditti（主持）、Rosanne Skirble（主持）、Lita Baker（嘉宾，UCLA 美国语言中心教师）",
    "英文来源：Whisper large-v3 转写原文（CPU / int8，未改写）",
    "整理日期：2026-09-27",
]

CJK = "微软雅黑"


def set_cjk(run, name=CJK):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:fill"): hex_color})
    tcPr.append(shd)


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = CJK
    normal.font.size = Pt(10.5)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CJK)

    title = doc.add_heading("", level=0)
    tr = title.add_run("《Thought Groups》逐句翻译")
    set_cjk(tr, CJK)

    sub = doc.add_paragraph()
    sr = sub.add_run("VOA WordMaster · 意群与英语发音")
    sr.font.size = Pt(11)
    sr.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    set_cjk(sr)

    for line in META:
        p = doc.add_paragraph()
        r = p.add_run(line)
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0x37, 0x41, 0x51)
        set_cjk(r)

    tip = doc.add_paragraph()
    tr2 = tip.add_run("说明：英文为语音识别转写原文（未改写）；中文为逐句译文。"
                      "「时间」为该句所在转写分段的时间区间，便于对照音频跟读。")
    tr2.font.size = Pt(9.5)
    tr2.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    set_cjk(tr2)

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(1.1), Cm(2.3), Cm(6.3), Cm(6.3)]
    heads = ["序号", "时间", "英文原句", "中文翻译"]
    for i, (cell, text) in enumerate(zip(table.rows[0].cells, heads)):
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(10)
        set_cjk(r)
        shade(cell, "E8EEF9")
        cell.width = widths[i]

    for n, (t, en, zh) in enumerate(SENTENCES, 1):
        cells = table.add_row().cells
        for i, val in enumerate([str(n), t, en, zh]):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(10)
            set_cjk(r)
            if i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
            if i == 2:
                r.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
            cells[i].width = widths[i]

    doc.add_paragraph()
    h = doc.add_heading("", level=1)
    hr = h.add_run("存疑片段与说明")
    set_cjk(hr)

    nt = doc.add_table(rows=1, cols=3)
    nt.style = "Table Grid"
    nwidths = [Cm(2.4), Cm(3.0), Cm(10.6)]
    for i, (cell, text) in enumerate(zip(nt.rows[0].cells, ["时间", "类型", "说明"])):
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(10)
        set_cjk(r)
        shade(cell, "FDECEC")
        cell.width = nwidths[i]

    for t, kind, desc in NOTES:
        cells = nt.add_row().cells
        for i, val in enumerate([t, kind, desc]):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            set_cjk(r)
            cells[i].width = nwidths[i]

    docx_path = OUT_DIR / DOCX_NAME
    doc.save(str(docx_path))
    print("已生成 Word:", docx_path, docx_path.stat().st_size, "字节")

    audio_dst = OUT_DIR / "Thought groups.mp3"
    audio_dst.write_bytes(AUDIO.read_bytes())
    print("已复制音频:", audio_dst, audio_dst.stat().st_size, "字节")


def main():
    global AUDIO, OUT_DIR

    ap = argparse.ArgumentParser(
        description="生成《Thought Groups》逐句翻译 Word 文档，并把原音频一并复制到输出目录")
    ap.add_argument("--audio", required=True,
                    help="源音频文件路径（会被复制到输出目录）")
    ap.add_argument("-o", "--outdir", default="out",
                    help="输出目录，默认 ./out")
    args = ap.parse_args()

    AUDIO = Path(args.audio).expanduser()
    if not AUDIO.is_file():
        raise SystemExit(f"[错误] 找不到音频文件：{AUDIO}")
    OUT_DIR = Path(args.outdir).expanduser()

    build()


if __name__ == "__main__":
    main()