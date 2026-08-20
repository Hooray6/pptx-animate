"""用 PowerPoint COM 给 .pptx 添加动画与切换效果。

设计目标:
    python-pptx / pptxgenjs 都无法写入动画,只有 PowerPoint COM 能做。
    本脚本把「预设方案」和「逐页自定义」两种用法收在一个入口里。

用法:
    # 套用内置预设(推荐)
    python animate.py deck.pptx --preset refined
    python animate.py deck.pptx --preset artistic

    # 用 JSON 精细控制每一页
    python animate.py deck.pptx --spec spec.json

    # 清除全部动画与切换
    python animate.py deck.pptx --clear

    # 列出可用效果名(不需要文件)
    python animate.py --list

JSON 规格示例:
    {
      "default_transition": "morph",
      "slides": [
        {"index": 1, "transition": "dissolve", "duration": 1.0},
        {"index": 2, "transition": "morph",
         "animations": [
           {"shape": 2, "effect": "fade",  "duration": 0.6, "trigger": "after"},
           {"shape": 3, "effect": "float", "duration": 0.6, "trigger": "after", "delay": 0.2}
         ]}
      ]
    }

    index 为 1 起算的页码;shape 为 1 起算的形状序号,也可写形状名字符串。
    省略 animations 时该页只设切换效果。
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import time
from contextlib import contextmanager

# ---------------------------------------------------------------------------
# 效果常量
#
# 下列取值均在本机 PowerPoint 上实测通过,并回读 slideN.xml 校验过写入结果。
# 动画:AddEffect 的 ID 在 <=30 区间与 OOXML presetID 一致。
# 切换:数值为 PpEntryEffect,注释里的名字来自实测读出的 XML 元素名。
# ---------------------------------------------------------------------------

ANIMATIONS = {
    "appear": 1,   # 出现,无过渡
    "fly": 2,      # 飞入
    "blinds": 3,   # 百叶窗
    "fade": 10,    # 淡入 —— 最克制,首选
    "float": 11,   # 浮入
    "grow": 12,    # 缩放并旋转
    "wipe": 22,    # 擦除
    "zoom": 23,    # 缩放
}

# 不支持设置时长的效果:AddEffect 后写 Timing.Duration 会被 COM 拒绝
_NO_DURATION = {"appear"}

TRANSITIONS = {
    "none": 0,
    "dissolve": 1537,     # p:dissolve
    "circle": 3845,       # p:circle
    "diamond": 3846,      # p:diamond
    "ripple": 3870,       # p14:ripple
    "gallery": 3880,      # p14:gallery
    "flythrough": 3890,   # p14:flythrough
    "reveal": 3895,       # p14:reveal
    "ferris": 3899,       # p14:ferris
    "shred": 3910,        # p14:shred
    "prism": 3920,        # p14:prism
    "pan": 3930,          # p14:pan
    "morph": 3954,        # p159:morph —— 按对象补间,高级感核心
    "morph_words": 3955,  # p159:morph 按词
    "morph_chars": 3956,  # p159:morph 按字符
}

# Morph 需要 PowerPoint 2019 / Microsoft 365,旧版本写入会失败
_MORPH = {"morph", "morph_words", "morph_chars"}

# 触发方式(MsoAnimTriggerType)
TRIGGERS = {
    "click": 1,     # 单击时
    "after": 2,     # 上一动画之后,自动连播
    "with": 3,      # 与上一动画同时
}

msoTrue = -1  # MsoTriState.msoTrue,Effect.Exit 这类属性要的是它而非 Python True

# ---------------------------------------------------------------------------
# 两套预设
#
# refined:  以 Morph + 淡入为主,不抢内容。商务汇报、提案、评审。
# artistic: 用上涟漪/棱柱/摩天轮等强效果。发布会、展示、创意提案。
# ---------------------------------------------------------------------------

PRESETS = {
    "refined": {
        "cover_transition": "dissolve",
        "body_transition": "morph",
        "section_transition": "reveal",
        "anim_effect": "fade",
        "anim_duration": 0.5,
        "anim_delay": 0.15,
        "transition_duration": 0.9,
    },
    "artistic": {
        "cover_transition": "ripple",
        "body_transition": "morph",
        "section_transition": "prism",
        "anim_effect": "float",
        "anim_duration": 0.7,
        "anim_delay": 0.2,
        "transition_duration": 1.2,
    },
}


class AnimateError(RuntimeError):
    """面向用户的错误:main() 只打印消息,不打印堆栈。"""


# ---------------------------------------------------------------------------
# COM 会话
# ---------------------------------------------------------------------------


def _require_win32():
    """导入 pywin32,失败时给出可操作的提示而非裸 ImportError。"""
    if sys.platform != "win32":
        raise AnimateError(
            "本脚本依赖 Windows 版 Microsoft PowerPoint 的 COM 接口,"
            f"无法在 {sys.platform} 上运行。"
        )
    try:
        import win32com.client as w32
    except ImportError as exc:
        raise AnimateError(
            "缺少 pywin32,请先安装: pip install pywin32"
        ) from exc
    return w32


def _same_file(a: str, b: str) -> bool:
    """比较两个路径是否指向同一个文件。Windows 下大小写不敏感。"""
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except Exception:
        return False


@contextmanager
def _session(path: str):
    """打开演示文稿并保证收尾干净。

    两个关键点,都是为了不破坏用户正在进行的工作:

    1. 只有本脚本启动的 PowerPoint 才由本脚本 Quit。若用户本来就开着
       PowerPoint,直接 Quit 会连带关掉他手上的其他文稿。
    2. 目标文件已在 PowerPoint 中打开时,复用那个 Presentation 对象,
       结束时不关闭它 —— 否则会丢掉用户未保存的编辑。
    """
    w32 = _require_win32()

    started_app = False
    try:
        # 已在运行则接管现有实例
        app = w32.GetActiveObject("PowerPoint.Application")
    except Exception:
        try:
            app = w32.Dispatch("PowerPoint.Application")
        except Exception as exc:
            raise AnimateError(
                "无法启动 PowerPoint。请确认 Windows 上已安装 Microsoft PowerPoint。"
                f"（{type(exc).__name__}: {exc}）"
            ) from exc
        started_app = True

    target = os.path.abspath(path)
    pres = None
    opened_here = False
    try:
        # 复用已打开的同名文稿
        for i in range(1, app.Presentations.Count + 1):
            cand = app.Presentations(i)
            if _same_file(cand.FullName, target):
                pres = cand
                print("提示: 该文件已在 PowerPoint 中打开,将直接修改并保存这个窗口。")
                break

        if pres is None:
            try:
                pres = app.Presentations.Open(target, WithWindow=False)
            except Exception as exc:
                raise AnimateError(
                    f"打开失败: {path}。文件可能损坏、被其他程序占用,"
                    f"或不是有效的 .pptx。（{type(exc).__name__}: {exc}）"
                ) from exc
            opened_here = True

        if pres.Slides.Count == 0:
            raise AnimateError("演示文稿没有幻灯片,无事可做。")

        yield pres

        try:
            pres.Save()
        except Exception as exc:
            raise AnimateError(
                f"保存失败,改动未写入。文件可能为只读或位于受保护目录。"
                f"（{type(exc).__name__}: {exc}）"
            ) from exc
    finally:
        if pres is not None and opened_here:
            try:
                pres.Close()
            except Exception:
                pass
        if started_app:
            try:
                app.Quit()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# 单页操作
# ---------------------------------------------------------------------------


def _resolve_shape(slide, ref):
    """把 shape 引用(序号或名称)解析成 Shape 对象。"""
    if isinstance(ref, bool):
        raise AnimateError(f"shape 不能是布尔值: {ref!r}")
    if isinstance(ref, int):
        if not 1 <= ref <= slide.Shapes.Count:
            raise AnimateError(
                f"第 {slide.SlideIndex} 页没有第 {ref} 个形状(该页共 {slide.Shapes.Count} 个)"
            )
        return slide.Shapes(ref)
    for i in range(1, slide.Shapes.Count + 1):
        if slide.Shapes(i).Name == ref:
            return slide.Shapes(i)
    names = ", ".join(slide.Shapes(i).Name for i in range(1, slide.Shapes.Count + 1))
    raise AnimateError(
        f"第 {slide.SlideIndex} 页找不到名为 {ref!r} 的形状;该页现有: {names}"
    )


def _clear_slide(slide) -> None:
    """清空一页的动画序列与切换效果。"""
    seq = slide.TimeLine.MainSequence
    # 逆序删除,避免边删边移位
    for i in range(seq.Count, 0, -1):
        seq(i).Delete()
    slide.SlideShowTransition.EntryEffect = TRANSITIONS["none"]


def _set_transition(slide, name: str, duration: float) -> None:
    if name not in TRANSITIONS:
        raise AnimateError(
            f"未知切换效果 {name!r};可用: {', '.join(sorted(TRANSITIONS))}"
        )
    try:
        slide.SlideShowTransition.EntryEffect = TRANSITIONS[name]
    except Exception as exc:
        hint = "（Morph 需要 PowerPoint 2019 或 Microsoft 365）" if name in _MORPH else ""
        raise AnimateError(
            f"第 {slide.SlideIndex} 页设置切换 {name!r} 失败{hint}: "
            f"{type(exc).__name__}: {exc}"
        ) from exc
    if name != "none":
        slide.SlideShowTransition.Duration = float(duration)
        # 保留单击换页,不强制自动播放
        slide.SlideShowTransition.AdvanceOnClick = True


def _add_animation(slide, spec: dict) -> None:
    """给一个形状加一条动画。"""
    effect = spec.get("effect", "fade")
    if effect not in ANIMATIONS:
        raise AnimateError(
            f"未知动画 {effect!r};可用: {', '.join(sorted(ANIMATIONS))}"
        )
    trigger = spec.get("trigger", "after")
    if trigger not in TRIGGERS:
        raise AnimateError(
            f"未知触发方式 {trigger!r};可用: {', '.join(sorted(TRIGGERS))}"
        )

    shape = _resolve_shape(slide, spec.get("shape", 1))
    eff = slide.TimeLine.MainSequence.AddEffect(shape, ANIMATIONS[effect])

    if effect not in _NO_DURATION:
        try:
            eff.Timing.Duration = float(spec.get("duration", 0.5))
        except Exception:
            # 个别效果不接受自定义时长,保留 PowerPoint 默认值即可
            pass
    eff.Timing.TriggerType = TRIGGERS[trigger]
    delay = float(spec.get("delay") or 0)
    if delay > 0:
        eff.Timing.TriggerDelayTime = delay
    if spec.get("exit"):
        eff.Exit = msoTrue


# 判定装饰件的尺寸阈值(单位:磅)。分隔线、强调条通常只有几磅厚。
_HAIRLINE_PT = 10


def _animatable_shapes(slide) -> list:
    """挑出值得加动画的形状序号。

    排除分隔线与强调条这类装饰件(任一边极薄),它们逐个淡入只会制造噪音。
    """
    picks = []
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        try:
            if s.Height < _HAIRLINE_PT or s.Width < _HAIRLINE_PT:
                continue
        except Exception:
            # 取不到尺寸的形状按可动画处理,宁可多动不可漏
            pass
        picks.append(i)
    return picks


# ---------------------------------------------------------------------------
# 三种入口
# ---------------------------------------------------------------------------


def apply_preset(path: str, preset_name: str, animate_body: bool = True) -> dict:
    """按预设批量套用切换与动画。"""
    if preset_name not in PRESETS:
        raise AnimateError(
            f"未知预设 {preset_name!r};可用: {', '.join(PRESETS)}"
        )
    cfg = PRESETS[preset_name]
    stats = {"slides": 0, "animations": 0, "preset": preset_name}

    with _session(path) as pres:
        for idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(idx)
            _clear_slide(slide)

            # 每页只算一次:形状枚举是 COM 调用,重复算既慢也无必要。
            # 判定封面/章节页不能只看形状总数 —— 章节页(两条满宽装饰条 + 标题)
            # 和只有一个正文块的内容页都可能是 3 个形状。区别在于去掉装饰件后,
            # 是否还剩标题之外的实体内容。
            movable = _animatable_shapes(slide)
            is_cover = len(movable) <= 1

            # 首页用封面切换;形状少的页当作章节页
            if idx == 1:
                tname = cfg["cover_transition"]
            elif is_cover:
                tname = cfg["section_transition"]
            else:
                tname = cfg["body_transition"]
            _set_transition(slide, tname, cfg["transition_duration"])

            # 正文页逐个元素渐次出现;封面/章节页保持干净
            if animate_body and not is_cover:
                for n, shape_idx in enumerate(movable):
                    _add_animation(slide, {
                        "shape": shape_idx,
                        "effect": cfg["anim_effect"],
                        "duration": cfg["anim_duration"],
                        # 第一个元素跟随换页,其余顺序自动连播
                        "trigger": "after",
                        "delay": 0 if n == 0 else cfg["anim_delay"],
                    })
                    stats["animations"] += 1
            stats["slides"] += 1

    return stats


def apply_spec(path: str, spec: dict) -> dict:
    """按 JSON 规格逐页设置。"""
    if not isinstance(spec, dict):
        raise AnimateError("规格文件的顶层必须是 JSON 对象。")
    entries = spec.get("slides", [])
    if not isinstance(entries, list):
        raise AnimateError("规格中的 slides 必须是数组。")

    stats = {"slides": 0, "animations": 0}
    with _session(path) as pres:
        total = pres.Slides.Count

        default_tr = spec.get("default_transition")
        if default_tr:
            for idx in range(1, total + 1):
                _set_transition(pres.Slides(idx), default_tr,
                                float(spec.get("default_duration", 0.9)))

        touched = set()
        for entry in entries:
            if not isinstance(entry, dict) or "index" not in entry:
                raise AnimateError("slides 的每一项都必须是带 index 字段的对象。")
            try:
                idx = int(entry["index"])
            except (TypeError, ValueError):
                raise AnimateError(f"页码不是整数: {entry['index']!r}") from None
            if not 1 <= idx <= total:
                raise AnimateError(f"页码 {idx} 超出范围(共 {total} 页)")
            slide = pres.Slides(idx)

            if entry.get("clear"):
                _clear_slide(slide)
            if entry.get("transition"):
                _set_transition(slide, entry["transition"],
                                float(entry.get("duration", 0.9)))
            for anim in entry.get("animations", []):
                if not isinstance(anim, dict):
                    raise AnimateError(f"第 {idx} 页的 animations 每一项都必须是对象。")
                _add_animation(slide, anim)
                stats["animations"] += 1
            touched.add(idx)

        # 统计去重后的页数:同一页写两条规格不该算成两页
        stats["slides"] = len(touched) if not default_tr else total

    return stats


def clear_all(path: str) -> dict:
    """清除整份文件的动画与切换。"""
    stats = {"slides": 0}
    with _session(path) as pres:
        for idx in range(1, pres.Slides.Count + 1):
            _clear_slide(pres.Slides(idx))
            stats["slides"] += 1
    return stats


def make_backup(path: str) -> str:
    """在原文件旁复制一份带时间戳的备份,返回备份路径。"""
    root, ext = os.path.splitext(path)
    dest = f"{root}.backup-{time.strftime('%Y%m%d-%H%M%S')}{ext}"
    shutil.copy2(path, dest)
    return dest


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------


def _force_utf8_output() -> None:
    """把 stdout/stderr 切到 UTF-8。

    Windows 的 Python 默认跟随控制台代码页(中文系统上是 cp936),
    但 Git Bash 与现代终端都按 UTF-8 解码,不改会让中文提示变乱码。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main(argv: list[str] | None = None) -> int:
    _force_utf8_output()
    ap = argparse.ArgumentParser(
        description="给 PPTX 添加动画与切换(PowerPoint COM)",
    )
    ap.add_argument("pptx", nargs="?", help=".pptx 文件路径")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--preset", choices=sorted(PRESETS), help="套用内置预设")
    g.add_argument("--spec", help="JSON 规格文件路径")
    g.add_argument("--clear", action="store_true", help="清除全部动画与切换")
    g.add_argument("--list", action="store_true", help="列出可用效果名,不需要文件")
    ap.add_argument("--no-body-anim", action="store_true",
                    help="配合 --preset:只设页面切换,不给元素加动画")
    ap.add_argument("--backup", action="store_true",
                    help="改动前在原文件旁存一份带时间戳的备份")
    args = ap.parse_args(argv)

    if args.list:
        print("动画效果:", ", ".join(sorted(ANIMATIONS)))
        print("切换效果:", ", ".join(sorted(TRANSITIONS)))
        print("触发方式:", ", ".join(sorted(TRIGGERS)))
        print("预设方案:", ", ".join(sorted(PRESETS)))
        return 0

    if not args.pptx:
        ap.error("除 --list 外都需要给出 .pptx 文件路径")
    if not os.path.isfile(args.pptx):
        print(f"找不到文件: {args.pptx}", file=sys.stderr)
        return 1

    try:
        if args.backup:
            print(f"已备份到: {make_backup(args.pptx)}")

        if args.preset:
            st = apply_preset(args.pptx, args.preset,
                              animate_body=not args.no_body_anim)
            print(f"预设 {st['preset']} 已套用: {st['slides']} 页, {st['animations']} 条动画")
        elif args.spec:
            try:
                with open(args.spec, encoding="utf-8") as f:
                    spec = json.load(f)
            except FileNotFoundError:
                print(f"找不到规格文件: {args.spec}", file=sys.stderr)
                return 1
            except json.JSONDecodeError as exc:
                print(f"规格文件不是合法 JSON(第 {exc.lineno} 行): {exc.msg}",
                      file=sys.stderr)
                return 1
            st = apply_spec(args.pptx, spec)
            print(f"规格已套用: {st['slides']} 页, {st['animations']} 条动画")
        else:
            st = clear_all(args.pptx)
            print(f"已清除 {st['slides']} 页的动画与切换")
    except AnimateError as exc:
        print(f"失败: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("已中断。", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"失败: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
