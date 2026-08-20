---
name: pptx-animate
description: 使用本机 Microsoft PowerPoint 的 COM 自动为 PPTX 添加真实动画和页面切换效果。
---

# PowerPoint 动画与转场

当用户要求为 `.pptx` 添加**动态效果、入场动画、页面转场**，或希望演示文稿更具高级感、艺术感、发布会感时，使用此技能。

此技能仅使用 Windows 上已安装的 **Microsoft PowerPoint** 与 `pywin32` COM 自动化；不生成 HTML，不依赖浏览器动画。动画和切换会被写入 PPTX 的 OOXML，可在 PowerPoint 中继续编辑和播放。

## 与其他 PPT 工具的分工

- 用 `pptx` MCP 或 `pptx` skill 创建、编辑、排版、检查 PPTX。
- 用 `pptx-design-styles` skill 选择视觉风格、色彩、字体与版式。
- 在视觉与内容定稿后，用本技能添加动画与切换。
- 若要检查视觉输出，优先通过 PowerPoint COM 导出 PNG；不要依赖 LibreOffice。

动画会修改源文件。对重要成品，加 `--backup` 保留副本，或明确获用户同意后再直接处理。

## 快速开始

脚本位于本 skill 目录下的 `scripts/animate.py`。在 Git Bash 中运行：

```bash
ANIMATOR="$HOME/.claude/skills/pptx-animate/scripts/animate.py"
python "$ANIMATOR" "deck.pptx" --preset refined
```

可用操作：

```bash
# 商务、提案、评审：克制、稳定的高级感
python "$ANIMATOR" "deck.pptx" --preset refined

# 发布会、展示、创意提案：更鲜明的动效
python "$ANIMATOR" "deck.pptx" --preset artistic

# 只添加转场，不逐项播放正文元素
python "$ANIMATOR" "deck.pptx" --preset refined --no-body-anim

# 改动前先存一份带时间戳的备份
python "$ANIMATOR" "deck.pptx" --preset refined --backup

# 使用 JSON 对每页、每个对象精确控制
python "$ANIMATOR" "deck.pptx" --spec animation-spec.json

# 清除整份演示文稿的动画与切换
python "$ANIMATOR" "deck.pptx" --clear

# 查看可用动画、切换和触发方式（不需要文件）
python "$ANIMATOR" --list
```

## 内置预设

### `refined`：克制高级感

适用于商务汇报、管理层评审、投融资材料、咨询提案。

- 封面：`dissolve`（溶解）切换
- 正文：`morph`（平滑切换）
- 章节页：`reveal`（揭示）切换
- 正文对象：依次 `fade`（淡入）
- 动画时长：0.5 秒；页面切换：0.9 秒

默认会跳过细线、强调条等装饰件，只让实质内容依次出现；封面和简洁章节页仅设置切换效果，避免喧宾夺主。

### `artistic`：创意展示感

适用于新品发布、品牌展示、创意提案、活动演讲。

- 封面：`ripple`（涟漪）切换
- 正文：`morph`（平滑切换）
- 章节页：`prism`（棱镜）切换
- 正文对象：依次 `float`（浮入）
- 动画时长：0.7 秒；页面切换：1.2 秒

此预设仍以清晰阅读为前提，但视觉存在感更强。不要在严肃财报、数据密集型汇报中默认使用。

## JSON 精细控制

规格文件必须使用 UTF-8 编码。页码和形状序号均从 **1** 开始。`shape` 既可以是形状序号，也可以是 PowerPoint 中的形状名称。

```json
{
  "default_transition": "morph",
  "default_duration": 0.9,
  "slides": [
    {
      "index": 1,
      "clear": true,
      "transition": "dissolve",
      "duration": 1.0
    },
    {
      "index": 2,
      "transition": "morph",
      "duration": 0.9,
      "animations": [
        {
          "shape": 2,
          "effect": "fade",
          "duration": 0.6,
          "trigger": "after"
        },
        {
          "shape": "正文图表",
          "effect": "float",
          "duration": 0.7,
          "trigger": "after",
          "delay": 0.2
        }
      ]
    }
  ]
}
```

字段说明：

| 字段 | 含义 |
| --- | --- |
| `default_transition` | 对所有页面先设置的默认切换；页级 `transition` 可覆盖它。 |
| `default_duration` | 默认切换时长（秒）。 |
| `slides` | 逐页设置数组。 |
| `index` | 页码，1 起算。 |
| `clear` | `true` 时先移除该页原有动画与切换。 |
| `transition` / `duration` | 页面切换效果与时长。 |
| `animations` | 该页的对象动画数组。 |
| `shape` | 形状序号或形状名称。 |
| `effect` / `duration` | 入场动画与时长。 |
| `trigger` | `click`（单击）、`after`（上一动画后）、`with`（与上一动画同时）。 |
| `delay` | 触发后的额外等待时间（秒）。 |
| `exit` | `true` 时将动画标为退出动画；仅在确有退场叙事需要时使用。 |

## 可用效果

### 入场动画

`appear`、`fly`、`blinds`、`fade`、`float`、`grow`、`wipe`、`zoom`

优先级建议：

1. 默认使用 `fade`；最稳妥，也最不易显得廉价。
2. 创意展示可使用 `float`、`wipe` 或 `zoom`。
3. `fly`、`grow`、`blinds` 只在效果能服务于叙事时使用，不应铺满整份演示文稿。

### 页面切换

`none`、`dissolve`、`circle`、`diamond`、`ripple`、`gallery`、`flythrough`、`reveal`、`ferris`、`shred`、`prism`、`pan`、`morph`、`morph_words`、`morph_chars`

`morph` 是高质量转场的首选，但要获得最佳效果，相邻两页应保留可被 PowerPoint 识别为同一对象的元素（例如复制上一页再调整位置/尺寸）。`morph_words` 和 `morph_chars` 更适合以文字变化为主体的连续页面。

## 使用规则

1. 动画必须服务于信息层级和叙事，不要让每个对象都使用不同效果。
2. 数据表、复杂图表和页脚通常不应逐项飞入；可使用 `--no-body-anim`，或通过 JSON 只动画关键标题、重点数据和图表。
3. 若导入图片、图表、复杂组合对象后形状序号不稳定，先在 PowerPoint 的“选择窗格”中为对象命名，再用名称引用。
4. 动画与切换建议在内容、版式、对象名称稳定后最后添加；后续大量重排可能削弱 Morph 的连续性。
5. 运行失败时，先确认 Windows 已安装 Microsoft PowerPoint，并确保运行 Python 环境可导入 `win32com.client`（`pywin32`）。
6. `morph`、`morph_words`、`morph_chars` 需要 PowerPoint 2019 或 Microsoft 365；旧版本会报错并附带版本提示。

## 对 PowerPoint 进程的处理

脚本不会打断用户正在进行的工作，因此**无需**要求用户先关闭 PowerPoint：

- PowerPoint 已在运行时接管现有实例，结束时不 Quit，用户手上的其他文稿不受影响。
- 目标文件已在 PowerPoint 中打开时，直接修改并保存那个窗口，不关闭它，用户未保存的编辑不会丢失。
- 只有脚本自己启动的 PowerPoint 实例才由脚本退出。
