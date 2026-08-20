# pptx-animate

给 `.pptx` 加**真实的**入场动画与页面切换效果，写进 OOXML，在 PowerPoint 里可以继续编辑和播放。

`python-pptx`、`pptxgenjs` 这类库都无法写入动画。这个 skill 走 Windows 上已安装的 Microsoft PowerPoint 的 COM 接口，因此产出的是原生动画，不是 HTML 模拟。

配合 [Claude Code](https://claude.com/claude-code) 作为 skill 使用，也可以当独立命令行脚本跑。

## 环境要求

| 项目 | 要求 |
| --- | --- |
| 操作系统 | Windows（COM 接口只在 Windows 上有） |
| PowerPoint | 已安装 Microsoft PowerPoint。Morph 切换需要 2019 或 Microsoft 365 |
| Python | 3.8+ |
| 依赖 | `pip install pywin32` |

## 安装

作为 Claude Code skill：

```bash
git clone https://github.com/<your-name>/pptx-animate.git ~/.claude/skills/pptx-animate
```

之后 Claude Code 会在你提到「给 PPT 加动画」时自动加载它。

只想当脚本用的话，clone 到任意位置，直接调用 `scripts/animate.py` 即可。

## 用法

```bash
A=~/.claude/skills/pptx-animate/scripts/animate.py

# 商务、提案、评审：克制稳定的高级感
python "$A" deck.pptx --preset refined

# 发布会、展示、创意提案：更鲜明的动效
python "$A" deck.pptx --preset artistic

# 只加页面切换，正文元素不逐项出现
python "$A" deck.pptx --preset refined --no-body-anim

# 改动前先存一份带时间戳的备份
python "$A" deck.pptx --preset refined --backup

# 用 JSON 精确控制每页每个对象
python "$A" deck.pptx --spec spec.json

# 清除整份文件的动画与切换
python "$A" deck.pptx --clear

# 列出可用效果名（不需要文件）
python "$A" --list
```

动画会**修改源文件**。重要成品建议加 `--backup`，或先自己复制一份。

## 两套预设

### `refined` — 克制高级感

商务汇报、管理层评审、投融资材料、咨询提案。

- 封面 `dissolve` 溶解，正文 `morph` 平滑切换，章节页 `reveal` 揭示
- 正文对象依次 `fade` 淡入
- 动画 0.5 秒，切换 0.9 秒

### `artistic` — 创意展示感

新品发布、品牌展示、创意提案、活动演讲。

- 封面 `ripple` 涟漪，正文 `morph`，章节页 `prism` 棱镜
- 正文对象依次 `float` 浮入
- 动画 0.7 秒，切换 1.2 秒

视觉存在感更强，但仍以清晰阅读为前提。严肃财报、数据密集型汇报不要默认用它。

## JSON 精细控制

规格文件用 UTF-8 编码。页码和形状序号都从 **1** 起算，`shape` 可以填序号，也可以填 PowerPoint「选择窗格」里的形状名。

```json
{
  "default_transition": "morph",
  "default_duration": 0.9,
  "slides": [
    { "index": 1, "clear": true, "transition": "dissolve", "duration": 1.0 },
    {
      "index": 2,
      "transition": "morph",
      "animations": [
        { "shape": 2, "effect": "fade", "duration": 0.6, "trigger": "after" },
        { "shape": "正文图表", "effect": "float", "duration": 0.7, "trigger": "after", "delay": 0.2 }
      ]
    }
  ]
}
```

| 字段 | 含义 |
| --- | --- |
| `default_transition` | 先给所有页设的默认切换，页级 `transition` 可覆盖 |
| `default_duration` | 默认切换时长（秒） |
| `index` | 页码，1 起算 |
| `clear` | `true` 时先清掉该页原有动画与切换 |
| `transition` / `duration` | 该页的切换效果与时长 |
| `shape` | 形状序号或形状名称 |
| `effect` / `duration` | 入场动画与时长 |
| `trigger` | `click` 单击、`after` 上一动画后、`with` 与上一动画同时 |
| `delay` | 触发后额外等待（秒） |
| `exit` | `true` 时标为退出动画 |

## 可用效果

**入场动画** `appear` `fly` `blinds` `fade` `float` `grow` `wipe` `zoom`

默认用 `fade`，最稳妥也最不容易显得廉价。创意场景可用 `float`、`wipe`、`zoom`。`fly`、`grow`、`blinds` 只在效果确实服务于叙事时用，别铺满全篇。

**页面切换** `none` `dissolve` `circle` `diamond` `ripple` `gallery` `flythrough` `reveal` `ferris` `shred` `prism` `pan` `morph` `morph_words` `morph_chars`

`morph` 是高质量转场的首选。想要最佳效果，相邻两页要保留能被 PowerPoint 认成同一对象的元素（比如复制上一页再调位置和尺寸）。`morph_words`、`morph_chars` 适合以文字变化为主体的连续页。

所有效果 ID 都在真机 PowerPoint 上实测过，并回读 `slideN.xml` 校验了写入结果。

## 自动判定行为

套用预设时脚本会做两件事，省掉手工挑选：

- **跳过装饰件**。任一边小于 10 磅的形状（分隔线、强调条）不加动画，它们逐个淡入只会制造噪音。
- **识别封面与章节页**。去掉装饰件后只剩 0 或 1 个实体形状的页，判为封面/章节页，只设切换、不加元素动画，避免喧宾夺主。

## 对 PowerPoint 进程的处理

脚本刻意避免打断你正在进行的工作：

- PowerPoint 本来就在运行时，接管现有实例，结束时**不** Quit——否则会连带关掉你手上的其他文稿。
- 目标文件已在 PowerPoint 中打开时，直接改那个窗口并保存，不关闭它，你未保存的编辑不会丢。
- 只有脚本自己启动的 PowerPoint 才由脚本退出。

## 已知限制

- 只能在 Windows + 已安装 PowerPoint 的环境跑，没有 LibreOffice 回退路径。
- Morph 需要 PowerPoint 2019 / Microsoft 365，旧版本会报错并提示。
- 动画作用于整个形状，不做段落级逐行播放。
- 动画建议在内容、版式、对象名称都稳定后最后加。后续大量重排会削弱 Morph 的连续性。
- 导入图片、图表、组合对象后形状序号可能不稳定，这时先在「选择窗格」里给对象命名，再按名称引用。

## License

MIT
