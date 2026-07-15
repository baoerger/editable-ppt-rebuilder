# CodexImage-PPT

CodexImage-PPT 是一个用于 Codex 的 PowerPoint 重建 Skill。它把图片化的 PPT/PPTX 页面或幻灯片截图，重建为可编辑、可审计、经过严格交付验证的 PowerPoint 演示文稿。

它把源页面当作视觉和文字事实依据：可读文字重建为 PowerPoint 文本框，简单结构重建为原生形状，复杂或风格敏感的图像保留为独立语义 PNG。Skill 不声称能够恢复原始文件中不可见的矢量对象、隐藏文字或原始图表数据。

## 能做什么

- 从嵌入的整页图片或高分辨率渲染结果开始，保留不可变的源页面。
- 通过 OCR 和区域分析发现文字、图形、图片及其坐标，生成基线分解结果。
- 将标题、正文、标签、图例、坐标轴、表格和公式等可读内容重建为可编辑文字。
- 将面板、色带、边框、分隔线、连接器、箭头、徽章和简单图形重建为原生 PowerPoint 对象。
- 将照片、人物、产品、复杂图表、网络图、示意图和复杂图标作为独立语义视觉资产处理。
- 对小图标执行强制的 ImageGen 视觉核心流程：生成无文字视觉核心、去除色键、检查透明边缘、嵌入 PPTX，并记录审核证据。
- 对复杂的大幅背景执行前景移除或重建审核；简单背景优先使用原生 PowerPoint 形状。
- 生成清单、裁剪审核、背景审核、质量报告、差异图和 PowerPoint 渲染预览。
- 在交付前检查文字可编辑性、宽高比、裁剪归属、透明度、重复资产、溢出、裁切、重叠和背景残留。

## 核心工作流

1. 提取或渲染源页面，并将其保存到 `source_pages/`。
2. 使用 `scripts/decompose_visual_elements.py` 完成 OCR、候选区域发现和基线输出。
3. 检查源页面、清单、审核覆盖图和粗略预览，不盲信错误或缺少语言模型的 OCR。
4. 为每个可见对象选择一种最终表达：`editable_text`、`native_shape`、`semantic_png`、`background` 或 `backup_only`。
5. 按源页面坐标确定性地重建 PPTX，并保持源页面宽高比。
6. 为复杂视觉和每个小图标生成独立资产，完成透明度、裁剪和语义独立性审核。
7. 通过 PowerPoint 或等效渲染器导出最终预览，检查实际渲染结果。
8. 写入清单和逐页质量报告，运行严格交付验证；验证器非零退出时不交付。

## 目录

```text
codeximage-ppt/
├── SKILL.md                         # Codex Skill 主规则
├── INSTALL.md                       # 安装和运行要求
├── config.example.yaml               # 渲染、分割、OCR 和质量检查示例配置
├── requirements.txt                  # Python 依赖
├── agents/openai.yaml                # Skill 的 Codex 代理元数据
├── references/
│   ├── refined_rebuild_workflow.md   # 精细重建流程
│   ├── manifest_schema.md            # 清单字段和不变量
│   ├── quality_review_workflow.md   # 质量审核和交付门禁
│   └── imagegen_visual_core_workflow.md # ImageGen 图标和背景流程
└── scripts/
    ├── decompose_visual_elements.py # 基线分解、OCR 和区域发现
    ├── run_batches.py                # 批量基线处理
    └── validate_delivery.py          # 严格交付验证
```

## 安装

将完整目录安装到 Codex Skills 目录：

```text
$CODEX_HOME/skills/codeximage-ppt
```

如果没有设置 `CODEX_HOME`，通常使用 `~/.codex/skills/codeximage-ppt`。安装后以 `$codeximage-ppt` 调用。若当前任务没有立即发现新 Skill，请重新加载 Codex 或新建任务。

安装 Skill 与运行完整工作流是两件事。先安装 `requirements.txt` 中的 Python 依赖；OCR 还需要 Tesseract 和相应语言数据（中文通常需要 `chi_sim`）；最终渲染检查需要 Microsoft PowerPoint 或等效的高保真渲染器；小图标的强制流程需要可用的 ImageGen 能力。

## 基线处理和验证

重复处理多个输入时使用批处理脚本：

```text
python scripts/run_batches.py ...
```

完成精细重建后，使用最终输出目录和 PPTX 运行交付验证：

```text
python scripts/validate_delivery.py <refined_output_dir> --pptx <refined_deck.pptx>
```

严格交付要求验证器返回 0。输出通常包含源页面、独立 PNG、CSV/JSON 清单、裁剪和背景审核证据、PowerPoint 预览、差异图和逐页质量报告。

## 交付原则

- 源页面始终保留且不修改。
- 所有可读文字必须是可编辑文本；文字 PNG 只能作为不可见备份。
- 简单结构必须使用原生 PowerPoint 对象。
- 复杂 PNG 必须有明确的语义对象、唯一归属和审核证据，不能用页面切片拼贴整页。
- 小图标不能直接裁剪、放大、描摹或复用源图标像素作为可见最终图标。
- 不把原始整页截图放入可见最终重建，也不把未审核的生成背景当作最终结果。
- 所有审核项、裁剪问题、重复内容、溢出和裁切问题都必须关闭后才能交付。

## 局限

图片化输入无法完美恢复原始矢量对象、隐藏内容、图表底层数据或原始编辑层级。ImageGen 生成的是经过审核的视觉核心重建，不应描述为源图标的逐像素复制。对于不可读文字或无法可靠重建的精确标志，应保留人工审核状态，而不是猜测内容。

