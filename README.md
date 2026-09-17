# DocumentFormat

一个面向正式 Word 文档的**模板驱动格式化 Skill**。

它的目标不是把文档“排得看起来差不多”，而是尽量保留或创建真正的 Word 原生结构，包括：

- Heading / Outline 标题层级
- 多级标题自动编号
- 章节段前分页
- 图、表题注 `SEQ` 域
- 自动目录 `TOC`
- 页码 `PAGE`
- 参考文献原生自动编号 `[1]`、`[2]`……
- 正文参考文献引用 `REF ... \n \h`
- 图表、公式等交叉引用
- Section、页边距、页码重启、页眉页脚
- 模板中的自定义样式、表格和排版规则

核心原则：

> **Word 模板是格式事实来源，profile 只保存模板难以可靠表达的语义规则。**

---

## 1. 最简单的使用方式

如果把这个仓库作为 `document-format-writer` Skill 使用，通常不需要自己运行脚本。

你只需要把内容或 Word 文件发给支持该 Skill 的 Agent，然后说明要按什么格式生成即可。

### 1.1 直接发文字，使用默认模板

例如直接发送：

```text
请使用默认格式生成 Word 文档。

题目：基于多模态感知的反无人机目标识别研究

摘要：……
关键词：多模态；反无人机；目标识别

1 绪论
正文……

1.1 研究背景
正文……

参考文献
[1] 张三. 示例文献[J]. 示例期刊, 2026.
```

没有额外提供模板时，Skill 默认使用：

```text
templates/generic-formal/template.docx
templates/generic-formal/profile.yaml
```

Agent 会先识别文字中的标题、正文、图表、参考文献等语义，再按默认 Word 模板生成最终 `.docx`。

为了让结构识别更稳定，直接发文字时建议明确写出：

```text
摘要
Abstract
目录
1 一级标题
1.1 二级标题
1.1.1 三级标题
参考文献
致谢
附录A
```

不需要你手工设置字体、字号、行距或 Word 多级列表。

---

### 1.2 只发一个已有 Word，使用默认模板重新格式化

上传：

```text
论文原稿.docx
```

然后说：

```text
请按默认模板重新格式化这个 Word，正文内容不要修改。
```

Skill 会使用默认 `generic-formal` 模板，将原文中的标题、正文、图表题注、参考文献等映射到目标 Word 样式，并尽量转换为 Word 原生结构。

---

### 1.3 直接发文字 + 你自己的 Word 模板

上传：

```text
学校论文格式模板.docx
```

同时发送正文文字，例如：

```text
请按照我上传的 Word 模板生成文档，内容如下：

1 绪论
……

1.1 研究背景
……
```

此时用户提供的模板优先于默认模板。

如果没有额外提供 `profile.yaml`，Skill 会尽量从模板中自动识别：

- Heading / Outline 层级
- 自定义标题样式，例如“章标题”“节标题”
- 正文样式
- 图题、表题样式
- 参考文献样式
- 页码位置和格式
- 前置部分 / 正文 Section
- 页眉和页脚
- 题注标签
- 参考文献是否使用数字自动编号

因此，一个规范制作的 Word 模板可以**不配 YAML 直接使用**。

---

### 1.4 已有 Word + 你自己的 Word 模板

上传两个文件：

```text
论文原稿.docx
学校论文格式模板.docx
```

然后说：

```text
请按照“学校论文格式模板.docx”重新格式化“论文原稿.docx”，内容不要改动。
```

这是最推荐的自定义格式使用方式。

处理优先级为：

```text
用户明确要求
    ↓
用户提供的 template.docx / .dotx
    ↓
用户提供的 profile.yaml（如果有）
    ↓
默认 generic-formal 模板
```

---

## 2. 默认模板

默认模板位于：

```text
templates/generic-formal/
├── template.docx
└── profile.yaml
```

`template.docx` 是真正的 Word 格式基线，负责视觉格式和大部分 Word 原生结构。

默认格式目前主要覆盖：

- 正式长文档结构
- 中文摘要 / 英文摘要
- 自动目录
- Heading 1–4
- `1 → 1.1 → 1.1.1 → 1.1.1.1` 原生多级标题编号
- 一级标题新页
- 正文样式
- 图 / 表按章题注
- `SEQ` 图表编号
- 前置部分罗马页码、正文阿拉伯页码
- 参考文献原生 `[1]`、`[2]` 自动编号
- 参考文献编号后真实 TAB
- 正文 `[1]` 转为可更新、可点击的 Word `REF` 引用
- 致谢
- 附录
- 公式和交叉引用相关规则
- Section、页眉、页脚和页码结构

`profile.yaml` 不是第二份“格式说明书”。

它主要保存语义配置，例如：

- 哪个样式对应一级标题
- 图、表题注标签
- 目录深度
- 参考文献编号规则
- Section 标题别名
- 某些无法仅从 Word 模板可靠判断的行为

字体、字号、行距、页边距、对齐、边框等视觉规则，应优先保存在 `template.docx` 中，而不是重复写进 YAML。

---

## 3. 新增一个长期复用的格式模板

如果某个格式以后会经常使用，可以新增：

```text
templates/
└── my-format/
    ├── template.docx
    └── profile.yaml     # 可选
```

例如：

```text
templates/
└── xtu-thesis/
    ├── template.docx
    └── profile.yaml
```

### 推荐的 Word 模板制作方式

尽量在 Word / WPS 中提前配置真正的原生结构：

- 标题 1、标题 2、标题 3、标题 4
- Heading 对应的 Outline Level
- 真正的多级列表
- 一级标题 `pageBreakBefore`
- 正文样式
- 图题、表题样式
- `SEQ` 题注
- TOC
- 页码 PAGE
- 前置部分 / 正文 Section
- 页码重启
- 页眉页脚
- 表格样式

模板越规范，自动识别越可靠。

不要只做这种“视觉模板”：

```text
1 绪论          ← 1 是手打文字
1.1 研究背景    ← 1.1 是手打文字
图1-1 示例图    ← 图1-1 是普通文字
```

更推荐真正配置 Word 原生编号和域。

---

## 4. 什么时候需要 profile.yaml

**大多数新模板可以先不写 profile。**

只有自动推断不够明确时，再增加一个很小的 profile。

例如：

```yaml
name: my-format

styles:
  heading_1: 章标题
  heading_2: 节标题
  heading_3: 小节标题
  body: 正文样式
  figure_caption: 图题
  table_caption: 表题
  references: 文献条目

captions:
  figure:
    label: 图
    chapter_heading_level: 1
    separator: "-"
  table:
    label: 表
    chapter_heading_level: 1
    separator: "-"

contents:
  depth: 3

references:
  numbering:
    native_list_required: true
    format: "[%1]"
    suffix: tab
```

不要在 profile 中重复：

```text
宋体
小四
1.5 倍行距
段前 12 磅
左边距 3 cm
```

这些应该直接配置在 Word 模板里。

---

## 5. 参考文献机制

对于数字编号参考文献，Skill 不会只生成看起来像：

```text
[1] 文献一
[2] 文献二
```

的普通文本。

正确结构是：

```text
Word 原生编号列表：[%1]
编号后缀：TAB
```

参考文献正文只保存：

```text
张三. 示例文献[J]. 示例期刊, 2026.
```

`[1]` 由 Word 编号系统自动生成。

正文中的：

```text
已有研究表明该方法有效[1]。
```

会转换为类似：

```text
REF _FmtRef0001 \n \h
```

的 Word 域。

这样调整参考文献顺序后，更新域即可同步更新正文引用，而不是重新手工修改编号。

---

## 6. 图、表和交叉引用

图表题注应使用 Word 原生 `SEQ`，例如：

```text
图1-1 系统结构图
表2-3 实验结果
```

正文中的“见图1-1”“如表2-3所示”应尽量使用 Word `REF` 域，而不是复制普通数字。

题注的字体、段距、对齐方式来自模板样式；格式化代码不会强行把所有题注居中。

---

## 7. Section、页眉和页脚

自定义模板可以包含多个 Section，例如：

```text
封面
摘要 / 目录     → I, II, III ...
正文             → 1, 2, 3 ...
附录
```

Skill 会尽量按照模板匹配前置部分与正文 Section，而不是简单把模板最后一个 Section 套到整篇文档。

同时支持迁移：

- 页面大小和页边距
- 页码格式 / 重启
- 默认页眉页脚
- 首页不同
- 奇偶页不同
- 页眉页脚中的图片和超链接关系

---

## 8. 命令行使用

如果你是在本地开发或调试这个 Skill，可以直接运行内部脚本。

### 安装依赖

```bash
python -m pip install -e ".[test]"
```

### 使用默认模板格式化已有 Word

```bash
python scripts/fast_format_docx.py \
  --source input.docx \
  --output output.docx \
  --ensure-toc \
  --strict-cross-references
```

未指定 `--template` 时使用默认模板。

### 使用自定义模板，不提供 profile

```bash
python scripts/fast_format_docx.py \
  --source input.docx \
  --template my-template.docx \
  --output output.docx \
  --ensure-toc \
  --strict-cross-references
```

程序会尽量从模板自身推断语义样式和行为。

### 使用自定义模板 + profile

```bash
python scripts/fast_format_docx.py \
  --source input.docx \
  --template my-template.docx \
  --profile my-profile.yaml \
  --output output.docx \
  --ensure-toc \
  --strict-cross-references
```

### 验证输出结构

默认模板：

```bash
python scripts/validate_docx.py output.docx \
  --require-toc \
  --require-captions \
  --require-multilevel \
  --require-crossrefs \
  --require-reference-numbering
```

自定义模板：

```bash
python scripts/validate_docx.py output.docx \
  --template my-template.docx \
  --require-crossrefs
```

### 运行测试

```bash
pytest -q
```

---

## 9. 项目结构

```text
DocumentFormat/
├── README.md
├── SKILL.md
├── pyproject.toml
├── references/
│   ├── template-contract.md
│   ├── validation.md
│   └── word-native-structures.md
├── scripts/
│   ├── fast_format_docx.py
│   ├── inspect_docx.py
│   ├── profile_config.py
│   ├── repair_cross_references.py
│   ├── restore_template_parts.py
│   ├── template_runtime.py
│   └── validate_docx.py
├── templates/
│   ├── README.md
│   └── generic-formal/
│       ├── template.docx
│       └── profile.yaml
└── tests/
    ├── test_formatter.py
    └── test_template_runtime.py
```

---

## 10. 设计原则

1. **模板优先**：Word 模板是真正的格式来源。
2. **语义优先**：标题、题注、参考文献、目录等使用 Word 原生结构。
3. **不静态伪装**：不接受“看起来一样，但内部只是普通文本”。
4. **尽量零配置**：新模板优先自动推断，只有必要时才写小型 profile。
5. **不重复维护格式**：能由 Word 表达的格式，不再复制成大段 Markdown / YAML。
6. **结构 + 视觉双重验收**：最终 `.docx` 既要看起来正确，也要能在 Word/WPS 中继续编辑、更新编号和交叉引用。
