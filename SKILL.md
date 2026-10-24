---
name: manuscript-format-standard
description: "学术论文 Word 稿件的格式标准落地与体检技能：规定并强制「Fig. N 链接 + 后缀字母为正文样式」的图引用格式、引用编号书签跳转（_ref_N/_fig_N/_tbl_N）、书签下划线前缀隐藏（消除 Word 在图片旁渲染的灰色方块标记）、DOI 外部链接、以及段落 XML 重建防乱码。当用户要求按既有规范整理论文格式、统一图表引用、给引用加跳转、处理 Fig. 4A 这类带子图引用、删除图片旁的正方形/方块标记、或对已排版稿件做格式体检与修复时使用。关键词：图注引用格式、Fig. N 后缀字母、书签跳转、隐藏书签、图片旁方块、格式标记、引用体检、稿件格式标准、manuscript house style、figure citation format。该技能直接操作 .docx（OOXML 级），提供只读审计与自动修复脚本，不改内容只改格式。"
agent_created: true
---

# 稿件格式标准落地与体检

## 目的

把一份已排版稿件**修正到既定格式标准**并**验证其达标**。标准定义在 `references/house-style.md`，本文件只讲怎么用。

核心约定（详见 house-style.md）：

1. 引用 `[n]`：括号是普通文本，数字包在指向 `_ref_N` 的超链接里，色值 `1F4E79`。
2. 图/表引用 `Fig. N` / `Table N`：链接文字**恰好**是 `Fig. N`；`Fig. 4A` 中的 `A` 是超链接**之外**的普通 run，与正文字体完全一致（不带链接色、不加粗）。
3. 所有书签以 `_` 开头（`_ref_N` / `_fig_N` / `_tbl_N`），Word 因此不显示灰色方块标记。
4. 图注/表注（`Fig. 4.` 开头）是标签不是引用，**不加链接**。
5. DOI 只对已核实的建 `https://doi.org/…` 外链。

## 何时使用

- 用户要求"按之前的规范/标准整理这份论文"、"把图引用格式统一一下"。
- 用户提出 `Fig. 4A` 这类带子图引用的格式问题。
- 用户要求删除图片旁、图注旁的"正方形/方块"标记（**先读 `references/word-pitfalls.md` §1**，那通常不是文档内容）。
- 用户要求给引用加跳转、核对书签与链接是否完好。
- 对已排版稿件做交付前格式体检。

文献重编号、映射表生成、DOI 批量核验属于 `academic-manuscript-formatting` 技能的范围，不在本技能内。

## 环境

```bash
PY=/Users/mac/.workbuddy-ai/binaries/python/envs/default/bin/python
S=~/.workbuddy-ai/skills/manuscript-format-standard/scripts
[ -x "$PY" ] || /Users/mac/.workbuddy-ai/binaries/python/versions/3.13.12/bin/python3 \
    -m venv /Users/mac/.workbuddy-ai/binaries/python/envs/default
"$PY" -m pip install -q python-docx lxml
```

只处理 `.docx`；`.doc` 先请用户在 Word 中另存为 `.docx`。

## 工作流

### 第 1 步 · 体检（只读，必做）

```bash
"$PY" "$S/format_audit.py" 稿件.docx -o audit.json
```

退出码 0 = 达标；1 = 有问题，按 JSON 中的分类处理。**不要跳过这一步直接改**——先量化差距。

会把正常内容正确排除的两类：图注段落（`Fig. 4.` 开头带句点）与 References 之后的条目编号。

### 第 2 步 · 修复（按需选择操作）

```bash
# 全修（幂等，可重复执行）
"$PY" "$S/format_fix.py" 稿件.docx -o 稿件_fixed.docx --all

# 或按需
"$PY" "$S/format_fix.py" 稿件.docx -o out.docx --hide-bookmarks      # 书签加 _ 前缀 + 同步锚点
"$PY" "$S/format_fix.py" 稿件.docx -o out.docx --fix-suffix-color    # 去掉链接外 run 的 1F4E79
"$PY" "$S/format_fix.py" 稿件.docx -o out.docx --split-link-text     # 把 Fig. 4A 拆成 Fig. 4 + A
```

输入文件永不被修改；输出为新文件。`--dry-run` 可预演。三个操作可组合，也可与手工 OOXML 编辑混用。

### 第 3 步 · 复检

```bash
"$PY" "$S/format_audit.py" 稿件_fixed.docx
```

必须为 0 问题。若仍有问题，回到第 2 步或按 `references/ooxml-recipes.md` 手工处理。

### 第 4 步 · 生成对照/清单（可选）

```bash
"$PY" "$S/docx_inventory.py" 稿件_fixed.docx --csv mapping.csv
```

输出每条文献的编号、书签是否就位、DOI 链接与目标；每张图/表的书签与被链接次数。

### 第 5 步 · 实机验收（不可省略）

静态通过不等于跳转可用。必须请用户在 Word 中抽查：点正文首/中/尾各一处引用确认落在对应条目；点 2–3 个 DOI 确认打开同一篇论文；保存→关闭→重开后再点一次。未抽查的项标注 NOT VERIFIED。

## 文档正在本地编辑器中打开时

host 注入了 `file_id` 时，直接用 python-docx 改磁盘会让编辑器内存副本过期，用户一保存就覆盖改动。顺序必须是：

1. python-docx 改磁盘文件；
2. `open_file file_path=... open_with_existing=true` 让服务端重载；
3. `get_pool_status` 取回新 `file_id`；
4. 用查询工具复核内容已更新；
5. **不要**再调 `save_file`。

注意：editor_sdk 的 `doc_insert_normal_link` 只接受 URL，**建不了指向书签的内部跳转**；`doc_replace_text` 也改不了超链接字段的锚点。这两类操作必须走 OOXML。完整说明见 `references/ooxml-recipes.md` §8。

## 铁律

1. **只改格式，不改科学内容。** 数值、化学计量、术语、作者、年份一律不动。
2. **不编造。** DOI 未核实就不建链接；书签目标不存在就不建跳转。
3. **改文字必改锚点。** "Fig. 4" 改成 "Fig. 3" 时必须同步把锚点从 `_fig_4` 改成 `_fig_3`，否则得到错链接。
4. **书签改名必同步锚点。** 漏一个 `w:hyperlink/@w:anchor` 就产生悬空链接。
5. **备份优先。** 动手前把原稿复制到备份目录。
6. **区分"文档有问题"与"编辑器显示效果"。** 用户看到的方块可能只是 Word 格式标记，动手前先穷举确认（`references/word-pitfalls.md`）。

## 参考资源

| 文件 | 内容 |
|---|---|
| `references/house-style.md` | **格式标准权威定义**：书签命名、引用/图引用格式、caption 规则、DOI、验收清单 |
| `references/ooxml-recipes.md` | OOXML 实现配方：内部/外部超链接、隐藏书签、段落重建防乱码、后缀字母移出链接、editor_sdk 同步 |
| `references/word-pitfalls.md` | 六个实际踩过的坑：图片旁方块、后缀字母假链接、图引用指错图、乱码、编辑器不同步、审计假阳性 |
| `references/ref-1018.md` | 本文稿（Úbeda pointing mortar）的文献编号对照、DOI 清单、图表清单 |

## 脚本清单

| 脚本 | 作用 | 是否修改文件 |
|---|---|---|
| `format_audit.py` | 8 类格式问题体检，输出 JSON/文本报告 | 否 |
| `format_fix.py` | 隐藏书签 / 去后缀字母链接色 / 拆分链接文字 | 是（输出新文件） |
| `docx_inventory.py` | 文献·图·表清单与 DOI 链接状态，可导 CSV | 否 |
