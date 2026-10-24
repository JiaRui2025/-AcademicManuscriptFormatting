# 稿件格式标准（House Style）

本文件是**格式规范的唯一权威定义**。判定一份稿件是否"整理完成"，以本文件为准，不以外观印象为准。
判定工具：`scripts/format_audit.py`（只读，输出问题清单，干净时退出码 0）。

## 1. 书签命名与可见性

| 对象 | 书签名 | 说明 |
|---|---|---|
| 第 N 条参考文献 | `_ref_N` | `_` 前缀 |
| 第 N 张图 | `_fig_N` | `_` 前缀 |
| 第 N 张表 | `_tbl_N` | `_` 前缀 |

**所有书签必须以 `_` 开头。** Word 把下划线前缀的书签视为内部书签，即使打开"显示书签"也不显示灰色方括号标记。非 `_` 前缀的书签会在图片旁、图注旁渲染成灰色小方块——这是用户反馈"图片旁边有正方形"的根因。

修复：`format_fix.py --hide-bookmarks`（重命名 + 同步全部锚点）。

## 2. 正文引用格式

### 2.1 数字引用 `[n]`

- 括号 `[` `]` 与逗号、连接号（`,` `–` `-`）是**普通正文文本**。
- 只有**数字本身**包在 `<w:hyperlink w:anchor="_ref_N">` 内。
- 区间引用 `[2–5]`：首、尾数字各链接到各自条目；中间数字不单独链接（技术上无法拆分一个数字）。
- 颜色 `1F4E79`，不加下划线、不加粗（Word 默认超链接样式不带下划线时保持这一外观）。
- 目标书签必须存在，否则为悬空锚点（`format_audit.py` 会报 `dangling_anchors`）。

示例结构：

```
[ 普通run  2  hyperlink→_ref_2  – 普通run  5  hyperlink→_ref_5  ] 普通run
```

### 2.2 图/表引用 `Fig. N` / `Table N`

**这是本标准的重点，也是最容易出错的地方。**

规则：

1. 链接文字**恰好**是 `Fig. N`（或 `Table N`），不含任何其它字符。
2. 带子图引用时写作 `Fig. 4A`：`Fig. 4` 在超链接内，`A` 是**独立的普通 run**，位于超链接**之后**。
3. 后缀字母必须与正文字体完全一致：**不得**设 `w:color=1F4E79`，不得在超链接内，不得加粗/斜体。
4. 空格规则：`Fig. 4A`（无空格）；`Fig. 4` 后接句读时正常加标点。

正确与错误对照：

| 形态 | 判定 |
|---|---|
| `[Fig. 4]A`（链接只含 Fig. 4，A 为无格式 run） | ✅ 正确 |
| `[Fig. 4A]`（字母在链接内） | ❌ 字母也会变蓝且整块可点 |
| `[Fig. 4]A`（A 是带 `1F4E79` 的 run） | ❌ 字母发蓝但点不动，颜色与正文不一致 |

修复：
- 字母被上了链接色 → `format_fix.py --fix-suffix-color`
- 字母被包进链接 → `format_fix.py --split-link-text`（自动把尾部文字移出链接并去掉格式）

### 2.3 图注与表注（caption）

- caption 段落以标签开头：`Fig. 4.` / `Table 2.`（标签后带句点）。
- caption **不是**引用，**不得**加超链接。`format_audit.py` 会把"标签开头 + 句点"的段落识别为 caption 并跳过检查。
- caption 中的 `(A)` `(B)` 等分区说明保持普通正文格式。

## 3. 外部 DOI 链接

- 形式：`https://doi.org/10.xxxx/...`，作为外部关系（`r:id`）挂在参考文献条目的 DOI 文本上。
- 只对**已核实**的 DOI 建链接；未核实的一律不建，不得根据标题猜链接。
- 无 DOI 的条目（学位论文、项目报告、标准、专著、早期论文）保持纯文本，属正常，不算问题。
- 核实方式：Crossref API `https://api.crossref.org/works/{doi}` 比对标题/作者/年份。

## 4. 版面与结构

- 图与其图注保持同页：图注段落设 `keepNext`，图段落设 `keepNext`/`keepLines`。
- 图片为 `wp:inline`（嵌入式），非浮动锚定，避免文字环绕标记干扰。
- 参考文献条目一书签一段落，一个 `_ref_N` 对应一条。

## 5. 验收清单

`format_audit.py` 退出码 0 即通过静态验收。必须为 0 的问题类型：

- `bookmarks_not_hidden` = 0
- `figure_citation_not_linked` = 0
- `table_citation_not_linked` = 0
- `link_text_not_pure` = 0
- `suffix_letter_styled` = 0
- `citation_not_linked` = 0
- `dangling_anchors` = 0
- `missing_fig_bookmarks` = 0

静态通过后仍需实机验收：在 Word 中点击若干引用与 DOI，保存→关闭→重开后再点一次。

## 6. 与 `academic-manuscript-formatting` 的分工

| 任务 | 使用 |
|---|---|
| 文献重编号、映射表、DOI 批量核验 | `academic-manuscript-formatting` |
| 本稿件/同类稿件的格式规范落地与体检 | 本 skill |
