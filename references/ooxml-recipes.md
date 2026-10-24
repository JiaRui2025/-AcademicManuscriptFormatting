# OOXML 实操要点

本文件记录在 .docx 上实现书签跳转、超链接、段落重建的可靠做法与陷阱。
python-docx 只提供基础的段落/run 读写；**超链接、书签、内部锚点跳转都要自己拼 XML**。

## 1. 内部跳转超链接（引用 → 书签）

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def make_internal_link(text, anchor, color="1F4E79"):
    h = OxmlElement('w:hyperlink')
    h.set(qn('w:anchor'), anchor)          # 指向书签名，无 r:id
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    c = OxmlElement('w:color'); c.set(qn('w:val'), color)
    rpr.append(c); r.append(rpr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t); h.append(r)
    return h
```

- `w:anchor` 是书签名，**不是** `r:id`。内部跳转不需要关系条目。
- 书签：`<w:bookmarkStart w:id="N" w:name="_ref_1"/>` + 配对的 `<w:bookmarkEnd w:id="N"/>`，`w:id` 只需在同一文档内唯一配对。
- 校验锚点：遍历 `w:hyperlink/@w:anchor`，与全部 `w:bookmarkStart/@w:name` 集合求差，空差即无悬空。

## 2. 外部 DOI 链接

```python
from docx.opc.constants import RELATIONSHIP_TYPE as RT

rid = doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
h = OxmlElement('w:hyperlink')
h.set(qn('r:id'), rid)        # 外部链接用 r:id
```

- 外部链接用 `r:id`，内部跳转用 `w:anchor`，二者互斥。
- `relate_to` 对同一 URL 会复用既有关系，不会重复创建。

## 3. 隐藏书签（解决"图片旁有正方形"）

Word 对**非下划线前缀**的书签，在开启"显示书签"时于图片/图注旁渲染灰色方括号标记。把书签名改为 `_` 开头即可隐藏，且跳转功能不受影响。

```python
mapping = {}
for b in body.findall('.//' + qn('w:bookmarkStart')):
    name = b.get(qn('w:name'))
    if name and not name.startswith('_'):
        new = '_' + name
        mapping[name] = new
        b.set(qn('w:name'), new)
for h in body.findall('.//' + qn('w:hyperlink')):
    a = h.get(qn('w:anchor'))
    if a in mapping:
        h.set(qn('w:anchor'), mapping[a])
```

**必须同步更新所有 `w:hyperlink/@w:anchor`**，漏一个就产生悬空链接。

## 4. 替换整段文字（避免 Fig./数字乱码）

**陷阱**：用 python-docx 只清空 `run.text` 或只删 `w:r`，会留下段落内原有的 `w:hyperlink` 元素。旧超链接里的文字（如 "Fig. 2"、"21"、"17"）会残留在段尾，与新文字拼接成 `Fig. 217171819Fig. 4…` 这类乱码。

**正确做法：重建整段 XML**，保留 `w:pPr`，删除其余全部子元素后按需重新生成：

```python
def rebuild_paragraph(p, new_children):
    pPr = p.find(qn('w:pPr'))
    for child in list(p):
        if child is not pPr:
            p.remove(child)
    for c in new_children:
        p.append(c)
```

- 段落的直接子元素只会有：`w:pPr`、`w:r`、`w:hyperlink`、`w:bookmarkStart/End`。
- 重建后书签若被删掉，需重新插入，否则引用该书签的链接全部悬空。
- 生成链接文字时按"普通文本 + 超链接"交替的顺序 append，见 `house-style.md` §2。

## 5. 把后缀字母移出超链接

当链接文字是 `Fig. 4A` 而标准要求 `Fig. 4` + 独立 `A`：

1. 把链接内所有 `w:t` 的文本改成 `Fig. 4`；
2. 深拷贝链接内第一个 `w:r`，**删掉其 `w:rPr`**（去掉链接色等直设格式），文字改为 `A`；
3. 用 `h.addnext(new_run)` 插到超链接之后。

删掉 `w:rPr` 后 run 继承段落/样式默认格式，与正文完全一致——这正是标准要求的效果。

## 6. 定位文本位置

- 不要用肉眼数字数定位。用字符串匹配遍历段落，或先 `doc_find` 拿坐标。
- 段落序号（`paragraph_index`）与字符坐标（UTF-16 offset）是两套坐标系，不可混用。
- 任何写操作后，之前取得的坐标全部失效，需重新查询。

## 7. python-docx 的已知限制

| 能力 | 状态 |
|---|---|
| 读写段落文字、样式 | ✅ |
| 内部锚点超链接 | ❌ 需手写 XML |
| 外部超链接 | ❌ 需 `relate_to` + 手写 XML |
| 书签创建/改名 | ❌ 需手写 XML |
| 删除段落内指定子元素 | ✅ 直接操作 `p` 的元素树 |
| 修改既有超链接的锚点 | ❌ 需手写 XML |

## 8. 与本地编辑器（editor_sdk）的同步

若文档正在本地编辑器中打开（host 注入了 `file_id`），直接用 python-docx 改磁盘文件会导致编辑器内存副本过期、用户保存时覆盖改动。

可靠顺序：

1. 先用 python-docx 改并保存磁盘文件；
2. 调 `open_file file_path=... open_with_existing=true` 让服务端发现磁盘已更新并重载；
3. 用 `get_pool_status` 取回新的 `file_id`，用查询工具（如 `doc_find`）复核内容确实已更新；
4. **不要**再调 `save_file`（会把过期内存写回磁盘，覆盖改动）。

注意：editor_sdk 的 `doc_insert_normal_link` 只接受 URL，**无法创建指向书签的内部跳转**；`doc_replace_text` 也改不了超链接字段的锚点。这两类操作必须走 OOXML。
