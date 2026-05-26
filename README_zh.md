# agentdocx

Claude Code 可集成的 Word (.docx) 文档编辑工具，支持修订模式、批注和格式设置。

## 概述

agentdocx 是一个 MCP（模型上下文协议）服务器，让 Claude Code 能够直接编辑 .docx 文件。与基于 python-docx 高层 API 的现有工具不同，agentdocx 通过 lxml 在 OOXML 层面直接操作，实现了 python-docx 不原生支持的功能：修订追踪（w:ins/w:del）、批注（w:comment）、精细的文本运行和段落属性控制。

### 解决的核心问题

在合同审核等场景中，AI 审核产出了几十条修改意见，但将这些意见落实到 .docx 文档上需要手工操作。agentdocx 让 Claude Code 能够直接执行这些修改——包括以修订模式插入/删除文字、添加批注、设置字体段落格式——全部可追踪、可审阅。

## 架构

```
Claude Code ── MCP stdio ── agentdocx server ── python-docx（.docx 打包层）
                                         └── lxml（OOXML 直接操控）
```

### 模块结构

```
src/agentdocx/
├── __init__.py              包初始化
├── oxml_helpers.py      352  OOXML 命名空间常量、元素工厂函数、文本/运行/段落遍历工具
├── document.py           431  文档包装器：打开/保存、文本搜索、段落访问、批注 ZIP 注入
├── track_changes.py      513  修订追踪：w:ins（插入标记）+ w:del（删除标记），含作者/日期元数据
├── comments.py           226  批注管理：w:commentRangeStart/End + w:commentReference + comments part 增删查
├── formatting.py         389  字体属性（名称、字号、粗斜体、下划线、颜色、高亮、删除线、上下标）和
                               段落属性（对齐、缩进、间距、行距、大纲级别）
├── operations.py         465  查找替换、表格、图片、标题、项目符号/编号列表
├── batch.py             350+  批量执行器，语义坐标（find-based）解析 + offset 自动重算
└── server.py            700+  MCP 服务器：22 个独立工具 + 1 个批量工具，工具调度、会话管理
```

## 安装

```bash
cd /path/to/agentdocx
pip install -e .
```

### 依赖

- `python-docx >= 0.8.11` — .docx 文件打包（OPC/ZIP 层）
- `lxml >= 4.9.0` — OOXML 元素树直接操控
- `mcp >= 1.0.0` — MCP 服务器框架（stdio 传输）

### Claude Code 集成

在 `.claude/settings.local.json` 中添加：

```json
{
  "mcpServers": {
    "agentdocx": {
      "command": "python3",
      "args": ["-m", "agentdocx.server"],
      "cwd": "/path/to/agentdocx"
    }
  }
}
```

重启 Claude Code 会话后，所有 `docx_*` 开头的工具即可使用。

## 编辑模型：三层抽象

agentdocx 提供三层编辑抽象，从粗到精：

### 第一层：语义层（find-based）—— 推荐 LLM/Agent 使用

调用方指定**要查找的文本**，服务端在内部解析为精确位置。

```json
{"op": "replace_text", "paragraph_index": 92, "find": "2天", "new": "7个工作日"}
{"op": "insert_text",  "paragraph_index": 92, "find": "不得拖延或延误", "position": "after", "text": "默示同意条款内容。"}
{"op": "delete_text",  "paragraph_index": 125, "find": "人民币人民币"}
{"op": "add_comment",  "paragraph_index": 59, "find": "42个月", "text": "建议修改为无限TOC挂钩"}
```

**多重匹配消歧**：当段落中同一文本出现多次时，服务端报错并列出所有匹配位置及上下文片段，让调用方指定 `occurrence`（第几次出现）或使用 `context_before`/`context_after` 缩小范围：

```json
{"find": {"text": "重新计算", "context_before": "修理或更换合格后"}, "new": "重新计算，无上限"}
```

**引号自动容错**：ASCII 直引号 `"` 和 Unicode 弯引号 `"` `"`（U+201C/D）自动互转。LLM 生成的 JSON 用 `"` 而合同文档用 `"` `"` 时，无需人工干预。

### 第二层：偏移层 —— 精确到字符

调用方指定精确的 `paragraph_index` 和 `offset`。

```json
{"op": "insert_text", "paragraph_index": 0, "offset": 5, "text": "Hello"}
{"op": "delete_text", "paragraph_index": 0, "start_offset": 10, "end_offset": 15}
```

### 第三层：批量层 —— 一次调用，多个操作

所有操作合并为一个 JSON 数组，一次 MCP 调用完成。**所有 `find` 字段在原始文档状态上解析，不依赖执行顺序。**偏移量随文本插入/删除自动重算。

工作流程：
```
JSON 操作数组
    │
    ▼
步骤 1: 解析所有 find → offset（在原始文档状态上完成）
         ├─ 精确匹配
         ├─ 引号容错（ASCII ↔ Unicode）
         └─ 多重匹配检测 → 报错并列出所有位置
    │
    ▼
步骤 2: 顺序执行 + 偏移自动调整
         ├─ 按段落追踪 (position, delta) 编辑历史
         ├─ 后续操作的 offset 穿过前面的编辑点修正
         └─ 批注范围失效时优雅跳过
```

## MCP 工具列表

### 文档生命周期
| 工具 | 说明 |
|------|------|
| `docx_open` | 打开 .docx 文件，返回文档结构概览 |
| `docx_save` | 保存（可选另存为新路径） |
| `docx_close` | 关闭（可选保存） |

### 信息查询
| 工具 | 说明 |
|------|------|
| `docx_get_info` | 段落数、样式、标题结构 |
| `docx_get_text` | 获取指定段落或全文文本 |
| `docx_search` | 搜索文本，返回段落索引 + offset + 上下文 |

### 修订模式
| 工具 | 说明 |
|------|------|
| `docx_insert_text` | 插入文本（可选 w:ins 修订标记） |
| `docx_delete_text` | 删除文本（可选 w:del 修订标记） |

### 批注
| 工具 | 说明 |
|------|------|
| `docx_add_comment` | 在指定文本范围添加批注 |
| `docx_list_comments` | 列出全部批注 |
| `docx_delete_comment` | 按 ID 删除批注 |

### 格式设置
| 工具 | 说明 |
|------|------|
| `docx_set_font` | 字体名称、字号（半点，24=12pt）、粗体、斜体、下划线、颜色（hex）、高亮、删除线、上下标 |
| `docx_set_paragraph_format` | 对齐方式、缩进（twips，1英寸=1440）、间距、行距、大纲级别、段前分页 |
| `docx_set_style` | 应用命名样式（如 "Heading 1"、"Normal"） |
| `docx_set_heading` | 设为标题（1-9 级） |

### 段落操作
| 工具 | 说明 |
|------|------|
| `docx_add_paragraph` | 末尾添加段落 |
| `docx_insert_paragraph` | 在指定位置插入段落 |
| `docx_delete_paragraph` | 按索引删除段落 |

### 其他
| 工具 | 说明 |
|------|------|
| `docx_find_and_replace` | 全文查找替换 |
| `docx_add_table` | 添加表格（含表头和数据行） |
| `docx_insert_image` | 插入图片 |
| `docx_set_list_format` | 设置项目符号或编号列表 |
| `docx_batch` | **批量执行多个操作（支持语义模式和偏移模式）** |

## OOXML 实现细节

### 修订模式

- **插入**：新文本包裹在 `w:ins` 元素中，携带 `w:id`、`w:author`、`w:date` 属性。`w:ins` 内含 `w:r`（运行）元素。
- **删除**：原始文本从 `w:r/w:t` 移入 `w:del/w:r/w:delText` 元素。运行在删除边界处拆分，保留剩余文本的运行属性。
- **边界情况**：在已有修订插入中再次插入、跨多个运行的删除、与已有 w:ins/w:del 区域重叠的删除。

### 批注

OOXML 批注分为两部分：
1. **文档正文中的标记**：`w:commentRangeStart`、`w:commentRangeEnd` 和 `w:r/w:commentReference` 运行
2. **独立 comments part 中的内容**：`word/comments.xml`

由于 python-docx 的 OPC 层未暴露创建新 Part 的干净 API，agentdocx 通过以下方式处理：
1. 在内存中构建 `w:comments` 元素树
2. 通过 python-docx 保存文档
3. 将 .docx 作为 ZIP 重新打开，注入 `word/comments.xml`，更新 `[Content_Types].xml` 和 `word/_rels/document.xml.rels`

### 格式

- **字体**：属性设置在 `w:r` 的 `w:rPr`（运行属性）子元素上。字号单位为半点（如 24 = 12pt）。
- **段落**：属性设置在 `w:pPr`（段落属性）上。缩进和间距单位为 twips（1/20 点；1 英寸 = 1440 twips）。

## 使用示例

### 合同审核场景

```python
# 1. 打开文档
docx_open("/path/to/contract.docx")

# 2. 批量执行审核修改（语义模式，无需计算 offset）
docx_batch({
  "doc_id": "/path/to/contract.docx",
  "operations": [
    # 修正定义
    {"op": "replace_text", "paragraph_index": 61, "find": "潜在缺陷", "new": "潜在缺陷是指"},
    {"op": "replace_text", "paragraph_index": 59, "find": "，或该批次合同设备到交货地点42个月，二者以先到为准",
     "new": "。若因任何原因项目整体TOC证书颁发延迟，质保期相应顺延，不以固定年限为限"},

    # 修正通知条款
    {"op": "replace_text", "paragraph_index": 92, "find": "2天", "new": "7个工作日"},
    {"op": "insert_text",  "paragraph_index": 92, "find": "不得拖延或延误", "position": "after",
     "text": "被通知方在前述期限内未书面答复且未提出合理异议的，视为对通知内容的认可和同意。"},

    # 修正错误引用
    {"op": "replace_text", "paragraph_index": 228, "find": "附件八", "new": "附件七"},

    # 调整违约金
    {"op": "replace_text", "paragraph_index": 240, "find": "10 %", "new": "3 %"},

    # 强化合同生效条件
    {"op": "replace_text", "paragraph_index": 476,
     "find": "本合同经双方签字盖章之日起生效。",
     "new": "本合同在以下条件全部满足之日起生效：(a)双方签字盖章；(b)需方收到供方提交的、需方认可的履约保函和预付款保函；(c)总合同生效；(d)需方向供方发出书面生效通知。"},

    # 加粗关键条款
    {"op": "set_font", "paragraph_index": 88, "bold": true},

    # 添加审核批注
    {"op": "add_comment", "paragraph_index": 63, "find": "延长质保期",
     "text": "建议删除定义16中的总上限'最多不超过质保期结束后2年'，改为引用第11.3条重新计算机制。"},
  ]
})

# 3. 另存为
docx_save({"doc_id": "...", "path": "/path/to/contract-reviewed.docx"})
```

## 局限性

- **图片嵌入**：创建 OOXML 绘图结构，但不在 OPC 包中注入图片二进制部分。如需完整图片支持，请对底层 Document 对象使用 python-docx 的 `add_picture()`。
- **字节流中的批注**：`save_to_bytes()` 方法不支持批注注入。使用基于文件的 `save()` 以获得完整的批注支持。
- **样式存在性检查**：`set_style` 应用样式名称引用，但不验证该样式是否存在于文档的 styles part 中。应用不存在的样式可能导致 Word 以默认格式显示段落。

## 设计决策

### 语义坐标优先

agentdocx 最重要的设计决策是将 `find`（语义定位）作为主要编辑接口，而非 `offset`（字符偏移）。理由：

1. LLM 不擅长数数——在 500 字的中文段落中精确数到第 428 个字符是不可靠的
2. 审核意见以语义形式表达（"把'2天'改成'7个工作日'"），而非"在第 92 段第 97 个字节处删除 2 个字符"
3. 信息同步问题——LLM 读到的文本和 `text_content()` 算出的文本之间可能存在差异

所有编辑操作的 `find` 字段都在执行任何修改之前的原始文档状态上解析，消除了 LLM 在多次编辑之间追踪偏移变化的负担。

### 评论持久化的 ZIP 后注入

python-docx 的 OPC 层是只读的——无法在运行时向包中添加新的 Part（如评论）。选择事后解决而非扩展 OPC 层：将文档保存为扁平 ZIP 文件，然后注入评论 XML 并更新清单文件。这种方法的移植性更好，即使 python-docx 的内部 API 发生变化也不会受影响。

### 单次遍历批量执行

批量执行器两次遍历操作列表：第一次将所有语义位置（`find` 字段）解析为偏移量，第二次执行并追踪偏移变化。这种分离确保了所有语义查询都在一致的文档快照上运行，不受执行顺序问题的影响。

## 许可证

MIT
