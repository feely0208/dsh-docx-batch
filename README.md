# @deepwhale-cn/dsh-docx-batch

[![Listed on dsh-plugin.org](https://dsh-plugin.org/badges/listed.svg)](https://dsh-plugin.org/plugins/feely0208/dsh-docx-batch)

**整目录的 Word 一次处理完**：批量替换文本、提取正文、按文件名合并、列页眉页脚、转 Markdown —— 不打开 Word，不动原文件。

**关键词**：Word 批量替换 · docx 批量处理 · 批量合并 Word · 提取 Word 文本 · Word 转 Markdown · 页眉页脚提取 · 合同批量改甲方 · 命令行处理 Word

## 不装会怎样

50 份合同要统一把甲方「深鲸科技」改成「深海鲸」：不装就得逐份打开 Word、Ctrl+H、保存、关掉，
半小时起步，改到第 37 份时手一滑漏一份也没人发现。

几十份周报要合成一份、把 Word 文档塞进知识库前要先转成 Markdown、需要核对每份文档的页眉页脚
写的是什么 —— 这些都属于「机械但量大」的活，人工做既慢又容易漏。

更麻烦的是：直接让 AI 改原文件，一旦改坏没有退路。这个插件只读输入，所有产物写到单独的输出目录，
同名文件自动加序号、不覆盖任何东西。

## 能做什么

| 子命令 | 作用 | 产物（可核对） |
|---|---|---|
| `replace` | 整目录批量替换文本，`--from`/`--to` 可多组 | 替换后的 `.docx` 副本；JSON 里给出每个文件替换了几处 |
| `extract` | 提取正文文本（含表格） | 每个文档一个 `.txt` |
| `merge` | 按文件名自然顺序合并成一个 docx（正文 + 表格） | `merged.docx`，段落数只增不减 |
| `headers` | 列出每份文档每一节的页眉、页脚文本 | `headers.txt` |
| `to-md` | 按「标题样式 → `#`、表格 → Markdown 表格」转换 | 每个文档一个 `.md` |

共同约定：

- **不修改输入文件**，默认输出到 `<输入目录>/_输出`，可用 `--out` 指定；
- **同名文件不覆盖**：已存在时自动变成 `合同_1.docx`、`合同_2.docx`；
- 目录递归处理，自动跳过 Word 临时文件（`~$` 开头）和 `_输出` 目录；
- `--json` 输出结构化结果，方便脚本 / CI 断言；出错时退出码非 0 并说明缺什么。

```bash
# 整目录替换，多组同时替换
dsh-docx-batch replace --input ./合同 --out ./合同_新 \
  --from 深鲸科技 --to 深海鲸 --from 2024 --to 2026

# 提取正文
dsh-docx-batch extract --input ./合同 --out ./文本

# 按文件名顺序合并
dsh-docx-batch merge --input ./周报 --out ./合并

# 页眉页脚清单 / 转 Markdown
dsh-docx-batch headers --input ./合同
dsh-docx-batch to-md --input ./合同 --out ./md
```

## 不能做什么

- ❌ **不保证排版 100% 不变**：替换跨多个 run（同一段里被拆开的文字块）时会合并该段落的 run，
  段落内的局部加粗、字号、颜色可能丢失。重要文档请先跑一份对比再批量用。
- ❌ 不处理 `.doc`（老格式）、`.docm`、加密/需要密码打开的文档；`.doc` 请先另存为 `.docx`。
- ❌ 不处理脚注、尾注、批注、文本框、SmartArt、图片里的文字 —— 只覆盖正文段落、表格、页眉页脚。
- ❌ `merge` 只合并正文和表格；页眉页脚、样式主题不合并，复杂排版（分栏、自动编号域）会退化为纯文本顺序。
- ❌ `headers` 只**读**页眉页脚，不改；`replace` 能改页眉页脚里的文字。
- ❌ **转 Markdown 会丢东西**：批注 / 修订痕迹 / 复杂表格样式（合并单元格降级为普通表格）/ 图片位置与浮动排版 / 脚注尾注 —— 只保留标题、段落、列表与基础表格文字，适合进知识库，不适合当排版成品。
- ❌ 不是格式转换器：不导出 PDF、不做 OCR、不做批量重命名。
- ❌ 替换是**字面匹配**，大小写敏感，不支持正则；`--from` 是空字符串之类无意义输入不做特殊处理。
- ✅ 全部在本机由 python-docx 处理，不联网、不上传文档。

## 安装

```bash
dsh plugin --profile web add @deepwhale-cn/dsh-docx-batch
```

运行依赖本机的 **python3 + python-docx**（插件不会替你装）。缺依赖时 CLI 会直接打印安装命令：

```bash
python3 -m pip install python-docx
# macOS 没有 python3 时
brew install python3
```

## 用法

```bash
dsh-docx-batch --help
dsh-docx-batch --demo        # 自造样例文件跑完整流程，打印产物路径
dsh-docx-batch <action> --input <文件或目录> [--out <目录>] [--json]
```

`--demo` 会在系统临时目录里造样例文档并依次跑 replace / extract / to-md / merge / headers，
不动你的任何文件，跑完直接看产物路径就知道它能干什么。

## License

MIT © 温州深鲸智能科技有限公司
