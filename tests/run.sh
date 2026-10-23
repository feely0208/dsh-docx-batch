#!/bin/bash
# @deepwhale-cn/dsh-docx-batch · 双向验证
# 正例：用 python-docx 现场造 docx，跑 replace/extract/merge/headers/to-md，
#       逐项断言产物存在且内容确实变了；反例：坏输入必须非 0 退出并说明原因。
set -uo pipefail
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
cd "$(dirname "$0")/.."

PASS=0; FAIL=0
ok(){ echo "  ✅ $1"; PASS=$((PASS+1)); }
bad(){ echo "  ❌ $1"; FAIL=$((FAIL+1)); }
check(){ # check <描述> <命令...>
  local desc="$1"; shift
  if "$@" >/dev/null 2>"$TMP/last_err"; then ok "$desc"
  else bad "$desc"; sed -n '1,3p' "$TMP/last_err" | sed 's/^/       /'; fi
}
say(){ echo ""; echo "  【$1】"; }

TMP="$(mktemp -d "${TMPDIR:-/tmp}/dsh-docx-batch-test.XXXXXX")"
trap 'rm -rf "$TMP"' EXIT
FIX="$TMP/输入"
OUT="$TMP/输出"
mkdir -p "$FIX"

echo ""; echo "  @deepwhale-cn/dsh-docx-batch · 双向验证"
echo "  ────────────────────────────────────────────────────────"

# ── 造正例样本（python-docx，不依赖手工二进制） ────────────────────
python3 - "$FIX" <<'PY' || { echo "  ❌ 造样例失败"; exit 1; }
import sys
from docx import Document
root = sys.argv[1]

def make(name, heading, lines, rows=None, header=None, footer=None):
    doc = Document()
    doc.add_heading(heading, level=1)
    for line in lines:
        doc.add_paragraph(line)
    if rows:
        table = doc.add_table(rows=len(rows), cols=len(rows[0]))
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                table.cell(i, j).text = value
    if header:
        doc.sections[0].header.paragraphs[0].text = header
    if footer:
        doc.sections[0].footer.paragraphs[0].text = footer
    doc.save(root + "/" + name)

make("A合同.docx", "第一章 总则",
     ["甲方：深鲸科技", "乙方：示例客户", "深鲸科技对本合同内容保密。"],
     [["条款", "内容"], ["保密", "深鲸科技对数据保密"]],
     header="机密 · 深鲸科技", footer="第 1 页")
make("B附件.docx", "附件 B", ["附件第一段", "附件第二段"])
PY

BEFORE=$(shasum -a 256 "$FIX"/*.docx | shasum -a 256)
say "CLI 形态"
node bin/cli.mjs --help >"$TMP/help.txt" 2>&1 && ok "--help 退出 0" || bad "--help 退出非 0"
grep -q "子命令" "$TMP/help.txt" && ok "--help 列出子命令" || bad "--help 缺少子命令列表"
grep -q "python-docx" "$TMP/help.txt" && ok "--help 说明 python 依赖" || bad "--help 未说明依赖"
node bin/cli.mjs >/dev/null 2>&1 && ok "无参数时打印 help 且退出 0" || bad "无参数未退出 0"
node bin/cli.mjs replace --input "$FIX" >/dev/null 2>&1; [ $? -ne 0 ] && ok "缺 --from 时非 0 退出" || bad "缺 --from 仍成功"
DEP=$(python3 -S lib/docx_tool.py replace '{}' 2>&1); RC=$?
[ $RC -eq 3 ] && ok "缺 python-docx 依赖时退出码 3" || bad "缺依赖退出码应为 3，实际 $RC"
echo "$DEP" | grep -q "pip install python-docx" && ok "缺依赖明确提示安装命令（不静默）" || bad "缺依赖没有安装提示"

say "正例 · replace 批量替换（正文+表格+页眉页脚）"
node bin/cli.mjs replace --input "$FIX" --out "$OUT/替换" \
  --from 深鲸科技 --to 深海鲸 --from 示例客户 --to 真实客户 --json >"$TMP/replace.json" 2>&1
RC=$?
[ $RC -eq 0 ] && ok "replace 退出 0" || bad "replace 退出 $RC"
cat > "$TMP/check_replace.py" <<'PY'
import glob, json, os, sys
from docx import Document
out, json_path = sys.argv[1], sys.argv[2]
data = json.load(open(json_path, encoding="utf-8"))
assert data.get("ok") is True, data
assert data["replacements"] >= 5, "替换次数过少: %s" % data["replacements"]
docs = sorted(glob.glob(os.path.join(out, "*.docx")))
assert len(docs) == 2, "期望 2 个输出文件，实际 %d" % len(docs)
found_new = False
for path in docs:
    doc = Document(path)
    texts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    for section in doc.sections:
        texts += [p.text for p in section.header.paragraphs]
        texts += [p.text for p in section.footer.paragraphs]
    joined = "\n".join(texts)
    assert "深鲸科技" not in joined, "%s 仍含旧文本" % os.path.basename(path)
    assert "示例客户" not in joined, "%s 仍含旧文本" % os.path.basename(path)
    found_new = found_new or ("深海鲸" in joined and "真实客户" in joined)
assert found_new, "输出里找不到新文本"
PY
check "替换后正文/表格/页眉页脚都已改变，JSON 计数 ≥5" python3 "$TMP/check_replace.py" "$OUT/替换" "$TMP/replace.json"

AFTER=$(shasum -a 256 "$FIX"/*.docx | shasum -a 256)
[ "$BEFORE" = "$AFTER" ] && ok "输入文件未被修改（sha256 一致）" || bad "输入文件被修改了"

say "正例 · extract 提取正文"
node bin/cli.mjs extract --input "$OUT/替换" --out "$OUT/文本" --json >"$TMP/extract.json" 2>&1
RC=$?
[ $RC -eq 0 ] && ok "extract 退出 0" || bad "extract 退出 $RC"
cat > "$TMP/check_extract.py" <<'PY'
import glob, json, os, sys
out, json_path = sys.argv[1], sys.argv[2]
assert json.load(open(json_path, encoding="utf-8")).get("ok") is True
txts = glob.glob(os.path.join(out, "*.txt"))
assert len(txts) == 2, "期望 2 个 txt，实际 %d" % len(txts)
content = "\n".join(open(p, encoding="utf-8").read() for p in txts)
assert "深海鲸" in content, "没有新文本"
assert "深鲸科技" not in content, "仍有旧文本"
assert "保密" in content, "表格内容没提取到"
PY
check "提取出 2 个 .txt，内容已变化且含表格" python3 "$TMP/check_extract.py" "$OUT/文本" "$TMP/extract.json"

say "正例 · merge 合并（段落数必须增加）"
node bin/cli.mjs merge --input "$FIX" --out "$OUT/合并" --json >"$TMP/merge.json" 2>&1
RC=$?
[ $RC -eq 0 ] && ok "merge 退出 0" || bad "merge 退出 $RC"
cat > "$TMP/check_merge.py" <<'PY'
import glob, json, os, sys
from docx import Document
out, json_path = sys.argv[1], sys.argv[2]
data = json.load(open(json_path, encoding="utf-8"))
assert data.get("ok") is True
merged = glob.glob(os.path.join(out, "*.docx"))
assert len(merged) == 1, "期望 1 个合并文件，实际 %d" % len(merged)
doc = Document(merged[0])
text = "\n".join(p.text for p in doc.paragraphs)
assert len(doc.paragraphs) > 4, "合并后段落数没有增加: %d" % len(doc.paragraphs)
assert "第一章 总则" in text and "附件第二段" in text, "两份文档内容没有都合进来"
tables_text = "\n".join(c.text for t in doc.tables for r in t.rows for c in r.cells)
assert "保密" in tables_text, "表格没有合并进来"
PY
check "合并后段落数 > 单份，两份正文和表格都在" python3 "$TMP/check_merge.py" "$OUT/合并" "$TMP/merge.json"

say "正例 · headers 页眉页脚"
node bin/cli.mjs headers --input "$FIX" --out "$OUT/页眉" >"$TMP/headers.txt" 2>&1
RC=$?
[ $RC -eq 0 ] && ok "headers 退出 0" || bad "headers 退出 $RC"
cat > "$TMP/check_headers.py" <<'PY'
import glob, os, sys
out = sys.argv[1]
files = glob.glob(os.path.join(out, "headers.txt"))
assert len(files) == 1, "没有 headers.txt"
text = open(files[0], encoding="utf-8").read()
assert "深鲸科技" in text, "没列出页眉文本"
assert "第 1 页" in text, "没列出页脚文本"
PY
check "headers.txt 同时列出页眉与页脚文本" python3 "$TMP/check_headers.py" "$OUT/页眉"

say "正例 · to-md 转 Markdown"
node bin/cli.mjs to-md --input "$FIX" --out "$OUT/md" >"$TMP/md.txt" 2>&1
RC=$?
[ $RC -eq 0 ] && ok "to-md 退出 0" || bad "to-md 退出 $RC"
cat > "$TMP/check_md.py" <<'PY'
import glob, os, sys
out = sys.argv[1]
files = sorted(glob.glob(os.path.join(out, "*.md")))
assert len(files) == 2, "期望 2 个 md，实际 %d" % len(files)
a = next(p for p in files if "A合同" in p)
b = next(p for p in files if "B附件" in p)
a_text = open(a, encoding="utf-8").read()
b_text = open(b, encoding="utf-8").read()
assert "# 第一章 总则" in a_text, "标题层级没转对:\n" + a_text[:200]
assert "| 条款 | 内容 |" in a_text, "表格没转成 Markdown 表格"
assert "# 附件 B" in b_text, "第二份标题没转对"
PY
check "Markdown 含标题/正文/表格" python3 "$TMP/check_md.py" "$OUT/md"

say "正例 · 不覆盖同名产物 & 默认输出目录"
node bin/cli.mjs replace --input "$FIX" --out "$OUT/替换" --from 深鲸科技 --to 深海鲸 >/dev/null 2>&1
[ -f "$OUT/替换/A合同.docx" ] && [ -f "$OUT/替换/A合同_1.docx" ] \
  && ok "同名产物自动加序号（A合同_1.docx）" || bad "同名产物被覆盖或未加序号"
mkdir -p "$TMP/默认目录" && cp "$FIX/A合同.docx" "$TMP/默认目录/"
node bin/cli.mjs extract --input "$TMP/默认目录" >/dev/null 2>&1
[ -f "$TMP/默认目录/_输出/A合同.txt" ] && ok "未给 --out 时默认写到 <输入目录>/_输出" || bad "默认输出目录不对"
node bin/cli.mjs extract --input "$FIX/A合同.docx" --out "$TMP/单文件" >/dev/null 2>&1
[ -f "$TMP/单文件/A合同.txt" ] && ok "支持单个 .docx 输入" || bad "单文件输入失败"

say "正例 · --demo 自造样例跑完整流程"
node bin/cli.mjs --demo --json >"$TMP/demo.json" 2>&1
RC=$?
[ $RC -eq 0 ] && ok "--demo 退出 0" || bad "--demo 退出 $RC"
cat > "$TMP/check_demo.py" <<'PY'
import json, os, sys
data = json.load(open(sys.argv[1], encoding="utf-8"))
assert data.get("ok") is True and data.get("action") == "demo", data
assert len(data.get("steps", [])) == 5, "演示步骤不足 5 个"
arts = data.get("artifacts", [])
assert len(arts) >= 8, "演示产物不足: %d" % len(arts)
for path in arts:
    assert os.path.exists(path), "演示产物不存在: " + path
print(data["demoRoot"])
PY
DROOT=$(python3 "$TMP/check_demo.py" "$TMP/demo.json" 2>"$TMP/last_err")
[ $? -eq 0 ] && ok "5 个动作全跑通，产物均真实存在" || bad "demo 产物校验失败"
[ -n "${DROOT:-}" ] && rm -rf "$DROOT"

say "反例 · 坏输入必须非 0 退出并说明缺什么"
NEG=$(node bin/cli.mjs extract --input "$TMP/根本不存在" 2>&1); RC=$?
[ $RC -ne 0 ] && ok "不存在的输入 → 非 0 退出（$RC）" || bad "不存在的输入仍退出 0"
echo "$NEG" | grep -q "不存在" && ok "提示「输入不存在」" || bad "缺少错误提示：$NEG"
mkdir -p "$TMP/空目录"
NEG=$(node bin/cli.mjs merge --input "$TMP/空目录" 2>&1); RC=$?
[ $RC -ne 0 ] && ok "空目录 → 非 0 退出（$RC）" || bad "空目录仍退出 0"
echo "$NEG" | grep -q "没有 .docx" && ok "提示「目录里没有 .docx」" || bad "空目录提示不对：$NEG"
NEG=$(node bin/cli.mjs replace --input "$FIX" --from 深鲸科技 2>&1); RC=$?
[ $RC -ne 0 ] && ok "--from/--to 不成对 → 非 0 退出（$RC）" || bad "from/to 不成对仍成功"
echo "$NEG" | grep -q "数量不一致" && ok "提示「from 与 to 数量不一致」" || bad "不成对提示不对：$NEG"
NEG=$(node bin/cli.mjs nosuch --input "$FIX" 2>&1); RC=$?
[ $RC -ne 0 ] && ok "未知子命令 → 非 0 退出（$RC）" || bad "未知子命令仍退出 0"
echo "$NEG" | grep -q "未知子命令" && ok "提示「未知子命令」" || bad "未知子命令提示不对"
NEG=$(node bin/cli.mjs extract 2>&1); RC=$?
[ $RC -ne 0 ] && ok "缺 --input → 非 0 退出（$RC）" || bad "缺 --input 仍退出 0"
node bin/cli.mjs extract --input "$TMP/根本不存在" --json >"$TMP/err.json" 2>&1; RC=$?
[ $RC -ne 0 ] && ok "--json 模式下错误也非 0 退出" || bad "--json 错误仍退出 0"
python3 -c "import json,sys; d=json.load(open(sys.argv[1])); assert d['ok'] is False and d['message']" "$TMP/err.json" \
  >/dev/null 2>"$TMP/last_err" && ok "--json 错误输出可解析且含 message" || bad "--json 错误输出不可解析"

echo ""
echo "  ────────────────────────────────────────────────────────"
echo "  通过 $PASS · 失败 $FAIL"
echo ""
[ "$FAIL" -eq 0 ] || exit 1
