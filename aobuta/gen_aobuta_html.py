# -*- coding: utf-8 -*-
"""从青ブタ巡礼 xlsx 生成移动端友好的单文件静态网页
读取 D:\AI\Default\青ブタ圣地巡礼时间表_2026-11-22.xlsx 全部 Sheet，
输出 D:\AI\Default\青ブタ圣地巡礼时间表_2026-11-22.html（自包含 CSS，无外部依赖）。
"""
import html
import re
from openpyxl import load_workbook

XLSX = r"D:\AI\Default\青ブタ圣地巡礼时间表_2026-11-22.xlsx"
OUT = r"D:\AI\Default\青ブタ圣地巡礼时间表_2026-11-22.html"

TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})(?:\s*[–\-~—]\s*(\d{1,2}):(\d{2}))?")

NAV_ICONS = {
    "主时间表": "🕘",
    "圣地清单": "📍",
    "说明与备选": "📋",
    "预约清单": "✅",
    "孤独摇滚圣地": "🎸",
}


def esc(v):
    return html.escape("" if v is None else str(v))


def cell_class(v):
    s = str(v) if v is not None else ""
    if s.startswith("✓") or "✓" in s:
        return "ok"
    if "★★★" in s:
        return "star3"
    if "★★" in s and "★★★" not in s:
        return "star2"
    if "⚠" in s:
        return "warn"
    return ""


def sheet_to_html(ws):
    rows = []
    for row in ws.iter_rows(values_only=True):
        cells = [c for c in row]
        if all(c is None or str(c).strip() == "" for c in cells):
            continue
        rows.append(cells)
    if not rows:
        return ""
    # 第一个非空行作为标题行（title_row 是合并大标题），第二个作为表头
    title = ""
    header_idx = 0
    if len(rows) >= 2:
        title = esc(rows[0][0]) if rows[0][0] else ""
        header_idx = 1
    header = rows[header_idx]
    body = rows[header_idx + 1:]
    parts = []
    if title:
        parts.append(f'<div class="sheet-title">{title}</div>')
    parts.append('<div class="tablewrap"><table>')
    parts.append("<thead><tr>" + "".join(f'<th>{esc(h)}</th>' for h in header) + "</tr></thead>")
    parts.append("<tbody>")
    for r in body:
        tds = []
        first_txt = "" if r[0] is None else str(r[0]).strip()
        tm = TIME_RE.match(first_txt)
        sub_txt = ""
        if tm and len(r) > 1 and r[1] is not None and str(r[1]).strip():
            sub_txt = str(r[1]).strip()
        tr_attrs = ""
        if tm:
            st = "%02d:%02d" % (int(tm.group(1)), int(tm.group(2)))
            tr_attrs = ' data-start="%s"' % st
            if tm.group(3):
                tr_attrs += ' data-end="%02d:%02d"' % (int(tm.group(3)), int(tm.group(4)))
        for i, c in enumerate(r):
            cls = cell_class(c)
            txt = "" if c is None else str(c).strip()
            if i > 0 and not txt:
                cls = (cls + " empty").strip()
            if i == 0:
                cls = (cls + " first").strip()
                if tm:
                    cls = (cls + " head").strip()
            if i == 1 and sub_txt:
                cls = (cls + " sub").strip()
            label = esc(header[i]) if i < len(header) else ""
            tds.append(f'<td class="{cls}" data-label="{label}">{esc(c)}</td>')
        parts.append(f"<tr{tr_attrs}>" + "".join(tds) + "</tr>")
    parts.append("</tbody></table></div>")
    return "\n".join(parts)


def main():
    wb = load_workbook(XLSX, data_only=True)
    sections = []
    nav = []
    for i, name in enumerate(wb.sheetnames, 1):
        ws = wb[name]
        icon = NAV_ICONS.get(name, "📄")
        anchor = f"s{i}"
        nav.append(f'<a href="#{anchor}" data-s="{anchor}">{icon} {esc(name)}</a>')
        hidden = ' class="hidden"' if i != 1 else ""
        sections.append(
            f'<section id="{anchor}"{hidden}><h2>{icon} {esc(name)}</h2>\n{sheet_to_html(ws)}</section>'
        )
    page = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>青ブタ×孤独摇滚 圣地巡礼时间表 2026-11-22</title>
<style>
:root {{ --c1:#0f6bb8; --bg:#f4f7fb; --card:#fff; --ok:#e8f5e9; --ok-b:#2e7d32;
  --star3:#fff8e1; --star3-b:#f9a825; --star2:#f3e5f5; --star2-b:#8e24aa; --warn:#fff3e0; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:#1c2733;
  font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif; font-size:15px; }}
header {{ background:linear-gradient(135deg,#0f6bb8,#0a4d85); color:#fff; padding:18px 16px 14px; }}
header h1 {{ margin:0 0 4px; font-size:19px; }}
header p {{ margin:0; font-size:12.5px; opacity:.85; }}
nav {{ position:sticky; top:0; z-index:9; background:#fff; display:flex; gap:6px; overflow-x:auto;
  border-bottom:1px solid #e1e8ef; padding:8px 10px; }}
nav a {{ flex:0 0 auto; padding:8px 12px; font-size:13.5px; color:var(--c1); text-decoration:none;
  white-space:nowrap; font-weight:600; background:#eef4fa; border-radius:999px; }}
nav a.active {{ background:var(--c1); color:#fff; }}
section {{ padding:14px 12px 4px; }}
section.hidden {{ display:none; }}
h2 {{ font-size:16.5px; margin:2px 2px 10px; color:var(--c1); }}
.sheet-title {{ font-size:12.5px; color:#5a6b7d; margin:0 2px 8px; line-height:1.5; }}
.tablewrap {{ background:var(--card); border-radius:12px; overflow-x:auto;
  box-shadow:0 1px 4px rgba(16,42,67,.08); }}
table {{ border-collapse:collapse; width:100%; font-size:13.5px; min-width:560px; }}
th {{ background:#eef4fa; color:#0a4d85; text-align:left; padding:9px 10px;
  position:sticky; top:50px; font-size:12.5px; white-space:nowrap; }}
td {{ padding:8px 10px; border-top:1px solid #eef2f6; vertical-align:top; line-height:1.55; }}
td.first {{ font-weight:700; white-space:nowrap; color:#0a4d85; }}
td.ok {{ background:var(--ok); }}
td.star3 {{ background:var(--star3); }}
td.star2 {{ background:var(--star2); }}
td.warn {{ background:var(--warn); }}
footer {{ text-align:center; color:#8296a9; font-size:11.5px; padding:16px 0 26px; }}
tbody tr {{ cursor:pointer; }}
tbody tr:active {{ background:#f0f6fc; }}
tbody tr.now {{ outline:2px solid var(--c1); outline-offset:-2px; background:#eaf4fc; }}
@keyframes locatePulse {{ 0% {{ outline-color:rgba(15,107,184,.9); box-shadow:0 0 0 8px rgba(15,107,184,.30); }}
  100% {{ box-shadow:0 0 0 0 rgba(15,107,184,0); }} }}
tbody tr.flash {{ animation:locatePulse 1.1s ease-out 2; }}
.fab {{ position:fixed; right:14px; bottom:18px; z-index:50; background:var(--c1); color:#fff;
  border:none; border-radius:999px; padding:11px 16px; font-size:13.5px; font-weight:700;
  box-shadow:0 4px 14px rgba(15,107,184,.45); cursor:pointer; }}
.fab:active {{ transform:scale(.96); }}
/* ---- 行程说明弹层 ---- */
.modal {{ position:fixed; inset:0; background:rgba(10,30,50,.45); display:none; z-index:99;
  align-items:flex-end; justify-content:center; }}
.modal.open {{ display:flex; }}
.modal-card {{ background:#fff; width:100%; max-width:560px; max-height:82vh; overflow-y:auto;
  border-radius:16px 16px 0 0; padding:14px 16px 22px; box-shadow:0 -4px 24px rgba(0,0,0,.18); }}
.modal-head {{ display:flex; align-items:center; justify-content:space-between; gap:10px;
  margin-bottom:12px; padding-bottom:10px; border-bottom:1px solid #eef2f6; }}
#m-title {{ font-size:16.5px; font-weight:700; color:var(--c1); line-height:1.4; }}
#m-close {{ border:none; background:#eef4fa; color:#0a4d85; width:32px; height:32px; flex:0 0 auto;
  border-radius:50%; font-size:17px; cursor:pointer; line-height:1; }}
.m-field {{ margin:0 0 12px; }}
.m-label {{ font-size:11px; font-weight:700; color:#8296a9; letter-spacing:.4px; margin-bottom:2px; }}
.m-val {{ font-size:14.5px; line-height:1.6; }}
@media (min-width: 641px) {{
  .modal {{ align-items:center; padding:20px; }}
  .modal-card {{ border-radius:16px; }}
}}
/* ---- 手机端：宽表转紧凑卡片式布局，不再横向滚动 ---- */
@media (max-width: 640px) {{
  body {{ font-size:14px; -webkit-text-size-adjust:100%; }}
  section {{ padding:10px 8px 2px; }}
  h2 {{ font-size:15px; margin:0 2px 8px; }}
  nav {{ padding:6px 8px; }}
  nav a {{ padding:7px 11px; font-size:13px; }}
  table {{ min-width:0; font-size:13.5px; }}
  .tablewrap {{ background:transparent; box-shadow:none; border-radius:0; overflow:visible; }}
  thead {{ display:none; }}
  tbody tr {{ display:block; background:var(--card); border-radius:10px; margin:0 0 8px;
    padding:1px 0; box-shadow:0 1px 3px rgba(16,42,67,.08); overflow:hidden; }}
  td {{ display:block; border-top:none; padding:4px 12px; line-height:1.45; }}
  td::before {{ content:attr(data-label)': '; display:inline; font-size:12px; font-weight:700;
    color:#8296a9; margin-right:2px; }}
  td.first {{ white-space:normal; font-size:14.5px; padding:8px 14px 2px; }}
  td.first.head {{ display:inline-block; padding:8px 6px 0 14px; }}
  td.first.head::before {{ display:none; }}
  td.sub {{ display:inline; padding:8px 14px 2px 0; color:#8296a9; font-size:13px; font-weight:600; }}
  td.sub::before {{ display:none; }}
  td.empty {{ display:none; }}
  td.ok, td.star3, td.star2, td.warn {{ margin:3px 6px; border-radius:6px; padding:4px 8px; }}
  footer {{ padding:12px 0 22px; }}
}}
</style>
</head>
<body>
<header>
  <h1>🎶 青ブタ × 孤独摇滚 圣地巡礼</h1>
  <p>2026-11-22（周日）｜住吉出发 → 湘南 → 龟户晚餐 → 羽田 23:55 航班</p>
</header>
<nav>{''.join(nav)}</nav>
<button id="fab-now" class="fab">📍 现在</button>
<div id="modal" class="modal"><div class="modal-card"><div class="modal-head"><span id="m-title"></span><button id="m-close" aria-label="关闭">×</button></div><div id="m-body"></div></div></div>
{chr(10).join(sections)}
<footer>由 gen_aobuta_html.py 从 Excel 自动生成 · 数据与 xlsx 同步 · 点顶部标签切换板块</footer>
<script>
(function () {{
  var links = Array.prototype.slice.call(document.querySelectorAll('nav a'));
  var sections = Array.prototype.slice.call(document.querySelectorAll('section'));
  var activeId = null;
  function show(id, scroll) {{
    var changed = activeId !== id;
    sections.forEach(function (s) {{ s.classList.toggle('hidden', s.id !== id); }});
    links.forEach(function (a) {{ a.classList.toggle('active', a.getAttribute('data-s') === id); }});
    activeId = id;
    if (scroll && changed) window.scrollTo({{ top: 0, behavior: 'smooth' }});
  }}
  links.forEach(function (a) {{
    a.addEventListener('click', function (e) {{
      e.preventDefault();
      var id = a.getAttribute('data-s');
      show(id, true);
      try {{ history.replaceState(null, '', '#' + id); }} catch (err) {{}}
    }});
  }});
  var hash = (location.hash || '').replace('#', '');
  var valid = sections.some(function (s) {{ return s.id === hash; }});
  show(valid ? hash : sections[0].id, false);

  var modal = document.getElementById('modal');
  var mTitle = document.getElementById('m-title');
  var mBody = document.getElementById('m-body');
  function openModal(tr) {{
    var tds = Array.prototype.slice.call(tr.querySelectorAll('td'));
    var title = tds[0].textContent.trim();
    var start = 1;
    if (tds.length > 1) {{
      var v1 = tds[1].textContent.trim();
      if (v1 && v1.length <= 30) {{ title += ' · ' + v1; start = 2; }}
    }}
    var out = '';
    tds.forEach(function (td, i) {{
      if (i < start) return;
      var v = td.textContent.trim();
      if (!v) return;
      out += '<div class="m-field"><div class="m-label">' + (td.getAttribute('data-label') || '') +
             '</div><div class="m-val">' + v.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</div></div>';
    }});
    if (!out) out = '<div class="m-val">（本行无附加说明）</div>';
    mTitle.textContent = title;
    mBody.innerHTML = out;
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';
  }}
  function closeModal() {{
    modal.classList.remove('open');
    document.body.style.overflow = '';
  }}
  Array.prototype.slice.call(document.querySelectorAll('tbody tr')).forEach(function (tr) {{
    tr.addEventListener('click', function () {{ openModal(tr); }});
  }});
  document.getElementById('m-close').addEventListener('click', closeModal);
  modal.addEventListener('click', function (e) {{ if (e.target === modal) closeModal(); }});
  document.addEventListener('keydown', function (e) {{ if (e.key === 'Escape') closeModal(); }});

  var fab = document.getElementById('fab-now');
  var rows = Array.prototype.slice.call(document.querySelectorAll('tr[data-start]'));
  function toMin(s) {{ var p = s.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10); }}
  function nowMin() {{ var d = new Date(); return d.getHours() * 60 + d.getMinutes(); }}
  function fmtNow() {{ var d = new Date(); return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }}
  function targetRow() {{
    var m = nowMin(), cur = null, i;
    for (i = 0; i < rows.length; i++) {{
      var s = toMin(rows[i].getAttribute('data-start'));
      var e = rows[i].getAttribute('data-end');
      e = e ? toMin(e) : s + 45;
      if (s <= m && m < e) cur = rows[i];
    }}
    if (cur) return cur;
    for (i = 0; i < rows.length; i++) {{
      if (toMin(rows[i].getAttribute('data-start')) > m) return rows[i];
    }}
    return rows.length ? rows[rows.length - 1] : null;
  }}
  function markNow() {{
    rows.forEach(function (r) {{ r.classList.remove('now'); }});
    var tr = targetRow();
    if (tr) tr.classList.add('now');
  }}
  function locateNow() {{
    show('s1', false);
    var tr = targetRow();
    if (!tr) return;
    setTimeout(function () {{
      tr.scrollIntoView({{ block: 'center', behavior: 'smooth' }});
      tr.classList.remove('flash'); void tr.offsetWidth; tr.classList.add('flash');
    }}, 180);
  }}
  fab.addEventListener('click', locateNow);
  function tick() {{ fab.textContent = '📍 ' + fmtNow() + ' 现在'; }}
  tick();
  markNow();
  setInterval(function () {{ tick(); markNow(); }}, 30000);
}})();
</script>
</body>
</html>
"""
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(page)
    print("saved:", OUT)


if __name__ == "__main__":
    main()
