# -*- coding: utf-8 -*-
"""从青ブタ巡礼 xlsx 生成移动端友好的单文件静态网页 v7（时间轴重构版）
读取 xlsx 全部 Sheet，输出单文件 HTML（自包含 CSS/JS，无外部依赖）。
主时间表 → 竖向时间轴布局；其余 Sheet → 徽章卡片布局。
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


def cell_kind(v):
    """返回 (类型, chip类)。类型: ok/star3/star2/warn/''"""
    s = str(v) if v is not None else ""
    if "★★★" in s:
        return "star3"
    if "★★" in s:
        return "star2"
    if "✓" in s:
        return "ok"
    if "⚠" in s:
        return "warn"
    return ""


def sheet_rows(ws):
    rows = []
    for row in ws.iter_rows(values_only=True):
        cells = list(row)
        if all(c is None or str(c).strip() == "" for c in cells):
            continue
        rows.append(cells)
    return rows


def render_timeline_item(r, header, is_time):
    """生成一个 .item。is_time=True 时带左侧时间轨。"""
    tds = []
    tm = TIME_RE.match(str(r[0]).strip()) if r[0] is not None else ""
    cls = ["item"]
    tr_attrs = ""
    title_parts = []
    fields = []
    chips = []
    sub = None

    if is_time and tm:
        cls.append("tl")
        st = "%02d:%02d" % (int(tm.group(1)), int(tm.group(2)))
        tr_attrs = ' data-start="%s"' % st
        if tm.group(3):
            tr_attrs += ' data-end="%02d:%02d"' % (int(tm.group(3)), int(tm.group(4)))
        time_txt = str(r[0]).strip()
        act = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
        title_parts = [time_txt, act] if act else [time_txt]
        start_idx = 2
        # 额外字段（第3列起）
        for i in range(start_idx, len(r)):
            c = r[i]
            txt = "" if c is None else str(c).strip()
            if not txt:
                continue
            k = cell_kind(c)
            if k:
                chips.append((k, txt))
            else:
                label = header[i] if i < len(header) else ""
                fields.append((label, txt))
    else:
        cls.append("solo")
        c0 = str(r[0]).strip() if r[0] is not None else ""
        title_parts = [c0]
        if len(r) > 1 and r[1] is not None and str(r[1]).strip():
            c1 = str(r[1]).strip()
            if len(c1) <= 18:
                sub = c1
                start_idx = 2
            else:
                start_idx = 1
        else:
            start_idx = 2
        for i in range(start_idx, len(r)):
            c = r[i]
            txt = "" if c is None else str(c).strip()
            if not txt:
                continue
            k = cell_kind(c)
            if k:
                chips.append((k, txt))
            else:
                label = header[i] if i < len(header) else ""
                fields.append((label, txt))

    data_title = " · ".join(p for p in title_parts if p)
    tr_attrs += ' data-title="%s"' % esc(data_title)

    # ---- 组装 HTML ----
    rail = ""
    if "tl" in cls:
        st = "%02d:%02d" % (int(tm.group(1)), int(tm.group(2)))
        rail = '<div class="rail"><span class="t-start">%s</span>' % st
        if tm.group(3):
            rail += '<span class="t-end">%02d:%02d</span>' % (int(tm.group(3)), int(tm.group(4)))
        rail += "</div>"

    ttl = esc(title_parts[0]) if title_parts else ""
    if "tl" in cls and len(title_parts) > 1:
        ttl = esc(title_parts[1])  # 时间进轨道，卡片标题=活动名
    sub_html = '<span class="sub">%s</span>' % esc(sub) if sub else ""
    chev = '<span class="chev">›</span>'

    f_html = "".join(
        '<div class="f"><span class="fl">%s</span><span class="fv">%s</span></div>' % (esc(l), esc(v))
        for l, v in fields
    )
    c_html = "".join('<span class="chip %s">%s</span>' % (k, esc(t)) for k, t in chips)
    chips_html = '<div class="chips">%s</div>' % c_html if c_html else ""

    return (
        '<div class="%s"%s>%s<div class="card">'
        '<div class="card-head"><span class="ttl">%s</span>%s%s</div>'
        '<div class="fields">%s</div>%s'
        "</div></div>"
    ) % (" ".join(cls), tr_attrs, rail, ttl, sub_html, chev, f_html, chips_html)


def sheet_to_html(ws, timeline=False):
    rows = sheet_rows(ws)
    if not rows:
        return ""
    title = ""
    header_idx = 0
    if len(rows) >= 2:
        title = esc(rows[0][0]) if rows[0][0] else ""
        header_idx = 1
    header = rows[header_idx]
    body = rows[header_idx + 1:]
    parts = []
    if title:
        parts.append('<div class="desc">%s</div>' % title)
    parts.append('<div class="%s">' % ("tl" if timeline else "list"))
    for r in body:
        first = str(r[0]).strip() if r[0] is not None else ""
        is_time = bool(TIME_RE.match(first))
        parts.append(render_timeline_item(r, header, is_time))
    parts.append("</div>")
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
        inner = sheet_to_html(ws, timeline=(i == 1))
        sections.append(
            f'<section id="{anchor}"{hidden}><h2><span class="ico">{icon}</span>{esc(name)}</h2>\n{inner}</section>'
        )
    page = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#1d4ed8">
<title>青ブタ×孤独摇滚 圣地巡礼时间表 2026-11-22</title>
<style>
:root {{
  --bg:#f2f5fa; --card:#fff; --ink:#16232e; --muted:#64748b;
  --primary:#2563eb; --primary-d:#1d4ed8; --sky:#38bdf8;
  --rail:#c7dbf7;
  --ok-bg:#e7f8f1; --ok-ink:#0b7a4b; --ok-bd:#bfe9d7;
  --amber-bg:#fef3df; --amber-ink:#93620a; --amber-bd:#f4dfae;
  --purple-bg:#f1e9fd; --purple-ink:#6d3fc4; --purple-bd:#e2d4fa;
  --warn-bg:#fff1e4; --warn-ink:#b34a0a; --warn-bd:#ffd9b8;
  --radius:16px;
}}
* {{ box-sizing:border-box; }}
html {{ scroll-behavior:smooth; }}
body {{ margin:0; background:var(--bg); color:var(--ink);
  font-family:-apple-system,"PingFang SC","HarmonyOS Sans SC","Microsoft YaHei",sans-serif;
  font-size:14.5px; -webkit-text-size-adjust:100%; }}
.page {{ max-width:760px; margin:0 auto; }}

/* ---------- 头部 ---------- */
header {{ position:relative; background:linear-gradient(140deg,#0b2f6b 0%,#1d4ed8 52%,#38bdf8 130%);
  color:#fff; padding:26px 20px 46px; overflow:hidden; }}
header::before {{ content:''; position:absolute; right:-70px; top:-70px; width:230px; height:230px;
  border-radius:50%; background:rgba(255,255,255,.10); }}
header::after {{ content:''; position:absolute; right:-30px; top:60px; width:120px; height:120px;
  border-radius:50%; background:rgba(255,255,255,.08); }}
.hx-tag {{ display:inline-block; font-size:12px; font-weight:700; letter-spacing:.5px;
  background:rgba(255,255,255,.18); border:1px solid rgba(255,255,255,.35);
  padding:4px 12px; border-radius:999px; margin-bottom:12px; }}
header h1 {{ margin:0; font-size:23px; line-height:1.35; font-weight:800; letter-spacing:.5px;
  text-shadow:0 2px 8px rgba(7,35,80,.35); }}
header p {{ margin:10px 0 0; font-size:12.5px; opacity:.92; line-height:1.6; }}
.wave {{ position:absolute; left:0; right:0; bottom:-1px; line-height:0; }}
.wave svg {{ width:100%; height:32px; display:block; }}

/* ---------- 导航 ---------- */
nav {{ position:sticky; top:0; z-index:20; background:rgba(255,255,255,.9);
  backdrop-filter:blur(10px); -webkit-backdrop-filter:blur(10px);
  border-bottom:1px solid #e4eaf2; }}
.navin {{ display:flex; gap:8px; overflow-x:auto; padding:10px 14px; scrollbar-width:none; }}
.navin::-webkit-scrollbar {{ display:none; }}
nav a {{ flex:0 0 auto; padding:7px 13px; font-size:13px; font-weight:700; color:#3b5878;
  text-decoration:none; white-space:nowrap; background:#fff; border:1px solid #e2e9f2;
  border-radius:999px; transition:all .15s; }}
nav a.active {{ background:linear-gradient(135deg,var(--primary-d),var(--primary)); color:#fff;
  border-color:transparent; box-shadow:0 3px 10px rgba(37,99,235,.35); }}

/* ---------- 区块 ---------- */
section {{ padding:14px 14px 4px; }}
section.hidden {{ display:none; }}
h2 {{ display:flex; align-items:center; gap:9px; font-size:16px; font-weight:800; color:#0f2b52;
  margin:2px 0 12px; }}
h2 .ico {{ width:27px; height:27px; border-radius:9px; flex:0 0 auto;
  background:linear-gradient(135deg,var(--primary-d),var(--sky)); color:#fff;
  display:flex; align-items:center; justify-content:center; font-size:14px;
  box-shadow:0 3px 8px rgba(37,99,235,.3); }}
.desc {{ font-size:12.5px; color:var(--muted); background:#f7fafe; border-left:3px solid #cfe0f7;
  padding:7px 11px; border-radius:8px; margin:0 0 12px; line-height:1.65; }}

/* ---------- 时间轴（主时间表） ---------- */
.tl {{ position:relative; padding:4px 0 8px; }}
.tl::before {{ content:''; position:absolute; left:28px; top:10px; bottom:14px; width:2px;
  background:linear-gradient(180deg,var(--rail) 75%,rgba(199,219,247,.15)); border-radius:2px; }}
.item {{ position:relative; display:flex; gap:12px; padding:0 0 10px; cursor:pointer; }}
.item .rail {{ flex:0 0 56px; text-align:center; padding-top:10px; z-index:1; }}
.t-start {{ display:block; font-weight:800; font-size:12.5px; color:var(--primary-d);
  background:#eaf2fe; border:1px solid #d3e4fd; padding:4px 0; border-radius:10px;
  font-variant-numeric:tabular-nums; transition:all .2s; }}
.t-end {{ display:block; font-size:10px; color:var(--muted); margin-top:3px;
  font-variant-numeric:tabular-nums; }}

/* ---------- 卡片 ---------- */
.card {{ flex:1; min-width:0; background:var(--card); border-radius:var(--radius);
  border:1px solid #edf1f7; padding:12px 14px;
  box-shadow:0 1px 3px rgba(22,40,60,.05),0 4px 14px rgba(22,40,60,.04);
  transition:transform .12s ease, box-shadow .15s ease; }}
.item:active .card {{ transform:scale(.985); }}
.card-head {{ display:flex; align-items:baseline; gap:8px; }}
.ttl {{ font-weight:800; font-size:15px; line-height:1.45; color:var(--ink); word-break:break-word; }}
.sub {{ font-size:12px; color:var(--muted); font-weight:700; flex:0 0 auto; }}
.chev {{ margin-left:auto; color:#b6c4d4; font-size:20px; line-height:1; flex:0 0 auto;
  font-weight:400; transform:translateY(-1px); }}
.fields {{ margin-top:5px; }}
.f {{ font-size:13px; line-height:1.65; color:#33465c; }}
.fl {{ color:#8496aa; font-size:11px; font-weight:800; letter-spacing:.3px; margin-right:7px; }}
.fl::after {{ content:'·'; margin-left:5px; color:#c3cfdd; font-weight:400; }}
.chips {{ margin-top:8px; display:flex; flex-wrap:wrap; gap:6px; }}
.chip {{ font-size:11px; font-weight:800; padding:3px 10px; border-radius:999px; line-height:1.5; }}
.chip.ok {{ color:var(--ok-ink); background:var(--ok-bg); border:1px solid var(--ok-bd); }}
.chip.star3 {{ color:var(--amber-ink); background:var(--amber-bg); border:1px solid var(--amber-bd); }}
.chip.star2 {{ color:var(--purple-ink); background:var(--purple-bg); border:1px solid var(--purple-bd); }}
.chip.warn {{ color:var(--warn-ink); background:var(--warn-bg); border:1px solid var(--warn-bd); }}

/* 非 s1 的列表（无轨道） */
.list .item {{ padding:0 0 10px; }}

/* ---------- 当前时段高亮 ---------- */
.item.now .card {{ border-color:#93c5fd;
  box-shadow:0 0 0 3px rgba(37,99,235,.10),0 4px 16px rgba(37,99,235,.12);
  background:linear-gradient(180deg,#f4f9ff,#fff 60%); }}
.item.now .t-start {{ background:var(--primary); border-color:var(--primary); color:#fff; }}
@keyframes locatePulse {{
  0% {{ box-shadow:0 0 0 0 rgba(37,99,235,.35); }}
  100% {{ box-shadow:0 0 0 14px rgba(37,99,235,0); }} }}
.item.flash .card {{ animation:locatePulse 1s ease-out 2; }}

/* ---------- 悬浮按钮 ---------- */
.fab {{ position:fixed; right:16px; bottom:calc(18px + env(safe-area-inset-bottom)); z-index:50;
  background:linear-gradient(135deg,var(--primary-d),var(--primary) 60%,var(--sky)); color:#fff;
  border:none; border-radius:999px; padding:12px 18px; font-size:13.5px; font-weight:800;
  box-shadow:0 6px 18px rgba(29,78,216,.42); cursor:pointer; letter-spacing:.3px;
  font-variant-numeric:tabular-nums; }}
.fab:active {{ transform:scale(.95); }}

/* ---------- 弹层 ---------- */
.modal {{ position:fixed; inset:0; background:rgba(9,25,47,.5); display:none; z-index:99;
  align-items:flex-end; justify-content:center; backdrop-filter:blur(2px); }}
.modal.open {{ display:flex; }}
.modal-card {{ background:#fff; width:100%; max-width:560px; max-height:82vh; overflow-y:auto;
  border-radius:20px 20px 0 0; padding:10px 18px calc(22px + env(safe-area-inset-bottom));
  box-shadow:0 -6px 30px rgba(0,0,0,.2); }}
.modal-card::before {{ content:''; display:block; width:38px; height:4px; border-radius:4px;
  background:#dde5ee; margin:0 auto 12px; }}
.modal-head {{ display:flex; align-items:center; justify-content:space-between; gap:10px;
  margin-bottom:4px; padding-bottom:10px; border-bottom:1px solid #f1f5f9; }}
#m-title {{ font-size:16.5px; font-weight:800; color:#0f2b52; line-height:1.45; }}
#m-close {{ border:none; background:#eef4fa; color:var(--primary-d); width:32px; height:32px;
  flex:0 0 auto; border-radius:50%; font-size:17px; cursor:pointer; line-height:1; }}
.m-field {{ padding:11px 0; border-bottom:1px solid #f4f7fb; }}
.m-field:last-child {{ border-bottom:none; }}
.m-label {{ font-size:11px; font-weight:800; color:#8496aa; letter-spacing:.4px; margin-bottom:3px; }}
.m-val {{ font-size:14px; line-height:1.7; color:#26384c; }}
@media (min-width:641px) {{
  .modal {{ align-items:center; padding:20px; }}
  .modal-card {{ border-radius:20px; }}
  .modal-card::before {{ display:none; }}
}}

footer {{ text-align:center; color:#8ea1b5; font-size:11.5px; padding:18px 0 calc(26px + env(safe-area-inset-bottom)); line-height:1.7; }}
</style>
</head>
<body>
<header>
  <div class="hx">
    <div class="hx-tag">2026-11-22（周日）</div>
    <h1>青ブタ × 孤独摇滚<br>圣地巡礼时间表</h1>
    <p>住吉 07:00 出发 → 藤沢 → 七里ヶ浜 → 江ノ島 日落 → 湘南モノレール → 辻堂夜景晚餐<br>✈ 羽田 T3 · 23:55 起飞 · 21:40 前到达</p>
  </div>
  <div class="wave"><svg viewBox="0 0 1440 70" preserveAspectRatio="none"><path d="M0,42 C240,68 480,12 760,32 C1000,48 1220,16 1440,44 L1440,70 L0,70 Z" fill="#f2f5fa"/></svg></div>
</header>
<nav><div class="navin">{''.join(nav)}</div></nav>
<button id="fab-now" class="fab">📍 现在</button>
<div id="modal" class="modal"><div class="modal-card"><div class="modal-head"><span id="m-title"></span><button id="m-close" aria-label="关闭">×</button></div><div id="m-body"></div></div></div>
<div class="page">
{chr(10).join(sections)}
</div>
<footer>由 gen_aobuta_html.py 从 Excel 自动生成 · 数据与 xlsx 同步<br>点卡片看详情 · 点顶部标签切换板块 · 「现在」跳转当前时段</footer>
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
  function openModal(item) {{
    var out = '';
    item.querySelectorAll('.f').forEach(function (f) {{
      var l = f.querySelector('.fl'), v = f.querySelector('.fv');
      var lt = l ? l.textContent.replace(/·\\s*$/, '') : '';
      var vt = v ? v.textContent : '';
      if (!vt) return;
      out += '<div class="m-field"><div class="m-label">' + lt +
             '</div><div class="m-val">' + vt.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</div></div>';
    }});
    var chips = item.querySelectorAll('.chip');
    if (chips.length) {{
      var ct = Array.prototype.map.call(chips, function (c) {{ return c.textContent; }}).join(' / ');
      out += '<div class="m-field"><div class="m-label">标记</div><div class="m-val">' +
             ct.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</div></div>';
    }}
    if (!out) out = '<div class="m-val">（本条无附加说明）</div>';
    mTitle.textContent = item.getAttribute('data-title') || '';
    mBody.innerHTML = out;
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';
  }}
  function closeModal() {{
    modal.classList.remove('open');
    document.body.style.overflow = '';
  }}
  Array.prototype.slice.call(document.querySelectorAll('.item')).forEach(function (item) {{
    item.addEventListener('click', function () {{ openModal(item); }});
  }});
  document.getElementById('m-close').addEventListener('click', closeModal);
  modal.addEventListener('click', function (e) {{ if (e.target === modal) closeModal(); }});
  document.addEventListener('keydown', function (e) {{ if (e.key === 'Escape') closeModal(); }});

  var fab = document.getElementById('fab-now');
  var items = Array.prototype.slice.call(document.querySelectorAll('.item[data-start]'));
  function toMin(s) {{ var p = s.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10); }}
  function nowMin() {{ var d = new Date(); return d.getHours() * 60 + d.getMinutes(); }}
  function fmtNow() {{ var d = new Date(); return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }}
  function targetItem() {{
    var m = nowMin(), cur = null, i;
    for (i = 0; i < items.length; i++) {{
      var s = toMin(items[i].getAttribute('data-start'));
      var e = items[i].getAttribute('data-end');
      e = e ? toMin(e) : s + 45;
      if (s <= m && m < e) cur = items[i];
    }}
    if (cur) return cur;
    for (i = 0; i < items.length; i++) {{
      if (toMin(items[i].getAttribute('data-start')) > m) return items[i];
    }}
    return items.length ? items[items.length - 1] : null;
  }}
  function markNow() {{
    items.forEach(function (r) {{ r.classList.remove('now'); }});
    var it = targetItem();
    if (it) it.classList.add('now');
  }}
  function locateNow() {{
    show('s1', false);
    var it = targetItem();
    if (!it) return;
    setTimeout(function () {{
      it.scrollIntoView({{ block: 'center', behavior: 'smooth' }});
      it.classList.remove('flash'); void it.offsetWidth; it.classList.add('flash');
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
