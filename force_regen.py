# -*- coding: utf-8 -*-
"""GitHub Actions CI 内强制重跑 gen_aobuta_xlsx.py + gen_aobuta_html.py。

两个生成脚本本来就没有 E:\Report 依赖（gen_aobuta_html.py 只用 openpyxl + 标准库），
但它们把路径写死为 D:\AI\Default\...，CI (ubuntu) 上没有这个盘。
这里的做法：在 CI 上把 D:\AI\Default 挂到等价位置不可行，所以改为
- 在仓库根创建目录结构，用 os.makedirs + 环境变量方式
- 直接用 monkeypatch 方式：runpy 跑脚本前把脚本文本里的路径替换成 CI 上的工作目录
- 跑完后把 index.html 输出到仓库根
"""
import re
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
XLSX_NAME = None
for p in ROOT.glob("*.xlsx"):
    XLSX_NAME = p.name
    break
if not XLSX_NAME:
    print("[force_regen] ERROR: no xlsx found in repo root", file=sys.stderr)
    sys.exit(1)

print(f"[force_regen] using xlsx: {XLSX_NAME}")

for script_name in ("gen_aobuta_xlsx.py", "gen_aobuta_html.py"):
    src_path = ROOT / script_name
    src = src_path.read_text(encoding="utf-8")
    # 把写死的 D:\AI\Default\ 与反斜杠路径整体替换为仓库根（CI 与本地语义一致）
    patched = src.replace(r"D:\AI\Default\fhlm.github.io\aobuta", str(ROOT))
    patched = patched.replace(r"D:\AI\Default\aobuta-pilgrimage-site", str(ROOT))
    patched = patched.replace(r"D:\AI\Default", str(ROOT))
    if script_name == "gen_aobuta_html.py":
        # 网页产物固定写为仓库根 index.html（Pages 直接发布）
        patched = re.sub(
            r'OUT = r"[^"]+"',
            'OUT = r"%s"' % (ROOT / "index.html").as_posix(),
            patched,
        )
    tmp = ROOT / f"_ci_{script_name}"
    tmp.write_text(patched, encoding="utf-8")
    print(f"[force_regen] run {script_name}")
    try:
        runpy.run_path(str(tmp), run_name="__main__")
    finally:
        tmp.unlink(missing_ok=True)

out = ROOT / "index.html"
if not out.exists() or out.stat().st_size < 1000:
    print("[force_regen] ERROR: index.html not regenerated", file=sys.stderr)
    sys.exit(1)
print(f"[force_regen] OK, index.html = {out.stat().st_size} bytes")
