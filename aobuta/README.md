# 青ブタ × 孤独摇滚 圣地巡礼时间表（2026-11-22）

单文件静态网页（`index.html`），手机友好，自包含 CSS/JS，无外部依赖。

> 通用规划流程见 [`../doc/行程规划工作流.md`](../doc/行程规划工作流.md)——本文档只写本项目特有信息。

## 在线查看

GitHub Pages 部署后访问：`https://<你的用户名>.github.io/<仓库名>/`

## 目录结构

```
index.html                 # 最终网页（Pages 入口）
青ブタ圣地巡礼时间表_2026-11-22.xlsx   # 数据源（4+1 Sheet）
gen_aobuta_xlsx.py         # xlsx 生成脚本（openpyxl）
gen_aobuta_html.py         # xlsx → html 生成脚本
force_regen.py             # CI 入口：依次跑上面两个脚本
.github/workflows/regen.yml # 推送 xlsx 后自动重建
```

## 更新流程

### 方式一：改 Excel 后本地重发（推荐）

1. 在 D:\AI\Default\ 改 gen_aobuta_xlsx.py → `python gen_aobuta_xlsx.py`
2. `python gen_aobuta_html.py` 重新生成 HTML
3. 把 xlsx 和 index.html 复制进本仓库根目录，git push

### 方式二：只推 Excel，让 GitHub 自动重建

直接把新 xlsx push 上来，Actions 会自动跑 force_regen.py 重建 index.html 并提交
（也可在 Actions 页手动 Run workflow）。
