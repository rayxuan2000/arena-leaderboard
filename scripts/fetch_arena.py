#!/usr/bin/env python3
"""
fetch_arena.py — 直接从 arena.ai 抓取全部榜单数据
=================================================================
不依赖任何第三方 API 或仓库（如 wulong.dev / oolong-tea repo）。

为什么不会截断:
  arena.ai 的榜单页是服务端渲染（SSR）的，完整的标准 <table> 就在初始 HTML 里，
  一次普通 HTTP GET 就能拿到全部模型行。无需 JS 渲染、无需 LLM 抽取、无字符上限，
  因此不存在"长榜被砍掉尾部"的问题（第三方仓库的截断来自它把页面渲染成 markdown
  后 text[:15000] 再喂给 LLM 抽取，长榜会溢出丢行）。

解析策略（自动降级）:
  1) 直连 arena.ai 原始 HTML，用 BeautifulSoup 按【表头驱动】确定性解析整张表
     ← 零第三方依赖，默认路径
  2) 仅当 (1) 解析不到且显式 --allow-jina 时，回退 Jina Reader 渲染成完整
     markdown 再按同样的表头逻辑解析（Jina 是免密钥通用渲染代理，仅作兜底）

表头驱动的意义:
  不同榜单列结构不同。text/code/vision/document/search 及图像视频生成榜是
  `Rank | Rank Spread | Model | Score(Elo±CI) | Votes | Price | Context`；
  agent 榜则是 `Rank | Model | Net Improvement | ... | Sessions`，没有 Elo/Votes。
  脚本读取表头来定位列，核心 7 字段对齐既有管线，其余列原样存入 extra 不丢数据。

输出（schema 与 export_arena_excel.py 一致）:
  <out-dir>/<slug>.json            每个榜单一个文件（最新）
  <out-dir>/arena_data.json        全部榜单合并
  <out-dir>/history/<date>/<slug>.json   按北京日期归档（可选 --archive）

用法:
  python fetch_arena.py                  # 抓全部榜单到 ./data
  python fetch_arena.py --only text      # 只抓某个榜单
  python fetch_arena.py --out-dir data --archive
  python fetch_arena.py --allow-jina     # 主路径失败时允许 Jina 兜底
"""

import argparse
import json
import re
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

import requests
from bs4 import BeautifulSoup

ARENA_BASE = "https://arena.ai/leaderboard"
JINA_BASE = "https://r.jina.ai/"

KNOWN_SLUGS = [
    "text", "code", "vision", "document", "search", "agent",
    "text-to-image", "image-edit",
    "text-to-video", "image-to-video", "video-edit",
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# 已知厂商名（含多词与带点的），用于把"模型名粘厂商名"的单元格切开。
# 按长度降序，优先匹配更长的（如 "Stability AI" 先于 "AI"）。
_VENDORS = [
    "Black Forest Labs", "Stability AI", "Tencent Hunyuan",
    "Luma AI", "Genmo AI", "Alibaba-ATH", "Alibaba",
    "Anthropic", "OpenAI", "Google", "Meta", "xAI", "Qwen",
    "Baidu", "Xiaomi", "Moonshot", "DeepSeek", "MiniMax", "Kuaishou",
    "Perplexity", "Adobe", "Runway", "Ideogram", "Midjourney", "Mistral",
    "Cohere", "Microsoft", "Amazon", "NVIDIA", "Nvidia", "Bytedance",
    "ByteDance", "Tencent", "Reka", "AI21", "Databricks", "IBM", "Reve",
    "Recraft", "Luma", "Pika", "Genmo", "Z.ai", "01.AI", "Zhipu",
    "Snowflake", "HiDream", "KlingAI", "Kling", "Vidu", "Hailuo",
    "Pixverse", "Pruna", "Kandinsky", "Lightricks",
]
KNOWN_VENDORS = sorted(set(_VENDORS), key=len, reverse=True)

# 这些厂商名同时是其型号品牌前缀（如 DeepSeek V4 Pro / Minimax M2.7），
# 解析无链接单元格时不可把前置的厂商词当徽标删掉。
SELF_BRANDED = {"DeepSeek", "MiniMax", "Minimax"}

# ── 解析用正则 ──
SCORE_RE = re.compile(r"(\d{3,4})\s*±\s*(\d+)")          # 1503±4（对称，text/视频等榜）
SCORE_ASYM_RE = re.compile(r"(\d{3,4})\s*\+\s*(\d+)\s*/\s*[−\-]\s*(\d+)")  # 1567+9/-9（非对称，code 榜）
SCORE_ONLY_RE = re.compile(r"\b(\d{3,4})\b")              # 退化：无置信区间
DATE_RE = re.compile(r"\b([A-Z][a-z]{2}\s+\d{1,2},\s+\d{4})\b")  # Jun 4, 2026
REL_DATE_RE = re.compile(r"\b(\d+\s+(?:hour|day|week|month)s?\s+ago)\b", re.I)
# markdown 链接：[text](url "title") —— 同时拿 text 与可选 title
MD_LINK_RE = re.compile(r'\[([^\]]+)\]\(\s*[^)\s]*(?:\s+"([^"]*)")?\s*\)')


def http_get(url: str, timeout: int = 45) -> str:
    last = None
    for attempt in range(3):
        try:
            resp = requests.get(
                url,
                headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"},
                timeout=timeout,
            )
            resp.raise_for_status()
            return resp.text
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"  [retry {attempt+1}/3] {url}: {e}", file=sys.stderr)
            time.sleep(3 * (attempt + 1))
    raise RuntimeError(f"GET failed after 3 attempts: {url} ({last})")


def normalize_license(lic):
    """归一为 'proprietary' | 'open' | None。"""
    if not isinstance(lic, str) or not lic.strip():
        return None
    s = lic.strip().lower()
    if "proprietary" in s or "closed" in s:
        return "proprietary"
    if s in ("open", "open source", "open-source") or any(
        kw in s for kw in ("mit", "apache", "bsd", "gpl", "llama", "gemma",
                           "cc-", "cc0", "community", "openrail",
                           "non-commercial", "noncommercial")
    ):
        return "open"
    return None


def _int_or_none(text):
    """从 '38,559' / '21,520' 取整数；只用于 votes/sessions 这类纯计数列。"""
    if not text:
        return None
    head = text.split("±")[0]
    digits = re.sub(r"[^\d]", "", head)
    return int(digits) if digits else None


def _parse_score_cell(text):
    """从 '1503±4' / '1567+9/-9' / '1489±6 Preliminary' / '1503' 提取 (score, ci)。
    code 榜用 '+上界/-下界' 写法；对称时 ci 取该值，非对称时取较大的界（保守的 ± 估计）。"""
    if not text:
        return None, None
    m = SCORE_RE.search(text)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = SCORE_ASYM_RE.search(text)
    if m:
        up, lo = int(m.group(2)), int(m.group(3))
        return int(m.group(1)), (up if up == lo else max(up, lo))
    m2 = SCORE_ONLY_RE.search(text)
    return (int(m2.group(1)), None) if m2 else (None, None)


def _vendor_suffix(s):
    """取字符串结尾处的厂商名（已知厂商大小写不敏感优先；否则回退末尾大写 token）。"""
    s = s.strip()
    low = s.lower()
    for v in KNOWN_VENDORS:                 # 已按长度降序，优先匹配更长的多词厂商
        if low.endswith(v.lower()):
            return s[len(s) - len(v):]      # 返回原文大小写（如 'lightricks'）
    toks = s.split()
    if toks and re.match(r"[A-Z0-9]", toks[-1]):
        return toks[-1]
    return None


def split_model_vendor_license(cell_text, link_model=None):
    """
    从模型单元格文本解析 (model, vendor, license)。兼容两种真实格式:
      A) 有链接（text/code 等榜）: link_model 给出干净模型名；尾部 'Vendor · License'
         例 '... claude-opus-4-6-thinking ... Anthropic · Proprietary'
      B) 无链接（agent 榜）: 'GPT 5.5 (High)OpenAI · Proprietary'，模型名与厂商粘连
      许可后可能还有托管商: 'GLM 5.1Z.ai · MIT · SiliconFlow'（取第一段为 license）
    """
    text = re.sub(r"\s+", " ", (cell_text or "")).strip()
    model = (link_model or "").strip() or None
    vendor = lic = None

    if "·" in text or "•" in text:
        parts = re.split(r"\s*[·•]\s*", text)
        head = parts[0].strip()              # 模型(可能粘厂商)
        lic = parts[1].strip() if len(parts) > 1 else None  # 第一段=许可
        vendor = _vendor_suffix(head)
        if model is None:                    # 无链接 → 从 head 去掉尾部厂商
            if vendor and head.endswith(vendor):
                model = head[:-len(vendor)].strip() or head
            else:
                model = head
    else:
        if model is None:
            model = text or None

    # 去掉模型名里残留的、与厂商重复的徽标
    if vendor and model:
        # 尾部粘连的厂商徽标：'...ProDeepSeek' / '...(Thinking)Anthropic'
        if model.endswith(vendor):
            model = model[:-len(vendor)].strip()
        # 前置文字徽标（Anthropic/Meta 等会在型号前渲染厂商名）——
        # 仅当厂商名不是该型号的品牌前缀时才删（避免把 'DeepSeek V4 Pro' 删成 'V4 Pro'）
        if vendor not in SELF_BRANDED and model.startswith(vendor + " "):
            model = model[len(vendor):].strip()
    model = (model or "").strip(" ·-") or None
    return model, vendor, lic


# 计算 extra 时要跳过的列（已并入核心字段或无需重复存储）
def _build_extra(headers, celltext, used_idxs):
    extra = {}
    for i, h in enumerate(headers):
        if i in used_idxs or not h:
            continue
        val = celltext(i)
        if val:
            extra[h] = val
    return extra


def parse_html_table(html: str):
    """
    主解析：在 SSR 的 HTML 里找到榜单 <table>，按表头定位列，逐行确定性抽取。
    自动适配不同列结构（Elo 榜 / agent 行为指标榜）。
    """
    soup = BeautifulSoup(html, "html.parser")
    best, best_key = [], (-1, -1)

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header_cells = rows[0].find_all(["th", "td"])
        headers = [c.get_text(" ", strip=True).lower() for c in header_cells]
        if not headers or not any("model" in h for h in headers):
            continue

        def find_col(pred):
            for i, h in enumerate(headers):
                if pred(h):
                    return i
            return None

        c_rank = find_col(lambda h: h == "rank")
        if c_rank is None:
            c_rank = find_col(lambda h: h.startswith("rank") and "spread" not in h)
        c_model = find_col(lambda h: "model" in h)
        c_score = find_col(lambda h: h in ("score", "elo", "rating")
                           or "score" in h or "elo" in h)
        c_votes = find_col(lambda h: "vote" in h)
        c_sessions = find_col(lambda h: "session" in h)
        if c_model is None:
            continue

        # 一个页面可能有多张表（如 Elo 完整 stats 表 vs 子类别排名表）。
        # 用"丰富度"挑出含 分数/票数/价格/上下文 的那张主表，避免误选纯排名表。
        has_price = any(("price" in h or "cost" in h) for h in headers)
        has_context = any("context" in h for h in headers)
        richness = ((c_score is not None) + (c_votes is not None)
                    + has_price + has_context)

        models = []
        for tr in rows[1:]:
            tds = tr.find_all(["td", "th"])
            if len(tds) < 2 or c_model >= len(tds):
                continue

            def celltext(i):
                if i is None or i >= len(tds):
                    return ""
                return re.sub(r"\s+", " ", tds[i].get_text(" ", strip=True)).strip()

            # rank
            rank = None
            rk = re.match(r"\s*(\d{1,4})", celltext(c_rank))
            if rk:
                rank = int(rk.group(1))

            # model / vendor / license
            md_td = tds[c_model]
            link = md_td.find("a")
            link_model = None
            if link is not None:
                link_model = (link.get("title")
                              or link.get_text(" ", strip=True) or "").strip() or None
            model, vendor, lic = split_model_vendor_license(
                celltext(c_model), link_model=link_model)
            if not model:
                continue

            # score / ci
            score, ci = (_parse_score_cell(celltext(c_score))
                         if c_score is not None else (None, None))

            # votes（无 votes 列时用 sessions 充当样本量）
            vote_src = c_votes if c_votes is not None else c_sessions
            votes = _int_or_none(celltext(vote_src)) if vote_src is not None else None

            used = {c_rank, c_model, c_score, vote_src}
            extra = _build_extra(headers, celltext, used)

            rec = {
                "rank": rank, "model": model, "vendor": vendor,
                "license": normalize_license(lic),
                "score": score, "ci": ci, "votes": votes,
            }
            if extra:
                rec["extra"] = extra
            models.append(rec)

        if (richness, len(models)) > best_key:
            best_key = (richness, len(models))
            best = models

    # 去重(rank+model) + 按 rank 排序
    dedup = {}
    for m in best:
        dedup.setdefault((m["rank"], m["model"]), m)
    return sorted(dedup.values(),
                  key=lambda x: (x["rank"] is None, x["rank"] or 0))


def parse_markdown_table(md: str):
    """兜底解析：解析 Jina 渲染出的完整 markdown 表格（同样表头驱动，兼容异构列）。"""
    lines = md.splitlines()
    header_idx = None
    for i, ln in enumerate(lines):
        low = ln.lower()
        if ln.lstrip().startswith("|") and "model" in low and "rank" in low:
            header_idx = i
            break
    if header_idx is None:
        return []

    def split_row(ln):
        return [c.strip() for c in ln.strip().strip("|").split("|")]

    headers = [h.lower() for h in split_row(lines[header_idx])]
    if len(headers) < 3:
        return []

    def find_col(pred):
        for i, h in enumerate(headers):
            if pred(h):
                return i
        return None

    c_rank = None
    for i, h in enumerate(headers):
        if h == "rank" or (h.startswith("rank") and "spread" not in h):
            c_rank = i
            break
    c_model = find_col(lambda h: "model" in h)
    c_score = find_col(lambda h: h in ("score", "elo", "rating")
                       or "score" in h or "elo" in h)
    c_votes = find_col(lambda h: "vote" in h)
    c_sessions = find_col(lambda h: "session" in h)
    if c_model is None:
        return []

    models = []
    for ln in lines[header_idx + 2:]:        # 跳过表头 + 分隔行
        if not ln.lstrip().startswith("|"):
            break
        cells = split_row(ln)
        if len(cells) < len(headers) or all(set(c) <= set("-:") for c in cells):
            continue

        def cell(i):
            return cells[i] if (i is not None and i < len(cells)) else ""

        model_cell = cell(c_model)
        model_cell = re.sub(r"!\[[^\]]*\]\([^)]*\)", " ", model_cell)  # 去掉图片徽标
        lm = MD_LINK_RE.search(model_cell)
        link_model = None
        if lm:
            link_model = (lm.group(2) or lm.group(1) or "").strip() or None
        clean = MD_LINK_RE.sub(lambda m: f" {m.group(2) or m.group(1)} ", model_cell)
        model, vendor, lic = split_model_vendor_license(clean, link_model=link_model)
        if not model:
            continue

        score, ci = (_parse_score_cell(cell(c_score)) if c_score is not None
                     else _parse_score_cell(model_cell))
        vote_src = c_votes if c_votes is not None else c_sessions
        votes = _int_or_none(cell(vote_src)) if vote_src is not None else None

        rank = None
        rk = re.match(r"\s*(\d{1,4})", cell(c_rank))
        if rk:
            rank = int(rk.group(1))

        used = {c_rank, c_model, c_score, vote_src}
        extra = {headers[i]: cell(i) for i in range(len(headers))
                 if i not in used and headers[i] and cell(i)}

        rec = {"rank": rank, "model": model, "vendor": vendor,
               "license": normalize_license(lic),
               "score": score, "ci": ci, "votes": votes}
        if extra:
            rec["extra"] = extra
        models.append(rec)
    return models


def extract_last_updated(text: str):
    """从页面文本提取数据更新时间（'Jun 4, 2026' 或 'X days ago'），可空。"""
    m = REL_DATE_RE.search(text)
    if m:
        return m.group(1)
    m = DATE_RE.search(text)
    return m.group(1) if m else None


def fetch_one(slug: str, allow_jina: bool = False) -> dict:
    url = f"{ARENA_BASE}/{slug}"
    now = datetime.now(timezone.utc)
    models, last_updated, method = [], None, None

    try:
        html = http_get(url)
        models = parse_html_table(html)
        last_updated = extract_last_updated(
            BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
        if models:
            method = "html"
    except Exception as e:  # noqa: BLE001
        print(f"  [html] {slug}: {e}", file=sys.stderr)

    if not models and allow_jina:
        try:
            md = http_get(f"{JINA_BASE}{url}", timeout=75)
            models = parse_markdown_table(md)
            last_updated = last_updated or extract_last_updated(md)
            if models:
                method = "jina"
        except Exception as e:  # noqa: BLE001
            print(f"  [jina] {slug}: {e}", file=sys.stderr)

    return {
        "meta": {
            "leaderboard": slug,
            "source_url": url,
            "fetched_at": now.isoformat(),
            "last_updated": last_updated,
            "model_count": len(models),
            "method": method,
        },
        "models": models,
    }


def main():
    ap = argparse.ArgumentParser(description="直连 arena.ai 抓取榜单数据")
    ap.add_argument("--out-dir", default="data", help="输出目录（默认 ./data）")
    ap.add_argument("--only", action="append", help="只抓指定 slug（可多次）")
    ap.add_argument("--allow-jina", action="store_true",
                    help="主路径失败时允许 Jina Reader 兜底（第三方渲染代理）")
    ap.add_argument("--archive", action="store_true",
                    help="同时按北京日期归档到 history/<date>/")
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    slugs = args.only or KNOWN_SLUGS

    combined, ok, fail = {}, [], []
    for slug in slugs:
        print(f"\n=== {slug} ===", file=sys.stderr)
        result = fetch_one(slug, allow_jina=args.allow_jina)
        n = result["meta"]["model_count"]
        if n > 0:
            (out_dir / f"{slug}.json").write_text(
                json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            combined[slug] = result
            ok.append((slug, n, result["meta"]["method"]))
            print(f"  ✅ {n} models via {result['meta']['method']}", file=sys.stderr)
        else:
            fail.append(slug)
            print("  ❌ no rows parsed", file=sys.stderr)
        time.sleep(1)

    # 合并写入：局部抓取（--only）只更新对应 slug，保留此前已抓的其它榜单，
    # 避免清空合并文件导致下游 export 漏掉未在本次抓取的榜单。
    combined_path = out_dir / "arena_data.json"
    merged = {}
    if combined_path.exists():
        try:
            merged = json.loads(combined_path.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            merged = {}
    merged.update(combined)
    combined_path.write_text(
        json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.archive and combined:
        bj = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d")
        adir = out_dir / "history" / bj
        adir.mkdir(parents=True, exist_ok=True)
        for slug, res in combined.items():
            (adir / f"{slug}.json").write_text(
                json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n{'='*40}", file=sys.stderr)
    print(f"Done: {len(ok)}/{len(slugs)} 榜单", file=sys.stderr)
    for slug, n, method in ok:
        print(f"  {slug}: {n} ({method})", file=sys.stderr)
    if fail:
        print(f"未解析到数据: {', '.join(fail)}", file=sys.stderr)
        if not args.allow_jina:
            print("  提示：可加 --allow-jina 启用兜底渲染再试。", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
