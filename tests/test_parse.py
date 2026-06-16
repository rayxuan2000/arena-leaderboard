#!/usr/bin/env python3
"""
离线解析回归测试 —— 不联网，用真实抓取的 arena.ai 数据构造夹具验证解析器。
运行: python tests/test_parse.py
覆盖: HTML 主路径 + markdown 兜底路径；text/agent/视频 三种列结构；
      以及真实数据里的各种坑（图片徽标、多词/小写厂商、自品牌前缀、
      Preliminary 后缀、千分位、三段许可、Apache/community 归一）。
"""
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("fa", ROOT / "scripts" / "fetch_arena.py")
fa = importlib.util.module_from_spec(spec)
sys.modules["fa"] = fa
spec.loader.exec_module(fa)


def test_text_html():
    html = """<table><thead><tr><th>Rank</th><th>Rank Spread</th><th>Model</th><th>Score</th><th>Votes</th><th>Price $/M</th><th>Context</th></tr></thead><tbody>
<tr><td>1</td><td>13</td><td><span>Anthropic</span><a href="x" title="claude-opus-4-6-thinking">claude-opus-4-6-thinking</a><span>Anthropic · Proprietary</span></td><td>1503±4</td><td>38,559</td><td>$5/$25</td><td>1M</td></tr>
<tr><td>5</td><td>311</td><td><span>Meta</span><a href="x" title="muse-spark">muse-spark</a><span>Meta · Proprietary</span></td><td>1489±6 Preliminary</td><td>12,925</td><td>N/A</td><td>N/A</td></tr>
<tr><td>6</td><td>49</td><td><a href="x" title="gemini-3.1-pro-preview">gemini-3.1-pro-preview</a><span>Google · Proprietary</span></td><td>1488±4</td><td>49,587</td><td>$2/$12</td><td>1M</td></tr>
<tr><td>12</td><td>825</td><td><a href="x" title="glm-5.1">glm-5.1</a><span>Z.ai · MIT</span></td><td>1475±6</td><td>15,019</td><td>$1.4/$4.4</td><td>202K</td></tr></tbody></table>"""
    r = {x["rank"]: x for x in fa.parse_html_table(html)}
    assert len(r) == 4
    assert r[1]["model"] == "claude-opus-4-6-thinking" and r[1]["vendor"] == "Anthropic"
    assert r[1]["score"] == 1503 and r[1]["ci"] == 4 and r[1]["votes"] == 38559
    assert r[1]["extra"]["context"] == "1M"
    assert r[5]["score"] == 1489  # Preliminary 后缀不干扰
    assert r[6]["vendor"] == "Google"  # 无前置徽标
    assert r[12]["vendor"] == "Z.ai" and r[12]["license"] == "open"
    print("  ✓ text 榜 HTML（前置徽标 / Preliminary / 多词厂商 Z.ai）")


def test_agent_html():
    html = """<table><thead><tr><th>Rank</th><th>Model</th><th>Net Improvement</th><th>Confirmed Success</th><th>Sessions</th></tr></thead><tbody>
<tr><td>1 14</td><td>GPT 5.5 (High)OpenAI · Proprietary</td><td>10.66%±1.60%</td><td>7.06%±2.70%</td><td>21,520</td></tr>
<tr><td>2 16</td><td><span>Anthropic</span>Claude Opus 4.7 (Thinking)Anthropic · Proprietary</td><td>9.47%±1.50%</td><td>7.95%±2.71%</td><td>21,524</td></tr>
<tr><td>8 710</td><td>GLM 5.1Z.ai · MIT · SiliconFlow</td><td>3.38%±2.00%</td><td>4.63%±3.41%</td><td>16,746</td></tr>
<tr><td>12 1014</td><td>DeepSeek V4 ProDeepSeek · MIT · SiliconFlow</td><td>1.88%±1.79%</td><td>1.57%±3.43%</td><td>16,898</td></tr>
<tr><td>15 1417</td><td>Minimax M2.7MiniMax · Modified MIT · Fireworks</td><td>8.52%±1.76%</td><td>8.00%±3.63%</td><td>16,935</td></tr></tbody></table>"""
    r = {x["rank"]: x for x in fa.parse_html_table(html)}
    assert len(r) == 5
    assert r[1]["model"] == "GPT 5.5 (High)" and r[1]["vendor"] == "OpenAI"
    assert r[1]["score"] is None and r[1]["votes"] == 21520  # 无 Elo，Sessions→votes
    assert r[2]["model"] == "Claude Opus 4.7 (Thinking)"     # 前置徽标已删
    assert r[8]["model"] == "GLM 5.1" and r[8]["license"] == "open"  # 三段许可取 MIT
    assert r[12]["model"] == "DeepSeek V4 Pro"               # 自品牌前缀保留
    assert r[15]["model"] == "Minimax M2.7"                  # 同上
    assert r[1]["extra"]["net improvement"] == "10.66%±1.60%"  # 行为指标入 extra
    print("  ✓ agent 榜 HTML（异构列 / Sessions→votes / 自品牌前缀 / 指标入 extra）")


def test_image_and_multiword_vendor_html():
    html = """<table><thead><tr><th>Rank</th><th>Rank Spread</th><th>Model</th><th>Score</th><th>Votes</th></tr></thead><tbody>
<tr><td>29</td><td>2232</td><td><img alt="Kandinsky" src="x.png"/><a href="x" title="kandinsky-5.0-t2v-pro">kandinsky-5.0-t2v-pro</a><span>Kandinsky · MIT</span></td><td>1176±21</td><td>2,020</td></tr>
<tr><td>24</td><td>1930</td><td><a href="x" title="ray-3">ray-3</a><span>Luma AI · Proprietary</span></td><td>1207±22</td><td>1,121</td></tr></tbody></table>"""
    r = {x["rank"]: x for x in fa.parse_html_table(html)}
    assert r[29]["model"] == "kandinsky-5.0-t2v-pro" and r[29]["vendor"] == "Kandinsky"
    assert r[24]["vendor"] == "Luma AI"  # 多词厂商不被截成 'AI'
    print("  ✓ HTML 图片徽标 + 多词厂商（Luma AI）")


def test_real_markdown_fixture():
    """对真实抓取的 text-to-video 榜（14 行样本）跑 markdown 兜底解析。"""
    md = (Path(__file__).resolve().parent / "sample_text_to_video.md").read_text(encoding="utf-8")
    rows = fa.parse_markdown_table(md)
    r = {x["rank"]: x for x in rows}
    assert len(rows) == 14, f"应 14 行，实际 {len(rows)}"
    assert r[1]["model"] == "dreamina-seedance-2.0-720p" and r[1]["vendor"] == "Bytedance"
    assert r[9]["votes"] == 121372                       # 无逗号大数
    assert r[24]["vendor"] == "Luma AI"                  # 多词厂商
    assert r[29]["model"] == "kandinsky-5.0-t2v-pro"     # 图片徽标行未丢
    assert r[30]["license"] == "open"                    # tencent-hunyuan-community
    assert r[33]["vendor"] == "lightricks"               # 小写厂商
    assert r[34]["license"] == "open"                    # Apache 2.0
    assert r[40]["vendor"] == "Genmo AI"                 # 多词厂商
    print("  ✓ 真实 text-to-video markdown（14/14，全部真实坑覆盖）")


def test_score_formats():
    """分数列的两种置信区间写法：text/视频用 '±'，code 榜用 '+上界/-下界'。"""
    assert fa._parse_score_cell("1503±4") == (1503, 4)              # 对称 ±
    assert fa._parse_score_cell("1567+9/-9") == (1567, 9)            # code 榜对称
    assert fa._parse_score_cell("1537+13/-13  Preliminary") == (1537, 13)  # 带后缀
    assert fa._parse_score_cell("1552+16/-12") == (1552, 16)         # 非对称取较大界
    assert fa._parse_score_cell("1500") == (1500, None)              # 无 CI
    # code 榜 HTML 端到端
    html = """<table><thead><tr><th>Rank</th><th>Rank Spread</th><th>Model</th><th>Score</th><th>Votes</th></tr></thead><tbody>
<tr><td>1</td><td>14</td><td><span>Anthropic</span><a href="x" title="claude-opus-4-7-thinking">claude-opus-4-7-thinking</a><span>Anthropic · Proprietary</span></td><td>1567+9/-9</td><td>6,234</td></tr>
<tr><td>8</td><td>312</td><td><a href="x" title="glm-5.1">glm-5.1</a><span>Z.ai · MIT</span></td><td>1532+11/-11</td><td>3,608</td></tr></tbody></table>"""
    r = {x["rank"]: x for x in fa.parse_html_table(html)}
    assert r[1]["score"] == 1567 and r[1]["ci"] == 9
    assert r[8]["score"] == 1532 and r[8]["ci"] == 11
    print("  ✓ 分数/CI 双格式（± 与 +a/-b，code 榜）")


def main():
    print("运行离线解析回归测试…")
    for fn in (test_text_html, test_agent_html,
               test_image_and_multiword_vendor_html, test_score_formats,
               test_real_markdown_fixture):
        fn()
    print("\n🎉 全部测试通过")


if __name__ == "__main__":
    main()
