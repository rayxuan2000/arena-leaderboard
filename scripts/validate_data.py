#!/usr/bin/env python3
"""
validate_data.py — 抓取后、构建前的数据健全性校验（守门员）
=================================================================
fetch_arena.py 已保证：任一榜单解析到 0 行就 sys.exit(1)，且 0 行不覆盖既有好数据。
本脚本补上更隐蔽的失败：**榜单解析出 n>0 行、但相比上一版已提交数据断崖式暴跌**
（例如官网改版后只勉强解析出几行垃圾），这种 fetch_arena 不会报错，却会污染线上榜单。

校验规则（任一硬失败 → 退出码 1 → 工作流中止，不提交、不部署，保住上一版）：
  1. 每个已知榜单的 data/<slug>.json 必须存在、可解析、models 非空；
  2. 核心榜单（text/code/vision）模型数 ≥ CORE_MIN；
  3. 全部榜单模型总数 ≥ TOTAL_MIN；
  4. 与上一版已提交数据（git show HEAD:data/<slug>.json）相比，
     原本 ≥ BASE_MIN 的榜单若跌到不足 DROP_RATIO，判定为断崖式损坏。
对比基线取不到（首次运行 / 无 git）时，跳过规则 4，仅做绝对值检查（规则 1-3）。

用法：
  python scripts/validate_data.py --data-dir data
  python scripts/validate_data.py --data-dir data --slugs text code   # 只校验部分
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

KNOWN_SLUGS = [
    "text", "code", "vision", "document", "search", "agent",
    "text-to-image", "image-edit",
    "text-to-video", "image-to-video", "video-edit",
]
CORE_SLUGS = ("text", "code", "vision")

CORE_MIN = 10      # 核心榜单的绝对下限
TOTAL_MIN = 50     # 所有榜单模型总数下限
BASE_MIN = 20      # 仅当上一版数量 ≥ 此值才做暴跌比较（小榜单波动不参与）
DROP_RATIO = 0.4   # 新数量 < 上一版 * 0.4（即丢失 >60%）判为断崖式损坏


def _count(path: Path) -> int:
    try:
        d = json.loads(path.read_text(encoding="utf-8"))
        return len(d.get("models", []))
    except (ValueError, OSError):
        return -1


def _prev_count(slug: str, data_dir: Path) -> int:
    """读取上一版已提交的该榜单模型数；取不到返回 None。"""
    rel = f"{data_dir.name}/{slug}.json"
    try:
        out = subprocess.run(
            ["git", "show", f"HEAD:{rel}"],
            capture_output=True, text=True, timeout=20,
            cwd=data_dir.parent if data_dir.parent != Path("") else None,
        )
        if out.returncode != 0 or not out.stdout.strip():
            return None
        return len(json.loads(out.stdout).get("models", []))
    except (ValueError, OSError, subprocess.SubprocessError):
        return None


def main() -> int:
    ap = argparse.ArgumentParser(description="抓取数据健全性校验")
    ap.add_argument("--data-dir", default="data")
    ap.add_argument("--slugs", nargs="*", help="仅校验指定榜单（默认全部已知榜单）")
    args = ap.parse_args()

    data_dir = Path(args.data_dir)
    slugs = args.slugs or KNOWN_SLUGS

    errors, warnings = [], []
    total = 0
    print(f"{'榜单':<16}{'本次':>8}{'上一版':>10}  状态")
    print("-" * 48)

    for slug in slugs:
        fp = data_dir / f"{slug}.json"
        n = _count(fp) if fp.exists() else -1
        prev = _prev_count(slug, data_dir)
        prev_disp = "—" if prev is None else str(prev)

        status = "ok"
        if n < 0:
            status = "缺失/损坏"
            errors.append(f"{slug}: 文件缺失或无法解析")
        elif n == 0:
            status = "空"
            errors.append(f"{slug}: 0 个模型")
        else:
            total += n
            if slug in CORE_SLUGS and n < CORE_MIN:
                status = f"过少(<{CORE_MIN})"
                errors.append(f"{slug}（核心榜）仅 {n} 个模型，低于下限 {CORE_MIN}")
            elif prev is not None and prev >= BASE_MIN and n < prev * DROP_RATIO:
                status = "断崖暴跌"
                errors.append(
                    f"{slug}: {n} 个模型，相比上一版 {prev} 暴跌超 "
                    f"{int((1 - DROP_RATIO) * 100)}%（疑似官网改版/解析损坏）")
            elif prev is not None and prev >= BASE_MIN and n < prev * 0.75:
                status = "明显减少"
                warnings.append(f"{slug}: {n} 个模型（上一版 {prev}），减少较多，请留意")

        print(f"{slug:<16}{(n if n>=0 else 'N/A'):>8}{prev_disp:>10}  {status}")

    print("-" * 48)
    print(f"模型总数：{total}")

    if total and total < TOTAL_MIN:
        errors.append(f"全部榜单模型总数仅 {total}，低于下限 {TOTAL_MIN}")

    for w in warnings:
        print(f"  ⚠️ {w}")
    if errors:
        print(f"\n❌ 校验未通过（{len(errors)} 项），中止构建/部署以保住上一版数据：")
        for e in errors:
            print(f"   · {e}")
        return 1

    print("\n✅ 数据校验通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
