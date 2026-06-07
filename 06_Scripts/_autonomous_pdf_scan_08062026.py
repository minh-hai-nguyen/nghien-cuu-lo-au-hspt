"""Autonomous PDF integrity scan against canonical_index.json (08/06/2026).

Checks:
1. Magic byte (%PDF)
2. Size > 20KB
3. First-page extraction works
4. Random 10 papers: Funding/Acknowledgment presence
5. Stub flag (< 50KB)
"""
from __future__ import annotations

import json
import os
import random
import re
import sys
from pathlib import Path

ROOT = Path(r"c:/Users/OS/OneDrive/read_books/Lo-au")
INDEX = ROOT / "02_Papers-goc" / "canonical_index.json"
REPORT = ROOT / "06_Scripts" / "AUTONOMOUS_PDF_SCAN_08062026.md"

try:
    import fitz  # pymupdf
except Exception as e:
    print("fitz import fail:", e)
    sys.exit(1)


def main() -> int:
    data = json.loads(INDEX.read_text(encoding="utf-8"))
    rows = list(data.values())
    total = len(rows)

    pass_magic = 0
    fail_magic: list[tuple[str, str]] = []
    pass_size = 0
    fail_size: list[tuple[str, int]] = []  # < 20KB
    stubs: list[tuple[str, int]] = []      # < 50KB
    pass_open = 0
    fail_open: list[tuple[str, str]] = []
    missing: list[str] = []

    # Funding/Ack pattern
    fund_re = re.compile(
        r"(funding|acknowledg(e?)ments?|grant\s+(no\.|number)|supported\s+by)",
        re.IGNORECASE,
    )

    # Random sample 10 from those that successfully open
    random.seed(20260608)
    sample_target = 10
    funding_results: list[tuple[str, bool, str]] = []

    # Build list of paths for sampling
    candidate_for_sample: list[tuple[str, Path]] = []

    for r in rows:
        qid = r.get("id", "?")
        rel = r.get("pdf_path") or ""
        if not rel:
            missing.append(qid + " :: no pdf_path")
            continue
        p = (ROOT / rel).resolve()
        if not p.exists():
            missing.append(f"{qid} :: MISSING {rel}")
            continue

        # 1. Magic byte
        try:
            with p.open("rb") as fh:
                head = fh.read(5)
            if head.startswith(b"%PDF"):
                pass_magic += 1
            else:
                fail_magic.append((qid, head.hex()))
        except Exception as e:
            fail_magic.append((qid, f"read-err:{e}"))

        # 2. Size
        sz = p.stat().st_size
        if sz > 20 * 1024:
            pass_size += 1
        else:
            fail_size.append((qid, sz))
        if sz < 50 * 1024:
            stubs.append((qid, sz))

        # 3. Open + first page
        try:
            with fitz.open(p) as doc:
                if doc.page_count < 1:
                    raise RuntimeError("zero pages")
                _ = doc.load_page(0).get_text("text")
            pass_open += 1
            candidate_for_sample.append((qid, p))
        except Exception as e:
            fail_open.append((qid, str(e)[:120]))

    # 4. Funding sample
    if candidate_for_sample:
        sample_n = min(sample_target, len(candidate_for_sample))
        sample = random.sample(candidate_for_sample, sample_n)
        for qid, p in sample:
            try:
                with fitz.open(p) as doc:
                    # Last 2 pages most likely to have funding info, plus first page acks
                    pages_to_scan: list[int] = []
                    n = doc.page_count
                    pages_to_scan.append(0)
                    if n >= 2:
                        pages_to_scan.append(n - 1)
                    if n >= 3:
                        pages_to_scan.append(n - 2)
                    text = ""
                    for pi in sorted(set(pages_to_scan)):
                        text += "\n" + doc.load_page(pi).get_text("text")
                m = fund_re.search(text)
                if m:
                    funding_results.append((qid, True, m.group(0)))
                else:
                    funding_results.append((qid, False, ""))
            except Exception as e:
                funding_results.append((qid, False, f"err:{e}"))

    pct_magic = pass_magic / total * 100 if total else 0
    pct_size = pass_size / total * 100 if total else 0
    pct_open = pass_open / total * 100 if total else 0

    lines: list[str] = []
    lines.append("# AUTONOMOUS PDF INTEGRITY SCAN — 08/06/2026")
    lines.append("")
    lines.append(f"- Index file: `{INDEX.as_posix()}`")
    lines.append(f"- Total entries in canonical_index.json: **{total}**")
    lines.append(f"- Missing/no-path on disk: **{len(missing)}**")
    lines.append("")
    lines.append("## Summary")
    lines.append("")
    lines.append("| Check | PASS | Total | % |")
    lines.append("|---|---:|---:|---:|")
    lines.append(f"| 1. Magic byte `%PDF` | {pass_magic} | {total} | {pct_magic:.1f}% |")
    lines.append(f"| 2. Size > 20 KB | {pass_size} | {total} | {pct_size:.1f}% |")
    lines.append(f"| 3. Open + page-1 extract | {pass_open} | {total} | {pct_open:.1f}% |")
    lines.append("")

    lines.append("## Missing files")
    if not missing:
        lines.append("- (none)")
    else:
        for m in missing:
            lines.append(f"- {m}")
    lines.append("")

    lines.append("## Magic-byte failures")
    if not fail_magic:
        lines.append("- (none)")
    else:
        for qid, h in fail_magic:
            lines.append(f"- {qid} :: head={h}")
    lines.append("")

    lines.append("## Size failures (< 20 KB)")
    if not fail_size:
        lines.append("- (none)")
    else:
        for qid, s in fail_size:
            lines.append(f"- {qid} :: {s} bytes")
    lines.append("")

    lines.append("## Stub suspects (< 50 KB)")
    if not stubs:
        lines.append("- (none)")
    else:
        for qid, s in sorted(stubs, key=lambda x: x[1]):
            lines.append(f"- {qid} :: {s} bytes")
    lines.append("")

    lines.append("## Open / first-page extract failures")
    if not fail_open:
        lines.append("- (none)")
    else:
        for qid, e in fail_open:
            lines.append(f"- {qid} :: {e}")
    lines.append("")

    lines.append("## Funding / Acknowledgment sample (random 10)")
    if not funding_results:
        lines.append("- (no candidates)")
    else:
        for qid, ok, hit in funding_results:
            mark = "FOUND" if ok else "NOT FOUND"
            extra = f" — `{hit}`" if hit else ""
            lines.append(f"- {qid} :: {mark}{extra}")
    lines.append("")

    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text("\n".join(lines), encoding="utf-8")

    # Machine-readable summary on stdout
    print(json.dumps({
        "total": total,
        "missing": len(missing),
        "pass_magic": pass_magic,
        "pass_size": pass_size,
        "pass_open": pass_open,
        "pct_magic": round(pct_magic, 2),
        "pct_size": round(pct_size, 2),
        "pct_open": round(pct_open, 2),
        "fail_magic": fail_magic,
        "fail_size": fail_size,
        "stubs_count": len(stubs),
        "stubs_smallest_5": sorted(stubs, key=lambda x: x[1])[:5],
        "fail_open": fail_open,
        "funding_found": sum(1 for _, ok, _ in funding_results if ok),
        "funding_total": len(funding_results),
        "report": REPORT.as_posix(),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
