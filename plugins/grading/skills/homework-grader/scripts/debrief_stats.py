#!/usr/bin/env python3
"""Per-rubric-item statistics for the post-assignment debrief.

    debrief_stats.py <assignment_dir> [--scores FILE] [--json]

Answers the question a debrief turns on: when students lost points on an item,
was the item hard, or was the instruction unclear?

The discriminator is whether the strong students missed it too. On a genuinely
difficult item the strong half does markedly better -- that is what "difficult"
means, an item that separates who understood the material. On a badly worded
one they lose about as much as everyone else, because reading the prompt the
other way had nothing to do with understanding.

Both measures leave the item out of its own comparison. The halves are ranked
by the rest of the student's score, never the total: ranking by the total lets
an item place students in its own halves, since anyone who lost five points
here drops to the bottom half *because* they lost them, and the gap that
manufactures says nothing about ability. The correlation is likewise between
the item and the rest of the score.

    high loss + low correlation + small gap   -> suspect the wording
    high loss + high correlation + big gap    -> genuinely hard, or under-taught
    the two disagreeing                        -> mixed; go read the feedback
    low loss                                   -> the item worked

None of this is a verdict. It says where to look; reading the feedback
documents is what tells you what actually went wrong.

The signal is withheld rather than printed when it cannot mean anything: under
about eight students, or when leaving one item out leaves fewer than three to
compare against. A correlation computed against one or two other scores is not
an ability estimate, and printing it would dress noise up as evidence.
"""

import argparse
import json
import sys
from pathlib import Path

MIN_FOR_SIGNAL = 8          # below this, the top/bottom split is noise
MIN_OTHER_ITEMS = 3         # "the rest of the score" needs enough items to mean anything
WEAK_CORRELATION = 0.3      # at or under this, the item barely tracks the rest
NOTABLE_LOSS = 0.15         # an item losing under 15% of its points worked


def mean(xs):
    return sum(xs) / len(xs) if xs else 0.0


def median(xs):
    if not xs:
        return 0.0
    s = sorted(xs)
    mid = len(s) // 2
    return s[mid] if len(s) % 2 else (s[mid - 1] + s[mid]) / 2


def pearson(xs, ys):
    """Correlation, or None when either series has no spread to correlate."""
    n = len(xs)
    if n < 3:
        return None
    mx, my = mean(xs), mean(ys)
    dx = [x - mx for x in xs]
    dy = [y - my for y in ys]
    num = sum(a * b for a, b in zip(dx, dy))
    den = (sum(a * a for a in dx) * sum(b * b for b in dy)) ** 0.5
    return None if den == 0 else num / den


def collect(scores_path):
    """Per-item and whole-class numbers from a recalculated scores.xlsx."""
    from sync_scores import load_scores, num, norm, NON_STUDENT

    wb, ws, layout = load_scores(scores_path, data_only=True)
    header_row, netid_col = layout["header_row"], layout["netid_col"]
    rubric_cols = layout["rubric_cols"]
    if not rubric_cols:
        raise ValueError("no rubric columns found between Net ID and Total -- "
                         f"is {scores_path.name} the score summary?")

    # The max-points row sits just under the header and is labelled in the
    # identity column.
    maxima = {}
    for r in range(header_row + 1, min(header_row + 4, ws.max_row + 1)):
        label = ws.cell(r, netid_col).value
        if isinstance(label, str) and norm(label).startswith("maxpoint"):
            for col, name in rubric_cols:
                maxima[col] = num(ws.cell(r, col).value)
            break
    if not maxima:
        raise ValueError("no 'Max Points' row found under the header")

    students = []
    for key, row in layout["rows"].items():
        if NON_STUDENT.match(key):
            continue
        earned = {col: num(ws.cell(row, col).value) for col, _ in rubric_cols}
        if any(v is None for v in earned.values()):
            continue        # not scored, or the workbook was never recalculated
        students.append({"key": key, "earned": earned,
                         "total": sum(earned.values())})
    if not students:
        raise ValueError(f"no scored student rows in {scores_path.name} -- "
                         "recalculate it first (recalc.py)")

    total_max = sum(v for v in maxima.values() if v)
    items = []
    for col, name in rubric_cols:
        mx = maxima.get(col) or 0
        got = [s["earned"][col] for s in students]
        rest = [s["total"] - s["earned"][col] for s in students]
        lost_frac = 1 - (mean(got) / mx) if mx else 0.0

        # Rank by everything EXCEPT this item. Ranking by the total would let
        # the item place students in its own halves: anyone who lost five
        # points here drops to the bottom half because they lost them, and the
        # gap that produces says nothing about ability.
        by_rest = sorted(students, key=lambda s: s["total"] - s["earned"][col],
                         reverse=True)
        half = len(by_rest) // 2
        top = mean([s["earned"][col] for s in by_rest[:half]]) / mx if mx and half else None
        bot = mean([s["earned"][col] for s in by_rest[-half:]]) / mx if mx and half else None

        # With few rubric items the leftover is one or two scores, which is not
        # an ability estimate. Correlating against it produces a number that
        # looks like evidence and is not.
        usable = (len(students) >= MIN_FOR_SIGNAL
                  and len(rubric_cols) - 1 >= MIN_OTHER_ITEMS)
        r = pearson(got, rest) if usable else None
        items.append({
            "column": name, "max": mx,
            "mean": mean(got), "median": median(got),
            "lost_fraction": lost_frac,
            "full_credit": sum(1 for g in got if mx and g >= mx),
            "any_deduction": sum(1 for g in got if mx and g < mx),
            "top_half_fraction": top, "bottom_half_fraction": bot,
            "correlation": r,
        })

    totals = [s["total"] for s in students]
    return {
        "n": len(students),
        "total_max": total_max,
        "mean": mean(totals), "median": median(totals),
        "min": min(totals), "max": max(totals),
        "items": items,
    }


def read(item, n, n_items):
    """What this item's numbers point at, in one phrase."""
    if item["lost_fraction"] < NOTABLE_LOSS:
        return "worked"
    if item["correlation"] is None:
        why = ("too few students" if n < MIN_FOR_SIGNAL
               else "too few other items to compare against")
        return f"lost points -- {why} to say why"

    # Two views of the same question, and they can disagree. The halves are a
    # blunt split; the correlation uses every student. Only call it when both
    # point the same way, and say so when they do not.
    gap = None
    if item["top_half_fraction"] is not None and item["bottom_half_fraction"] is not None:
        gap = item["top_half_fraction"] - item["bottom_half_fraction"]
    weak_r = item["correlation"] <= WEAK_CORRELATION
    weak_gap = gap is not None and gap <= 0.10

    if weak_r and weak_gap:
        return "SUSPECT THE WORDING -- strong students missed it too"
    if not weak_r and not weak_gap:
        return "genuinely hard, or under-taught"
    return "mixed signal -- read the feedback before deciding"


def report(stats):
    n = stats["n"]
    out = [
        f"\n{n} student(s), out of {stats['total_max']:g} points",
        f"  mean {stats['mean']:.1f}   median {stats['median']:.1f}   "
        f"range {stats['min']:g}-{stats['max']:g}",
        "",
        f"  {'Rubric item':<26} {'max':>4} {'mean':>6} {'lost':>6} "
        f"{'top½':>6} {'bot½':>6} {'r':>6}   reading",
        f"  {'-' * 26} {'-' * 4} {'-' * 6} {'-' * 6} {'-' * 6} {'-' * 6} {'-' * 6}   {'-' * 7}",
    ]
    for it in stats["items"]:
        pct = lambda v: f"{v * 100:5.0f}%" if v is not None else "     -"
        r = f"{it['correlation']:6.2f}" if it["correlation"] is not None else "     -"
        out.append(
            f"  {it['column'][:26]:<26} {it['max'] or 0:>4g} {it['mean']:>6.1f} "
            f"{pct(it['lost_fraction'])} {pct(it['top_half_fraction'])} "
            f"{pct(it['bottom_half_fraction'])} {r}   "
            f"{read(it, n, len(stats['items']))}")

    out += [
        "",
        "  lost = share of the item's points the class did not earn",
        "  top½ / bot½ = share of the item earned by the better and weaker halves",
        "  r = how well the item tracks the rest of the student's score",
        "",
        "  An item losing real points that the top half missed about as often as the",
        "  bottom half is not measuring understanding. Read what students actually did",
        "  before concluding anything -- these numbers say where to look, not why.",
        "",
    ]
    if n < MIN_FOR_SIGNAL:
        out.insert(2, f"  Fewer than {MIN_FOR_SIGNAL} students: the split is withheld as noise.")
    elif len(stats["items"]) - 1 < MIN_OTHER_ITEMS:
        out.insert(2, f"  Only {len(stats['items'])} rubric items: leaving one out leaves too "
                      "little to compare against, so the split is withheld.")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("assignment_dir", type=Path)
    ap.add_argument("--scores", type=Path, help="scores.xlsx (default: <dir>/scores.xlsx)")
    ap.add_argument("--json", action="store_true", help="emit the numbers as JSON")
    args = ap.parse_args()

    root = args.assignment_dir.expanduser().resolve()
    scores = (args.scores or root / "scores.xlsx").expanduser().resolve()
    if not scores.is_file():
        sys.exit(f"error: no scores workbook at {scores}")

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    try:
        stats = collect(scores)
    except ImportError as exc:
        sys.exit(f"error: {exc}")
    except Exception as exc:
        sys.exit(f"error: {exc}")

    print(json.dumps(stats, indent=2) if args.json else report(stats))
    return 0


if __name__ == "__main__":
    sys.exit(main())
