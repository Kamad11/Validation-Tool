import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.server import InvoiceService  # noqa: E402


def metric_summary(inv: Dict[str, Any]) -> Dict[str, Any]:
    mpans = inv.get("mpans") or {}
    energy_rows = sum(len((m.get("energy_rates") or [])) for m in mpans.values())
    standing_rows = sum(1 for m in mpans.values() if m.get("standing_charge"))
    nonzero_energy_cost_rows = 0
    total_energy_cost = 0.0
    total_usage_kwh = 0.0

    for m in mpans.values():
        for r in (m.get("energy_rates") or []):
            units = r.get("units_kwh")
            cost = r.get("cost")
            if isinstance(units, (int, float)):
                total_usage_kwh += float(units)
            if isinstance(cost, (int, float)):
                total_energy_cost += float(cost)
                if cost > 0:
                    nonzero_energy_cost_rows += 1

    return {
        "invoice_number": inv.get("invoice_number"),
        "source": inv.get("extraction_source"),
        "di_available": bool(inv.get("document_intelligence_available")),
        "di_error": inv.get("document_intelligence_error"),
        "mpan_count": len(mpans),
        "energy_rows": energy_rows,
        "standing_rows": standing_rows,
        "nonzero_energy_cost_rows": nonzero_energy_cost_rows,
        "total_usage_kwh": round(total_usage_kwh, 4),
        "total_energy_cost": round(total_energy_cost, 4),
        "period_start": inv.get("invoice_period_start"),
        "period_end": inv.get("invoice_period_end"),
    }


def run_di(pdf: Path) -> Dict[str, Any]:
    inv = InvoiceService.parse_pdf(pdf)
    return metric_summary(inv)


def run_pypdf_only(pdf: Path) -> Dict[str, Any]:
    original = InvoiceService._analyze_with_document_intelligence
    InvoiceService._analyze_with_document_intelligence = staticmethod(lambda _p: None)
    try:
        inv = InvoiceService.parse_pdf(pdf)
    finally:
        InvoiceService._analyze_with_document_intelligence = original
    return metric_summary(inv)


def choose_verdict(di: Dict[str, Any], py: Dict[str, Any]) -> str:
    di_key = (di["energy_rows"], di["standing_rows"], di["mpan_count"], di["nonzero_energy_cost_rows"])
    py_key = (py["energy_rows"], py["standing_rows"], py["mpan_count"], py["nonzero_energy_cost_rows"])
    if di_key > py_key:
        return "DI better"
    if py_key > di_key:
        return "pypdf better"
    return "tie"


def default_pdfs() -> List[Path]:
    bills = ROOT / "Bills"
    if not bills.exists():
        return []
    return sorted([p for p in bills.glob("*.pdf")])[:5]


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare Document Intelligence vs pypdf extraction quality.")
    parser.add_argument("--pdf", action="append", default=[], help="Specific PDF path(s) to test. Can be passed multiple times.")
    parser.add_argument("--limit", type=int, default=5, help="When no --pdf is provided, compare up to this many files from Bills/.")
    parser.add_argument("--json-out", default="", help="Optional path to write full comparison JSON.")
    args = parser.parse_args()

    if args.pdf:
        pdfs = [Path(p).expanduser().resolve() for p in args.pdf]
    else:
        pdfs = default_pdfs()[: max(1, args.limit)]

    missing = [p for p in pdfs if not p.exists()]
    if missing:
        raise SystemExit(f"Missing PDF(s): {', '.join(str(p) for p in missing)}")
    if not pdfs:
        raise SystemExit("No PDFs found to compare.")

    rows: List[Dict[str, Any]] = []

    print("DI vs pypdf extraction comparison")
    print("=" * 120)
    for pdf in pdfs:
        di = run_di(pdf)
        py = run_pypdf_only(pdf)
        verdict = choose_verdict(di, py)
        row = {
            "pdf": str(pdf),
            "di": di,
            "pypdf": py,
            "verdict": verdict,
        }
        rows.append(row)

        print(f"PDF: {pdf.name}")
        print(
            "  DI    -> "
            f"source={di['source']} di_available={di['di_available']} mpans={di['mpan_count']} "
            f"energy_rows={di['energy_rows']} standing_rows={di['standing_rows']} "
            f"nonzero_energy_rows={di['nonzero_energy_cost_rows']} "
            f"usage_kwh={di['total_usage_kwh']} energy_cost={di['total_energy_cost']} "
            f"period={di['period_start']} -> {di['period_end']}"
        )
        if di.get("di_error"):
            print(f"           di_error={di['di_error']}")
        print(
            "  PYPDF -> "
            f"source={py['source']} mpans={py['mpan_count']} "
            f"energy_rows={py['energy_rows']} standing_rows={py['standing_rows']} "
            f"nonzero_energy_rows={py['nonzero_energy_cost_rows']} "
            f"usage_kwh={py['total_usage_kwh']} energy_cost={py['total_energy_cost']} "
            f"period={py['period_start']} -> {py['period_end']}"
        )
        print(f"  Verdict: {verdict}")
        print("-" * 120)

    if args.json_out:
        out = Path(args.json_out).expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(rows, indent=2), encoding="utf-8")
        print(f"JSON report written: {out}")


if __name__ == "__main__":
    main()

