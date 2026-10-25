#!/usr/bin/env python3
"""HMRC travel expense claim pipeline (standard library, plus openpyxl for the workbook).

Everything runs locally. Typical order for a new person:
    prepare   rename pdf2md output to YYYY-MM_<employer>.md, report duplicates and gaps
    extract   payslips.csv and payslip_line_codes.md
    init      draft config.yaml from TIS text and homes.csv (fill the TODOs)
    check     validate the config against the payslips before calculating
    attendance, calc, evidence   the claim itself
    all       run the whole chain (see run_all.sh)
"""
import argparse, csv, collections, glob, math, os, re
from datetime import date, datetime, timedelta

NUM = re.compile(r"^-?\d[\d,]*\.?\d*R?$")


def num(s):
    if s is None:
        return None
    neg = s.endswith("R")  # ESR marks refunds with a trailing R
    s = s.replace(",", "").rstrip("R")
    try:
        return -float(s) if neg else float(s)
    except ValueError:
        return None


def is_num(tok):
    return bool(NUM.match(tok))


def is_blank(tok):
    return tok.replace(" ", "") == "''"


def parse_table(lines):
    """Return (pay_lines, ded_lines); each is (desc, wkd, paid, rate, amount)."""
    start = next(i for i, l in enumerate(lines) if "BALANCE C/F" in l)
    end = next(i for i, l in enumerate(lines) if "Year To Date Balances" in l)
    pay, ded = [], []
    for l in lines[start + 1:end]:
        if not l.strip():
            continue
        indent = len(l) - len(l.lstrip())
        toks = re.split(r"\s{2,}", l.strip())
        i = 0
        if indent < 40:
            desc = toks[0]
            i = 1
            nums = []
            while i < len(toks) and (is_num(toks[i]) or is_blank(toks[i])):
                nums.append(toks[i])
                i += 1
            amt = num(nums[-1]) if nums else None
            vals = [None if is_blank(n) else num(n) for n in nums[:-1]]
            vals += [None] * (3 - len(vals))
            pay.append((desc, vals[0], vals[1], vals[2], amt))
        if i < len(toks):
            ddesc = toks[i]
            damt = num(toks[i + 1]) if i + 1 < len(toks) else None
            ded.append((ddesc, damt))
    return pay, ded


def find(pat, text, cast=num, flags=0):
    m = re.search(pat, text, flags)
    return cast(m.group(1)) if m else None


def tax_year(d):
    y = d.year if (d.month, d.day) >= (4, 6) else d.year - 1
    return f"{y}/{str(y + 1)[2:]}"


# Classification of line names. First match wins. Order matters.
PAY_RULES = [
    ("mileage_or_excess_travel", r"mileage|excess travel|\bmiles\b|exc.*trav"),
    ("tunnel_or_toll", r"tunnel|toll|mersey"),
    ("misc_travel", r"misc travel|travel"),
    ("course_expense", r"course|study|exam|subs"),
    ("basic", r"^basic pay$"),
    ("basic_arrears", r"basic pay arrs"),
    ("additional_hours", r"addn ros|additional"),
    ("enhanced_night", r"night"),
    ("weekend", r"weekend|w/?e\b"),
    ("on_call", r"on.?call|standby|availability"),
    ("locum_bank", r"locum|patchwork|bank"),
    ("relocation", r"relocation"),
    ("other_expense", r"expense|reimb|allowance"),
    ("salary_sacrifice_electronics", r"home elec"),
    ("bank_dit_collaborative", r"\bdit\b"),
]
DED_RULES = [
    ("tax", r"^paye"),
    ("ni", r"^ni\b"),
    ("pension", r"pension|nest|nhs pen|pen arr|pension arrs"),
    ("parking", r"car p|park"),
    ("salary_sacrifice", r"sacrifice|\bss\b|lease|cycle|bike|car scheme|\bevs\b"),
    ("student_loan", r"student|\bsl\b"),
    ("union_or_subs", r"bma|union|subscription|gmc|defence|mps|mdu"),
    ("recovery_of_advance", r"recovery|advance|refund|overpay|recover"),
]


def classify(desc, rules):
    d = desc.lower()
    for cls, pat in rules:
        if re.search(pat, d):
            return cls
    return "UNCLASSIFIED"


def parse_payslip(path):
    text = open(path).read()
    lines = text.splitlines()
    i = next(k for k, l in enumerate(lines) if "ASSIGNMENT NUMBER" in l)
    row = re.split(r"\s{3,}", lines[i + 2].strip())
    j = next(k for k, l in enumerate(lines) if "DEPARTMENT" in l and "JOB TITLE" in l)
    drow = re.split(r"\s{3,}", lines[j + 2].strip())
    k = next(k for k, l in enumerate(lines) if "TAX OFFICE NAME" in l)
    trow = re.split(r"\s{3,}", next(l for l in lines[k + 1:k + 4] if "/" in l or "UPON" in l).strip())
    pay_date = datetime.strptime(
        re.search(r"PAY DATE\s+(\d+ \w+ \d{4})", text).group(1).title(), "%d %b %Y")
    pay, ded = parse_table(lines)

    r = collections.OrderedDict()
    r["file"] = os.path.basename(path)
    r["pay_date"] = pay_date.date().isoformat()
    r["tax_period"] = find(r"TAX PERIOD\s+(\d+)", text, int)
    r["tax_year"] = tax_year(pay_date)
    r["assignment"] = row[0]
    r["location"] = row[-1] if len(row) > 2 else ""
    r["department"] = drow[0] if drow else ""
    r["job_title"] = drow[1] if len(drow) > 1 else ""
    r["payscale"] = drow[-1] if len(drow) > 2 else ""
    tblock = " ".join(lines[k + 1:k + 4])
    tcode = re.search(r"\b(\d+L|BR|D0|D1|0T|NT|K\d+)\b", tblock)
    tmode = re.search(r"\b(NONCUM|CUMUL|W1/M1)\b", tblock)
    r["tax_code"] = f"{tcode.group(1)} {tmode.group(1) if tmode else ''}".strip() if tcode else ""
    r["tax_office"] = trow[0] if not re.match(r"^\d+L$", trow[0]) else ""
    r["total_payments"] = find(r"TOTAL PAYMENTS\s+(-?[\d.,]+)", text)
    r["non_taxable_pay"] = find(r"NON-TAXABLE PAY\s+(-?[\d.,]+)", text)
    r["taxable_pay_period"] = find(r"PENSIONABLE PAY\s+[\d.,]+\s+TAXABLE PAY\s+(-?[\d.,]+)", text)
    r["total_deductions"] = find(r"TOTAL DEDUCTIONS\s+(-?[\d.,]+)", text)
    r["net_pay"] = find(r"NET PAY\s+(-?[\d.,]+)", text)
    r["gross_ytd"] = find(r"GROSS PAY\s+(-?[\d.,]+)", text)
    r["taxable_ytd"] = find(r"\bTAXABLE PAY\s+(-?[\d.,]+)\s+PENSIONABLE", text)
    r["tax_paid_ytd"] = find(r"TAX PAID\s+(-?[\d.,]+)", text)
    r["ni_ytd"] = find(r"NI CONTS\s+(-?[\d.,]+)", text)

    sums = collections.defaultdict(float)
    hours = collections.defaultdict(float)
    names = collections.defaultdict(list)
    for desc, wkd, paid, rate, amt in pay:
        c = classify(desc, PAY_RULES)
        sums["pay:" + c] += amt or 0
        names["pay:" + c].append(f"{desc}={amt}")
        if c == "basic":
            hours["basic"] += wkd or 0
        elif c == "additional_hours":
            hours["additional"] += wkd or 0
        elif c == "enhanced_night":
            hours["enhanced"] += paid if paid is not None else (wkd or 0)
    dsums = collections.defaultdict(float)
    for desc, amt in ded:
        c = classify(desc, DED_RULES)
        dsums[c] += amt or 0
        names["ded:" + c].append(f"{desc}={amt}")

    r["paye_period"] = round(dsums["tax"], 2)
    r["ni_period"] = round(dsums["ni"], 2)
    r["basic_hours"] = round(hours["basic"], 2)
    r["additional_hours"] = round(hours["additional"], 2)
    r["enhanced_hours"] = round(hours["enhanced"], 2)
    r["weekend_allowance"] = round(sums["pay:weekend"], 2)
    r["on_call"] = round(sums["pay:on_call"], 2)
    r["mileage_or_excess_travel_paid"] = round(sums["pay:mileage_or_excess_travel"], 2)
    r["tunnel_or_toll_reimbursement"] = round(sums["pay:tunnel_or_toll"], 2)
    r["misc_travel_amount"] = round(sums["pay:misc_travel"], 2)
    r["parking_deduction"] = round(dsums["parking"], 2)
    r["pension_deduction"] = round(dsums["pension"], 2)
    r["salary_sacrifice_lines"] = "; ".join(names["ded:salary_sacrifice"])
    r["other_expense_lines"] = "; ".join(
        names["pay:misc_travel"] + names["pay:tunnel_or_toll"]
        + names["pay:course_expense"] + names["pay:other_expense"]
        + names["pay:mileage_or_excess_travel"])
    r["unclassified_lines"] = "; ".join(
        names["pay:UNCLASSIFIED"] + names["ded:UNCLASSIFIED"])

    # Parking tax treatment: taxable pay = total payments - non-taxable pay
    # - pre-tax deductions. Compare the gap with pension and with pension+parking.
    gap = None
    if None not in (r["total_payments"], r["taxable_pay_period"]):
        gap = round(r["total_payments"] - (r["non_taxable_pay"] or 0)
                    - r["taxable_pay_period"], 2)
    r["pretax_gap"] = gap
    pk, pen = r["parking_deduction"], r["pension_deduction"]
    if pk == 0:
        r["parking_tax_treatment"] = ""
    elif gap is None:
        r["parking_tax_treatment"] = "unknown"
    elif abs(gap - pen) <= 0.02:
        r["parking_tax_treatment"] = "post-tax"
    elif abs(gap - pen - pk) <= 0.02:
        r["parking_tax_treatment"] = "pre-tax"
    else:
        r["parking_tax_treatment"] = "unclear"

    # Reconciliation of components to totals
    problems = []
    pay_total = round(sum(a or 0 for *_, a in pay), 2)
    ded_total = round(sum(a or 0 for _, a in ded), 2)
    if r["total_payments"] is not None and abs(pay_total - r["total_payments"]) > 0.02:
        problems.append(f"pay lines {pay_total} != total payments {r['total_payments']}")
    if r["total_deductions"] is not None and abs(ded_total - r["total_deductions"]) > 0.02:
        problems.append(f"deduction lines {ded_total} != total deductions {r['total_deductions']}")
    if None not in (r["total_payments"], r["total_deductions"], r["net_pay"]):
        if abs(r["total_payments"] - r["total_deductions"] - r["net_pay"]) > 0.02:
            problems.append("total payments - total deductions != net pay")
    if gap is not None and r["parking_tax_treatment"] == "unclear":
        problems.append(f"pre-tax gap {gap} matches neither pension {pen} nor pension+parking")
    r["reconciliation"] = "OK" if not problems else " | ".join(problems)
    return r, pay, ded


def write_line_codes(path, seen):
    with open(path, "w") as f:
        f.write("# Payslip line codes\n\nEvery distinct line name found, its classification, "
                "how many payslips it appears on, and total amount.\n\n")
        for side, rules in (("Pay and allowances", "pay"), ("Deductions", "ded")):
            f.write(f"## {side}\n\n| Line name | Class | Payslips | Total amount |\n|---|---|---|---|\n")
            rows = sorted((k for k in seen if k[0] == rules), key=lambda k: (k[2], k[1]))
            for k in rows:
                n, tot = seen[k]
                f.write(f"| {k[1]} | {k[2]} | {n} | {tot:.2f} |\n")
            f.write("\n")
        unc = [k for k in seen if k[2] == "UNCLASSIFIED"]
        f.write("## Unclassified\n\n")
        f.write("\n".join(f"- {k[1]} ({'pay' if k[0] == 'pay' else 'deduction'})" for k in unc)
                if unc else "None.")
        f.write("\n")


def cmd_extract(a):
    rows, seen = [], collections.defaultdict(lambda: [0, 0.0])
    for p in sorted(glob.glob(os.path.join(a.md, "*.md"))):
        r, pay, ded = parse_payslip(p)
        rows.append(r)
        for desc, *_, amt in pay:
            k = ("pay", re.sub(r"\s+", " ", desc), classify(desc, PAY_RULES))
            seen[k][0] += 1
            seen[k][1] += amt or 0
        for desc, amt in ded:
            k = ("ded", re.sub(r"\s+", " ", desc), classify(desc, DED_RULES))
            seen[k][0] += 1
            seen[k][1] += amt or 0
    with open(os.path.join(a.out, "payslips.csv"), "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    write_line_codes(os.path.join(a.out, "payslip_line_codes.md"), seen)
    bad = [r for r in rows if r["reconciliation"] != "OK"]
    print(f"{len(rows)} payslips parsed, {len(bad)} failed reconciliation")


# ---------------------------------------------------------------------------
# Config reader: tiny YAML subset (PyYAML is not assumed to be installed).
# Supports nested mappings, lists of scalars or one-line {k: v} items,
# [a, b] lists, dates (YYYY-MM-DD), numbers, quoted strings, null and bools.
# ---------------------------------------------------------------------------
def _scalar(v):
    v = v.strip()
    if v in ("", "~", "null"):
        return None
    if v in ("true", "false"):
        return v == "true"
    if v[0] in "\"'" and v[-1] == v[0]:
        return v[1:-1]
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return date.fromisoformat(v)
    if v.startswith("["):
        return [_scalar(x) for x in _split(v[1:-1])]
    if v.startswith("{"):
        d = {}
        for part in _split(v[1:-1]):
            k, _, val = part.partition(":")
            d[k.strip().strip("\"'")] = _scalar(val)
        return d
    try:
        return int(v)
    except ValueError:
        try:
            return float(v)
        except ValueError:
            return v


def _split(s):
    out, cur, q, depth = [], "", None, 0
    for ch in s:
        if q:
            cur += ch
            q = None if ch == q else q
        elif ch in "\"'":
            q = ch
            cur += ch
        elif ch in "[{":
            depth += 1
            cur += ch
        elif ch in "]}":
            depth -= 1
            cur += ch
        elif ch == "," and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur)
    return out


def load_config(path):
    lines = [l.rstrip() for l in open(path) if l.strip() and not l.lstrip().startswith("#")]

    def block(i, indent):
        if lines[i].lstrip().startswith("- "):
            res = []
            while i < len(lines) and len(lines[i]) - len(lines[i].lstrip()) == indent \
                    and lines[i].lstrip().startswith("- "):
                res.append(_scalar(lines[i].lstrip()[2:]))
                i += 1
            return res, i
        res = {}
        while i < len(lines) and len(lines[i]) - len(lines[i].lstrip()) == indent:
            k, _, v = lines[i].strip().partition(":")
            i += 1
            if v.strip():
                res[k] = _scalar(v)
            elif i < len(lines) and len(lines[i]) - len(lines[i].lstrip()) > indent:
                res[k], i = block(i, len(lines[i]) - len(lines[i].lstrip()))
            else:
                res[k] = None
        return res, i

    return block(0, 0)[0]


# ---------------------------------------------------------------------------
# Attendance (Phase 5)
# ---------------------------------------------------------------------------
def is_training_slip(r):
    """Payslips for the training post, not bank, locum or collaborative-bank work."""
    t = (r.get("job_title", "") + " " + r.get("department", "")).lower()
    return not any(w in t for w in ("bank", "locum", "patchwork", "as and when", "collaborative"))


def is_high_risk(label):
    """Placements labelled risk: "high..." in the config are dropped in the lower case."""
    return str(label or "").lower().startswith("high")


def tax_year_of(d):
    y = d.year if (d.month, d.day) >= (4, 6) else d.year - 1
    return f"{y}/{str(y + 1)[2:]}"


def home_on(cfg, d):
    for h in cfg["homes"]:
        if h["from"] <= d and (h["to"] is None or d <= h["to"]):
            return h["postcode"]
    return None


def pick_evenly(n_items, k):
    """Indices of k items spread evenly through n_items (deterministic)."""
    if k <= 0:
        return set()
    k = min(k, n_items)
    return {int((i + 1) * n_items / (k + 1)) for i in range(k)}


def attendance_rows(cfg):
    sites = {s["code"]: s for s in cfg["sites"]}
    miles = {(d["home"], d["site"]): d["miles"] for d in cfg["distances"]["return_miles"]}
    bh = set(cfg["attendance"]["bank_holidays"])
    excl = cfg["attendance"]["excluded_periods"]
    end_cap = cfg["claim_window"]["end"]
    al = cfg["attendance"]["annual_leave_days_per_year"]
    sl = cfg["attendance"]["study_leave_days_per_year"]
    rows, plan = [], []
    for p in cfg["placements"]:
        if not p["claim"]:
            plan.append(dict(p=p, working=0, leave=0, claimed=0, cal=0))
            continue
        start, end = p["start"], min(p["end"] or end_cap, end_cap)
        days = []
        d = start
        while d <= end:
            if d.weekday() < 5 and d not in bh and not any(e["from"] <= d <= e["to"] for e in excl):
                days.append(d)
            d += timedelta(days=1)
        cal = (end - start).days + 1
        raw = (al + sl) * cal / 365.25 * p["wte"]
        leave = math.ceil(raw) if cfg["attendance"]["leave_rounding"] == "up" else round(raw)
        skip = pick_evenly(len(days), leave)
        kept = [x for i, x in enumerate(days) if i not in skip]
        plan.append(dict(p=p, working=len(days), leave=len(days) - len(kept), claimed=len(kept), cal=cal))
        for d in kept:
            home = home_on(cfg, d)
            m = miles.get((home, p["site"]))
            rows.append(dict(
                date=d.isoformat(), tax_year=tax_year_of(d), from_postcode=home,
                to_site=f"{sites[p['site']]['name']} ({sites[p['site']]['postcode']})",
                purpose="home to temporary workplace", return_miles=m if m is not None else "",
                method_used=cfg["attendance"]["method_default"], placement_id=p["id"],
                risk=p.get("risk", "")))
    return rows, plan


def crosscheck(cfg, rows, plan, payslips_csv, path):
    """Compare the day estimate with payslip evidence and write a report."""
    slips = [r for r in csv.DictReader(open(payslips_csv)) if is_training_slip(r)]
    toll = {(d["home"], d["site"]): bool(d.get("toll")) for d in cfg["distances"]["return_miles"]}
    out = ["# Attendance cross-check against payslips\n",
           "Method for every placement: `fallback_working_days` (no rota or work schedule supplied). "
           "Estimate = Mon to Fri, minus bank holidays, minus annual and study leave pro-rated, minus excluded periods.\n",
           "| Placement | Site | Claimed days | Payslips in range | Avg basic hrs | Avg basic+addn hrs | Est hrs at 8h/day per month | Night hrs (sum) | Tunnel reimb. (non-taxable) | Taxable travel | Route has toll | Flags |",
           "|---|---|---|---|---|---|---|---|---|---|---|---|"]
    for x in plan:
        p = x["p"]
        if not p["claim"]:
            continue
        end = p["end"] or cfg["claim_window"]["end"]
        ss = [r for r in slips if p["start"].isoformat() <= r["pay_date"] <= end.isoformat()]
        n = len(ss)
        basic = sum(float(r["basic_hours"] or 0) for r in ss) / n if n else 0
        paid = sum(float(r["basic_hours"] or 0) + float(r["additional_hours"] or 0) for r in ss) / n if n else 0
        night = sum(float(r["enhanced_hours"] or 0) for r in ss)
        nt = sum(float(m) for r in ss for m in re.findall(r"Misc Travel NP NT=([\d.]+)", r["other_expense_lines"]))
        tx = sum(float(m) for r in ss for m in re.findall(r"Misc Travel NP=([\d.]+)", r["other_expense_lines"]))
        months = x["cal"] / 30.44
        est_h = x["claimed"] / months * 8 if months else 0
        sample = next((r for r in rows if r["placement_id"] == p["id"]), None)
        tolled = toll.get((sample["from_postcode"], p["site"])) if sample else None
        flags = []
        if n == 0:
            flags.append("no payslip falls inside this placement")
        if n and basic < 160:
            flags.append("basic hours below full time")
        if n and est_h > paid * 1.1:
            flags.append("estimate exceeds paid hours")
        if tolled and n and nt == 0:
            flags.append("tolled route but no tunnel reimbursement in range")
        if tolled is False and nt > 0:
            flags.append("tunnel reimbursement but route has no toll (may relate to an earlier site)")
        if tx > 0:
            flags.append("taxable travel line present")
        out.append(f"| {p['id']} | {p['site']} | {x['claimed']} | {n} | {basic:.1f} | {paid:.1f} | {est_h:.1f} | "
                   f"{night:.1f} | {nt:.2f} | {tx:.2f} | {'yes' if tolled else 'no'} | {'; '.join(flags) or 'none'} |")
    out.append("")
    out.append("Reading the table: paid hours per month (basic plus additional) are well above the estimated hours, "
               "so the working-day estimate is conservative compared with hours paid. Night and weekend pay show the rota "
               "included nights and weekends, which a Mon to Fri estimate does not count.\n")
    # per tax year totals
    by = collections.defaultdict(lambda: [0, 0.0, 0, 0.0])
    for r in rows:
        v = by[r["tax_year"]]
        v[0] += 1
        v[1] += float(r["return_miles"] or 0)
        if is_high_risk(r["risk"]):
            v[2] += 1
            v[3] += float(r["return_miles"] or 0)
    out += ["## Days and miles by tax year (base case)\n",
            "| Tax year | Days | Miles | of which high-risk placements: days | miles |", "|---|---|---|---|---|"]
    for y in sorted(by):
        d, m, rd, rm = by[y]
        out.append(f"| {y} | {d} | {m:.1f} | {rd} | {rm:.1f} |")
    open(path, "w").write("\n".join(out) + "\n")


def cmd_attendance(a):
    cfg = load_config(a.config)
    rows, plan = attendance_rows(cfg)
    missing = [r for r in rows if r["return_miles"] == ""]
    out = os.path.join(a.out, "mileage_log.csv")
    with open(out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    pc = os.path.join(a.out, "payslips.csv")
    if os.path.exists(pc):
        crosscheck(cfg, rows, plan, pc, os.path.join(a.out, "attendance_crosscheck.md"))
    print(f"{len(rows)} attendance rows written; {len(missing)} rows without a distance")
    for x in plan:
        p = x["p"]
        if p["claim"]:
            print(f"  placement {p['id']:>2} {p['site']:<6} {p['start']} to {p['end'] or 'open'}: "
                  f"{x['working']} working days, {x['leave']} leave/study, {x['claimed']} claimed")


# ---------------------------------------------------------------------------
# Claim calculation (Phase 6)
# ---------------------------------------------------------------------------
def income_tax(income, year, cfg):
    t = cfg["tax"]
    pa = t["personal_allowance"]
    if income > t["taper_start"]:
        pa = max(0.0, pa - (income - t["taper_start"]) / 2)
    taxable = max(0.0, income - pa)
    r = t["rates"]
    basic = min(taxable, t["basic_band"])
    higher = min(max(taxable - t["basic_band"], 0), t["additional_threshold"][year] - t["basic_band"])
    addl = max(taxable - t["additional_threshold"][year], 0)
    return basic * r["basic"] + higher * r["higher"] + addl * r["additional"]


def mileage_amount(miles, cfg):
    r = cfg["rates"]
    band = r["mileage_first_band_miles"]
    return min(miles, band) * r["mileage_first_band_rate"] + max(miles - band, 0) * r["mileage_after_rate"]


def calc_claim(cfg, rows, slips):
    """Return per-year results for the base case and the lower case (high-risk placements excluded)."""
    res = {}
    for y in cfg["claim_window"]["tax_years"]:
        yr = [r for r in rows if r["tax_year"] == y]
        income = sum(float(s["taxable_pay_period"] or 0) for s in slips if s["tax_year"] == y)
        parking, parking_low = 0.0, 0.0
        for s in slips:
            if s["tax_year"] != y or not is_training_slip(s):
                continue
            if s["parking_tax_treatment"] != "post-tax":
                continue
            amt = float(s["parking_deduction"] or 0)
            parking += amt
            d = date.fromisoformat(s["pay_date"])
            hit = [p for p in cfg["placements"] if p["start"] <= d <= (p["end"] or date.max)]
            if not any(is_high_risk(p.get("risk")) for p in hit):
                parking_low += amt
        out = {"income": income}
        for name, sel, pk in (("base", yr, parking),
                              ("low", [r for r in yr if not is_high_risk(r["risk"])], parking_low)):
            miles = sum(float(r["return_miles"] or 0) for r in sel)
            approved = mileage_amount(miles, cfg)
            mapr = cfg["extras"]["mileage_allowance_received"][y]
            tolls = cfg["extras"]["unreimbursed_tolls"][y]
            allowable = max(0.0, approved - mapr + tolls + pk)
            refund = income_tax(income, y, cfg) - income_tax(max(0.0, income - allowable), y, cfg)
            out[name] = dict(days=len(sel), miles=miles, approved=approved, extras=tolls + pk,
                             parking=pk, tolls=tolls, mapr=mapr, allowable=allowable, refund=refund,
                             eff=(refund / allowable if allowable else 0.0),
                             route="iForm / P87" if allowable <= cfg["tax"]["p87_limit"] else "Self Assessment")
        out["deadline"] = cfg["tax"]["deadlines"][y]
        res[y] = out
    return res


def marginal_label(income, y, cfg):
    t = cfg["tax"]
    pa = t["personal_allowance"]
    if income > t["taper_start"]:
        pa = max(0.0, pa - (income - t["taper_start"]) / 2)
    taxable = income - pa
    if income > t["taper_start"] and income < t["taper_start"] + 2 * t["personal_allowance"]:
        return "60% (taper)"
    if taxable > t["additional_threshold"][y]:
        return "45%"
    return "40%" if taxable > t["basic_band"] else "20%"


def write_summary(cfg, res, path):
    L = ["# Claim summary\n",
         "Calculation and evidence aid, not tax advice. All figures are estimates from payslips and a working-day attendance estimate (no rota evidence yet).\n",
         "Base case = every claimable training placement. Rates: 45p for the first 10,000 miles a year, 25p after. "
         "No mileage allowance payments were received. Parking is post-tax staff parking deducted on payslips. "
         "Unreimbursed tolls are 0 until you supply figures.\n",
         "## Base case\n",
         "| Tax year | Days | Miles | Approved amount | Extras (parking, tolls) | Allowable expenses | Marginal rate | Est. refund | Claim route | Deadline |",
         "|---|---|---|---|---|---|---|---|---|---|"]
    tot = collections.defaultdict(float)
    for y, o in res.items():
        b = o["base"]
        L.append(f"| {y} | {b['days']} | {b['miles']:.1f} | {b['approved']:.2f} | {b['extras']:.2f} | {b['allowable']:.2f} | "
                 f"{marginal_label(o['income'], y, cfg)} | {b['refund']:.2f} | {b['route']} | {o['deadline'].strftime('%d %b %Y')} |")
        for k in ("days", "miles", "approved", "extras", "allowable", "refund"):
            tot[k] += b[k]
    L.append(f"| **Total** | {tot['days']:.0f} | {tot['miles']:.1f} | {tot['approved']:.2f} | {tot['extras']:.2f} | "
             f"{tot['allowable']:.2f} | | **{tot['refund']:.2f}** | | |")
    L += ["", "## Lower case: high-risk placements excluded\n",
          "Excludes every day at placements labelled high risk in config.yaml (for example a site you returned to repeatedly, which HMRC may treat as one permanent workplace) and the parking deducted during them.\n",
          "| Tax year | Days | Miles | Approved amount | Extras | Allowable expenses | Est. refund | Claim route |",
          "|---|---|---|---|---|---|---|---|"]
    tl = collections.defaultdict(float)
    for y, o in res.items():
        b = o["low"]
        L.append(f"| {y} | {b['days']} | {b['miles']:.1f} | {b['approved']:.2f} | {b['extras']:.2f} | {b['allowable']:.2f} | {b['refund']:.2f} | {b['route']} |")
        for k in ("days", "miles", "approved", "extras", "allowable", "refund"):
            tl[k] += b[k]
    L.append(f"| **Total** | {tl['days']:.0f} | {tl['miles']:.1f} | {tl['approved']:.2f} | {tl['extras']:.2f} | {tl['allowable']:.2f} | **{tl['refund']:.2f}** | |")
    L += ["", "## Notes\n",
          "- Taxable income per year is the sum of taxable pay on all payslips found for that year (all employments). Missing locum months would understate it, so the rate is on the low side.",
          "- Refund = tax on that income minus tax on income less the allowable expenses, with the personal allowance taper above 100,000.",
          "- 2026/27 is provisional: it uses taxable pay to September 2026 only and attendance to 30 September 2026. Recalculate after 5 April 2027.",
          "- Placements set to claim: false in config.yaml (permanent workplace risk, out of time, or leave) are not counted.",
          "- The 2.5k P87 limit applies per year; every year is below it in both cases.",
          "- Days at high-risk placements are the least certain part of the claim. Check them with HMRC. See issues.md."]
    open(path, "w").write("\n".join(L) + "\n")


def write_xlsx(cfg, res, rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    a = wb.active
    a.title = "Assumptions"
    bold = Font(bold=True)
    a.append(["Input", "Value", "Note"])
    a.append(["Rate, first band (GBP/mile)", cfg["rates"]["mileage_first_band_rate"], "HMRC AMAP"])
    a.append(["Rate, after band (GBP/mile)", cfg["rates"]["mileage_after_rate"], ""])
    a.append(["Band (miles per tax year)", cfg["rates"]["mileage_first_band_miles"], ""])
    a.append(["P87 limit (GBP)", cfg["tax"]["p87_limit"], "Self Assessment above this"])
    a.append([])
    a.append(["Tax year", "Mileage allowance received", "Unreimbursed tolls", "Post-tax parking (base)",
              "Post-tax parking (high-risk excluded)", "Effective tax rate on expenses", "Deadline"])
    for y, o in res.items():
        a.append([y, o["base"]["mapr"], o["base"]["tolls"], round(o["base"]["parking"], 2),
                  round(o["low"]["parking"], 2), round(o["base"]["eff"], 4), o["deadline"]])
    a.append([])
    for t in ("The effective rate is calculated by build_claim.py from payslip taxable pay and the income tax bands.",
              "Change any input on this sheet (tolls, parking, rate) and the Summary sheet recalculates.",
              "Include = 0 on a Log row removes that day. Attendance dates are estimates, not rota records."):
        a.append([t])
    for c in a[1] + a[7]:
        c.font = bold
    lg = wb.create_sheet("Log")
    hdr = ["date", "tax_year", "from_postcode", "to_site", "purpose", "return_miles", "method_used",
           "placement_id", "risk", "include"]
    lg.append(hdr)
    for r in rows:
        lg.append([date.fromisoformat(r["date"]), r["tax_year"], r["from_postcode"], r["to_site"], r["purpose"],
                   r["return_miles"], r["method_used"], r["placement_id"], r["risk"], 1])
    n = len(rows) + 1
    sm = wb.create_sheet("Summary", 0)
    sm.append(["Tax year", "Miles", "Approved amount", "Extras (parking + tolls)", "Mileage allowance received",
               "Allowable expenses", "Effective rate", "Est. refund", "Claim route", "Deadline",
               "Miles (high-risk excluded)", "Allowable (high-risk excluded)", "Est. refund (high-risk excluded)"])
    for i, y in enumerate(res):
        r, ar = i + 2, i + 8
        sm.append([
            y,
            f'=SUMIFS(Log!$F$2:$F${n},Log!$B$2:$B${n},A{r},Log!$J$2:$J${n},1)',
            f'=MIN(B{r},Assumptions!$B$4)*Assumptions!$B$2+MAX(B{r}-Assumptions!$B$4,0)*Assumptions!$B$3',
            f'=Assumptions!C{ar}+Assumptions!D{ar}',
            f'=Assumptions!B{ar}',
            f'=MAX(0,C{r}-E{r}+D{r})',
            f'=Assumptions!F{ar}',
            f'=F{r}*G{r}',
            f'=IF(F{r}<=Assumptions!$B$5,"iForm / P87","Self Assessment")',
            f'=Assumptions!G{ar}',
            f'=B{r}-SUMIFS(Log!$F$2:$F${n},Log!$B$2:$B${n},A{r},Log!$I$2:$I${n},"high*",Log!$J$2:$J${n},1)',
            f'=MAX(0,MIN(K{r},Assumptions!$B$4)*Assumptions!$B$2+MAX(K{r}-Assumptions!$B$4,0)*Assumptions!$B$3-E{r}+Assumptions!C{ar}+Assumptions!E{ar})',
            f'=L{r}*G{r}'])
    t = len(res) + 2
    sm.append(["Total", f"=SUM(B2:B{t-1})", f"=SUM(C2:C{t-1})", f"=SUM(D2:D{t-1})", f"=SUM(E2:E{t-1})",
               f"=SUM(F2:F{t-1})", "", f"=SUM(H2:H{t-1})", "", "", f"=SUM(K2:K{t-1})", f"=SUM(L2:L{t-1})", f"=SUM(M2:M{t-1})"])
    for c in sm[1] + sm[t]:
        c.font = bold
    for col in "BCDEFHKLM":
        for c in sm[col]:
            if c.row > 1:
                c.number_format = "#,##0.00"
    for c in sm["G"][1:]:
        c.number_format = "0.0%"
    for c in sm["J"][1:]:
        c.number_format = "dd mmm yyyy"
    for c in lg["A"][1:]:
        c.number_format = "dd/mm/yyyy"
    wb.save(path)


def cmd_calc(a):
    cfg = load_config(a.config)
    rows, plan = attendance_rows(cfg)
    slips = list(csv.DictReader(open(os.path.join(a.out, "payslips.csv"))))
    res = calc_claim(cfg, rows, slips)
    write_summary(cfg, res, os.path.join(a.out, "summary.md"))
    write_xlsx(cfg, res, rows, os.path.join(a.out, "claim_summary.xlsx"))
    tb = sum(o["base"]["refund"] for o in res.values())
    tl = sum(o["low"]["refund"] for o in res.values())
    print(f"base refund {tb:.2f}, lower case {tl:.2f}")


def cmd_evidence(a):
    """Write evidence_index.md: payslips per tax year plus what is still missing."""
    cfg = load_config(a.config)
    slips = list(csv.DictReader(open(os.path.join(a.out, "payslips.csv"))))
    rows, plan = attendance_rows(cfg)
    L = ["# Evidence index\n",
         "Per tax year: payslip evidence on file, placement evidence, and what is still missing. "
         "Originals stay in your source folders; `payslips_md/` holds the converted copies.\n"]
    for y in cfg["claim_window"]["tax_years"]:
        ss = sorted((r for r in slips if r["tax_year"] == y), key=lambda r: r["pay_date"])
        main = [r for r in ss if is_training_slip(r)]
        pls = sorted({p["placement_id"] for p in rows if p["tax_year"] == y})
        L.append(f"## {y}\n")
        L.append(f"- Payslips on file: {len(ss)} ({len(main)} training post). Months: "
                 + (", ".join(r["pay_date"][:7] for r in main) or "none") + ".")
        L.append("- Placement evidence: TIS placements printout and ESR assignment history. "
                 + ("Claimed placement ids: " + ", ".join(map(str, pls)) + "." if pls else "No claimed placements."))
        L.append("- Parking: post-tax deductions on payslips "
                 + f"({sum(float(r['parking_deduction'] or 0) for r in main):.2f} GBP).")
        nt = [r["pay_date"] for r in main if "Misc Travel NP NT" in r["other_expense_lines"]]
        L.append("- Tunnel reimbursement lines (attendance evidence): " + (", ".join(nt) or "none") + ".")
        L.append("- Still missing: rota or e-roster export, annual and study leave records, home address move dates, unreimbursed toll receipts.\n")
    open(os.path.join(a.out, "evidence_index.md"), "w").write("\n".join(L) + "\n")


# ---------------------------------------------------------------------------
# prepare: rename converted payslips, split multi-slip files, report gaps
# ---------------------------------------------------------------------------
def _segments(text):
    lines = text.splitlines()
    idx = [i for i, l in enumerate(lines) if "ASSIGNMENT NUMBER" in l]
    if len(idx) <= 1:
        return [text]
    head = lines[:idx[0]]
    return ["\n".join(head + lines[i:(idx[k + 1] if k + 1 < len(idx) else len(lines))])
            for k, i in enumerate(idx)]


def cmd_prepare(a):
    emap = dict(x.split("=", 1) for x in (a.employer_map or "").split(",") if "=" in x)
    os.makedirs(a.out, exist_ok=True)
    seen, rows, nonpay, dups = {}, [], [], []
    files = sorted(glob.glob(os.path.join(a.raw, "**", "*.md"), recursive=True))
    for f in files:
        for seg in _segments(open(f).read()):
            m = re.search(r"PAY DATE\s+(\d+ \w+ \d{4})", seg)
            if not m or "ASSIGNMENT NUMBER" not in seg:
                nonpay.append(f)
                continue
            lines = seg.splitlines()
            i = next(k for k, l in enumerate(lines) if "ASSIGNMENT NUMBER" in l)
            asg = re.split(r"\s{3,}", lines[i + 2].strip())[0]
            net = re.search(r"NET PAY\s+(-?[\d.,]+)", seg)
            net = net.group(1) if net else "?"
            d = datetime.strptime(m.group(1).title(), "%d %b %Y").date()
            key = (d.isoformat(), asg, net)
            if key in seen:
                dups.append((f, seen[key]))
                continue
            seen[key] = f
            base, _, suf = asg.partition("-")
            emp = emap.get(base, "ASG" + base) + ("-" + suf if suf else "")
            rows.append(dict(src=f, text=seg, date=d, emp=emp))
    names = collections.Counter()
    for r in rows:
        stem = f"{r['date'].strftime('%Y-%m')}_{r['emp']}"
        names[stem] += 1
        r["dst"] = stem + ("" if names[stem] == 1 else f"_{names[stem]}") + ".md"
        open(os.path.join(a.out, r["dst"]), "w").write(r["text"])
    # gaps per employer label (main assignments only, not "-2" style extras)
    by = collections.defaultdict(set)
    for r in rows:
        by[r["emp"]].add((r["date"].year, r["date"].month))
    rep = ["# Conversion report\n", f"- Payslips written: {len(rows)}",
           f"- Duplicates skipped: {len(dups)}", f"- Non-payslip documents (P60s, statements, timesheets): {len(set(nonpay))}\n",
           "## Months present per employer label, and gaps\n"]
    for emp, ms in sorted(by.items()):
        lo, hi = min(ms), max(ms)
        allm, y, m = [], lo[0], lo[1]
        while (y, m) <= hi:
            allm.append((y, m))
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
        gaps = [f"{y}-{m:02d}" for y, m in allm if (y, m) not in ms]
        rep.append(f"- {emp}: {len(ms)} months, {lo[0]}-{lo[1]:02d} to {hi[0]}-{hi[1]:02d}; "
                   + (f"missing {', '.join(gaps)}" if gaps else "no gaps"))
    if dups:
        rep += ["", "## Duplicates skipped\n"] + [f"- {os.path.basename(x)} = {os.path.basename(y)}" for x, y in dups]
    if nonpay:
        rep += ["", "## Not payslips\n"] + [f"- {x}" for x in sorted(set(nonpay))]
    path = os.path.join(os.path.dirname(os.path.abspath(a.out)), "conversion_report.md")
    open(path, "w").write("\n".join(rep) + "\n")
    print(f"{len(rows)} payslips written, {len(dups)} duplicates, report: {path}")


# ---------------------------------------------------------------------------
# init and check: draft the config from TIS text, then validate it
# ---------------------------------------------------------------------------
BANK_HOLIDAYS = {
    2022: ["2022-01-03", "2022-04-15", "2022-04-18", "2022-05-02", "2022-06-02", "2022-06-03", "2022-08-29", "2022-09-19", "2022-12-26", "2022-12-27"],
    2023: ["2023-01-02", "2023-04-07", "2023-04-10", "2023-05-01", "2023-05-08", "2023-05-29", "2023-08-28", "2023-12-25", "2023-12-26"],
    2024: ["2024-01-01", "2024-03-29", "2024-04-01", "2024-05-06", "2024-05-27", "2024-08-26", "2024-12-25", "2024-12-26"],
    2025: ["2025-01-01", "2025-04-18", "2025-04-21", "2025-05-05", "2025-05-26", "2025-08-25", "2025-12-25", "2025-12-26"],
    2026: ["2026-01-01", "2026-04-03", "2026-04-06", "2026-05-04", "2026-05-25", "2026-08-31", "2026-12-25", "2026-12-28"],
    2027: ["2027-01-01", "2027-03-26", "2027-03-29", "2027-05-03", "2027-05-31", "2027-08-30", "2027-12-27", "2027-12-28"],
}


def parse_tis(path):
    """Placements from a TIS Self-Service placements printout converted by pdf2md."""
    t = open(path).read()
    blocks = re.split(r"\n(?=\s{10,}[A-Z][^\n]*\n\s+Need a CCT)", t)
    out = []
    for b in blocks[1:]:
        def f(k):
            m = re.search(r"\n\s+" + k + r"\s{2,}(\S[^\n]*(?:\n\s{40,}\S[^\n]*)*)", b)
            return re.sub(r"\s+", " ", m.group(1)) if m else ""
        try:
            st = datetime.strptime(f("Starts"), "%d/%m/%Y").date()
            en = datetime.strptime(f("Ends"), "%d/%m/%Y").date()
        except ValueError:
            continue
        out.append(dict(site=f("Site"), where=f("Site Location"), start=st, end=en,
                        wte=f("Whole Time Equivalent") or "1", grade=f("Grade"),
                        spec=f("Specialty"), kind=f("Placement Type")))
    return sorted(out, key=lambda x: x["start"])


def cmd_init(a):
    slips = list(csv.DictReader(open(os.path.join(a.out, "payslips.csv"))))
    today = date.today()
    years = sorted({r["tax_year"] for r in slips})
    years = [y for y in years if date(int("20" + y[-2:]), 4, 5) + timedelta(days=366 * 4) >= today
             and int(y[:4]) >= today.year - 5]
    pls = parse_tis(a.tis) if a.tis else []
    codes, sites = {}, []
    for p in pls:
        if p["site"] not in codes:
            code = re.sub(r"[^A-Z]", "", "".join(w[0] for w in p["site"].upper().split()))[:6] or "SITE"
            while code in codes.values():
                code += "X"
            codes[p["site"]] = code
            sites.append((code, p["site"], p["where"]))
    homes = []
    if a.homes and os.path.exists(a.homes):
        for line in open(a.homes):
            if line.strip() and not line.startswith("#"):
                pc, fr, to = [x.strip() for x in line.split(",")[:3]]
                homes.append((pc, fr, to or "null"))
    L = ["# Draft config. Fix every TODO, then run: python3 build_claim.py check --config config.yaml", "",
         "person:", '  label: "TODO name"', '  employer_lead: "TODO lead employer"', "",
         "rates:", "  mileage_first_band_miles: 10000", "  mileage_first_band_rate: 0.45", "  mileage_after_rate: 0.25",
         '  source: "HMRC Approved Mileage Allowance Payments"', "",
         "claim_window:", f"  end: {max(r['pay_date'] for r in slips)}  # last payslip date; edit to cut off sooner",
         "  tax_years: [" + ", ".join(f'"{y}"' for y in years) + "]", "",
         "attendance:", '  method_default: "fallback_working_days"', "  annual_leave_days_per_year: 27",
         "  # ASSUMPTION, not from records. Replace with real figures.", "  study_leave_days_per_year: 10",
         '  leave_rounding: "up"', "  excluded_periods:"]
    for p in pls:
        if "leave" in p["kind"].lower():
            L.append(f'    - {{from: {p["start"]}, to: {p["end"]}, reason: "{p["kind"]} (TIS)"}}')
    L.append("  bank_holidays: [" + ", ".join(d for y in BANK_HOLIDAYS.values() for d in y) + "]")
    L += ["", "homes:  # postcode, from, to (null for current). Supply with --homes homes.csv"]
    L += [f'  - {{postcode: "{pc}", from: {fr}, to: {to}}}' for pc, fr, to in homes] or \
         ['  - {postcode: "TODO", from: 2020-01-01, to: null}']
    L += ["", "sites:"]
    L += [f'  - {{code: "{c}", name: "{n}", postcode: "TODO", note: "{w}"}}' for c, n, w in sites] or \
         ['  - {code: "TODO", name: "TODO", postcode: "TODO"}']
    L += ["", "distances:", '  source: "TODO e.g. Google Maps, driving, postcode to postcode, date"', "  return_miles:"]
    pairs = []
    for pc, fr, to in homes:
        for c, _, _ in sites:
            pairs.append((pc, c))
    L += [f'    - {{row: {i + 1}, home: "{h}", site: "{c}", miles: TODO}}' for i, (h, c) in enumerate(pairs)] or \
         ['    - {row: 1, home: "TODO", site: "TODO", miles: TODO}']
    # Flag repeated returns to one site: more than 40 percent of the 24 months from the first
    # start, over more than one separate placement, or any single stay over 24 months.
    auto = {}
    for code in {c for c in codes.values()}:
        mine = [p for p in pls if codes[p["site"]] == code and "leave" not in p["kind"].lower()]
        if not mine:
            continue
        s0 = mine[0]["start"]
        w_end = s0 + timedelta(days=730)
        inside = sum(max(0, (min(p["end"], w_end) - max(p["start"], s0)).days + 1) for p in mine)
        for p in mine:
            if (p["end"] - p["start"]).days + 1 > 730:
                auto[id(p)] = "high_over_24_months"
            elif len(mine) > 1 and inside / 730 > 0.4:
                auto[id(p)] = "high_repeat_site"
    L += ["", "placements:  # claim: true to include. risk starting with high is dropped in the lower case. Add placements missing from TIS."]
    cutoff = date(2022, 4, 6)
    for i, p in enumerate(pls, 1):
        leave = "leave" in p["kind"].lower()
        claim = (p["end"] >= cutoff) and not leave
        note = "Parental or other leave, no attendance" if leave else ("Before the claim window" if p["end"] < cutoff else "")
        risk = auto.get(id(p), "")
        if risk:
            note = (note + " " if note else "") + "Auto-flagged: repeated returns or over 24 months at this site, check with HMRC"
        L.append(f'  - {{id: {i}, start: {p["start"]}, end: {p["end"]}, site: "{codes[p["site"]]}", grade: "{p["grade"]}", '
                 f'specialty: "{p["spec"]}", wte: {float(p["wte"]):.2f}, claim: {str(claim).lower()}, risk: "{risk}", note: "{note}"}}')
    if not pls:
        L.append('  - {id: 1, start: 2020-01-01, end: 2020-01-01, site: "TODO", grade: "TODO", specialty: "TODO", wte: 1.0, claim: true}')
    L += ["", "tax:", "  personal_allowance: 12570", "  taper_start: 100000", "  basic_band: 37700",
          "  rates: {basic: 0.20, higher: 0.40, additional: 0.45}",
          "  additional_threshold: {" + ", ".join(f'"{y}": {150000 if y < "2023/24" else 125140}' for y in years) + "}",
          "  deadlines: {" + ", ".join(f'"{y}": {int(y[:4]) + 5}-04-05' for y in years) + "}",
          "  p87_limit: 2500", "", "extras:",
          "  unreimbursed_tolls: {" + ", ".join(f'"{y}": 0' for y in years) + "}",
          "  mileage_allowance_received: {" + ", ".join(f'"{y}": 0' for y in years) + "}"]
    dest = os.path.join(a.out, "config.yaml")
    if os.path.exists(dest) and not a.force:
        dest = os.path.join(a.out, "config.draft.yaml")
    open(dest, "w").write("\n".join(L) + "\n")
    print(f"draft written to {dest} ({len(pls)} placements from TIS, {len(homes)} homes, {len(pairs)} distance rows to fill)")


def cmd_check(a):
    text = open(a.config).read()
    problems = [f"line {n}: TODO left in config" for n, l in enumerate(text.splitlines(), 1)
                if "TODO" in l and not l.lstrip().startswith("#")]
    if not problems:
        cfg = load_config(a.config)
        miles = {(d["home"], d["site"]) for d in cfg["distances"]["return_miles"]}
        for p in cfg["placements"]:
            if not p["claim"]:
                continue
            end = p["end"] or cfg["claim_window"]["end"]
            for d in (p["start"], min(end, cfg["claim_window"]["end"])):
                h = home_on(cfg, d)
                if h is None:
                    problems.append(f"placement {p['id']}: no home address on {d}")
                elif (h, p["site"]) not in miles:
                    problems.append(f"placement {p['id']}: no distance for {h} to {p['site']}")
        sites = {s["code"] for s in cfg["sites"]}
        problems += [f"placement {p['id']}: unknown site {p['site']}" for p in cfg["placements"] if p["site"] not in sites]
        pc = os.path.join(a.out, "payslips.csv")
        if os.path.exists(pc):
            slips = [r for r in csv.DictReader(open(pc)) if r["job_title"] and "bank" not in r["job_title"].lower()
                     and r["payscale"].lower().startswith(("spec", "found", "clin", "st"))]
            first = cfg["claim_window"]["tax_years"][0]
            window_start = date(int(first[:4]), 4, 6)
            for r in slips:
                d = date.fromisoformat(r["pay_date"])
                if d >= window_start and not any(p["start"] <= d <= (p["end"] or date.max) for p in cfg["placements"]):
                    problems.append(f"payslip {r['file']} ({r['pay_date']}, {r['location'][:25]}) falls outside every placement")
    if problems:
        print("\n".join(problems[:60]))
        print(f"{len(problems)} problem(s). Fix config.yaml and re-run check.")
        raise SystemExit(1)
    print("config OK")


def cmd_all(a):
    a.config = a.config or os.path.join(a.out, "config.yaml")
    cmd_extract(argparse.Namespace(md=a.md, out=a.out))
    if not os.path.exists(a.config):
        print("no config.yaml yet: run init, fill the TODOs, then run all again")
        raise SystemExit(1)
    cmd_check(argparse.Namespace(config=a.config, out=a.out))
    for fn in (cmd_attendance, cmd_calc, cmd_evidence):
        fn(argparse.Namespace(config=a.config, out=a.out))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("extract")
    e.add_argument("--md", required=True, help="folder of converted payslip .md files")
    e.add_argument("--out", default=".", help="output folder")
    e.set_defaults(fn=cmd_extract)
    t = sub.add_parser("attendance")
    t.add_argument("--config", required=True)
    t.add_argument("--out", default=".")
    t.set_defaults(fn=cmd_attendance)
    c = sub.add_parser("calc")
    c.add_argument("--config", required=True)
    c.add_argument("--out", default=".")
    c.set_defaults(fn=cmd_calc)
    v = sub.add_parser("evidence")
    v.add_argument("--config", required=True)
    v.add_argument("--out", default=".")
    v.set_defaults(fn=cmd_evidence)
    pr = sub.add_parser("prepare")
    pr.add_argument("--raw", required=True, help="folder of pdf2md output (searched recursively)")
    pr.add_argument("--out", required=True, help="payslips_md folder to write")
    pr.add_argument("--employer-map", default="", help='e.g. "12345678=TRUST1,87654321=TRUST2"')
    pr.set_defaults(fn=cmd_prepare)
    i = sub.add_parser("init")
    i.add_argument("--out", default=".")
    i.add_argument("--tis", help="TIS placements printout converted to .md")
    i.add_argument("--homes", help="homes.csv lines: postcode,from,to (to blank if current)")
    i.add_argument("--force", action="store_true", help="overwrite an existing config.yaml")
    i.set_defaults(fn=cmd_init)
    ck = sub.add_parser("check")
    ck.add_argument("--config", required=True)
    ck.add_argument("--out", default=".")
    ck.set_defaults(fn=cmd_check)
    al = sub.add_parser("all")
    al.add_argument("--md", required=True)
    al.add_argument("--out", default=".")
    al.add_argument("--config", default=None)
    al.set_defaults(fn=cmd_all)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
