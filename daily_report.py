"""Daily special-situations workbook in the Event Monitor layout.

Runs the terminal's public-source event scan and the SEC deal sheet, applies the screening
standards (standards.py), and writes one workbook with the Event Monitor's tabs:

    Funnel     how many names survive each standard
    Ideas      deals clearing the 25% risk-adjusted hurdle (tier 1 and verify)
    Assessed   every priced deal, then the unpriced hard catalysts, with values, returns and risks
    Events     hard-catalyst leads with a ticker
    Pipeline   hard-catalyst leads without a ticker yet (UK offer periods, Form 10 spin-offs)
    Rejected   everything screened out, with the stage and reason
    Config     the parameters behind the run

Usage: py daily_report.py [--out FOLDER] [--source public] [--days 270] [--email [ADDRESS]] [--open]
       py daily_report.py --email-only WORKBOOK    (resend an existing workbook)
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

import deal_sheet
import ibkr_market_scanner as scanner
import standards

BAND_FILL = PatternFill("solid", fgColor="002060")
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
PRICE_FORMAT = '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)'
CAP_FORMAT = '"$"#,##0_);[Red]\\("$"#,##0\\)'
PCT_FORMAT = "0.0%;[Red]-0.0%"
MISSING = "--"

# (column, header) for the Ideas and Assessed tabs; spacer columns A, G, K, Q, X, AB and AJ stay empty.
IDEA_COLUMNS = [
    ("B", "Company"), ("C", "Ticker"), ("D", "Closing Price"), ("E", "Market Cap ($ mn)"), ("F", "Country"),
    ("H", "Event Type(s)"), ("I", "Situation"), ("J", "Screen Flags"),
    ("L", "Base Value"), ("M", "Bull Value"), ("N", "Bear Value"), ("O", "Horizon (mo)"), ("P", "p(Base / Bull / Bear)"),
    ("R", "Base Gain (%)"), ("S", "Base Annualized (%)"), ("T", "Bear Drawdown (%)"), ("U", "Prob-Wtd Gain (%)"),
    ("V", "Verdict"), ("W", "Hurdle Reason"),
    ("Y", "Catalyst"), ("Z", "Status"), ("AA", "Date"),
    ("AC", "Thesis"), ("AD", "Value Basis"), ("AE", "Kill Criteria"), ("AF", "What Could I Be Wrong About?"),
    ("AG", "Facts (source phrase)"), ("AH", "Estimates (labeled)"), ("AI", "Diligence Next"),
    ("AK", "Filing Links"),
]
IDEA_BANDS = [("B", "Price Related"), ("H", "Event"), ("L", "Analyst Values (per share)"),
              ("R", "Returns (hurdle applied in code)"), ("Y", "Catalyst"), ("AC", "Thesis & Risk"), ("AK", "Sources")]
IDEA_WIDTHS = {"A": 1.7, "B": 26, "C": 8.7, "D": 10.7, "E": 10.7, "F": 7.7, "G": 1.7, "H": 18, "I": 40, "J": 18,
               "K": 1.7, "L": 9.7, "M": 9.7, "N": 9.7, "O": 8.7, "P": 11, "Q": 1.7, "R": 9.7, "S": 10.7, "T": 9.7,
               "U": 9.7, "V": 14, "W": 26, "X": 1.7, "Y": 32, "Z": 10, "AA": 10.7, "AB": 1.7, "AC": 50, "AD": 40,
               "AE": 36, "AF": 40, "AG": 50, "AH": 36, "AI": 30, "AJ": 1.7, "AK": 60}
WRAP = {"I", "W", "Y", "Z", "AC", "AD", "AE", "AF", "AG", "AH", "AI"}
COUNTRY_CODES = {
    "United States": "US", "Canada": "CA", "United Kingdom": "UK", "Ireland": "IE", "Germany": "DE", "France": "FR",
    "Netherlands": "NL", "Belgium": "BE", "Austria": "AT", "Finland": "FI", "Italy": "IT", "Spain": "ES",
    "Portugal": "PT", "Switzerland": "CH", "Sweden": "SE", "Norway": "NO", "Denmark": "DK", "Israel": "IL",
    "Japan": "JP", "Australia": "AU", "New Zealand": "NZ", "Hong Kong": "HK", "Singapore": "SG",
}
VERDICT_ORDER = {"CLEARS": 0, "VERIFY": 1, "BELOW_HURDLE": 2, "SCREENED": 3, "INSUFFICIENT_DATA": 4}


def bullets(items) -> str:
    items = [str(i) for i in items if i]
    return "\n".join(f"• {i}" for i in items) if items else MISSING


def country_code(basis: str) -> str:
    """'Domicile United States' / 'Listed in Australia; ...' -> 'US' / 'AU'."""
    match = re.match(r"(?:Domicile|Listed in) ([A-Za-z ,.'-]+?)(?:;|\.|$| is )", basis or "")
    name = match.group(1).strip() if match else ""
    return COUNTRY_CODES.get(name, name or MISSING)


def headline_form(item: dict) -> str:
    provider = item.get("provider", "")
    return provider.split(" · ", 1)[1].replace(" (bidder)", "") if " · " in provider else provider


# ---------------------------------------------------------------- rows

def deal_idea(row: dict) -> dict:
    """A priced deal-sheet row in Ideas/Assessed terms."""
    risk = row["risk"]
    if row["tier"] == deal_sheet.TIER_CLEARS:
        verdict = "CLEARS"
    elif row["tier"] == deal_sheet.TIER_VERIFY:
        verdict = "VERIFY"
    elif row["screen_reason"].startswith("Below hurdle"):
        verdict = "BELOW_HURDLE"
    else:
        verdict = "SCREENED"
    p = risk["probability"]
    bull = row["consideration_value"]   # plus the CVR's maximum when there is one
    cvr_max = None
    match = re.search(r"up to ([\d.]+)", row["deal_terms"])
    if "CVR" in row["deal_terms"] and match:
        cvr_max = float(match.group(1))
        bull = row["consideration_value"] + cvr_max
    reason = (f"risk-adj. {risk['risk_adjusted_annualized']:+.0%} annualized vs {standards.HURDLE:.0%} hurdle "
              f"({risk['days_expected']}d incl. {risk['slip_days']}d slip)")
    if row["screen_reason"]:
        reason = row["screen_reason"] + "; " + reason
    market = f"{risk['market_probability']:.0%}" if risk["market_probability"] is not None else "n/a"
    estimates = [f"Completion probability {p:.0%} (model: {risk['probability_basis']})",
                 f"Market-implied probability {market}",
                 f"Bear value = {risk['break_basis']}",
                 f"Horizon = guided close plus {risk['slip_days']}d regulatory slip",
                 f"Model size {deal_sheet.sizing_note(risk)}"]
    if risk["probability_needed"] is not None:
        estimates.insert(2, f"Hurdle needs p >= {max(0, risk['probability_needed']):.0%}"
                            + (f", close by {risk['latest_close']}" if risk["latest_close"] else ""))
    diligence = ["Confirm consideration and conditions in the source filing",
                 "Regulatory approvals and timetable (HSR, foreign filings, sector regulators)"]
    if "stock" in row["situation_type"]:
        diligence.append("Borrow availability and cost on the acquirer shares")
    if row["situation_type"].startswith("Tender"):
        diligence.append("Tender expiry, minimum condition and extension history")
    if cvr_max or "CVR" in row["deal_terms"]:
        diligence.append("CVR milestones, probability and timing")
    situation = f"{row['company']}: {row['situation_type'].lower()} by {row['counterparty'] or 'buyer not named'} - {row['deal_terms']}"
    return {
        "verdict": verdict, "sort": -risk["risk_adjusted_annualized"],
        "B": row["company"], "C": row["ticker"], "D": row["price"],
        "E": row["market_cap"] / 1e6 if row["market_cap"] else MISSING, "F": country_code(row["domicile"]),
        "H": row["situation_type"], "I": situation,
        "J": "; ".join(f["label"] for f in risk["flags"]) or MISSING,
        "L": round(row["consideration_value"], 4), "M": round(bull, 4), "N": round(risk["break_price"], 4),
        "O": round(risk["days_expected"] / 30.44, 1), "P": f"{p:.0%} / 0% / {1 - p:.0%}", "p": p,
        "V": verdict, "W": reason,
        "Y": row["key_dates_catalyst"], "Z": row["status"] or MISSING, "AA": row["close_date"],
        "AC": risk["bull"] + ". " + risk["bear"],
        "AD": f"Deal consideration {row['consideration_text']}; bear = break price ({risk['break_basis']})",
        "AE": bullets(risk["kill_criteria"]), "AF": bullets(risk["what_could_be_wrong"]),
        "AG": bullets(f"{f['label']} ('{f['evidence']}')" for f in risk["flags"]),
        "AH": bullets(estimates), "AI": bullets(diligence), "AK": row["sources"],
    }


def lead_idea(candidate) -> dict:
    """An unpriced hard catalyst: listed in Assessed so it is not lost, but without values."""
    research = candidate.research
    first = candidate.headlines[0]
    basis = research.get("risks", "").split(". Price it", 1)[0]
    symbol = str(getattr(candidate.contract, "localSymbol", "") or getattr(candidate.contract, "symbol", ""))
    return {
        "verdict": "INSUFFICIENT_DATA", "sort": 0,
        "B": candidate.description, "C": symbol, "D": candidate.price if candidate.price is not None else MISSING,
        "E": MISSING, "F": country_code(basis),
        "H": research.get("tier", "").replace("Hard catalyst · ", ""),
        "I": "; ".join(h.get("headline", "") for h in candidate.headlines[:3]),
        "J": basis if "MSCI" in basis else MISSING,
        "L": MISSING, "M": MISSING, "N": MISSING, "O": MISSING, "P": "0% / 0% / 0%", "p": None,
        "V": "INSUFFICIENT_DATA", "W": "no priced terms: enter terms, probability and break price in the deal calculator",
        "Y": research.get("key_dates_catalyst") or first.get("headline", ""),
        "Z": research.get("status_approvals") or MISSING, "AA": first.get("time", "") or MISSING,
        "AC": f"Primary-source {research.get('tier', '').replace('Hard catalyst · ', '').lower()} in a developed market; "
              "not yet priced against the hurdle.",
        "AD": "No base value: terms are not read automatically for this source",
        "AE": MISSING, "AF": bullets(["Offer terms, conditions and timetable are unverified until read from the announcement"]),
        "AG": bullets(f"{h.get('provider', '')}: {h.get('headline', '')}" for h in candidate.headlines[:3]),
        "AH": MISSING,
        "AI": bullets(["Read the offer announcement or scheme document for terms",
                       "Price the target and set a break price",
                       f"Counterparty: {research['counterparty']}" if research.get("counterparty") else ""]),
        "AK": first.get("url", ""),
    }


def has_ticker(symbol: str) -> bool:
    return bool(symbol) and not standards.ISIN_RE.match(symbol.upper()) and not symbol.upper().startswith("CIK") \
        and symbol != "Unverified"


# ---------------------------------------------------------------- workbook

def _header(ws, row: int, labels: list[str], widths: list[float]):
    for column, (label, width) in enumerate(zip(labels, widths), 1):
        cell = ws.cell(row=row, column=column, value=label)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        ws.column_dimensions[cell.column_letter].width = width


def write_ideas_sheet(ws, title: str, ideas: list[dict], empty_note: str, footer: str, autofilter: bool):
    for column, width in IDEA_WIDTHS.items():
        ws.column_dimensions[column].width = width
    ws["B2"] = title
    ws["B2"].font = Font(bold=True)
    for column, label in IDEA_BANDS:
        cell = ws[f"{column}4"]
        cell.value = label
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = BAND_FILL
    for column, label in IDEA_COLUMNS:
        cell = ws[f"{column}5"]
        cell.value = label
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "D6"
    r = 6
    if not ideas:
        ws[f"B{r}"] = empty_note
        r += 1
    for idea in ideas:
        p = idea["p"]
        formulas = {
            "R": f'=IFERROR(L{r}/D{r}-1,"--")',
            "S": f'=IFERROR((L{r}/D{r})^(12/O{r})-1,"--")',
            "T": f'=IFERROR(N{r}/D{r}-1,"--")',
            "U": f'=IFERROR({p:.4f}*L{r}/D{r}+{1 - p:.4f}*N{r}/D{r}-1,"--")' if p is not None else MISSING,
        }
        for column, _ in IDEA_COLUMNS:
            value = formulas.get(column, idea.get(column, MISSING))
            cell = ws[f"{column}{r}"]
            cell.value = MISSING if value in (None, "") else value
            cell.alignment = Alignment(vertical="top", wrap_text=column in WRAP)
            if column == "D":
                cell.number_format = PRICE_FORMAT
            elif column == "E":
                cell.number_format = CAP_FORMAT
            elif column in ("L", "M", "N"):
                cell.number_format = "0.00"
            elif column == "O":
                cell.number_format = "0.0"
            elif column in ("R", "S", "T", "U"):
                cell.number_format = PCT_FORMAT
            elif column == "V":
                cell.font = Font(bold=True)
            elif column == "AK" and str(value).startswith("http"):
                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")
        r += 1
    if autofilter and ideas:
        ws.auto_filter.ref = f"B5:AK{r - 1}"
    ws[f"B{r + 1}"] = footer
    ws[f"B{r + 1}"].font = Font(italic=True, color="666666")


def write_table(ws, headers: list[str], widths: list[float], rows: list[list], formats: dict[int, str] | None = None,
                link_column: int | None = None):
    _header(ws, 1, headers, widths)
    ws.freeze_panes = "A2"
    for r, values in enumerate(rows, 2):
        for column, value in enumerate(values, 1):
            cell = ws.cell(row=r, column=column, value=value if value not in (None, "") else None)
            if formats and column in formats and isinstance(value, (int, float)):
                cell.number_format = formats[column]
            if column == link_column and isinstance(value, str) and value.startswith("http"):
                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")


def build_workbook(path: Path, pulled: datetime, candidates: list, warnings: list[str], summary: dict,
                   deal_rows: list[dict], config: dict) -> Path:
    run_date = f"{pulled:%Y-%m-%d}"
    leads = [c for c in candidates if c.headlines and c.research.get("tier", "").startswith("Hard catalyst")]
    screened = summary.get("screened_out") or []
    deal_ideas = [deal_idea(row) for row in deal_rows]
    priced = {idea["C"] for idea in deal_ideas}
    tradable = [c for c in leads if has_ticker(str(getattr(c.contract, "localSymbol", "") or c.contract.symbol))]
    pipeline = [c for c in leads if c not in tradable]
    unpriced = [lead_idea(c) for c in tradable if str(c.contract.localSymbol or c.contract.symbol) not in priced]

    ideas = sorted((i for i in deal_ideas if i["verdict"] in ("CLEARS", "VERIFY")),
                   key=lambda i: (VERDICT_ORDER[i["verdict"]], i["sort"]))
    assessed = sorted(deal_ideas, key=lambda i: (VERDICT_ORDER[i["verdict"]], i["sort"])) + unpriced

    workbook = Workbook()
    funnel = workbook.active
    funnel.title = "Funnel"
    universe_rejects = sum(1 for s in screened if s["reason"].startswith("Outside universe"))
    clears = sum(1 for i in deal_ideas if i["verdict"] == "CLEARS")
    verify = sum(1 for i in deal_ideas if i["verdict"] == "VERIFY")
    live = sum(1 for i in deal_ideas if i["verdict"] != "SCREENED")
    funnel_rows = [
        ["Event leads detected (all sources)", len(leads) + len(screened)],
        ["Primary-source hard catalysts", len(leads) + universe_rejects],
        ["In MSCI developed markets", len(leads)],
        ["Tradable (ticker known)", len(tradable)],
        ["Pipeline (no ticker yet)", len(pipeline)],
        ["SEC deals priced (deal sheet)", len(deal_rows)],
        ["Priced deals in universe and live", live],
        ["Ideas clearing hurdle", clears + verify],
        ["  of which tier 1 (clears)", clears],
        ["  of which verify (market disagrees)", verify],
        ["", ""],
        ["Hard catalysts by type", ""],
    ]
    by_type: dict[str, int] = {}
    for c in leads:
        kind = c.research["tier"].replace("Hard catalyst · ", "")
        by_type[kind] = by_type.get(kind, 0) + 1
    funnel_rows += [[f"  {kind}", count] for kind, count in sorted(by_type.items(), key=lambda kv: -kv[1])]
    funnel_rows += [["", ""], ["Top screen-out reasons", ""]]
    reasons: dict[str, int] = {}
    for s in screened:
        key = s["reason"].split(":")[0]
        reasons[key] = reasons.get(key, 0) + 1
    funnel_rows += [[f"  {reason}", count] for reason, count in sorted(reasons.items(), key=lambda kv: -kv[1])[:8]]
    funnel_rows += [["", ""], ["Run pulled", pulled.strftime("%Y-%m-%d %H:%M %Z")]]
    funnel_rows += [["Warning", w] for w in warnings[:6]]
    write_table(funnel, ["Stage", "Names remaining"], [70, 16], funnel_rows)

    hurdle = f"{standards.HURDLE:.0%}"
    footer = (f"Cleared names are ranked leads for full diligence, not recommendations. Hurdle: probability-weighted "
              f"return of at least {hurdle} annualized over the guided close plus regulatory slip; terms and risk "
              "flags are read by pattern from filings, so verify against Filing Links.")
    write_ideas_sheet(workbook.create_sheet("Ideas"),
                      f"Special Situations - run {run_date}: {len(ideas)} idea(s) clear the hurdle, "
                      "ranked by risk-adjusted annualized return", ideas, "Nothing cleared the hurdle today.", footer, False)
    write_ideas_sheet(workbook.create_sheet("Assessed"),
                      f"Special Situations - run {run_date}: all {len(assessed)} assessed names (cleared, verify, "
                      "below hurdle, screened, then unpriced hard catalysts)", assessed, "Nothing assessed today.",
                      footer, True)

    def event_row(c):
        first = c.headlines[0]
        symbol = str(c.contract.localSymbol or c.contract.symbol)
        basis = c.research.get("risks", "").split(". Price it", 1)[0]
        deal = next((row for row in deal_rows if row["ticker"] == symbol), None)
        return [symbol, c.description, c.research["tier"].replace("Hard catalyst · ", ""), headline_form(first),
                str(first.get("time", ""))[:10], country_code(basis),
                deal["price"] if deal else c.price, deal["market_cap"] / 1e6 if deal and deal["market_cap"] else None,
                None, basis if "MSCI" in basis else "", first.get("url", "")]

    write_table(workbook.create_sheet("Events"),
                ["Ticker", "Company", "Category", "Form", "Filed", "Country", "Closing Price", "Market Cap ($ mn)",
                 "Size % Mkt Cap", "Flags", "URL"],
                [11, 34, 20, 11, 11, 11, 15, 19, 16, 34, 70],
                [event_row(c) for c in tradable], {7: PRICE_FORMAT, 8: CAP_FORMAT}, link_column=11)
    write_table(workbook.create_sheet("Pipeline"),
                ["Company", "Category", "Form", "Filed", "Country", "URL"], [40, 20, 11, 11, 11, 70],
                [[c.description, c.research["tier"].replace("Hard catalyst · ", ""), headline_form(c.headlines[0]),
                  str(c.headlines[0].get("time", ""))[:10],
                  country_code(c.research.get("risks", "").split(". Price it", 1)[0]), c.headlines[0].get("url", "")]
                 for c in pipeline], link_column=6)

    rejected = []
    for s in screened:
        stage = "Universe" if s["reason"].startswith("Outside universe") else "Catalyst"
        source = s.get("source", "")
        rejected.append([s["ticker"], s["company"], source.split(" · ")[0], stage, s["reason"],
                         source.split(" · ", 1)[1] if " · " in source else "", "", s.get("url", "")])
    for row in deal_rows:
        if row["tier"] in deal_sheet.TIER_ORDER:
            continue
        rejected.append([row["ticker"], row["company"], row["situation_type"], "Deal sheet", row["screen_reason"],
                         "", "", row["sources"]])
    write_table(workbook.create_sheet("Rejected"),
                ["Ticker", "Company", "Category", "Stage", "Reason", "Form", "Filed", "URL"],
                [11, 34, 20, 11, 60, 11, 11, 70], rejected, link_column=8)
    write_table(workbook.create_sheet("Config"), ["Parameter", "Value"], [28, 40],
                [[key, value] for key, value in config.items()])

    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(path)
    except PermissionError as error:
        raise RuntimeError(f"Close {path.name} in Excel and run again.") from error
    return path


def run_config(source_mode: str, days: int) -> dict:
    return {
        "source_mode": source_mode,
        "deal_lookback_days": days,
        "hurdle_risk_adjusted_annualized": standards.HURDLE,
        "nav_usd": standards.NAV_USD,
        "break_loss_budget_pct_nav": standards.BREAK_LOSS_BUDGET,
        "max_position_pct_nav": standards.MAX_POSITION,
        "adv_participation": standards.ADV_PARTICIPATION,
        "build_days": standards.BUILD_DAYS,
        "break_haircut": standards.BREAK_HAIRCUT,
        "fallback_break_loss": standards.FALLBACK_BREAK_LOSS,
        "edge_verify": standards.EDGE_VERIFY,
        "verify_spread": standards.VERIFY_SPREAD,
        "base_probability_merger": standards.BASE_PROBABILITY["merger"],
        "base_probability_tender": standards.BASE_PROBABILITY["tender"],
        "base_slip_days_merger": standards.BASE_SLIP_DAYS["merger"],
        "base_slip_days_tender": standards.BASE_SLIP_DAYS["tender"],
        "universe": "MSCI developed markets: " + ", ".join(sorted(standards.DM_COUNTRIES)),
    }


def run(out_dir: Path, source_mode: str = "public", days: int = 270, log=print) -> Path:
    pulled = datetime.now().astimezone()
    log(f"Scanning {scanner.SOURCE_MODES[source_mode]}...")
    candidates, warnings, summary = scanner._scan_selected_sources(source_mode)
    log(f"  {sum(1 for c in candidates if c.headlines)} hard catalysts kept, "
        f"{len(summary.get('screened_out') or [])} screened out")
    log("Building the SEC deal sheet...")
    # Its own folder, so a deal sheet left open in Excel does not stop the daily run.
    _, deal_rows = deal_sheet.build_deal_sheet(days, out_dir=out_dir / "deal_sheets", log=lambda m: log("  " + m))
    path = out_dir / f"special_situations_monitor_{pulled:%Y-%m-%d}.xlsx"
    build_workbook(path, pulled, candidates, warnings, summary, deal_rows, run_config(source_mode, days))
    log(f"Daily workbook: {path}")
    return path


def email_body(path: Path) -> str:
    """Plain-text summary for the email: the funnel and the ideas that clear, read back from the workbook."""
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True)
    lines = [f"Special situations monitor - {path.stem.rsplit('_', 1)[-1]}", ""]
    for stage, count in workbook["Funnel"].iter_rows(min_row=2, max_row=11, values_only=True):
        if stage:
            lines.append(f"{stage}: {count}")
    lines += ["", "Ideas clearing the risk-adjusted hurdle:"]
    ideas = [row for row in workbook["Ideas"].iter_rows(min_row=6, values_only=True) if row[2] and row[2] != MISSING]
    for row in ideas:
        values = dict(zip([c for c, _ in IDEA_COLUMNS], [row[_column_index(c)] for c, _ in IDEA_COLUMNS]))
        lines.append(f"- {values['C']} ({values['B']}): {values['V']} - {values['H']}; {values['W']}")
    if not ideas:
        lines.append("- None today.")
    lines += ["", "Verify-tier names clear only because the model is far more confident than the market: check the "
              "terms in the filing links before acting. Research leads, not recommendations.",
              "Full detail (Assessed, Events, Pipeline, Rejected) is in the attached workbook."]
    workbook.close()
    return "\n".join(lines)


def _column_index(letter: str) -> int:
    index = 0
    for char in letter:
        index = index * 26 + ord(char) - 64
    return index - 1


def send_email(path: Path, to: str = "") -> None:
    """Email the workbook through desktop Outlook (send_outlook_mail.ps1)."""
    import subprocess
    import tempfile

    with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as body:
        body.write(email_body(path))
    script = Path(__file__).resolve().parent / "send_outlook_mail.ps1"
    command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
               "-Subject", f"Special situations monitor {path.stem.rsplit('_', 1)[-1]}",
               "-BodyFile", body.name, "-Attachment", str(path)]
    if to:
        command += ["-To", to]
    try:
        result = subprocess.run(command, capture_output=True, text=True, timeout=180)
    finally:
        os.unlink(body.name)
    if result.returncode:
        raise RuntimeError(f"Outlook send failed: {(result.stderr or result.stdout).strip()}")
    print(result.stdout.strip())


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", type=Path, default=Path(os.environ.get(
        "DAILY_REPORT_DIR", Path(__file__).resolve().parent / "reports" / "daily")), help="folder for the workbook")
    parser.add_argument("--source", default="public", choices=sorted(scanner.SOURCE_MODES), help="scan sources")
    parser.add_argument("--days", type=int, default=270, help="deal-sheet lookback in days")
    parser.add_argument("--open", action="store_true", help="open the workbook when done")
    parser.add_argument("--email", nargs="?", const="", default=None, metavar="ADDRESS",
                        help="email the workbook through desktop Outlook (no address: the Outlook account's own)")
    parser.add_argument("--email-only", type=Path, metavar="WORKBOOK", help="email an existing workbook and stop")
    args = parser.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    if args.email_only:
        send_email(args.email_only, args.email or "")
        return 0
    path = run(args.out, args.source, args.days)
    if args.email is not None:
        send_email(path, args.email)
    if args.open:
        os.startfile(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
