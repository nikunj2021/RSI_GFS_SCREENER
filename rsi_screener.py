import pandas as pd
import yfinance as yf
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import time

# ─── CONFIG ───────────────────────────────────────────────────────────────────
RSI_PERIOD      = 14
MONTHLY_RSI_MIN = 60
WEEKLY_RSI_MIN  = 60
DAILY_RSI_MIN   = 40
DAILY_RSI_MAX   = 45
OUTPUT_DIR      = "output"
# ──────────────────────────────────────────────────────────────────────────────


def compute_rsi(series: pd.Series, period: int = 14) -> float:
    """Compute RSI from a price series, return the latest value."""
    delta = series.diff()
    gain  = delta.clip(lower=0)
    loss  = -delta.clip(upper=0)

    avg_gain = gain.ewm(com=period - 1, min_periods=period).mean()
    avg_loss = loss.ewm(com=period - 1, min_periods=period).mean()

    rs  = avg_gain / avg_loss.replace(0, np.nan)
    rsi = 100 - (100 / (1 + rs))
    return round(float(rsi.iloc[-1]), 2) if not rsi.empty else None


def load_symbols(filepath: str) -> list:
    """Read symbols from an Excel file. Accepts 'Symbol' or 'SYMBOL' column."""
    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip().str.upper()
    col = next((c for c in df.columns if "SYMBOL" in c), None)
    if col is None:
        raise ValueError(f"No 'Symbol' column found in {filepath}")
    symbols = df[col].dropna().astype(str).str.strip().str.upper().tolist()
    return symbols


def fetch_rsi(symbol: str) -> dict:
    """Fetch daily/weekly/monthly RSI and current price for a single symbol."""
    ticker = f"{symbol}.NS"
    try:
        raw = yf.download(ticker, period="5y", interval="1d",
                          auto_adjust=True, progress=False)
        if raw.empty or len(raw) < RSI_PERIOD + 5:
            return None

        close_daily = raw["Close"].squeeze()

        daily_rsi = compute_rsi(close_daily)

        weekly_close = close_daily.resample("W").last().dropna()
        weekly_rsi   = compute_rsi(weekly_close) if len(weekly_close) >= RSI_PERIOD + 5 else None

        monthly_close = close_daily.resample("ME").last().dropna()
        monthly_rsi   = compute_rsi(monthly_close) if len(monthly_close) >= RSI_PERIOD + 5 else None

        current_price = round(float(close_daily.iloc[-1]), 2)

        return {
            "Symbol":        symbol,
            "Current Price": current_price,
            "Daily RSI":     daily_rsi,
            "Weekly RSI":    weekly_rsi,
            "Monthly RSI":   monthly_rsi,
        }

    except Exception as e:
        print(f"  WARNING: Error fetching {symbol}: {e}")
        return None


def matches_filter(row: dict) -> bool:
    """Return True if stock passes all RSI conditions."""
    d = row.get("Daily RSI")
    w = row.get("Weekly RSI")
    m = row.get("Monthly RSI")
    if any(v is None for v in [d, w, m]):
        return False
    return (m > MONTHLY_RSI_MIN and
            w > WEEKLY_RSI_MIN  and
            DAILY_RSI_MIN < d < DAILY_RSI_MAX)


def tv_link(symbol: str) -> str:
    return f"https://www.tradingview.com/chart/?symbol=NSE:{symbol}"


def _thin_border(color="D0D0D0"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def style_header(cell, bg_hex="1F4E79"):
    cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    cell.fill      = PatternFill("solid", start_color=bg_hex)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _thin_border("FFFFFF")


def style_data_cell(cell, align="center"):
    cell.font      = Font(name="Arial", size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border    = _thin_border()


def highlight_rsi(cell, value, low, high):
    if value > high:
        cell.fill = PatternFill("solid", start_color="C6EFCE")   # green
    elif value >= low:
        cell.fill = PatternFill("solid", start_color="FFEB9C")   # yellow


def build_sheet(ws, results: list, sheet_label: str):
    run_date = datetime.now().strftime("%d-%b-%Y %H:%M IST")

    # ── Title row ──────────────────────────────────────────────────────────────
    ws.merge_cells("A1:G1")
    title = ws["A1"]
    title.value     = (f"{sheet_label}  |  RSI Screener  |  "
                       f"Run: {run_date}  |  {len(results)} stocks matched")
    title.font      = Font(bold=True, name="Arial", size=11, color="FFFFFF")
    title.fill      = PatternFill("solid", start_color="243F60")
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── Header row ─────────────────────────────────────────────────────────────
    headers    = ["#", "Symbol", "TradingView Link", "Current Price (₹)",
                  "Daily RSI", "Weekly RSI", "Monthly RSI"]
    col_widths = [5, 14, 22, 20, 12, 12, 13]

    for col_idx, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        style_header(cell)
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[2].height = 20

    # ── Data rows ──────────────────────────────────────────────────────────────
    for row_num, stock in enumerate(results, start=3):
        ws.row_dimensions[row_num].height = 18
        sym = stock["Symbol"]

        row_data = [
            row_num - 2,
            sym,
            tv_link(sym),
            stock["Current Price"],
            stock["Daily RSI"],
            stock["Weekly RSI"],
            stock["Monthly RSI"],
        ]

        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            style_data_cell(cell, align="left" if col_idx in (2, 3) else "center")

            if col_idx == 4:   # Price — currency format
                cell.number_format = '#,##0.00'
            elif col_idx == 5: # Daily RSI
                highlight_rsi(cell, value, DAILY_RSI_MIN, DAILY_RSI_MAX)
            elif col_idx == 6: # Weekly RSI
                highlight_rsi(cell, value, WEEKLY_RSI_MIN, 80)
            elif col_idx == 7: # Monthly RSI
                highlight_rsi(cell, value, MONTHLY_RSI_MIN, 80)

    # ── Freeze below header ────────────────────────────────────────────────────
    ws.freeze_panes = "A3"

    # ── Legend (below data) ────────────────────────────────────────────────────
    legend_row = len(results) + 4
    ws.cell(row=legend_row, column=1, value="Legend:").font = Font(bold=True, name="Arial", size=9)
    for col, (txt, color) in enumerate([
        ("Monthly RSI > 60", "C6EFCE"),
        ("Weekly RSI > 60",  "C6EFCE"),
        ("Daily RSI 40–45",  "FFEB9C"),
    ], start=2):
        c = ws.cell(row=legend_row, column=col, value=txt)
        c.fill      = PatternFill("solid", start_color=color)
        c.font      = Font(name="Arial", size=9)
        c.alignment = Alignment(horizontal="center")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    sources = {
        "NSE500":      "data/nse500list.xlsx",
        "MicroCap250": "data/Micro250list.xlsx",
    }

    wb          = Workbook()
    all_matched = []
    sheet_index = 0

    for sheet_name, filepath in sources.items():
        if not os.path.exists(filepath):
            print(f"WARNING: File not found: {filepath} — skipping")
            continue

        print(f"\n{'='*55}")
        print(f"Loading: {filepath}")
        symbols = load_symbols(filepath)
        seen    = set()
        symbols = [s for s in symbols if not (s in seen or seen.add(s))]
        print(f"   {len(symbols)} unique symbols loaded")

        matched = []
        for i, sym in enumerate(symbols, 1):
            print(f"  [{i:>3}/{len(symbols)}] {sym}", end="  ")
            data = fetch_rsi(sym)
            if data and matches_filter(data):
                matched.append(data)
                print(f"MATCH  D:{data['Daily RSI']}  W:{data['Weekly RSI']}  M:{data['Monthly RSI']}")
            else:
                print("-")
            time.sleep(0.3)

        print(f"\n  {len(matched)} stocks matched from {sheet_name}")
        all_matched.extend(matched)

        ws = wb.active if sheet_index == 0 else wb.create_sheet(title=sheet_name)
        if sheet_index == 0:
            ws.title = sheet_name
        build_sheet(ws, matched, sheet_name)
        sheet_index += 1

    # Combined sheet
    if sheet_index > 1 and all_matched:
        ws_all = wb.create_sheet(title="ALL_MATCHES")
        build_sheet(ws_all, all_matched, "All Sources Combined")

    date_str    = datetime.now().strftime("%Y-%m-%d")
    output_path = os.path.join(OUTPUT_DIR, f"RSI_Screener_{date_str}.xlsx")
    wb.save(output_path)
    print(f"\nReport saved: {output_path}")


if __name__ == "__main__":
    main()
