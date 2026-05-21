# NSE RSI Screener 📊

Automated weekly RSI screener for NSE stocks — runs every **Saturday at 9:05 AM IST** via GitHub Actions.  
Scans **NSE 500** and **Microcap 250** stocks and generates a formatted Excel report.

---

## Scan Conditions

| Timeframe | Condition |
|-----------|-----------|
| Monthly RSI | > 60 |
| Weekly RSI  | > 60 |
| Daily RSI   | > 40 and < 45 |

All three conditions must be **true simultaneously** for a stock to appear in the report.

---

## Report Columns

| Col | Header | Description |
|-----|--------|-------------|
| 1 | Symbol | NSE ticker symbol |
| 2 | TradingView Link | Direct chart link |
| 3 | Current Price (₹) | Last closing price |
| 4 | Daily RSI | 14-period RSI on daily candles |
| 5 | Weekly RSI | 14-period RSI on weekly candles |
| 6 | Monthly RSI | 14-period RSI on monthly candles |

---

## Folder Structure

```
nse-rsi-screener/
├── .github/
│   └── workflows/
│       └── rsi_screener.yml   ← Automated schedule
├── data/
│   ├── nse500list.xlsx        ← Must have a "Symbol" column
│   └── Micro250list.xlsx      ← Must have a "Symbol" column
├── output/                    ← Reports saved here (auto-created)
├── rsi_screener.py            ← Main script
├── requirements.txt
└── README.md
```

---

## Setup Instructions

### 1. Clone & prepare repo
```bash
git clone https://github.com/YOUR_USERNAME/nse-rsi-screener.git
cd nse-rsi-screener

# Place your Excel files in data/
cp /path/to/nse500list.xlsx   data/
cp /path/to/Micro250list.xlsx  data/
```

### 2. Excel file format required

Your `nse500list.xlsx` and `Micro250list.xlsx` must contain a column named **`Symbol`**:

| Symbol | (other columns are ignored) |
|--------|-----------------------------|
| RELIANCE | ... |
| TCS | ... |
| INFY | ... |

### 3. Push to GitHub
```bash
git add .
git commit -m "Initial setup"
git push origin main
```

### 4. Enable GitHub Actions
- Go to your repo → **Actions** tab
- If prompted, click **"I understand my workflows, go ahead and enable them"**

### 5. Test manually
- Go to **Actions** → `NSE RSI Screener — Weekly Saturday Run` → **Run workflow**

### 6. Download the report
After the workflow completes:
- Go to **Actions** → click the latest run
- Scroll to **Artifacts** section at the bottom
- Download `RSI-Screener-Report-XXXXX.zip`

The report is also **committed directly** to the `output/` folder in your repo.

---

## Schedule

| Setting | Value |
|---------|-------|
| Cron | `35 3 * * 6` |
| UTC time | Saturday 03:35 AM |
| IST time | Saturday 09:05 AM |

> GitHub Actions may run up to 10–15 minutes late due to runner availability. This is normal behaviour.

---

## Local Testing

```bash
# Install dependencies
pip install -r requirements.txt

# Run script
python rsi_screener.py

# Check output
ls output/
```

---

## Colour Legend in Report

| Colour | Meaning |
|--------|---------|
| 🟢 Green cell | RSI above threshold (strong signal) |
| 🟡 Yellow cell | RSI at or near boundary |
| White | Normal |
