# NSE RSI Screener 📊

Automated weekly RSI screener for NSE stocks.
Runs every **Saturday at 9:05 AM IST** via GitHub Actions.
Sends results via **Telegram message** + **Email with Excel attachment**.

---

## Scan Conditions

| Timeframe   | Condition              |
|-------------|------------------------|
| Monthly RSI | > 60                   |
| Weekly RSI  | > 60                   |
| Daily RSI   | > 40  **and**  < 45    |

---

## Alerts

| Alert    | What you receive                                         |
|----------|----------------------------------------------------------|
| Telegram | Formatted table of all matched stocks with RSI values    |
| Email    | HTML email with table + Excel report attached            |

---

## GitHub Secrets Required

Go to your repo → **Settings → Secrets and variables → Actions → New repository secret**

| Secret Name          | Value                                     |
|----------------------|-------------------------------------------|
| `TELEGRAM_BOT_TOKEN` | Your bot token from @BotFather            |
| `TELEGRAM_CHAT_ID`   | Your chat/group ID (use @userinfobot)     |
| `EMAIL_SENDER`       | Your Gmail address                        |
| `EMAIL_PASSWORD`     | Gmail App Password (16-char)              |
| `EMAIL_RECEIVER`     | Recipient email address                   |

### How to get Telegram credentials
1. Open Telegram → search **@BotFather** → send `/newbot`
2. Follow prompts → copy the **bot token**
3. Search **@userinfobot** → send any message → copy your **chat id**
4. Send a message to your bot first (so it can message you back)

### How to get Gmail App Password
1. Go to myaccount.google.com → Security
2. Enable **2-Step Verification** (required)
3. Search "App passwords" → create one for "Mail"
4. Copy the **16-character password** (no spaces) → use as `EMAIL_PASSWORD`

---

## Folder Structure

```
nse-rsi-screener/
├── .github/workflows/rsi_screener.yml
├── data/
│   ├── nse500list.xlsx        ← Must have a "Symbol" column
│   └── Micro250list.xlsx      ← Must have a "Symbol" column
├── output/                    ← Reports saved here automatically
├── rsi_screener.py
├── requirements.txt
└── README.md
```

---

## First-Time Setup

```bash
git clone https://github.com/nikunj2021/nse-rsi-screener.git
cd nse-rsi-screener

# Replace sample files with your actual stock lists
cp /your/path/nse500list.xlsx   data/
cp /your/path/Micro250list.xlsx data/

git add .
git commit -m "Add stock data files"
git push origin main
```

Then add all 5 secrets in GitHub → Settings → Secrets.

---

## Test Manually

GitHub → Actions tab → `NSE RSI Screener` → **Run workflow**

After it finishes:
- Check your **Telegram** for the alert message
- Check your **Email** for the HTML report + Excel attachment
- Download Excel from the **Artifacts** section of the workflow run

---

## Schedule

| Setting  | Value             |
|----------|-------------------|
| Cron     | `35 3 * * 6`      |
| UTC      | Saturday 03:35 AM |
| IST      | Saturday 09:05 AM |

> GitHub Actions may run up to 10–15 min late on free-tier — this is normal.
