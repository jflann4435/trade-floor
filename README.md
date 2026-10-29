# Trade Floor: class stock site

A small website for the class: trade buttons, a stock sheet for every company built around the course's terms, charts, and a replay that runs about 22 hours behind so a morning class can watch the previous day's market.

A script runs every weekday evening on GitHub's free servers. It pulls the day's prices from Alpaca and company numbers from SEC filings, saves them in the `data` folder, and the site reads them the next morning.

## What's in this folder

| Item | What it is |
|---|---|
| `index.html` | The website |
| `stocks/Trade_Stock_List.xlsx` | The stock list. Edit this to change stocks, trades, hooks, or 1y target estimates |
| `scripts/update_data.py` | The nightly update script |
| `.github/workflows/nightly.yml` | Tells GitHub when to run the script |
| `data/` | Sample data so the site works before the first real run. The script replaces it |
| `requirements.txt`, `robots.txt` | Support files (`robots.txt` asks search engines not to list the site) |

---

## One-time setup (about 30 minutes)

### 1. Get free Alpaca data keys

1. Sign up at **alpaca.markets**. You only need a **paper trading** account. Don't fund anything.
2. In the paper-trading dashboard, find **API Keys** and click **Generate New Keys**.
3. Copy the **Key ID** and **Secret Key** somewhere safe. The secret is shown only once.

### 2. Create the GitHub repository

1. Sign in at **github.com** (create a free account if needed).
2. Click **New repository**. Name it `trade-floor`, choose **Public**, and click **Create repository**. Free GitHub Pages hosting needs a public repository.
3. On the new repository page, click **uploading an existing file**.
4. Unzip this folder and drag in everything **except the `.github` folder**: `index.html`, `README.md`, `requirements.txt`, `robots.txt`, and the `data`, `scripts`, and `stocks` folders. Click **Commit changes**.

   Your computer may hide `.github` because its name starts with a dot, so add it by hand:

5. Click **Add file → Create new file**. Type the name `.github/workflows/nightly.yml` (the slashes create the folders). Open `nightly.yml` from the zip in a text editor, copy everything, paste it in, and click **Commit changes**.

### 3. Add your keys as secrets

In the repository: **Settings → Secrets and variables → Actions → New repository secret**. Add three:

| Name | Value |
|---|---|
| `ALPACA_KEY_ID` | Your Alpaca Key ID |
| `ALPACA_SECRET_KEY` | Your Alpaca Secret Key |
| `SEC_CONTACT` | Your name and school email, e.g. `Jane Smith jsmith@school.org`. The SEC requires a contact on every request |

Secrets stay hidden, even in a public repository.

### 4. Let the script save data

**Settings → Actions → General → Workflow permissions**: choose **Read and write permissions** and click **Save**.

### 5. Turn on the website

**Settings → Pages**. Under **Build and deployment**, set Source to **Deploy from a branch**, Branch to **main** and **/ (root)**, then **Save**. After a minute or two the page shows your address, like `https://yourname.github.io/trade-floor/`.

### 6. Run the first real update

**Actions** tab → **Nightly market data** → **Run workflow** → **Run workflow**.

It takes about 20 minutes (most of that is collecting bid/ask quotes). A green check means it worked. Click the run to see a short report, including any stocks it had trouble with. A minute or two later the site's banner changes from **Sample data** to **Real data**.

From then on it runs by itself at 5:40 PM Eastern in winter and 6:40 PM in summer, Monday through Friday.

### 7. Check it against Yahoo Finance

Open two or three stocks and compare with Yahoo:

- **Previous Close, Day's Range, 52-Week Range** should match closely.
- **Bid and Ask sizes**: if ours are 100 times smaller than Yahoo's, open `scripts/update_data.py`, change `QUOTE_SIZE_MULTIPLIER = 1` to `100`, and commit.
- **EPS, Market Cap, Book Value**: these come from SEC filings, so they can differ a little from Yahoo's (different share counts or adjusted earnings). That gap is a good class discussion.

---

## Everyday use

**Changing stocks or targets.** Edit `Trade_Stock_List.xlsx` on your computer, then in GitHub open the `stocks` folder → **Add file → Upload files**, drop in the file (same name), and commit. The next nightly run uses it. To see changes right away, run the workflow by hand from the Actions tab.

- Only rows with Status **Include** appear on the site.
- **1y Target Est** isn't in any free data feed. Copy it from Yahoo Finance into column J (and the date into K) whenever you like. Blank cells show "Not entered."

**The replay.** Replay settings on the site control the delay (22 hours by default), switch to a closing snapshot, and let you scrub through the day by hand. Settings are remembered per browser.

**Holidays and weekends.** On days the market is closed there's no new data, so the site keeps replaying the last trading day.

**If a run fails** (red X on the Actions tab), the site keeps the previous night's data. Open the run to see why:

| Message | Fix |
|---|---|
| `HTTP 401` or `HTTP 403` from Alpaca | The keys are wrong or were regenerated. Update the two Alpaca secrets |
| `SEC_CONTACT must be set` | Add the `SEC_CONTACT` secret |
| `Only N of M stocks got prices` | Usually a temporary Alpaca problem. Run it again later |
| `isn't on the Categories tab` (in the report) | A stock's Category doesn't match a Button Label on the Categories tab |

GitHub pauses scheduled jobs in repositories that go 60 days without activity. If that happens, the Actions tab shows a button to turn it back on.

---

## Things to know

- **Who can see it.** Anyone with the link can open the site, but search engines are asked not to list it. Share the link through your class page rather than publicly.
- **Data terms.** Read Alpaca's market-data terms before sharing the site widely. By default the script uses Alpaca's **SIP** feed (all U.S. exchanges), which matches Yahoo best. To use only the free IEX exchange feed, go to **Settings → Secrets and variables → Actions → Variables** and add `ALPACA_FEED` with the value `iex`. Prices stay close, but volume will look far smaller than Yahoo's.
- **School filters.** Open the site on a student Chromebook early to make sure `github.io` isn't blocked.
- **Charts** use TradingView Lightweight Charts, which requires the credit line at the bottom of the page. Keep it.
- **Testing without keys.** On a computer with Python, `pip install -r requirements.txt` then `python scripts/update_data.py --demo` writes sample data to `data/`.
