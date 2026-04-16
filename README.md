# Satguru SA Outreach Dashboard

Campaign performance dashboard for Satguru SA multi-channel sales outreach. Tracks email and LinkedIn engagement metrics across sequences, with real-time analytics powered by Google Sheets and Looker Studio.

**Live dashboard:** [lumif-ai.github.io/satguru-sa-dashboard](https://lumif-ai.github.io/satguru-sa-dashboard/)

## Overview

This dashboard monitors B2B outreach campaigns across multiple channels (email, LinkedIn connections, LinkedIn DMs, Aimfox automation). It transforms raw outreach data from Google Sheets into actionable metrics and visualizations.

### Key Features

- **Multi-channel tracking** -- email delivery/opens/clicks, LinkedIn connections, DMs, and Aimfox automation
- **Sequence analytics** -- step-level performance, progression tracking, reply attribution
- **Campaign metrics** -- delivery rates, open rates, reply rates, engagement funnels
- **A/B testing** -- version tracking per sequence for message comparison
- **Send time analysis** -- day/hour bucketing (UTC+2 SAST) for optimal timing
- **Week-over-week tracking** -- weekly bucketing for trend analysis
- **Dual views** -- client-facing (`index.html`) and internal team (`internal.html`) dashboards

## Project Structure

```
satguru-sa-dashboard/
  index.html          # Client-facing dashboard (password-protected)
  internal.html       # Internal team dashboard (password-protected)
  appscript/
    Code.gs           # Google Apps Script for data transformation
  .nojekyll           # Bypass Jekyll processing on GitHub Pages
```

## How It Works

### Data Pipeline

1. **Source:** Raw outreach step data lives in a Google Sheet (`conversation_dump_latest`)
2. **Transform:** `appscript/Code.gs` runs as a Google Apps Script, aggregating data into summary sheets:
   - **Lead Summary** -- per-prospect metrics across all channels
   - **Sequence Summary** -- per-sequence aggregate performance
   - **Campaign Summary** -- high-level campaign health
   - **Step Detail** -- granular per-step, per-lead activity
   - **Step Performance** -- step-level conversion metrics
3. **Visualize:** Looker Studio connects to the generated sheets, embedded in the HTML dashboards

### Derived Metrics

| Metric | Formula |
|--------|---------|
| Sequence Progress | (steps with activity / total steps) x 100 |
| Delivery Rate | (emails delivered / sent) x 100 |
| Open Rate | (opens / delivered) x 100 |
| Reply Rate | per-channel and overall calculations |

### Integrations

- **SendGrid** -- email delivery status, opens, clicks
- **Aimfox** -- LinkedIn automation connection tracking
- **Google Sheets** -- data source and transformation layer
- **Looker Studio** -- embedded visualizations

## Security

Both dashboards are protected with AES-256-GCM password encryption. Content is decrypted client-side on correct password entry.

## Tech Stack

- HTML5 / CSS3 / vanilla JavaScript (frontend)
- Google Apps Script (data transformation)
- Google Sheets (data layer)
- Looker Studio (visualization)
- GitHub Pages (hosting)
- Inter font (typography)

## Setup

### Apps Script

1. Open your Google Sheet with outreach data
2. Go to **Extensions > Apps Script**
3. Paste the contents of `appscript/Code.gs`
4. Run `generateAllSummaries()` to process data
5. Optionally set an hourly trigger for auto-refresh

### Dashboard

The dashboard is served via GitHub Pages from the `main` branch. Push changes to `main` to deploy.

---

Built by [lumif.ai](https://lumif.ai)
