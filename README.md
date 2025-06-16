# Forgd - Growth Capital Prototype v0.3

## 🧠 Overview

**Forgd - Growth Capital v0.3** is a modular, script-enhanced Google Sheets prototype designed to help startup teams prioritize and identify top-fit investors based on real data. Built on structured scoring logic, it ingests raw fundraising and investor data (via RootData or CSV), dynamically ranks investors, and produces narrative-style recommendations tailored to your raise.

## 📌 Key Features

- **Modular Scoring Engine**  
  Each matching parameter (e.g. vertical, region, round, raise amount) has its own App Script module. This makes the architecture extremely plug-and-play.

- **Weighted Aggregation**  
  Users can apply weights to both universal and context-dependent scoring dimensions to fully control how final investor rankings are generated.

- **Dynamic Data Sanitization**  
  RootData's CSV and API outputs often contain malformed fields. This system auto-sanitizes fields like "investor roles" or "tags" into structured, analysable form.

- **Auto-Recalculation via Triggers**  
  Scripts are linked to Google Sheet triggers so scores are refreshed automatically — no manual intervention required.

- **Narrative Generation**  
  Narrative summaries are dynamically generated for investor matches, helping founders build compelling outreach copy.

## 🧱 System Architecture

**RootData Inputs**
- FetchProjects
- FetchInvestorInfo
- FetchFundraising

**Sanitization Layers**
- ProjectsView
- InvestorView
- RoundView

**Scoring Modules**
- ScoringMatchIndexInvestStats
- ScoringMatchIndexTags
- ScoringMatchIndexEcosystem
- ScoringMatchIndexTotalFunding
- ScoringMatchIndexArea
- ScoringMatchIndexValuation
- ScoringMatchIndexRaise
- ScoringMatchIndexRound

**Scoring Layers**
- ScoreInvestorsUV.js (Universal)
- ScoreInvestorsNUV.js (Non-Universal)
- ScoreInvestorsWeighted.js

**Output Modules**
- GenerateNarrativeC22.js
- Growth Capital Key Insights (Sheet tab)

## 🔍 How It Works

The system operates on two main scoring layers: **Universal** and **Non-Universal** scoring dimensions.

**Universal scoring** evaluates investors based on activity patterns that apply regardless of your specific project:
- Number of investments made in the past 12 months
- Total historical investments
- Lead investments
- Recency of activity

**Non-universal scoring** dynamically adjusts based on your project's characteristics:
- Vertical and subvertical alignment
- Raise amount and funding round stage
- Ecosystem and geographic region
- Investment size preferences

Parameters like verticals, regions, raise ranges, and ecosystems are pulled from real data and normalized automatically. The system handles data inconsistencies (such as multiple investors packed in a single cell) through custom sanitization logic.

Users can adjust weighted coefficients through a simple matrix interface, allowing full control over how different factors influence the final investor rankings. The final output includes a ranked investor list, personalized narrative summaries for outreach, and visual insights within the sheet interface.

## 🧰 Setup Instructions

### 1. Prerequisites

- Google account with Apps Script API enabled
- Node.js + npm installed via Homebrew:
  ```bash
  brew install node
  ```
- Install clasp globally:
  ```bash
  npm install -g @google/clasp
  ```

### 2. Get RootData API Key

1. Visit [RootData API](https://www.rootdata.com/Api)
2. Choose a plan (Basic free tier includes 1000 monthly API credits)
3. Contact RootData to get your API key
4. Set up the API key in Google Apps Script:
   - Open your Apps Script project
   - Go to Project Settings (gear icon)
   - Click "Script Properties" tab
   - Add new property:
     - Property: `ROOTDATA_API_KEY`
     - Value: Your RootData API key

### 3. Clone and Link Project

```bash
git clone https://github.com/imajinl/growth-capital-v0.3.git
cd growth-capital-v0.3
clasp login
clasp clone 1hvRY2EAINznP4N4cw6NuT0X7zBCrp03EODD4cVnnYDAk6xI3Bjt60ece
```

### 4. Push or Pull Changes

To sync code with the cloud script editor:

```bash
clasp pull    # Pull down changes
clasp push    # Push local changes
```

## 📁 File Guide

| Filename | Purpose |
|----------|---------|
| `UnifiedTriggers.js` | Triggers recomputation of scoring modules |
| `GenerateNarrativeC22.js` | Concatenates investor fit narratives |
| `ScoreInvestorsUV.js` | Universal scoring layer |
| `ScoreInvestorsNUV.js` | Non-universal scoring layer |
| `ScoreInvestorsWeighted.js` | Final weighted score aggregator |
| `ScoringMatchIndex*.js` | Dedicated modules for each scoring parameter |
| `appsscript.json` | Script project metadata |

## 💡 Example Use Cases

- Founders seeking investor alignment based on raise size, stage, and ecosystem
- Analysts looking to match VCs to startup cohorts with customized weightings
- Automating outreach prep with investor bios and fit narratives

## 🛠️ Tech Stack

- **Google Sheets** — UI and data layer
- **Google Apps Script** — Backend logic and automation
- **RootData** — Investor and fundraising data source

---

Questions? DM [@imajinl](https://t.me/imajinl) on Telegram.
