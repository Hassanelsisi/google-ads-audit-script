# Google Ads Account Audit Script V1.0 🚀📊🔍

**Author:** [Hassan El-Sisi](https://www.linkedin.com/in/hassan-elsisi)

## 👋 Welcome\! What's This All About?

Hey there, fellow Google Ads enthusiast\! Ever wished you had a super-powered assistant to dig through your Google Ads account and find all those juicy insights, potential problems, and hidden opportunities? Well, you're in the right place\!

This script is your automated buddy for conducting a deep-dive audit of your Google Ads account. It crunches the numbers, analyzes performance, and lays it all out for you in a neat Google Sheet, complete with color-coding and actionable recommendations.

**Who is this for?**

  * **Advertisers** managing their own accounts.

  * **Digital Marketing Agencies** looking to streamline client audits.

  * Anyone who wants a clearer, data-backed understanding of what's really going on in their Google Ads.

**Why is it awesome?**

  * ✅ **Saves You Tons of Time:** No more manually pulling dozens of reports.

  * 🎯 **Pinpoints Key Issues & Opportunities:** Quickly see what's working, what's not, and where to focus your efforts.

  * 📈 **Data-Driven Insights:** Get recommendations based on common best practices and *your* defined performance targets.

  * 📄 **Organized Reporting:** All data is neatly presented in a multi-tab Google Sheet.

## ✨ Features at a Glance ✨

This script generates a comprehensive set of reports, each designed to give you specific insights:

  * 📄 **README**: (This file\!) Your guide to the script.

  * ⚙️ **CONFIG (Informational - Hidden by Default)**: Shows the hardcoded settings the script is using. You'll edit these settings directly in the script code (see "Getting Started").

  * ❗ **ERRORS**: If anything goes wrong during the script run, details will be logged here. (Hopefully, it stays empty\!)

  * 🏆 **Executive Summary**: A high-level snapshot of key findings, opportunities, and warnings. Your starting point\! (Colored Pastel Blue)

  * 🌍 **Master Overview**: A bird's-eye view of performance across all your advertising channels (Search, Display, Shopping, etc.). (Colored Light Teal)

  * 💡 **Master Recommendations**: Quick, actionable advice based on the Master Overview. (Colored Light Teal)

  * 📅 **Account Heatmap**: See when your account performs best with an hour-of-day and day-of-week breakdown. (Colored Light Green)

  * 📢 **Channel-Specific Reports**: Detailed performance and recommendations for each of your active channels:

      * Channel – Search

      * Channel – Display

      * Channel – YouTube

      * Channel – Shopping

      * Channel – PMax

      * Channel – App Installs

      * Channel – App Engagement
        *(These tabs are colored Light Green)*

  * 🛍️ **Product Performance Reports**:

      * **Shopping Product Performance**: Deep dive into how individual products are doing in your Shopping campaigns. (Colored Light Purple)

      * **PMax Product Titles Performance**: See product title performance within Performance Max campaigns. (Colored Light Purple)

      * **PMax Asset Groups**: Analyze the performance of your Performance Max asset groups. (Colored Light Purple)

  * 🔍 **Deep Dive Audit Reports**: Get granular with these specific analyses (most are colored Pale Yellow):

      * **Search Terms Report**: Uncover what people are *actually* searching for to trigger your ads. Find new keywords and negative keyword opportunities\!

      * **Keyword Deep Dive (All Search)**: A complete breakdown of all your Search keywords.

      * **Quality Score (Search)**: Understand your keyword Quality Scores and their components (Expected CTR, Ad Relevance, Landing Page Experience).

      * **Search IS Analysis**: See your Search Impression Share, and how much you're losing to budget or rank.

      * **Ads Performance (Search)**: Analyze the performance of your individual Search ads.

      * **Audience Insights**: See how different audiences are performing.

      * **Device Performance**: Compare performance across Desktops, Mobiles, and Tablets.

      * **Budget Pacing**: Check if your campaigns are on track with their budgets.

  * 📝 **KW - {Campaign Name}**: Individual, detailed keyword reports for each active Search campaign. These appear towards the end of your spreadsheet.

## ⚙️ How It Works (The "Magic" Explained Simply)

Think of this script as a very smart, very fast data analyst. Here's the basic idea:

1.  **Connects to Google Ads:** It uses Google's own tools (specifically, Google Ads Query Language or GAQL) to securely access your account data (don't worry, it only *reads* data, it doesn't change anything in your account\!).

2.  **Pulls the Data:** It runs a series of pre-defined queries to gather performance metrics for campaigns, keywords, search terms, etc., based on the date range you set.

3.  **Crunching Numbers:** It then processes this data, calculating things like CTR, CPA, ROAS, and comparing them against the benchmarks you've set in the script.

4.  **Creates Your Report:** Finally, it neatly organizes all this information and the recommendations into a new Google Spreadsheet, creating different tabs for each report section and even color-coding them for easier navigation.

## 🚀 Getting Started: Your Step-by-Step Guide

Ready to get your audit on? Follow these steps carefully. It might look like a lot, but if you take it one step at a time, you'll be up and running in no time\!

### Prerequisites:

  * ✅ Access to a Google Ads account.

  * 📄 A Google Sheet (either create a new one or you can use an existing one, but the script will add many new tabs).

### Step 1: Copy the Script Code

  * Grab all the code from the `GoogleAdsAuditScriptV1.js` file (or whatever you've named the file containing this script). Select everything (Ctrl+A or Cmd+A) and copy it (Ctrl+C or Cmd+C).

### Step 2: Open Google Ads Scripts

  * Log in to your Google Ads account.

  * In the top menu, click on **Tools & Settings** (it looks like a wrench 🔧).

  * Under "BULK ACTIONS," click on **Scripts**.

### Step 3: Paste and Name Your Script

  * Click the big blue **+** button to create a new script.

  * You'll see a script editor. Delete any existing code in there (usually a simple `function main() {}`).

  * Paste the entire code you copied in Step 1 into the script editor (Ctrl+V or Cmd+V).

  * At the top, where it says "Untitled script," give your script a memorable name, like "My Awesome Ads Audit" or "Account Audit V1".

### Step 4: Configure the Script (Super Important\! 🛠️)

This is where you tell the script where to put the report and what benchmarks to use.
*Scroll to the very top of the script code you just pasted.* You're looking for **SECTION 1.0 – SCRIPT DEFAULTS & CORE BENCHMARKS**.

1.  **`SPREADSHEET_ID`**: This tells the script which Google Sheet to use.

      * Open the Google Sheet you want the report to be generated in (or create a new one).

      * Look at the URL (the web address) in your browser. It will look something like this:
        `https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit#gid=0`

      * You need to copy the long string of letters and numbers that's between `/d/` and `/edit`. That's your Spreadsheet ID\!

      * In the script, find this line:
        `var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';`

      * Replace `'YOUR_SPREADSHEET_ID_HERE'` with *your* actual Spreadsheet ID, keeping the single quotes.

2.  **`DEFAULT_DATE_FROM` and `DEFAULT_DATE_TO`**: This sets the date range for your audit.

      * Find these lines:
        `var DEFAULT_DATE_FROM = '2025-01-01';`
        `var DEFAULT_DATE_TO   = '2025-04-30';`

      * Change the dates (inside the single quotes) to the start and end dates you want for your report. Use the `YYYY-MM-DD` format.

3.  **PERFORMANCE THRESHOLDS & BRAND PATTERNS**: This is where you customize the audit's logic\!

      * Still in `SECTION 1.0`, you'll see variables like `DEFAULT_CLICK_THRESHOLD`, `DEFAULT_CTR_LOW`, `DEFAULT_TARGET_ROAS`, etc.

      * **Read the comments next to each one carefully.** These explain what each setting does.

      * Adjust these numbers to match what *you* consider good or bad performance for *your* account or your client's account. For example, if your target ROAS is 400%, you'd set `DEFAULT_TARGET_ROAS = 4.0;`.

      * **`DEFAULT_BRAND_PATTERNS`**: This is an array (a list) of your brand names or keywords. Edit this list to include terms that identify your brand searches. This helps the script separate brand vs. non-brand performance.

          * Example: `var DEFAULT_BRAND_PATTERNS = ['my awesome company', 'myproduct x', 'mac'];`

### Step 5: Authorize the Script (Give it Permission)

Scripts need your permission to access your Ads data and write to your Google Sheet.

  * Click the **Save** icon (💾) at the top of the script editor.

  * Now, click the **Run** button (▶️).

  * A pop-up titled "Authorization required" will appear. Click **Review permissions**.

  * Choose the Google account that has access to the Google Ads account you want to audit.

  * You might see a screen saying "Google hasn’t verified this app." This is normal for custom scripts you've written yourself.

      * Click on **Advanced** (it might be a small link).

      * Then click on **"Go to \[Your Script Name\] (unsafe)"**.

  * Review the permissions the script needs (it will ask for access to your Google Ads data and your Google Spreadsheets). Click **Allow**.

### Step 6: Run the Script\! 🏃💨

  * After authorizing, you might be taken back to the script editor. Click the **Run** button (▶️) again.

  * **Be Patient\!** This script is doing a lot of work. It might take several minutes to run, especially for larger accounts or long date ranges (Google Ads Scripts have a 30-minute time limit).

  * You can see the script's progress in the **Logs** section below the code editor. It will print messages like "Running: Master Overview", "Finished: Master Overview", etc.

### Step 7: Check Your Google Sheet\! 🎉

  * Once the logs say something like "Google Ads Audit Script V1.0 ... finished successfully...", open the Google Sheet you specified in `SPREADSHEET_ID`.

  * You should see a whole bunch of new tabs filled with your audit data, color-coded and ready for review\!

## 📊 Understanding the Report Tabs

The script creates many tabs. Here's a quick guide to the color-coding to help you navigate:

  * ⚪ **Default/White (README, CONFIG, KW-Campaigns):** General info or very granular data.

  * ❤️ **Red (ERRORS):** If this tab appears and has content, check it first for any script execution problems\!

  * 💙 **Pastel Blue (Executive Summary):** Your high-level overview. Start here\!

  * 💚 **Light Teal (Master Overview, Master Recs):** Account-wide channel performance.

  * 🍏 **Light Green (Account Heatmap, Channel Tabs):** Performance patterns and channel-specific details.

  * 💜 **Light Purple (Product Tabs):** For Shopping and PMax product insights.

  * 💛 **Pale Yellow (Deep Dive Reports):** Granular reports on Search Terms, Keywords, Quality Score, etc.

## ⚠️ Important Notes & Disclaimer

  * **Execution Time:** This script is powerful and pulls a lot of data. For very large accounts or very long date ranges, it might hit Google Ads Scripts' 30-minute execution limit. If this happens, try running it with a shorter date range.

  * **Read-Only:** This script **DOES NOT MAKE ANY CHANGES** to your Google Ads account. It's purely for reporting and analysis. You are always in control.

  * **Data Accuracy:** The script relies on data from the Google Ads API. Sometimes, metrics might differ slightly from what you see in the Google Ads UI due to reporting lags or different ways data is attributed.

  * **Review Recommendations Critically:** The recommendations provided are based on common best practices and the thresholds *you* set in the script. **Always use your own expertise and judgment** before making any changes to your campaigns. This script is a tool to help you, not a replacement for a skilled account manager.

  * **Permissions:** Ensure you've granted the necessary permissions during the authorization step for the script to access your Ads data and Google Sheets.

## 🤔 Troubleshooting Common Issues

  * **"Script timed out"**: The script took longer than 30 minutes. Try reducing the `DEFAULT_DATE_FROM` and `DEFAULT_DATE_TO` range in Section 1.0 of the script.

  * **"\#ERROR\!" or "\#VALUE\!" in cells**: This usually means the script had trouble fetching or calculating a specific piece of data.

      * Check the "ERRORS" tab in your spreadsheet for any messages.

      * Ensure your `SPREADSHEET_ID` is correct.

      * Double-check that the hardcoded benchmarks in Section 1.0 are valid numbers or patterns.

  * **"No campaigns/data found" on a tab**:

      * The script might have correctly found no data for that specific report and date range.

      * For channel-specific tabs, if no campaigns of that type ran in the period, the tab might be automatically deleted by the script.

      * Double-check your `DEFAULT_DATE_FROM` and `DEFAULT_DATE_TO`.

Happy Auditing\! May your insights be plentiful and your optimizations impactful\!