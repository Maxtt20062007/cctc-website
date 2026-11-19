# CCTC Website + Automatic Lead System

This folder has everything for the CCTC website and its automatic customer system.

| File | What it is |
|---|---|
| `index.html` | The whole website: services, booking form, pricing calculator, FAQ, service-area checker |
| `backend/Code.gs` | The automatic "back office" that runs free in your Google account |
| `backend/appsscript.json` | Settings file for the back office |
| `.nojekyll` | Tells GitHub to serve the site as-is |

## How it works

```
Customer fills the form on your website
        │
        ▼
Google Apps Script (free, runs 24/7 in your Google account)
        ├── saves the customer to your Google Sheet "Leads"
        ├── emails YOU instantly (name, phone, job, call-by day)
        └── emails the CUSTOMER a confirmation
        │
Every work-day at 7 am → you get a "CCTC today" email:
   who to call, who to call back, surveys, quote follow-ups,
   review requests, plus your daily checklist
```

It costs $0. A normal Gmail account can send about 100 emails a day, which is plenty to start with.

---

## Step 1: Put the website on GitHub (about 5 min)

1. Sign in at **github.com**. Click **+** (top right) → **New repository**.
2. Name it `cctc-website`, choose **Public**, then click **Create repository**.
3. Click **uploading an existing file**. Drag in everything from this folder, including the `backend` folder. Then click **Commit changes**.
4. Open **Settings** → **Pages**. Under *Build and deployment*, set Source to **Deploy from a branch**, Branch to **main**, and folder to **/ (root)**. Click **Save**.
5. Wait 1–2 minutes. Your site will be live at `https://YOUR-USERNAME.github.io/cctc-website/`.

> The `.nojekyll` file starts with a dot, so your computer may hide it. If it doesn't upload, the site still works.

## Step 2: Create the Google Sheet + back office (about 10 min)

1. Go to **sheets.google.com** and create a blank sheet. Name it **CCTC Leads**.
2. In the menu, click **Extensions** → **Apps Script**.
3. Delete the sample code. Paste in **all** of `backend/Code.gs`.
4. At the top of the code, change `OWNER_EMAIL` to your email. Add your phone in `BUSINESS_PHONE` if you have a business number.
5. Click **Project Settings** (⚙️ on the left) and tick **Show "appsscript.json" manifest file in editor**. Go back to the editor, open `appsscript.json`, and replace its contents with `backend/appsscript.json`.
6. Click **Save** (💾). In the function dropdown at the top, choose **setup**, then click **Run**.
7. Google will ask for permission. Click **Review permissions**, choose your account, then **Advanced** → **Go to (project) (unsafe)** → **Allow**.
   *It says "unsafe" only because you wrote the script yourself and Google hasn't reviewed it. It runs only in your account.*
8. Go back to the sheet and reload the page. You'll see a **Leads** tab and a **CCTC** menu.

## Step 3: Turn the back office on for the website

1. In Apps Script, click **Deploy** → **New deployment**.
2. Click ⚙️ next to "Select type" and choose **Web app**.
3. Set **Execute as** to **Me**, and **Who has access** to **Anyone**.
4. Click **Deploy**, then copy the **Web app URL** (it ends in `/exec`).

## Step 4: Connect the form

1. On GitHub, open `index.html` and click the ✏️ pencil to edit.
2. Search for `CONFIG` (near the bottom) and fill in:
   ```js
   formEndpoint: "https://script.google.com/macros/s/....../exec",
   phone: "+13065550123",        // optional
   phoneDisplay: "306-555-0123", // optional
   email: "you@example.com",     // optional
   ```
3. Click **Commit changes**. The site updates in about a minute.

## Step 5: Test it

1. Open your site, fill in the form with your own email, and send it.
2. Check that a new row appears in the **Leads** sheet, that you got a "🔔 New lead" email, and that you got the customer confirmation.
3. In the sheet, click **CCTC** → **Send me today's to-do email now** to see the daily email.
4. Delete the test row when you're done.

---

## Your daily routine

1. **7 am:** read the "CCTC today" email.
2. **Call each person.** After each call, change the **Status** cell:

| Status | What happens automatically |
|---|---|
| New | On today's list until you call |
| No answer | Shows up again next work day |
| Survey booked | Type the visit date in **Next follow-up**. It shows up that morning |
| Quoted | Follow-up reminder 3 work days later |
| Won | Counted as a job in progress |
| Done | Reminder to ask for a review 2 work days later |
| Lost | Removed from lists |

3. **Follow-up emails:** click a customer's row, then **CCTC** → *quote follow-up* or *thank-you + review request*. You see the email first and it only sends when you click **Yes**.

**Tip:** Install the Gmail and Google Sheets apps on your phone. New-lead emails then reach you instantly, and the call links open your dialer.

## Changing things later

- **Work days, cutoff time, follow-up delays, daily checklist:** edit `SETTINGS` at the top of `Code.gs`, then click Save. If you changed `DIGEST_HOUR`, run `setup` again.
- **If you change `Code.gs` after deploying:** go to **Deploy** → **Manage deployments** → ✏️ → Version: **New version** → **Deploy**. The URL stays the same.
- **Service area:** in `index.html`, edit `servedPrefixes`. Moose Jaw uses S6H, S6J and S6K.
- **Website text:** edit `index.html` directly on GitHub.

## Moving to a real domain (e.g. cctc.ca)

- **Easiest:** buy the domain, then in GitHub **Settings** → **Pages** → **Custom domain**, type it in and follow GitHub's DNS instructions. Everything keeps working.
- **Other host** (Netlify, Hostinger, etc.): upload `index.html`. The form still sends to the same Google Sheet, so nothing else changes.

## Before you go fully public

- Add **real** customer reviews once you have them. Don't use made-up ones.
- If you put "licensed" or "insured" on the site, make sure it's true for CCTC.
- Electrical and plumbing work in Saskatchewan has licensing rules. The site intentionally doesn't offer them.
