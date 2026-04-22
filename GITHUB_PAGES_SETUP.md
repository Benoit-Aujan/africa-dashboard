# Publishing the Dashboard to GitHub Pages

This is a one-time setup that takes about 10 minutes.
Once done, the dashboard will have a permanent public URL you can share with anyone.

---

## Step 1 — Create a GitHub account

1. Go to https://github.com/signup
2. Enter your email, create a password, choose a username (e.g. `benoit-agh` or similar)
3. Verify your email

---

## Step 2 — Create a new repository

1. Once logged in, click the **+** button (top right) → **New repository**
2. Name it: `africa-dashboard` (or any name you like)
3. Set it to **Public** (required for free GitHub Pages)
4. Leave everything else unchecked — click **Create repository**

---

## Step 3 — Install Git (if not already installed)

Open PowerShell and run:
```
git --version
```
If it's not installed, download from: https://git-scm.com/download/win

---

## Step 4 — Push the dashboard files

Open PowerShell and run these commands one by one, replacing `YOUR-USERNAME` with your GitHub username:

```powershell
cd "C:\Claude Projects\projects\africa-dashboard"

git init
git add .
git commit -m "Initial dashboard release"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/africa-dashboard.git
git push -u origin main
```

When prompted, enter your GitHub username and password (or a personal access token if prompted).

---

## Step 5 — Enable GitHub Pages

1. Go to your repository on GitHub: `https://github.com/YOUR-USERNAME/africa-dashboard`
2. Click **Settings** (top menu of the repo)
3. In the left sidebar, click **Pages**
4. Under **Source**, select **Deploy from a branch**
5. Branch: select **main**, folder: **/ (root)**
6. Click **Save**

After about 1-2 minutes, your dashboard will be live at:
```
https://YOUR-USERNAME.github.io/africa-dashboard/
```

---

## Updating the dashboard with new data

Every time you want to update the data (after running the email automation script), run:

```powershell
cd "C:\Claude Projects\projects\africa-dashboard"
git add data/data.json
git commit -m "Update data"
git push
```

GitHub Pages will automatically republish within 1-2 minutes.

---

## Daily workflow (once everything is set up)

1. Open Outlook (Classic)
2. Run the update script: `python scripts/update_from_email.py`
3. If discrepancies are found, review `data/discrepancy_draft_email.txt`
4. Push the updated data.json to GitHub: `git add data/data.json && git commit -m "Update" && git push`
5. Share the link: `https://YOUR-USERNAME.github.io/africa-dashboard/`
