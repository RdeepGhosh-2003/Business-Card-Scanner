# 🪪 Business Card Scanner — Standalone App

> **NEW (Aug 2026)**: Completely rebuilt as a standalone browser app. No Google Forms, no Sheet setup, no popup issues. Just double-click and scan.

---

## ⚡ Quick Start (3 steps)

### Step 1 — Get a free Gemini API Key

Go to **[aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)** → Create API key → Copy it.

> Free tier: 1,500 requests/day — more than enough for a full day of card scanning.

---

### Step 2 — Launch the app

**Double-click `start.bat`**

This starts a local server and opens the app in your browser automatically.

> **Requires Python** (pre-installed on most Windows machines). If Python is not found, the script opens `index.html` in Microsoft Edge as a fallback.

---

### Step 3 — Enter your API key

Paste your Gemini API key in the yellow bar at the top → click **Save**. It's stored in your browser and persists across sessions.

---

## 📱 Using the App

### 📹 Camera Mode (Webcam or Phone)

1. Click **▶️ Start Camera** → allow camera permission when prompted
2. Align the business card within the dashed guide frame
3. Click **📸 Capture** → the photo appears as a preview
4. Click **🔍 Scan Card** — Gemini extracts all details
5. Review and edit the filled-in fields
6. Click **💾 Save Card**

**Scan from your phone's camera:**
- Keep `start.bat` running
- Connect your phone to the same WiFi
- Open `http://<your-laptop-IP>:8080` on your phone
- Tap **Start Camera** → use the rear camera to scan cards

---

### 📁 Upload Mode

1. Switch to the **📁 Upload File** tab
2. Drag & drop a card image or click to browse
3. Click **🔍 Scan Card** → review → save

---

### 📧 Send Follow-up Email

- After scanning: click **📧 Send Email** in the review panel
- From the table: click **📧 Email** on any saved card
- Your default email client (Outlook / Gmail) opens pre-filled with the contact's email and a greeting message. Just hit Send.

---

### ⬇️ Export at End of Day

Click **⬇️ Export CSV** (top-right or above the table) → downloads a `.csv` file named `BusinessCards_DD-MM-YYYY.csv`

The CSV includes: Timestamp, Name, Designation, Company, Phone, Email, Website, Address, LinkedIn, Other Info, Notes, Source.

---

## 📋 Feature Summary

| Feature | Details |
|---------|---------|
| Camera | Live webcam feed with card alignment guide |
| Phone camera | Open on phone via local WiFi |
| Upload | Drag & drop or browse — JPEG, PNG, WEBP, HEIC |
| AI extraction | Gemini 2.0 Flash — name, title, company, phone, email, address, website, LinkedIn |
| Edit | All fields editable before saving |
| Duplicate detection | Warns if same phone/email already scanned |
| Email | Opens Outlook/Gmail pre-filled |
| Storage | `localStorage` — persists through the day automatically |
| Export | CSV download with all scanned cards |
| Clear | "Clear Session" button (exports first!) |

---

## 🔑 API Key

- Stored in your browser's `localStorage` — never sent anywhere except Google's Gemini API
- You can change it anytime from the top bar
- Get it free: [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)

---

## 🛠️ Troubleshooting

| Problem | Fix |
|---------|-----|
| Camera not working | Make sure you clicked **▶️ Start Camera** and allowed permission in the browser popup |
| "Not set" API key status | Paste your key and click **Save** |
| Scan returns wrong data | Take a better-lit, closer photo of the card |
| `start.bat` says Python not found | Install Python from [python.org](https://python.org) or open `index.html` directly in Microsoft Edge |
| App doesn't open on phone | Make sure phone and laptop are on the same WiFi. Use the IP shown in the `start.bat` window |
| Want to keep Google Apps Script version | `Code.gs` and the GAS setup still works — see notes at the bottom |

---

## 📂 Files in This Folder

| File | Purpose |
|------|---------|
| `index.html` | The entire app — HTML + CSS + JS, self-contained |
| `start.bat` | Windows launcher — starts server + opens browser |
| `start.sh` | Mac/Linux launcher |
| `Code.gs` | (Legacy) Google Apps Script backend |
| `appsscript.json` | (Legacy) GAS manifest |

---

## 🔄 Switching Back to Google Apps Script Mode

The `Code.gs` + old `index.html` approach still works if you need form-based scanning or sheet integration.
To restore it, copy the previous `index.html` from your git history or the Apps Script editor.
