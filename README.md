# 🤖 Automation with PyAutoGUI and Playwright

A collection of practical **Python automation projects** using **PyAutoGUI** and **Playwright** to automate repetitive desktop and browser-based tasks.

This repository demonstrates how Python can be used to automate real-world activities such as **data collection, Excel report generation, screenshots, browser interaction, and web application automation**.

---

## 📌 About This Repository

The purpose of this repository is to learn and demonstrate **practical automation using Python**.

It contains two automation projects:

```text
Automation-with-Pyautogui-and-Playwright
│
├── 📁 daily_report_bot
│   └── Automated Daily Report
│
└── 📁 Play_wright_bot
    └── Browser / WhatsApp Automation
```

---

# 🛠️ Technologies Used

| Technology       | Purpose                    |
| ---------------- | -------------------------- |
| 🐍 Python        | Automation programming     |
| 🖱️ PyAutoGUI    | Desktop and GUI automation |
| 🌐 Playwright    | Browser automation         |
| 📊 Excel         | Report/data storage        |
| 📄 BeautifulSoup | Web page parsing           |
| 🌍 Requests      | HTTP requests              |
| 📋 Pyperclip     | Clipboard automation       |
| 🔧 Git & GitHub  | Version control            |

---

# 📂 Projects

## 1️⃣ Daily Report Bot

📁 `daily_report_bot`

An automation script designed to collect online information and generate a daily report.

### Workflow

```text
Start
  ↓
Open / Access Web Data
  ↓
Extract Required Information
  ↓
Process Data
  ↓
Create Excel Report
  ↓
Save Screenshot
  ↓
Complete
```

### Key Features

* 🌐 Access web information
* 🔎 Extract required data
* 📊 Create Excel reports
* 📸 Capture screenshots
* 📁 Save output automatically
* 🤖 Reduce repetitive manual work

### Example Use Case

A daily business report can be automated instead of manually:

```text
Open Website
     ↓
Find Required Information
     ↓
Copy Data
     ↓
Update Excel
     ↓
Save Report
     ↓
Take Screenshot
```

---

# 2️⃣ Playwright Automation Bot

📁 `Play_wright_bot`

A browser automation project using **Playwright with Python**.

The project demonstrates automating browser interaction and searching for contacts using information stored in Excel.

### Workflow

```text
Excel Contact List
        ↓
Python Script
        ↓
Launch Browser
        ↓
Open Web Application
        ↓
Search Contact
        ↓
Validate Search Result
        ↓
Perform Automation
```

### Key Concepts

* Browser automation
* Web element interaction
* Search automation
* Excel-based input
* Automated browser control
* Error handling
* Screenshots / debugging
* Selector-based automation

---

# 🔄 PyAutoGUI vs Playwright

| Feature                    | PyAutoGUI | Playwright |
| -------------------------- | --------- | ---------- |
| Desktop automation         | ✅         | ❌          |
| Browser automation         | ✅         | ✅          |
| Mouse control              | ✅         | Limited    |
| Keyboard control           | ✅         | Limited    |
| Web element selectors      | ❌         | ✅          |
| Browser testing            | ❌         | ✅          |
| GUI-based automation       | ✅         | ❌          |
| Web application automation | Basic     | Advanced   |

### Simple Explanation

**PyAutoGUI**

```text
Python
  ↓
Mouse + Keyboard
  ↓
Screen / Desktop
```

**Playwright**

```text
Python
  ↓
Browser
  ↓
Web Elements
  ↓
Automation
```

---

# ⚙️ Installation

Clone the repository:

```bash
git clone https://github.com/Seema777111-git/Automation-with-Pyautogui-and-Playwright.git
```

Navigate to the project:

```bash
cd Automation-with-Pyautogui-and-Playwright
```

Create a Python virtual environment:

```bash
py -3.11 -m venv .venv
```

Activate the environment:

```bash
.venv\Scripts\activate
```

Install the required packages:

```bash
pip install -r requirements.txt
```

For Playwright:

```bash
playwright install
```

---

# ▶️ Running the Projects

## Daily Report Bot

Navigate to:

```bash
cd daily_report_bot
```

Run:

```bash
python daily_report_bot.py
```

---

## Playwright Bot

Navigate to:

```bash
cd Play_wright_bot
```

Run:

```bash
python playwright_assign.py
```

---

# 📊 Automation Architecture

```text
                 ┌─────────────────┐
                 │      USER       │
                 └────────┬────────┘
                          ↓
                 ┌─────────────────┐
                 │     Python      │
                 └────────┬────────┘
                          ↓
              ┌───────────┴───────────┐
              ↓                       ↓
       ┌──────────────┐        ┌──────────────┐
       │  PyAutoGUI   │        │  Playwright  │
       └──────┬───────┘        └──────┬───────┘
              ↓                       ↓
       Desktop / GUI             Web Browser
              │                       │
              └───────────┬───────────┘
                          ↓
                 ┌─────────────────┐
                 │ Automated Output │
                 └─────────────────┘
```

---

# 🎯 Learning Objectives

Through these projects, I am building practical knowledge in:

* Python automation
* Desktop automation
* Browser automation
* Web scraping
* Excel automation
* UI interaction
* Error handling
* Debugging automation scripts
* Repetitive task automation
* End-to-end automation workflows

---

# 🚀 Future Enhancements

Planned improvements include:

* [x] PyAutoGUI automation
* [x] Playwright browser automation
* [x] Excel integration
* [x] Screenshot generation
* [x] Web data extraction
* [ ] Better error handling
* [ ] Logging framework
* [ ] Automated email reports
* [ ] Scheduled execution
* [ ] AI-powered automation
* [ ] n8n integration
* [ ] End-to-end business process automation

---

# 💡 Real-World Applications

These automation techniques can be applied to:

* 📊 Daily business reporting
* 📧 Email automation
* 📑 Excel data processing
* 🌐 Web application automation
* 🔍 Data collection
* 🧾 Report generation
* 🛒 E-commerce operations
* 📦 Logistics operations
* 👥 CRM automation
* 🔄 Repetitive business processes

---

# 🎓 Project Goal

The goal of this repository is to progress from **simple Python automation → browser automation → business process automation → AI-powered automation**.

```text
Python
   ↓
PyAutoGUI
   ↓
Playwright
   ↓
Web Automation
   ↓
Business Automation
   ↓
AI Automation
   ↓
AI Agents
```

---

## 👩‍💻 Author

**Seema**

Interested in:

**Python Automation | AI | Generative AI | LangChain | RAG | AI Agents | Business Automation**

---

⭐ If you find this repository useful, feel free to explore the projects and follow the learning journey.
