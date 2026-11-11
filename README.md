# 🛡️ Compliance Automation & Dashboard System

[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-Framework-009688?logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![Selenium](https://img.shields.io/badge/Selenium-Automation-43B02A?logo=selenium&logoColor=white)](https://www.selenium.dev/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft_Graph-API-0078D4?logo=microsoft&logoColor=white)](https://learn.microsoft.com/en-us/graph/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

An enterprise-grade automation platform designed to streamline compliance management across **ESIC** (Employees' State Insurance Corporation) and **EPFO** (Employees' Provident Fund Organisation) portals. Features automated downloads, data reconciliation, CAPTCHA handling, status dashboards, and automated email notifications via Microsoft Graph API.

---

## 🌟 Key Features

- **🤖 Automated Portal Processing**: Automated authentication and data extraction from ESIC and EPFO portals with integrated CAPTCHA handling.
- **📊 Real-Time Status Dashboard**: Interactive web dashboard (`dashboard/`) providing live monitoring, task status, execution logs, and compliance analytics.
- **🔄 Smart Data Reconciliation**: Automatic merging and comparison of portal downloads against HRMS master records (`compare_data.py`, `excel_merge.py`).
- **📧 Automated Email Notifications**: Secure automated dispatch of compliance reports and discrepancy alerts via **Microsoft Graph API / Outlook** (`mailer.py`).
- **🛡️ Enterprise Security**: Strict environment isolation with `.env` credential management, token caching, and zero data leakage.

---

## 📂 Project Architecture

```
Complience/
├── dashboard/                      # Web dashboard application
│   ├── server.py                   # Dashboard backend server
│   ├── requirements.txt            # Dashboard-specific dependencies
│   └── static/                     # Dashboard UI (HTML, CSS, JS)
│
├── New_EPFO_AUTO_software/         # EPFO automation suite
│   └── EPFO_AUTO/
│       ├── main.py                 # EPFO orchestration entrypoint
│       ├── config.py               # EPFO configuration settings
│       ├── run_epfo_txt_download.py# EPFO text download handler
│       └── modules/                # Modules: download, merge, compare, email
│
├── NEW_ESIC_AUTO_EPFO_SOFTWARE/    # ESIC automation suite
│   └── ESIC_AUTO_project_structure/
│       └── ESIC_AUTO/
│           ├── Main.py             # ESIC orchestration entrypoint
│           ├── dbdownload.py       # DB download utility
│           ├── esic_portal_download.py # Portal download handler
│           ├── read_captcha/       # CAPTCHA reading & OCR module
│           ├── merge.py            # Data merge engine
│           └── send_esic_email.py  # Email dispatch service
│
├── mailer.py                       # Microsoft Graph API email delivery service
├── create_desktop_shortcut.ps1     # 1-Click launcher creation script
├── start_dashboard.bat             # Batch launcher for dashboard
├── .env.example                    # Environment variable template
├── ms_graph_config.example.json    # Microsoft Graph configuration template
└── requirements.txt                # Core Python dependencies
```

---

## 🚀 Getting Started

### Prerequisites
- Python 3.10 or higher
- Google Chrome & compatible ChromeDriver (or webdriver-manager)
- Microsoft Azure App Registration (for MS Graph email integration)

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/KDI4U/Complience.git
   cd Complience
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python -m venv .venv
   # Windows:
   .venv\Scripts\activate
   # Linux/macOS:
   source .venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   pip install -r dashboard/requirements.txt
   ```

4. **Configure Environment Variables:**
   Copy `.env.example` to `.env` and fill in your credentials:
   ```bash
   cp .env.example .env
   ```
   ```ini
   MS_GRAPH_CLIENT_ID=your_azure_client_id
   MS_GRAPH_TENANT_ID=your_azure_tenant_id
   MS_GRAPH_CLIENT_SECRET=your_azure_client_secret
   MS_GRAPH_USER_EMAIL=your_sender_email@company.com
   ```

5. **Configure MS Graph settings:**
   Copy `ms_graph_config.example.json` to `ms_graph_config.json`:
   ```bash
   cp ms_graph_config.example.json ms_graph_config.json
   ```

---

## 🖥️ Running the Application

### 1. Launch the Status Dashboard
```bash
python dashboard/server.py
# Or double-click: start_dashboard.bat
```
Visit `http://localhost:8000` (or configured port) in your web browser.

### 2. Run EPFO Automation
```bash
python New_EPFO_AUTO_software/EPFO_AUTO/main.py
```

### 3. Run ESIC Automation
```bash
python NEW_ESIC_AUTO_EPFO_SOFTWARE/ESIC_AUTO_project_structure/ESIC_AUTO/Main.py
```

---

## 🔒 Security & Privacy Notice
This repository contains source code only. All operational runtime data, employee spreadsheets (`.xlsx`, `.csv`), authentication tokens, and downloaded reports are protected via `.gitignore` and are never committed.

---

## 📄 License
This project is licensed under the MIT License - see the LICENSE file for details.

---

<div align="center">
  <b>Maintained by <a href="https://github.com/KDI4U">Kundan Kumar</a></b>
</div>
