# Compliance Automation & Dashboard

This repository contains automated compliance management and reporting tools for ESIC and EPFO workflows, including dashboard reporting and automated notifications.

## Project Structure
- `dashboard/`: Web dashboard server and UI assets for monitoring compliance status.
- `ESIC_AUTOMATION_V2/`: Automation scripts and workflows for ESIC processes.
- `New_EPFO_AUTO_software/`: Automation scripts for EPFO data downloads and reconciliation.
- `NEW_ESIC_AUTO_EPFO_SOFTWARE/`: Integrated ESIC & EPFO automation components.
- `mailer.py`: Microsoft Graph email notification service.

## Setup Instructions
1. Clone this repository.
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in your Microsoft Graph API credentials.
4. Copy `ms_graph_config.example.json` to `ms_graph_config.json` with appropriate API keys.
5. Run the dashboard or automation scripts as needed.
