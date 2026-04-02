# 🚀 Azure VM Intelligent Monitoring

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![Azure](https://img.shields.io/badge/Azure-Monitor-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Production--Ready-success)

> AI-powered Azure VM monitoring, reporting, and automated email
> notification system.

------------------------------------------------------------------------

## 🧠 Overview

**Azure VM Intelligent Monitoring** is an advanced automation solution
that collects, analyzes, and reports Azure Virtual Machine performance
metrics.

It integrates multiple Azure services and intelligent filtering
techniques to generate **clean, actionable, and enterprise-ready
reports**, and automatically delivers them via email using Microsoft
Graph API.

------------------------------------------------------------------------

## ✨ Key Features

### 📊 Performance Monitoring

-   CPU, Memory, Disk, Network analysis
-   Hourly average & peak metrics
-   Clean visualization charts

### 🧠 Intelligent Data Filtering

-   Removes noisy mounts (`/snap`, `/proc`, `/sys`, `/dev`)
-   Focuses only on meaningful disks
-   Eliminates false alerts

### 🔍 Multi-Source Data Handling

-   Supports AMA (`InsightsMetrics`)
-   Supports legacy (`Perf`)
-   Falls back to Azure Metrics API

### 🛡️ Security Insights

-   NSG rule analysis
-   Risk classification
-   Security recommendations

### 📧 Automated Email Notifications

-   Sends reports via Microsoft Graph API
-   Fully automated delivery workflow
-   Supports enterprise mail systems

### 🧾 Report Generation

-   Professional Word reports
-   Executive summary
-   VM-wise breakdown

------------------------------------------------------------------------

## 🏗️ Architecture

Azure Monitor → Data Collection → Intelligent Processing → Report
Generation → Email Delivery

------------------------------------------------------------------------

## 📦 Repository Structure

azure-vm-intelligent-monitoring/ │ ├── main.py \# Orchestrates complete
workflow ├── collect_metrics.py \# Data collection from Azure ├──
generate_report.py \# Report generation ├── graph_mailer.py \# Email
notifications via Graph API ├── DEPLOYMENT_GUIDE.md ├── requirements.txt
└── README.md

------------------------------------------------------------------------

## ⚙️ Installation

``` bash
git clone https://github.com/<your-username>/azure-vm-intelligent-monitoring.git
cd azure-vm-intelligent-monitoring
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

------------------------------------------------------------------------

## 🔐 Configuration

Provide: - Subscription ID - Tenant ID - Client ID - Client Secret -
Workspace ID - Graph API permissions

------------------------------------------------------------------------

## ▶️ Usage

``` bash
python main.py
```

------------------------------------------------------------------------

## 🔄 Workflow

1.  Collect metrics from Azure
2.  Process & filter data
3.  Generate report
4.  Send email via Graph API

------------------------------------------------------------------------

## 📊 Output

-   Word report (.docx)
-   Charts & analysis
-   Security findings
-   Email delivery

------------------------------------------------------------------------

## 🛠️ Tech Stack

-   Python
-   Azure Monitor APIs
-   Log Analytics
-   Microsoft Graph API
-   python-docx
-   matplotlib

------------------------------------------------------------------------

## 📘 Documentation

See: DEPLOYMENT_GUIDE.md

------------------------------------------------------------------------

## 📄 License

MIT License

------------------------------------------------------------------------

## ⭐ Support

If useful, give a star ⭐
