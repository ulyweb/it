# 🛡️ IT_HealthCheck — Windows 11 Diagnostic Dashboard

![Version](https://img.shields.io/badge/Version-V7-gold?style=flat-square)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows%2011-0078D6?style=flat-square)
![License](https://img.shields.io/badge/License-Internal-lightgrey?style=flat-square)

> A comprehensive, all-in-one PowerShell diagnostic dashboard for IT support teams. Captures Windows 11 endpoint telemetry, event logs, resource anomalies, and session snapshots — then generates Copilot-ready prompts for AI-assisted root cause analysis, safe self-healing, and Ticketing-ready documentation.

---

## 📑 Table of Contents

- [Features](#-features)
- [Requirements](#-requirements)
- [Quick Start](#-quick-start)
- [Architecture Overview](#-architecture-overview)
- [Maximum Data Capture Workflow](#-maximum-data-capture-workflow)
- [What Gets Captured](#-what-gets-captured)
- [Standard Diagnostic Workflow (Prompts 1–3)](#-standard-diagnostic-workflow-prompts-13)
- [Advanced Diagnostic Workflow (Prompts 4–7)](#-advanced-diagnostic-workflow-prompts-47)
- [Proactive Monthly Workflow](#-proactive-monthly-workflow)
- [Coverage Matrix](#-coverage-matrix)
- [Quick Reference](#-quick-reference)
- [Design Principles](#-design-principles)
- [Interactive Workflow Guide](#-interactive-workflow-guide)
- [Pro Tips](#-pro-tips)
- [File Structure](#-file-structure)

---

## ✨ Features

- **Multi-Tab SPA Web Dashboard** — Real-time telemetry overview, live user activity, processes, services, installed apps, Device Manager, event logs, hardware events, and power events
- **Local & Remote Execution** — Run against the local machine or connect to a remote endpoint via WinRM
- **Session Cache Snapshots** — Processes, services, installed applications, and Device Manager snapshots exported to CSV/HTML
- **Resource Anomaly Detection** — RAM exhaustion, disk capacity warnings, CPU/thermal/processor power spike events
- **Hibernate Status & Enable Action** — Detect and enable hibernate directly from the dashboard
- **Device Removal Action** — Remove non-OK devices from Device Manager via the UI
- **Day-by-Day Capture Reports** — Historical CSV matrix + offline HTML dashboard with event counts, context, and expanded details
- **Copilot Diagnostic Prompts (V6)** — Two automated copy-ready prompts: Root Cause + Safe Self-Healing, and ServiceNow/User Communication
- **Dynamic Persona Gating (V7)** — Auto-enables Zoom, Network, Thermal, and Outlook diagnostic modules only when evidence is present
- **Visual HTML Report Generator (Prompt 3)** — Converts Copilot analysis into an interactive, animated, self-contained HTML report
- **Agent Builder Blueprint** — Complete package for creating a dedicated IT_HealthCheck agent in Microsoft 365 Copilot Agent Builder or Copilot Studio
- **Interactive Landing Overlay** — Professional full-screen guide with shield icon, starfield effect, and gold-accented UI

---

## 📋 Requirements

| Requirement | Details |
|---|---|
| **PowerShell** | 5.1 or later |
| **OS** | Windows 11 (Windows 10 compatible) |
| **Privileges** | Run as Administrator |
| **Remote** | WinRM enabled on target (for remote mode) |
| **Browser** | Any modern browser (for SPA dashboard) |

---

## 🚀 Quick Start

1. **Download** the script to your IT workstation.
2. **Run as Administrator** in PowerShell:

```powershell
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "IT_HealthCheck.ps1"
```

3. **Choose target**: `[L]ocal` machine or `[R]emote` machine.
4. The SPA dashboard opens automatically in the default browser.
5. Navigate tabs, fetch data, and generate capture reports.

---

## 🏗️ Architecture Overview

```
┌──────────────────────────────────┐
│  Master Deployment Tool          │
│  (Outer Wrapper)                 │
│  ┌────────────────────────────┐  │
│  │  Embedded Dashboard Payload│  │
│  │  (SPA Web App)             │  │
│  │  ┌──────────────────────┐  │  │
│  │  │ HTTP Listener Server │  │  │
│  │  │ (Port 15500-15999)   │  │  │
│  │  └──────────────────────┘  │  │
│  │  ┌──────────────────────┐  │  │
│  │  │ HTML/CSS/JS Frontend │  │  │
│  │  └──────────────────────┘  │  │
│  │  ┌──────────────────────┐  │  │
│  │  │ Capture Engine +     │  │  │
│  │  │ HTML Report Generator│  │  │
│  │  └──────────────────────┘  │  │
│  └────────────────────────────┘  │
└──────────────────────────────────┘
```

---

## 📥 Maximum Data Capture Workflow

Follow these 5 steps **before** generating the capture report to maximize diagnostic evidence:

| Step | Action | Dashboard Tab |
|:---:|---|---|
| **1** | Fetch **Active Processes** | ⚙ Active Processes |
| **2** | Fetch **System Services** | ⚙ System Services |
| **3** | Fetch **Installed Apps** | 💾 Installed Apps |
| **4** | Fetch **Device Manager** | 💻 Device Manager |
| **5** | Generate **Capture Report** | 📊 Capture & Reports |

> **Why?** Steps 1–4 populate the Session Cache. When you generate the report in Step 5, all cached snapshots are embedded in both the CSV and HTML outputs — giving Copilot the richest possible evidence for analysis.

---

## 📊 What Gets Captured

| Category | Details |
|---|---|
| **System Identity** | Hostname, manufacturer, model, CPU, RAM, BIOS version, boot time |
| **Security & Virtualization** | VBS status, Hyper-V status, hibernate status |
| **Network** | Active adapter, SSID, link speed, IPv4 address |
| **Live Resources** | CPU %, RAM %, GPU %, NPU %, disk usage %, free space |
| **Historical Events** | Unexpected shutdowns, clean reboots, BSODs, app crashes, app hangs, network drops |
| **Resource Anomalies** | RAM exhaustion events, disk capacity warnings, CPU/thermal/processor power spikes |
| **User Context** | Active foreground windows, recently launched processes, uptime, pending reboot, recent patches |
| **Session Snapshots** | Processes (with CPU/RAM), services (with status), installed apps (with versions), Device Manager (with status) |

---

## 🔧 Standard Diagnostic Workflow (Prompts 1–3)

| Prompt | Name | Time | Difficulty | Description |
|:---:|---|:---:|:---:|---|
| **1** | Root Cause + Safe Self-Healing | ~3 min | ⭐ Easy | Analyzes the full report, identifies confirmed findings vs. likely causes, recommends safe built-in Windows recovery steps, labels admin-only actions, and provides escalation criteria. Includes dynamic persona modules (Zoom, Network, Thermal, Outlook) auto-enabled by evidence. |
| **2** | ServiceNow + User Communication | ~2 min | ⭐ Easy | Generates ITSM-ready ticket documentation with category, subject, business impact, work notes (9 fields), status, and a polished non-technical user-facing message. |
| **3** | Convert to Visual HTML Report | ~5 min | ⭐⭐ Medium | Follow-up prompt in the **same** Copilot conversation after Prompt 1. Converts the full analysis into an interactive, animated, self-contained HTML file with landing overlay, resource tiles, event badges, and diagnostic cards. |

### Prerequisites
- **Prompt 1**: Generated HTML report with capture data
- **Prompt 2**: Generated HTML report with capture data
- **Prompt 3**: Must be used as a follow-up in the same conversation as Prompt 1

---

## 🔬 Advanced Diagnostic Workflow (Prompts 4–7)

| Prompt | Phase | Name | Time | Difficulty | Description |
|:---:|:---:|---|:---:|:---:|---|
| **4** | V8 | Driver Audit + Update Health + Boot Config | ~5 min | ⭐⭐ Medium | Deep analysis of driver versions, Windows Update history, boot configuration, and firmware currency. Identifies outdated drivers, failed updates, and boot anomalies. |
| **5** | V9 | Disk & Storage Health (SMART + SSD) | ~5 min | ⭐⭐ Medium | Storage subsystem analysis including SMART data, SSD wear leveling, disk latency, and storage controller health. Identifies pre-failure indicators. |
| **6** | V10 | Baseline Comparison + Drift Detection | ~8 min | ⭐⭐⭐ Advanced | Compares a Day 1 baseline capture against a Day 30 current capture. Detects configuration drift, new software, removed services, resource trend changes, and emerging issues. |
| **7** | V11 | Profile Health + Software Conflicts + Startup | ~5 min | ⭐⭐ Medium | User profile integrity check, software conflict detection, startup program analysis, and add-in health review. Identifies profile corruption and application-layer conflicts. |

### Prerequisites
- **Prompts 4–5**: Generated HTML report with Device Manager and services snapshots
- **Prompt 6**: Two captures — one baseline (Day 1) and one current (Day 30)
- **Prompt 7**: Generated HTML report with installed apps and processes snapshots

---

## 📅 Proactive Monthly Workflow

| Phase | Timing | Action | Tool |
|:---:|---|---|---|
| **1** | Day 1 | Generate **baseline** capture report | Dashboard → Capture & Reports |
| **2** | Day 30 | Generate **current** capture report | Dashboard → Capture & Reports |
| **3** | Day 30 | Run **Prompt 6** — Baseline vs. Current drift detection | Copilot with both reports |
| **4** | Day 30 | Run **Prompt 7** — Profile health + software conflicts | Copilot with current report |
| **5** | Day 30 | **Archive** both reports and Copilot analysis results | Save to ticket / shared drive |

> **Use case**: Proactive fleet health monitoring, quarterly device reviews, pre-escalation evidence gathering, and executive device stability audits.

---

## 📋 Coverage Matrix

| Diagnostic Area | P1 | P2 | P3 | P4 | P5 | P6 | P7 | Agent |
|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| Unexpected Shutdowns | ✅ | ✅ | ✅ | — | — | ✅ | — | ✅ |
| BSODs / Bugcheck | ✅ | ✅ | ✅ | — | — | ✅ | — | ✅ |
| Application Crashes | ✅ | ✅ | ✅ | — | — | ✅ | ✅ | ✅ |
| Application Hangs | ✅ | ✅ | ✅ | — | — | ✅ | ✅ | ✅ |
| Network Drops | ✅ | ✅ | ✅ | — | — | ✅ | — | ✅ |
| RAM Exhaustion | ✅ | ✅ | ✅ | — | — | ✅ | — | ✅ |
| Disk Warnings | ✅ | ✅ | ✅ | — | ✅ | ✅ | — | ✅ |
| CPU / Thermal Spikes | ✅ | ✅ | ✅ | — | — | ✅ | — | ✅ |
| Driver / Firmware | — | — | — | ✅ | — | ✅ | — | ✅ |
| Windows Update Health | — | — | — | ✅ | — | ✅ | — | ✅ |
| Boot Configuration | — | — | — | ✅ | — | — | — | ✅ |
| Storage / SMART / SSD | — | — | — | — | ✅ | — | — | ✅ |
| Configuration Drift | — | — | — | — | — | ✅ | — | ✅ |
| Profile / Add-ins | — | — | — | — | — | — | ✅ | ✅ |
| Startup Programs | — | — | — | — | — | — | ✅ | ✅ |

---

## 📖 Quick Reference

| Tool | Purpose | Output |
|---|---|---|
| **Prompt 1** | Root cause + safe self-healing | Copilot analysis text |
| **Prompt 2** | ServiceNow ticket + user message | ITSM-ready documentation |
| **Prompt 3** | Visual HTML diagnostic report | Self-contained `.html` file |
| **Prompt 4** | Driver/update/boot audit | Copilot analysis text |
| **Prompt 5** | Disk & storage SMART analysis | Copilot analysis text |
| **Prompt 6** | Baseline drift detection | Copilot comparison report |
| **Prompt 7** | Profile + software conflicts | Copilot analysis text |
| **Agent Builder** | Create dedicated Copilot agent | Agent configuration package |

---

## 🎯 Design Principles

| Principle | Description |
|---|---|
| **Evidence-First** | Never claim a root cause without supporting evidence from the report. Separate confirmed findings from likely causes. |
| **Safe Before Destructive** | Recommend built-in Windows tools and reversible actions first. Label admin-only steps clearly. Never recommend OS reset, profile rebuild, registry edits, or driver removal unless evidence supports it. |
| **Audit-Ready** | All output is professional, concise, and suitable for ServiceNow tickets, escalation notes, and compliance reviews. |
| **Business Continuity** | Prioritize user productivity and data safety. Minimize disruption during troubleshooting. |

### Safe Remediation Model

```
Detect → Summarize → Recommend → Approve → Remediate → Document
```

---

## 🖥️ Interactive Workflow Guide

The exported HTML dashboard includes a full-screen **Interactive Workflow Guide** landing overlay featuring:

- 🛡️ **Shield icon** with gold accent and drop-shadow glow
- ✨ **Starfield background** with subtle CSS particle effects
- 🕐 **Live local time badge** updating every second
- 📋 **Persona Evidence Gate (V7)** — Visual pills showing which diagnostic modules (Zoom, Network, Thermal, Outlook) are auto-enabled based on report evidence
- 🔘 **"I'm Ready to Begin"** call-to-action button with fade transition to the report

---

## 💡 Pro Tips

1. **Maximize evidence** — Always fetch Processes, Services, Apps, and Device Manager **before** generating the capture report
2. **Use Prompt 3 as a follow-up** — Run it in the **same** Copilot conversation as Prompt 1 for best results
3. **Adjust lookback period** — Use Dashboard Settings to increase the telemetry lookback from 2 days up to 365 days for deeper historical analysis
4. **Compare baselines monthly** — Use the Proactive Monthly Workflow with Prompt 6 to catch configuration drift before users report issues
5. **Open HTML reports locally** — All HTML reports are fully self-contained with zero external dependencies — they work offline with no internet connection
6. **Use the Agent Builder Blueprint** — Create a permanent Copilot agent for your team so any technician can analyze reports without copying prompts manually

---

## 📁 File Structure

```
C:\IT_Scripts\
└── IT_HealthCheck.ps1          # Main script (Master Deployment + Embedded Payload)

C:\Temp\
├── HOSTNAME_capture_log_YYYYMMDD_HHMMSS.csv    # Day-by-day CSV matrix
├── HOSTNAME_YYYYMMDD_HHMMSS.html               # Offline HTML dashboard
├── active_windows.txt                           # Temporary (auto-cleaned)
├── GetWindows.ps1                               # Temporary (auto-cleaned)
└── RunHidden.vbs                                # Temporary (auto-cleaned)
```

---

<p align="center">
  <sub>Built for IT support teams · Powered by PowerShell + Copilot · Evidence-first diagnostics</sub>
</p>
