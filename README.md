# Automated Pixel Privacy Audit

An automated privacy auditing tool that identifies third-party tracking technologies on publicly accessible healthcare websites, analyzes their behavior, and generates structured Excel reports for security and privacy assessments.

---

## Overview

Healthcare websites frequently integrate third-party services for analytics, advertising, appointment scheduling, and marketing. These integrations can introduce privacy risks when tracking technologies execute on patient-facing pages.

Automated Pixel Privacy Audit performs passive browser-based analysis using Playwright to detect these technologies, analyze their behavior, and generate evidence-based reports.

The tool only analyzes publicly accessible resources and does not require authentication or interact with protected patient information.

---

## Key Features

### Tracker Detection

Detects major third-party tracking technologies including:

* Meta Pixel
* Google Analytics (GA4)
* Google Tag Manager
* Google Ads
* DoubleClick

---

### Automated Page Discovery

Automatically discovers pages such as:

* Patient Portal
* Login
* Registration
* Appointment Booking
* Telehealth
* Contact Pages

---

### Redirect Analysis

Detects HTTP and browser redirects while crawling target websites.

---

### Consent Analysis

Identifies:

* Visible consent banners
* Tracker execution timing
* Trackers firing before a visible consent mechanism

---

### Payload Inspection

Analyzes tracker requests for:

* Browser identifiers
* Event types
* Page metadata
* Request parameters

---

### Risk Assessment

Automatically assigns a risk score based on:

* Number of trackers
* Consent observations
* Page sensitivity
* Payload characteristics

---

### Excel Report Generation

Produces structured Excel reports with multiple worksheets including:

* Executive Summary
* Page Findings
* Tracker Details

---

## Workflow

```
                Target Website
                       │
                       ▼
          Discover Public Pages
                       │
                       ▼
      Identify Healthcare Pages
                       │
                       ▼
      Launch Playwright Browser
                       │
                       ▼
      Capture Network Requests
                       │
                       ▼
     Detect Third-Party Trackers
                       │
                       ▼
      Analyze Consent Behaviour
                       │
                       ▼
      Inspect Tracker Payloads
                       │
                       ▼
         Calculate Risk Score
                       │
                       ▼
        Generate Excel Report
```

---

## Technologies Used

* Python
* Playwright
* Requests
* BeautifulSoup4
* Pandas
* OpenPyXL
* lxml

---

## Installation

Clone the repository:

```bash
git clone https://github.com/KushwanthD/Automated-Pixel-Privacy-Audit.git
```

Move into the project:

```bash
cd Automated-Pixel-Privacy-Audit
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Install Playwright browsers:

```bash
playwright install
```

---

## Usage

Run the tool:

```bash
python Pixel_Tracking.py
```

The tool will automatically:

1. Launch a browser session.
2. Discover publicly accessible pages.
3. Capture network requests.
4. Detect third-party trackers.
5. Analyze consent behavior.
6. Calculate a risk score.
7. Generate an Excel report.

---

## Report Contents

### Executive Summary

* Website information
* Total pages analyzed
* Tracker summary
* Overall risk assessment

### Page Findings

* Page URL
* Page classification
* Redirect information
* Consent observations
* Detected trackers
* Risk score

### Tracker Details

* Tracker vendor
* Event type
* Payload observations
* Network request information

---

## Project Scope

This project focuses on passive analysis of publicly accessible websites.

It does **not**:

* bypass authentication
* access patient portals
* collect protected health information (PHI)
* exploit vulnerabilities
* modify target systems
* perform intrusive security testing

---

## Intended Use

This project is intended for:

* Security Research
* Privacy Assessments
* Responsible Disclosure
* Internal Security Reviews
* Educational Purposes

---

## Future Improvements

* Support for additional tracker vendors
* Cookie consent framework detection
* Interactive web dashboard
* Historical scan comparison
* PDF report generation
* Automated screenshots
* Expanded healthcare page classification

---

## Disclaimer

This tool is intended solely for educational, research, and authorized security assessment purposes.

Users are responsible for ensuring that all assessments comply with applicable laws, organizational policies, and responsible disclosure practices.
