# Automated Pixel Privacy Audit

An automated privacy auditing tool that performs passive analysis of publicly accessible healthcare websites to identify third-party tracking technologies, evaluate tracker behavior, and generate structured privacy assessment reports.

---

## Overview

Automated Pixel Privacy Audit is a Python-based browser automation tool designed to help security researchers, privacy auditors, and organizations identify third-party tracking technologies deployed on healthcare websites.

The tool performs passive browser-based analysis without interacting with authenticated resources or protected patient information. It focuses on understanding how tracking technologies behave on publicly accessible patient-facing pages.

The assessment includes:

- Third-party tracker detection
- Consent banner detection
- Tracker execution timing
- Healthcare page classification
- Browser identifier detection
- Risk assessment
- Excel report generation

---

## Features

### Tracker Detection

Automatically detects common tracking technologies including:

- Meta Pixel (Facebook)
- Google Analytics (GA4)
- Google Tag Manager
- Google Ads / DoubleClick

---

### Healthcare Page Discovery

Automatically identifies healthcare-related pages such as:

- Patient Portal
- Telemedicine
- Appointment Booking
- Contact Forms
- Login Pages
- Registration Pages

---

### Passive Network Analysis

The tool captures browser network activity to determine:

- Tracker vendors
- Events fired
- Request timing
- Browser identifiers
- Page metadata
- Tracking behavior

---

### Consent Analysis

Evaluates whether trackers execute:

- Before user interaction
- Before consent
- Without a visible consent mechanism

---

### Automated Reporting

Generates structured Excel reports containing:

- Website information
- Pages analyzed
- Detected trackers
- Privacy observations
- Risk score
- Supporting evidence

---

## Detection Workflow

```
Target Website
        │
        ▼
Launch Playwright Browser
        │
        ▼
Discover Patient-Facing Pages
        │
        ▼
Capture Network Requests
        │
        ▼
Identify Third-Party Trackers
        │
        ▼
Analyze Consent Behaviour
        │
        ▼
Evaluate Privacy Exposure
        │
        ▼
Generate Audit Report
```

---

## Technologies Used

- Python
- Playwright
- Requests
- BeautifulSoup
- Pandas
- OpenPyXL

---

## Reported Information

The generated report may include:

- Redirect Information
- Tracker Vendors
- Tracker Events
- Healthcare Page Classification
- Consent Banner Presence
- Browser Identifier Detection
- Request Timing
- Privacy Risk Assessment

---

## Example Findings

The scanner is capable of identifying observations such as:

- Meta Pixel executing on patient-facing pages.
- Google Tag Manager loaded before user interaction.
- Third-party trackers active without a visible consent mechanism.
- Browser identifiers transmitted to third-party analytics providers.
- Multiple tracking technologies operating simultaneously on healthcare pages.

---

## Ethical Use

This project is intended for:

- Security research
- Privacy assessments
- Internal security reviews
- Responsible disclosure
- Educational purposes

The scanner only analyzes publicly accessible resources and passive browser network activity.

It **does not**:

- bypass authentication
- exploit vulnerabilities
- access patient accounts
- collect protected health information (PHI)
- modify target systems
- perform intrusive testing

Always obtain appropriate authorization before conducting assessments beyond publicly accessible resources.

---

## Limitations

Current capabilities include client-side privacy analysis only.

The tool does **not**:

- determine regulatory compliance
- identify server-side tracking
- inspect authenticated user sessions
- access protected patient records
- infer legal violations

All findings should be manually validated before disclosure.

---

## Installation

Clone the repository:

```bash
git clone https://github.com/KushwanthD/Automated-Pixel-Privacy-Audit.git
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

Run the tool:

```bash
python Pixel_Tracking.py
```
## Usage

Run the scanner:

```bash
python Pixel_Tracking.py
```

Provide the target website when prompted.

The tool will automatically:

1. Launch a browser session.
2. Discover relevant healthcare pages.
3. Capture network requests.
4. Detect third-party trackers.
5. Analyze privacy-related observations.
6. Generate an Excel report.

---

## Future Enhancements

Planned improvements include:

- Additional tracker support
- Cookie consent framework identification
- Privacy policy correlation
- Tracker visualization dashboard
- Web application interface
- Historical scan comparison
- Automated evidence collection
- Expanded healthcare page classification

---

## Disclaimer

This project is intended solely for educational, research, and authorized security assessment purposes.

The author is not responsible for misuse of this software. Users are responsible for ensuring that all assessments comply with applicable laws, organizational policies, and responsible disclosure practices.