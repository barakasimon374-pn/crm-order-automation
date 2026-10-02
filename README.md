# CRM Order Automation

A Python-based CRM order automation system that reads sales-order data from Excel, searches CRM products, checks inventory, posts orders, collects payment, and generates an Excel posting report.

## Project Overview

This project automates a repetitive order-entry workflow between Excel-based sales data and a browser-based CRM system.

The workflow includes:

* Excel order-file validation and processing
* Product-code extraction from item numbers and descriptions
* Support for 5-digit and 6-digit MRP product codes
* CRM product search and inventory checking
* Quantity allocation across available CRM inventory
* Automated CRM customer/order processing
* Payment collection workflow
* Customer-name rotation for successive orders
* Monitoring and exclusion of selected item codes
* Excel posting-report generation
* Inventory-status highlighting in the generated report
* Local web interface for uploading Excel files and starting the automation
* Downloadable posting report from the local web interface

## Main Files

### `test_brave_payment_working_normal_items.py`

The main automation engine.

Responsibilities include:

* Reading Excel sales data
* Validating order lines
* Extracting CRM search codes
* Searching CRM products
* Checking inventory
* Allocating order quantities
* Creating CRM orders
* Processing payment
* Creating the final Excel posting report

### `excel_upload_page.py`

A local FastAPI web application used to:

1. Select an Excel file
2. Upload it to the automation workflow
3. Start CRM processing
4. Monitor execution status
5. Download the generated posting report

## Reporting

The generated Excel report identifies inventory conditions using formatting:

* Enough inventory — Excel Description highlighted bright yellow
* Zero inventory — product description shown in red
* Insufficient inventory — product description shown in red and relevant quantity/inventory cells highlighted
* Monitored item codes — displayed in a blue monitoring section and excluded from CRM posting

## Technology

* Python
* Playwright
* FastAPI
* Uvicorn
* OpenPyXL
* PowerShell
* Excel

## Local Setup

Create and activate a Python virtual environment, then install the required packages.

The CRM URL is supplied through an environment variable rather than being stored directly in the source code.

PowerShell:

```powershell
$env:CRM_URL="YOUR_CRM_URL"
```

Start the local upload application:

```powershell
python .\excel_upload_page.py
```

Open:

```text
http://127.0.0.1:8000
```

Select the Excel order file and start the CRM posting process.

## Project Structure

```text
crm_order_automation/
│
├── excel_upload_page.py
├── test_brave_payment_working.py
├── .gitignore
└── README.md
```

## Notes

This repository contains the cleaned portfolio version of the automation project. Local Excel files, virtual-environment files, generated reports, caches, and runtime state are excluded from version control.

## Portfolio Project

This project demonstrates practical automation using Python, browser automation, Excel processing, inventory logic, reporting, and a local web interface.
