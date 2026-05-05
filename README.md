# Amazon Seller Central Data Extraction & SharePoint Automation

## Overview
This project automates the extraction of **Amazon Seller Central financial and order data** using the **Selling Partner API (SP-API)** and uploads the processed daily report to **Microsoft SharePoint**.

The script is designed for marketplace reporting, reconciliation, and operational dashboards.

Currently configured for:

- South Africa Marketplace (Amazon ZA)

---

## Features

### Amazon SP-API Integration
- Fetches **Financial Events**
- Retrieves:
  - Orders
  - Charges
  - Fees
  - Taxes
  - Promotions
  - Fulfillment Fees

### Orders API Integration
Fetches additional order-level details:

- Product Title
- Quantity Ordered
- Shipping Address
- City / State
- Tax Collection Model

### Data Cleaning & Transformation
Creates business-ready report with columns such as:

- Date/Time
- Order ID
- SKU
- Product Sales
- Shipping Credits
- Gift Wrap Credits
- Promotional Rebates
- Selling Fees
- FBA Fees
- Final Total

### SharePoint Automation
Automatically uploads final Excel report to SharePoint folder.

### Daily Automation Ready
Can be scheduled via:

- Windows Task Scheduler
- Cron Jobs
- Python Scheduler

---

## Tech Stack

- Python
- Pandas
- Requests
- Amazon SP-API
- Office365 / SharePoint API
- OpenPyXL

---

## Folder Output Example

```text
AmazonSouthAfrica_20260505.xlsx
