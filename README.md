# Parsonage Tenant Management System

This Google Apps Script-based solution provides a robust and automated system for managing tenants in a parsonage or multi-room rental property. Leveraging the power of Google Sheets, Google Forms, and Google Apps Script, it aims to streamline administrative tasks for landlords and improve communication with tenants.

## Features

### **Tenant & Room Management (Google Sheet: `Tenants`)**
* **Room Tracking:** Easily track room numbers, their standard rental price, and negotiated prices.
* **Occupancy Status:** Monitor which rooms are occupied, vacant, or pending.
* **Tenant Details:** Store essential tenant information including name, email, move-in date.
* **Security Deposit Tracking:** Record and verify security deposit payments.
* **Payment Status:** Automatically track monthly payment status (Paid, Due, Overdue).

### **Financial Overview (Google Sheet: `Budget`)**
* **Income & Expense Tracking:** Log all rental income and property-related expenses (utilities, maintenance, etc.).
* **Budget Analysis:** Tools to analyze profitability and visualize financial performance over time with charts.

### **Automated Workflows (Google Apps Script)**

#### **For Landlords/Managers:**
* **Rent Payment Reminders:** Automated email reminders for tenants when rent is due or overdue.
* **Late Payment Alerts:** Automatic notifications to the house manager when a tenant is significantly overdue on rent (e.g., more than one month).
* **Monthly Rent Invoicing:** Generate and automatically email personalized PDF rent invoices to tenants.
* **Payment Status Updates:** Functions to easily mark payments received.
* **Custom Menus:** Integrated directly into the Google Sheet for quick access to key functions.

#### **For Tenants (via Google Forms & Automated Emails):**
* **Online Application Form:** A dedicated Google Form for prospective tenants to submit applications, including document uploads (e.g., proof of income).
    * *Automation:* Upon submission, an automatic welcome email is sent to the applicant containing house rules, cultural vision, and rental agreement details.
* **Move-Out Request Form:** A Google Form for tenants to formally submit their move-out date.
    * *Automation:* Upon submission, an automatic email is sent to the tenant outlining move-out expectations and procedures.

## How It Works

The system is built primarily on:
* **Google Sheets:** As the central database for all tenant, room, and financial data.
* **Google Apps Script:** The automation engine that connects Google Sheets with Google Forms, sends emails, generates documents, and performs scheduled tasks.
* **Google Forms:** Provides user-friendly interfaces for tenant applications and move-out requests, with data flowing directly into the Google Sheet.
* **(Optional) Google Sites:** Can be used to create a simple, public-facing portal for tenants to access forms or general information.

## Setup & Installation

1.  **Create a Google Sheet:** Create a new Google Sheet (e.g., "Parsonage Tenant Manager") and set up the following sheets with their respective columns:
    * `Tenants`: `Room Number`, `Rental Price`, `Negotiated Price`, `Current Tenant Name`, `Tenant Email`, `Move-In Date`, `Security Deposit Paid`, `Room Status`, `Last Payment Date`, `Payment Status - Current Month`, `Move-Out Date (Planned)`, `Notes`
    * `Budget`: `Date`, `Type`, `Description`, `Amount`, `Category`
2.  **Open Apps Script:** Go to `Extensions > Apps Script` from your Google Sheet.
3.  **Copy & Paste Code:** Copy the Apps Script code (from the `Code.gs` file in this repository) into the Apps Script editor.
4.  **Save Project:** Save the Apps Script project.
5.  **Authorize Script:** Run any function from the `Parsonage Tools` custom menu in the Google Sheet (e.g., `Send Rent Reminders (Test)`). You will be prompted to authorize the script's permissions. Grant the necessary permissions.
6.  **Set Up Google Forms:**
    * Create a Google Form for "Tenant Application" and link its responses to a new tab in your Google Sheet.
    * Create a Google Form for "Move-Out Request" and link its responses to another new tab in your Google Sheet.
7.  **Configure Triggers:** (Detailed instructions will follow for specific automations like daily payment checks, form submission triggers).

## Usage

Once set up, the Google Sheet will be your primary interface.
* Manually update tenant information and payment statuses as needed.
* Use the "Parsonage Tools" custom menu for quick actions.
* Monitor the `Budget` sheet for financial insights.
* Incoming applications and move-out requests will automatically populate from Google Forms.
