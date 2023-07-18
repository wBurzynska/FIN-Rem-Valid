# Repository Name: SAP Unpaid Invoices Checker

---

### Table of Contents
You're sections headers will be used to reference location of destination.

- [Description](#description)
- [Features](#features)
- [How to Use](#how-to-use)
- [Other](#other)
- [Author Info](#author-info)

---

## Description

This repository contains VBA/VBS code that connects to SAP ERP R3 and verifies the status of unpaid invoices. The code automates the process of checking and retrieving information about invoices that have not been paid, providing a streamlined solution for monitoring outstanding payments.

## Features

- Connects to SAP ERP R3 using VBA/VBS code.
- Retrieves information about unpaid invoices.
- Verifies the status of invoices (paid or unpaid).
- Generates a report in Excel format, stored in the "Tracker" sheet.
- Handles different currencies.
- Provides error handling for troubleshooting.


## How to Use

#### Prerequisites

Before using the code in this repository, make sure you have the following:

1. Access to a SAP ERP R3 system.
2. Microsoft Excel (compatible with different versions, including MS 365 Excel).
3. SAP window opened and logged in.


#### Installation

1. Clone the repository to your local machine.
2. Open the Excel file containing the VBA code.
3. Enable the Developer tab in Excel (if not already enabled).
4. Open the VBA editor by clicking on "Developer" → "Visual Basic".
5. In the VBA editor, go to "File" → "Import File" and select the VBA/VBS code file from the repository.
6. Adjust the code to your specific requirements (e.g., SAP layouts, company codes, path to store reminders).
7. Save the changes and close the VBA editor.


#### Usage

1. Open the Excel file and navigate to the sheet where you want to run the code.
2. Ensure that the SAP window is open and you are logged in.
3. Press the assigned shortcut key or click on the assigned button to execute the code.
4. The code will connect to the SAP ERP R3 system, retrieve information about unpaid invoices from multiple vendors, and generate a report in the "Tracker" sheet.
5. Review the generated report to identify unpaid invoices, including details such as invoice numbers, amounts, and currency.

## Other

#### Contributing
Contributions are welcome! If you encounter any issues or have suggestions for improvements, please open an issue or submit a pull request to this repository.

#### Questions
If you have any further questions or need assistance, please feel free to reach out. Happy invoicing!

---

## Author Info

- LinkedIn - [@Weronika Burzynska](https://www.linkedin.com/in/laskaweronika/)

[Back To The Top](#read-me-template)
