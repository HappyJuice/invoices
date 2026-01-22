# Invoices

I use this project to generate all my invoices for work I have done.

> This project was my first vibe-coding project since I can't code in JavaScript nor in HTML even though I can understand the written code.

## Spreadsheet setup

### Invoices sheet

The Invoices sheet consists of the following columns:
 - **Invoice Number** is my reference number for the invoice, it's generated automatically using this formula: `=IF(ISBLANK(B2); ""; YEAR(E2) & "_" & LEFT(UPPER(B2); 2) & "_" & TEXT(COUNTIFS(B$2:B2; B2; ARRAYFORMULA(YEAR(E$2:E2)); YEAR(E2)); "00"))` which generates `YYYY_CC_NN`, where **YYYY** is the year, **CC** are the first two letters of the client name, and **NN** is the order number of the invoice for the given partner in the given year.
 - **Client Name** is the name of the client, to be filled by the user
 - **Client Address** is the address of the client, to be filled by the user; multi-line input is expected (street, city, country)
 - **Client VAT ID** is the VAT ID of the client, to be filled by the user; it can be multi-line (eg. when using both international and national IDs)
 - **Issue Date** is the date when the invoice is issued, to be filled by the user
 - **Due Date** is the date by which the invoice should be paid by the client, by default 90 days from the issue date: `=IF(ISBLANK(E9);"";E9+90)`
 - **Purchase Order** is the reference number of the invoice provided by the client, if omitted, it won't appear on the invoice at all, to be filled by the user
 - **Items** is a list of items the invoice is for, each line is a new item, to be filled by the user
 - **Quantity** is the quantity of items from the previous column, number of lines must match, to be filled by the user
 - **Unit Price** is the price per item from the previous column, number of lines must match, to be filled by the user
 - **Total** is the total price, summed automatically using this formula: `=IFERROR(SUMPRODUCT(SPLIT(I9; CHAR(10)) * SPLIT(J9; CHAR(10))); "")`
 - **Currency** is the invoice currency, accepts any text input, to be filled by the user
 - **Generate** is a check-box; when ticked, the App Script will generate an invoice, paste the link to it into the **Invoice File** column, and untick the check-box
 - **Invoice File** is filled automatically after generating the invoice; it contains the URL to the generated invoice saved on Drive
 - **Paid** is just for the user's reference and is meant to contain the date when the client paid

### Personal Data sheet

The sheet contains user data in the following cells:
 - **B2** the user's full name
 - **B3** the user's address, should be multi-line in the format: street, city, country
 - **B4** the user's email address
 - **B5** the user's VAT ID, can be multi-line, eg. if the user wants to include both international and national IDs
 - **E2** the user's bank name
 - **E3** the user's bank account IBAN
 - **E4** the user's bank account SWIFT
 - **E5** the user's bank account name (usually same as **B2**)

## App Script setup

### Editor

Create files [Code.gs](Code.gs) and [invoice-template.html](invoice-template.html) and paste there the content of the files from this repository.

### Triggers

Create a new trigger with following settings:

**Choose which function to run:** generateInvoiceOnCheck

**Which runs at deployment:** Head

**Select event source:** From spreadsheet

**Select event type:** On edit
