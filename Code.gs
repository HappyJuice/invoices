function generateInvoiceOnCheck(e) {
  Logger.log("Loaded");
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // Only run on the correct sheet
  if (sheet.getName() !== "Invoices") return;

  const col = range.getColumn();
  const row = range.getRow();

  // Assuming checkbox is in column M (column 13)
  if (col !== 13 || range.getValue() !== true) return;

  const invoiceNumber = sheet.getRange(row, 1).getValue();     // A
  const clientName    = sheet.getRange(row, 2).getValue();     // B
  const clientAddress = sheet.getRange(row, 3).getValue();     // C (comma-separated)
  const clientVATID   = sheet.getRange(row, 4).getValue();     // D
  const issueDate     = sheet.getRange(row, 5).getValue();     // E
  const dueDate       = sheet.getRange(row, 6).getValue();     // F
  const purchaseOrder = sheet.getRange(row, 7).getValue();     // G
  const products      = sheet.getRange(row, 8).getValue();     // H (newline-separated)
  const quantities    = sheet.getRange(row, 9).getValue();     // I (newline-separated)
  const unitPrices    = sheet.getRange(row, 10).getValue();    // J (newline-separated)
  const total         = sheet.getRange(row, 11).getValue();    // K
  const currency      = sheet.getRange(row, 12).getValue();    // L

  Logger.log("Values extracted");

  // Parse multi-line strings
  const productList = products.toString().split('\n');
  const quantityList = quantities.toString().split('\n').map(Number);
  const priceList = unitPrices.toString().split('\n').map(Number);

  let itemsHtml = '';

  for (let i = 0; i < productList.length; i++) {
    const qty = quantityList[i] || 0;
    const price = priceList[i] || 0;
    const itemTotal = qty * price;
    itemsHtml += `<tr>
      <td>${productList[i]}</td>
      <td>${qty}</td>
      <td>${price.toFixed(2)}</td>
      <td>${itemTotal.toFixed(2)}</td>
    </tr>`;
  }

  // 3. Create invoice HTML
  const invoice = {
    number: invoiceNumber,
    issue_date: Utilities.formatDate(issueDate, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    due_date: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    purchaseOrder: (purchaseOrder !== "") ? `<div><strong>Purchase Order:</strong> ${purchaseOrder}</div>` : "",
    total: total.toFixed(2),
    currency: currency
  };

  const customer = {
    name: clientName,
    address: clientAddress.toString().split("\n").join("<br>"),
    vat: clientVATID.toString().split("\n").join("<br>")
  };

  // Personal data

  const personal_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal Data");
  const personal_name = personal_sheet.getRange(2, 2).getValue();
  const personal_address = personal_sheet.getRange(3, 2).getValue();
  const personal_email = personal_sheet.getRange(4, 2).getValue();
  const personal_vat = personal_sheet.getRange(5, 2).getValue();
  const bank_name = personal_sheet.getRange(2, 5).getValue();
  const bank_iban = personal_sheet.getRange(3, 5).getValue();
  const bank_swift = personal_sheet.getRange(4, 5).getValue();
  const bank_owner = personal_sheet.getRange(5, 5).getValue();

  const supplier = {
    name: personal_name,
    address: personal_address.toString().split("\n").join("<br>"),
    email: personal_email,
    vat: personal_vat.toString().split("\n").join("<br>")
  };

  const bank = {
    name: bank_name,
    iban: bank_iban,
    swift: bank_swift,
    owner: bank_owner
  };

  Logger.log("Template values ready");

  const template = HtmlService.createTemplateFromFile("invoice-template");

  // Handling data to template
  template.invoice = invoice;
  template.supplier = supplier;
  template.customer = customer;
  template.items = itemsHtml;
  template.bank = bank;

  // HTML creation
  const htmlOutput = template.evaluate().getContent();

  Logger.log("html ready");

  // Convert HTML do PDF
  const blob = Utilities.newBlob(htmlOutput, "text/html", "invoice.html");
  const pdf = blob.getAs(MimeType.PDF);

  // Save to Drive
  const folder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  const fileName = `invoice_${invoice.number}.pdf`;
  const file = folder.createFile(pdf.setName(fileName));

  Logger.log("File saved");

  sheet.getRange(row, 13).setValue(false);
  sheet.getRange(row, 14).setValue(file.getUrl());
  Logger.log("Job finished");

  SpreadsheetApp.getUi().alert("Invoice created.");
  Logger.log("Alert closed");
}
