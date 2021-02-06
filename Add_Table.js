// Working add range
function main(workbook: ExcelScript.Workbook) {
  // Get the currently used range.
  var selectedSheet = workbook.getActiveWorksheet();
  var range = workbook.getActiveWorksheet().getUsedRange();
  var MEWSItems_table = selectedSheet.addTable(range, true);
}

// Working apply filter
function main(workbook: ExcelScript.Workbook) {
  let MEWSItems_table = workbook.getTable("MEWSItems_table");
  // Apply filter
  MEWSItems_table.getColumnByName("Accounting category")
    .getFilter()
    .applyCustomFilter("=*910 Adyen*");
}

// Working clear filter
function main(workbook: ExcelScript.Workbook) {
  let MEWSItems_table = workbook.getTable("MEWSItems_table");
  // Clear filter
  MEWSItems_table.clearFilters();
}

//Working sort

function main(workbook: ExcelScript.Workbook) {
  let MEWSItems_table = workbook.getTable("MEWSItems_table");
  // Sort on table: 'MEWSItems_table' column index: '11'
  MEWSItems_table.getSort()
    .apply([{
      key: 11,
      ascending: true
    }]); 
}

//Add consumed date value
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert at range M:M, move existing cells right
  selectedSheet.getRange("M:M")
    .insert(ExcelScript.InsertShiftDirection.right);
  // Set range M1:M2 on selectedSheet
  selectedSheet.getRange("M1:M2")
    .setValues([
      ["Consumed Date Value"],
      ["=value(text(L2,\"dd/mm/yyyy\"))"]
    ]);
  // Set number formats for Items!M:M
  selectedSheet.getRange("M:M")
    .setNumberFormatLocal("dd/mm/yyyy");
  // Set width of column(s) M1 to 141.75
  selectedSheet.getRange("M1")
    .getFormat()
    .setColumnWidth(141.75);
}
