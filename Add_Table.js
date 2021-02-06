// Working snippet
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Items");
    var ItemsTable = sheet.tables.add(WorkingArea)
    ItemsTable.name = "MEWS_Items";
}).catch(errorHandlerFunction);

// Working snippet

function main(workbook: ExcelScript.Workbook) {
  // Get the currently used range.
  var range = workbook.getActiveWorksheet().getUsedRange();
  var selectedSheet = workbook.getActiveWorksheet();
  var MEWSItems_table = selectedSheet.addTable(range, true);
  MEWSItems_table.setName("MEWSItems_table");
    MEWSItems_table.addColumn(null /*add columns to the end of the table*/, [
      ["Type of the Day"],
      ['=IF(OR((TEXT([Consumed], "dddd") = "Saturday"), (TEXT([Consumed],"dddd") = "Sunday")), "Weekend", "Weekday")'],
    ]);
}
