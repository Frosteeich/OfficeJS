function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getActiveWorksheet().getUsedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();
  let selectedSheet = workbook.getActiveWorksheet();
  const srange = 'A' + 1 + ':' + 'C' + rows;
  let table = selectedSheet.addTable(srange, true);

console.log(rows.toString())
console.log(cols.toString())
}
