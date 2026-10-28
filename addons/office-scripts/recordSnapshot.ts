// Record Snapshot (B-705): this status date's numbers (TakeSnapshot > SNAPSHOT ROW) pasted as values into the first
// empty row of History's table, as TakeSnapshot > SNAPSHOT does by hand.
function main(workbook: ExcelScript.Workbook) {
  const values = workbook.getNamedItem("snap_Row").getRange().getValues();
  const history = workbook.getTable("StatusHistory");
  const body = history.getRangeBetweenHeaderAndTotal();
  const first = body.getValues().findIndex(row => row[0] === "" || row[0] === null);
  if (first < 0) {
    history.addRow(-1, values[0]);
  } else {
    body.getCell(first, 0).getResizedRange(0, values[0].length - 1).setValues(values);
  }
}
