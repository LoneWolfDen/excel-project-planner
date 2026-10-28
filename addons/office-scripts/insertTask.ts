// Insert Task (B-705): a new Tasks row below the selected one, at its level, with Settings > NextFreeUID, as
// Start > Inserting tasks describes. Inserting cells inside the table grows it, so Excel fills the formula columns.
function main(workbook: ExcelScript.Workbook) {
  const tasks = workbook.getTable("Tasks");
  const body = tasks.getRangeBetweenHeaderAndTotal();
  const at = workbook.getActiveCell().getRowIndex() - body.getRowIndex();       // 0 = the first task
  if (at < 0 || at >= body.getRowCount()) {
    throw new Error("Select a task on the Tasks sheet first.");
  }
  const level = tasks.getColumnByName("Level").getRangeBetweenHeaderAndTotal().getCell(at, 0).getValue();
  body.getRow(at + 1).insert(ExcelScript.InsertShiftDirection.down);
  const column = (name: string) => tasks.getColumnByName(name).getRangeBetweenHeaderAndTotal().getCell(at + 1, 0);
  column("UID").setValue(workbook.getNamedItem("NextFreeUID").getRange().getValue());
  column("Level").setValue(level);
  column("Task Name").select();
}
