// Indent / outdent (B-705): the selected tasks one level deeper (by = 1) or shallower (by = -1). Level is an input;
// the outline, WBS and roll-ups follow.
function main(workbook: ExcelScript.Workbook, by: number = 1) {
  const tasks = workbook.getTable("Tasks");
  const levels = tasks.getColumnByName("Level").getRangeBetweenHeaderAndTotal();
  const selected = workbook.getSelectedRange();
  const first = selected.getRowIndex() - levels.getRowIndex();
  for (let k = 0; k < selected.getRowCount(); k++) {
    const r = first + k;
    if (r < 0 || r >= levels.getRowCount()) continue;
    const cell = levels.getCell(r, 0);
    cell.setValue(Math.max(1, Number(cell.getValue()) + by));
  }
}
