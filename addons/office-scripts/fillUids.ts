// Fill missing UIDs (Excel for the web): every named task without a UID gets the next free one, top to bottom, as
// Check's "Named tasks without a UID" asks (rows pasted or inserted without one). The same as typing
// Settings > NextFreeUID into each, by hand.
function main(workbook: ExcelScript.Workbook) {
  const tasks = workbook.getTable("Tasks");
  const names = tasks.getColumnByName("Task Name").getRangeBetweenHeaderAndTotal().getValues();
  const uidRange = tasks.getColumnByName("UID").getRangeBetweenHeaderAndTotal();
  const uids = uidRange.getValues();
  let next = Math.max(0, ...uids.map(u => Number(u[0]) || 0)) + 1;
  names.forEach((n, r) => {
    if (String(n[0]).trim() !== "" && (uids[r][0] === "" || uids[r][0] === null)) {
      uids[r][0] = next++;
    }
  });
  uidRange.setValues(uids);
}
