// Set Baseline (B-705, Excel for the web: Automate > New script, paste, Save). The same as the TakeSnapshot sheet's
// BASELINE steps: Tasks' BL Source columns pasted as values onto the Baseline columns, then TakeBaseline = No.
function main(workbook: ExcelScript.Workbook) {
  const source = workbook.getNamedItem("tsk_BLSource").getRange();
  const target = workbook.getNamedItem("tsk_Baseline").getRange();
  target.setValues(source.getValues());
  workbook.getNamedItem("TakeBaseline").getRange().setValue("No");
}
