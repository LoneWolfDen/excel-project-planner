// Keep a baseline (template 1.2, B-201): the current baseline's dates copied into Baseline 2 (or 3: change `which`)
// before re-baselining, as TakeSnapshot > KEEP A BASELINE does by hand.
function main(workbook: ExcelScript.Workbook, which: number = 2) {
  const source = workbook.getNamedItem("tsk_BaseDates").getRange();
  const target = workbook.getNamedItem(`tsk_Baseline${which}`).getRange();
  target.setValues(source.getValues());
}
