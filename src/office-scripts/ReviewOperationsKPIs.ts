function main(workbook: ExcelScript.Workbook): {
  totalRecords: number,
  uniqueOrderIds: number,
  totalOrders: number,
  totalUnits: number,
  overallErrorRate: number,
  lateOrderRate: number,
  returnRate: number,
  ordersPerHour: number,
  exceptionCount: number,
  status: string
} {
  const sheet = workbook.getWorksheet("KPI_Summary");

  if (!sheet) {
    throw new Error("KPI_Summary worksheet was not found.");
  }

  const totalRecords = Number(sheet.getRange("B4").getValue());
  const uniqueOrderIds = Number(sheet.getRange("B5").getValue());
  const totalOrders = Number(sheet.getRange("B6").getValue());
  const totalUnits = Number(sheet.getRange("B7").getValue());

  const overallErrorRate = Number(sheet.getRange("B8").getValue());
  const lateOrderRate = Number(sheet.getRange("B9").getValue());
  const returnRate = Number(sheet.getRange("B10").getValue());

  const ordersPerHour = Number(sheet.getRange("B11").getValue());
  const exceptionCount = Number(sheet.getRange("B12").getValue());

  const status =
    exceptionCount > 0
      ? "Review Required"
      : "No Exceptions";

  // Clean formatting for cloud/reporting use.
  sheet.getRange("B8:B10").setNumberFormat("0.00%");
  sheet.getRange("B11").setNumberFormat("0.00");

  // Add cloud review information.
  sheet.getRange("A14").setValue("Cloud Review Status");
  sheet.getRange("B14").setValue(status);

  sheet.getRange("A15").setValue("Last Cloud Review");
  sheet.getRange("B15").setValue(new Date().toISOString());

  sheet.getRange("A14:A15").getFormat().getFont().setBold(true);
  sheet.getRange("A1:B15").getFormat().autofitColumns();

  return {
    totalRecords,
    uniqueOrderIds,
    totalOrders,
    totalUnits,
    overallErrorRate,
    lateOrderRate,
    returnRate,
    ordersPerHour,
    exceptionCount,
    status
  };
}
