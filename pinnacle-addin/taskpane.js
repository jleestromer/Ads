/* global Excel */

Office.onReady(function() {
  OfficeRuntime.storage.setItem("buildingImageBase64","data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQMAAAAl21bKAAAABlBMVEUAAP8A/wD///+2deuCAAAACklEQVQI12NgYAAAAAMAASsJTYQAAAAASUVORK5CYII=");
  document.getElementById('overrideBtn').addEventListener('click', overrideExpenses);
});

async function run() {
  await Excel.run(async (context) => {
    const outputSheet = context.workbook.worksheets.getItem('Outputs');
    const cashFlowSheet = context.workbook.worksheets.getItem('Software Engineer Cash Flow');

    // Range containing operating expenses in red - adjust as needed
    const expensesRange = outputSheet.getRange('B5:B20');
    expensesRange.load(['values', 'format/*']);

    await context.sync();

    // Force calculation even in manual mode
    expensesRange.calculate();
    await context.sync();

    // Copy values to cash flow sheet, preserving formatting
    const destRange = cashFlowSheet.getRange('B5');
    destRange.getResizedRange(expensesRange.rowCount - 1, expensesRange.columnCount - 1)
      .copyFrom(expensesRange, Excel.RangeCopyType.all);

    await context.sync();

    // Bonus: create chart of operating expenses on Outputs sheet
    const chart = outputSheet.charts.add('ColumnClustered', expensesRange, 'Auto');
    chart.setPosition(outputSheet.getRange('D5'), outputSheet.getRange('L20'));

    // Insert building image if available
    try {
      const img = await OfficeRuntime.storage.getItem('buildingImageBase64');
      if (img) {
        outputSheet.shapes.addImage(img).left = chart.left + chart.width + 20;
      }
    } catch (e) {
      console.log('No image available', e);
    }

    await context.sync();
  });
}

async function overrideExpenses() {
  const val = parseFloat(document.getElementById('overrideInput').value);
  if (isNaN(val)) {
    return;
  }
  await Excel.run(async (context) => {
    const outputSheet = context.workbook.worksheets.getItem('Outputs');
    const overrideCell = outputSheet.getRange('B21'); // example total cell
    overrideCell.values = [[val]];
    await context.sync();
  });
}

