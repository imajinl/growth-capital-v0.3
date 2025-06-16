function populateInvestorViewFromInvestRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('FetchInvestorInfo');
  const targetSheet = ss.getSheetByName('InvestorView');

  const investRangeColumn = sourceSheet.getRange("I2:I").getValues(); // invest_range column
  const outputRangeStartRow = 2;
  const outputRangeStartCol = 7; // Column G = 7

  const rangeMap = {
    "1-3M": 0,
    "3-5M": 3,
    "5-10M": 6,
    "10-20M": 9,
    "20-50M": 12,
    "50M+": 15
  };

  const output = [];

  for (let row = 0; row < investRangeColumn.length; row++) {
    const cellValue = investRangeColumn[row][0];
    const rowOutput = Array(18).fill("");

    if (cellValue) {
      try {
        const parsedArray = JSON.parse(cellValue);
        parsedArray.forEach(entry => {
          const baseIndex = rangeMap[entry.amount_range];
          if (baseIndex !== undefined) {
            rowOutput[baseIndex] = entry.invest_num || 0;
            rowOutput[baseIndex + 1] = entry.lead_invest_num || 0;
            rowOutput[baseIndex + 2] = entry.lead_not_invest_num || 0;
          }
        });
      } catch (e) {
        Logger.log(`Invalid JSON at row ${row + 2}`);
      }
    }

    output.push(rowOutput);
  }

  targetSheet.getRange(outputRangeStartRow, outputRangeStartCol, output.length, 18).setValues(output);
}
