function populateInvestorViewFromInvestOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('FetchInvestorInfo');
  const targetSheet = ss.getSheetByName('InvestorView');

  const investOverviewColumn = sourceSheet.getRange("K2:K").getValues(); // invest_overview
  const outputStartRow = 2;
  const outputStartCol = 26; // Column Z = 26

  const output = [];

  for (let row = 0; row < investOverviewColumn.length; row++) {
    const cell = investOverviewColumn[row][0];
    const rowOutput = ["", "", "", ""]; // Z to AC

    if (cell) {
      try {
        const json = JSON.parse(cell);
        rowOutput[0] = json.lead_invest_num || 0;
        rowOutput[1] = json.last_invest_round || 0;
        rowOutput[2] = json.his_invest_round || 0;
        rowOutput[3] = json.invest_num || 0;
      } catch (e) {
        Logger.log(`Invalid JSON in K${row + 2}`);
      }
    }

    output.push(rowOutput);
  }

  targetSheet.getRange(outputStartRow, outputStartCol, output.length, 4).setValues(output);
}
