function scoreInvestorsUniversalToNewTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matchSheet = ss.getSheetByName('ScoringMatchIndexUniversal');
  const engineSheet = ss.getSheetByName('ScoringEngine');
  const outputSheetName = 'ScoreInvestorsUV';

  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clearContents();
  }

  const headers = matchSheet.getRange('A1:AV1').getValues()[0];
  const data = matchSheet.getRange(2, 1, matchSheet.getLastRow() - 1, headers.length).getValues();
  const weights = engineSheet.getRange('J2:J25').getValues().flat();

  const valueColIndices = [];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase().includes('- value')) {
      valueColIndices.push(i);
    }
  }

  if (valueColIndices.length !== weights.length) {
    throw new Error(`Mismatch: Found ${valueColIndices.length} value columns but ${weights.length} weights in ScoringEngine.`);
  }

  const results = data.map(row => {
    const name = row[0];
    let score = 0;
    for (let j = 0; j < valueColIndices.length; j++) {
      const value = parseFloat(row[valueColIndices[j]]) || 0;
      score += value * weights[j];
    }
    return [name, score];
  });

  outputSheet.getRange(1, 1, 1, 2).setValues([['Investor Name', 'Total Score']]);
  outputSheet.getRange(2, 1, results.length, 2).setValues(results);
}
