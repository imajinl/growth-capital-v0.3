function scoreInvestorsNUV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheetName = 'ScoreInvestorsNUV';
  const engineSheet = ss.getSheetByName('ScoringEngine');
  const investorIndex = 0;

  const scoreSources = [
    { sheet: 'ScoringMatchIndexInvestStats', weightCell: 'J26' },
    { sheet: 'ScoringMatchIndexTags', weightCell: 'J27' },
    { sheet: 'ScoringMatchIndexEcosystem', weightCell: 'J28' },
    { sheet: 'ScoringMatchIndexTotalFunding', weightCell: 'J29' },
    { sheet: 'ScoringMatchIndexArea', weightCell: 'J30' },
    { sheet: 'ScoringMatchIndexValuation', weightCell: 'J31' },
    { sheet: 'ScoringMatchIndexRaise', weightCell: 'J32' },
    { sheet: 'ScoringMatchIndexRound', weightCell: 'J33' }
  ];

  let scoreMap = new Map();

  for (const source of scoreSources) {
    const tab = ss.getSheetByName(source.sheet);
    if (!tab) throw new Error(`Sheet "${source.sheet}" not found`);
    const values = tab.getDataRange().getValues();
    const headers = values[0];
    const weight = parseFloat(engineSheet.getRange(source.weightCell).getValue()) || 0;

    const valueCols = [];
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toLowerCase().includes('value')) valueCols.push(i);
    }

    for (let i = 1; i < values.length; i++) {
      const investor = values[i][investorIndex];
      if (!investor) continue;

      let score = 0;
      for (let col of valueCols) {
        const val = parseFloat(values[i][col]);
        if (!isNaN(val)) score += val;
      }

      const current = scoreMap.get(investor) || 0;
      scoreMap.set(investor, current + score * weight);
    }
  }

  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clearContents();
  }

  const output = [['Investor Name', 'Total Score']];
  for (const [investor, score] of scoreMap.entries()) {
    output.push([investor, score]);
  }

  outputSheet.getRange(1, 1, output.length, 2).setValues(output);
}
