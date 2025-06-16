function writeDynamicNarrativeToC22() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');

  const verticals = sheet.getRange('D11').getValue();
  const subverticals = sheet.getRange('D12').getValue();
  const regions = sheet.getRange('D13').getValue();
  const ecosystems = sheet.getRange('D14').getValue();
  const round = sheet.getRange('D15').getValue();
  const raiseRange = sheet.getRange('D16').getValue();
  const fdv = sheet.getRange('D17').getValue();
  const raised = sheet.getRange('D18').getValue();

  const verticalArr = verticals.split(',').map(e => e.trim());
  const subverticalArr = subverticals.split(',').map(e => e.trim());
  const ecosystemArr = ecosystems.split(',').map(e => e.trim());
  const raiseArr = raiseRange.split(',').map(e => e.trim());
  const fdvArr = fdv.split(',').map(e => e.trim());
  const raisedArr = raised.split(',').map(e => e.trim());

  const stats = aggregateByColumnPairs('ScoringMatchIndexInvestStats', verticalArr);
  const tags = aggregateByColumnPairs('ScoringMatchIndexTags', subverticalArr);
  const ecosystem = aggregateByColumnPairs('ScoringMatchIndexEcosystem', ecosystemArr);
  const raise = aggregateByColumnPairs('ScoringMatchIndexRaise', raiseArr);
  const valuation = aggregateByColumnPairs('ScoringMatchIndexValuation', fdvArr);
  const funding = aggregateByColumnPairs('ScoringMatchIndexTotalFunding', raisedArr);

  const narrative =
    "In the '" + verticals + "' vertical(s), the most prolific investors are " + stats.names + ", having made " + stats.values + " investments respectively in these vertical(s). " +
    "In terms of '" + subverticals + "' subvertical(s), the most prolific investors are " + tags.names + ", having made " + tags.values + " investments respectively in these subvertical(s). " +
    "The investors most active in '" + ecosystems + "' ecosystem(s) are " + ecosystem.names + ", having made " + ecosystem.values + " investments respectively in these ecosystem(s). " +
    "The investors who are most active in the '" + raiseRange + "' raise amount ranges are " + raise.names + ", having made " + raise.values + " investments in this range. " +
    "Furthermore, the investors most active in the '" + fdv + "' fundraise FDV are " + valuation.names + ", having made " + valuation.values + " investments in this range. " +
    "Lastly, the investors most actively investing in projects with '" + raised + "' capital raised to date are " + funding.names + ", having made " + funding.values + " investments in this range.";

  sheet.getRange('C22').setValue(narrative);
}

function aggregateByColumnPairs(sheetName, keywords) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const scoreMap = new Map();

  for (let col = 0; col < headers.length - 1; col++) {
    const rankHeader = headers[col];
    const valueHeader = headers[col + 1];
    if (!rankHeader || !valueHeader) continue;

    const rankMatch = keywords.some(keyword => rankHeader.toLowerCase().includes(keyword.toLowerCase()));
    const valueMatch = valueHeader.toLowerCase().includes('value');
    if (!rankMatch || !valueMatch) continue;

    for (let row = 1; row < data.length; row++) {
      const name = data[row][col];
      const value = parseFloat(data[row][col + 1]);
      if (name && !isNaN(value)) {
        const current = scoreMap.get(name) || 0;
        scoreMap.set(name, current + value);
      }
    }
  }

  const sorted = Array.from(scoreMap.entries()).sort((a, b) => b[1] - a[1]);
  const topN = sorted.slice(0, Math.min(5, sorted.length));
  const names = topN.map(e => e[0]).join(', ');
  const values = topN.map(e => Math.round(e[1])).join(', ');

  return { names, values };
}
