function updateScoringMatchIndexInvestStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');
  const sourceSheet = ss.getSheetByName('InvestorView');
  const targetSheet = ss.getSheetByName('ScoringMatchIndexInvestStats');

  const selectedVerticals = inputSheet.getRange('D11').getValue().split(',').map(v => v.trim()).filter(Boolean);
  if (selectedVerticals.length === 0) return;

  const fundNames = sourceSheet.getRange('B2:B').getValues().flat();
  const verticalStats = sourceSheet.getRange('AF2:AF').getValues().flat();

  const verticalData = {};

  selectedVerticals.forEach(vertical => {
    verticalData[vertical] = [];
    for (let i = 0; i < fundNames.length; i++) {
      const cell = verticalStats[i];
      if (!cell || !fundNames[i]) continue;
      const match = cell.match(new RegExp(`${vertical}:\\s*(\\d+)`, 'i'));
      if (match) {
        verticalData[vertical].push({ name: fundNames[i], value: parseInt(match[1], 10) });
      }
    }
    verticalData[vertical].sort((a, b) => b.value - a.value);
  });

  const headers = [];
  selectedVerticals.forEach(v => {
    headers.push(`${v} - rank`);
    headers.push(`${v} - value`);
  });
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const maxRows = Math.max(...Object.values(verticalData).map(list => list.length));
  const output = Array.from({ length: maxRows }, () => Array(headers.length).fill(""));

  selectedVerticals.forEach((vertical, vi) => {
    const colRank = vi * 2;
    const colValue = colRank + 1;
    verticalData[vertical].forEach((entry, ri) => {
      output[ri][colRank] = entry.name;
      output[ri][colValue] = entry.value;
    });
  });

  targetSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}
