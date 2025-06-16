function writeTopInvestorScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUV = ss.getSheetByName('ScoreInvestorsUV');
  const sheetNUV = ss.getSheetByName('ScoreInvestorsNUV');
  const targetSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');

  const dataUV = sheetUV.getRange(2, 1, sheetUV.getLastRow() - 1, 2).getValues();
  const dataNUV = sheetNUV.getRange(2, 1, sheetNUV.getLastRow() - 1, 2).getValues();

  const combined = [...dataUV, ...dataNUV].filter(row => row[0]);

  const scoreMap = new Map();
  for (const [name, score] of combined) {
    const numericScore = parseFloat(score) || 0;
    scoreMap.set(name, (scoreMap.get(name) || 0) + numericScore);
  }

  const sorted = [...scoreMap.entries()].sort((a, b) => b[1] - a[1]);
  const top = sorted.slice(0, 1000);

  const output = top.map(([name, score]) => [name, score]);

  targetSheet.getRange(35, 3, output.length, 2).setValues(output);
}