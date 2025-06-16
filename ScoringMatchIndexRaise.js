function updateScoringMatchIndexRaise() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');
  const roundsSheet = ss.getSheetByName('RoundsView');
  const investorSheet = ss.getSheetByName('InvestorView');
  const targetSheet = ss.getSheetByName('ScoringMatchIndexRaise');

  const rangeLabel = String(inputSheet.getRange('D16').getValue()).trim();
  if (!rangeLabel || rangeLabel === 'N / A') return;

  let min = 0;
  let max = Infinity;
  const rangeMatch = rangeLabel.match(/\$?([\d,]+)(?:\s*-\s*\$?([\d,]+))?/);

  if (rangeMatch) {
    min = parseInt(rangeMatch[1].replace(/,/g, ''), 10);
    if (rangeMatch[2]) {
      max = parseInt(rangeMatch[2].replace(/,/g, ''), 10);
    }
  }

  const raiseValues = roundsSheet.getRange('E2:E').getValues().map(row => row[0]);
  const projectNames = roundsSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const investorProjects = investorSheet.getRange('AD2:AD').getValues().map(row => row[0] || "");
  const investorNames = investorSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const investorCounts = {};

  for (let i = 0; i < raiseValues.length; i++) {
    const val = raiseValues[i];
    if (typeof val !== 'number' || isNaN(val) || val < min || val > max) continue;

    const project = String(projectNames[i] || "").toLowerCase();
    if (!project) continue;

    for (let j = 0; j < investorProjects.length; j++) {
      const investorProject = String(investorProjects[j] || "").toLowerCase();
      if (investorProject.includes(project)) {
        const investor = investorNames[j];
        if (!investor) continue;
        investorCounts[investor] = (investorCounts[investor] || 0) + 1;
      }
    }
  }

  const sorted = Object.entries(investorCounts).sort((a, b) => b[1] - a[1]);
  const headers = [`${rangeLabel} - rank`, `${rangeLabel} - value`];

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const output = sorted.map(row => [row[0], row[1]]);
  if (output.length > 0) {
    targetSheet.getRange(2, 1, output.length, 2).setValues(output);
  }
}