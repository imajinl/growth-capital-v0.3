function updateScoringMatchIndexRound() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');
  const roundsSheet = ss.getSheetByName('RoundsView');
  const investorSheet = ss.getSheetByName('InvestorView');
  const targetSheet = ss.getSheetByName('ScoringMatchIndexRound');

  const selectedRound = inputSheet.getRange('D15').getValue().toString().trim();
  if (!selectedRound || selectedRound === 'N / A') return;

  const roundValues = roundsSheet.getRange('C2:C').getValues().map(row => row[0]?.toString().trim() || "");
  const projectNames = roundsSheet.getRange('B2:B').getValues().map(row => row[0]?.toString().trim() || "");

  const investorProjects = investorSheet.getRange('AD2:AD').getValues().map(row => row[0]?.toLowerCase() || "");
  const investorNames = investorSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const investorCounts = {};

  for (let i = 0; i < roundValues.length; i++) {
    if (roundValues[i].toLowerCase() !== selectedRound.toLowerCase()) continue;

    const project = projectNames[i];
    if (!project) continue;

    for (let j = 0; j < investorProjects.length; j++) {
      if (investorProjects[j].includes(project.toLowerCase())) {
        const investor = investorNames[j];
        if (!investor) continue;
        investorCounts[investor] = (investorCounts[investor] || 0) + 1;
      }
    }
  }

  const sorted = Object.entries(investorCounts).sort((a, b) => b[1] - a[1]);
  const headers = [`${selectedRound} - rank`, `${selectedRound} - value`];

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const output = sorted.map(row => [row[0], row[1]]);
  if (output.length > 0) {
    targetSheet.getRange(2, 1, output.length, 2).setValues(output);
  }
}
