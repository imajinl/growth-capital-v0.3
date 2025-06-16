function updateScoringMatchIndexTotalFunding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');
  const projectsSheet = ss.getSheetByName('ProjectsView');
  const investorSheet = ss.getSheetByName('InvestorView');
  const targetSheet = ss.getSheetByName('ScoringMatchIndexTotalFunding');

  const rangeLabel = inputSheet.getRange('D18').getValue().toString().trim();
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

  const fundingValues = projectsSheet.getRange('G2:G').getValues().map(row => row[0]);
  const projectNames = projectsSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const investorProjects = investorSheet.getRange('AD2:AD').getValues().map(row => row[0] || "");
  const investorNames = investorSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const investorCounts = {};

  for (let i = 0; i < fundingValues.length; i++) {
    const val = fundingValues[i];
    if (typeof val !== 'number' || isNaN(val) || val < min || val > max) continue;

    const project = projectNames[i];
    if (!project) continue;

    for (let j = 0; j < investorProjects.length; j++) {
      if (investorProjects[j].toLowerCase().includes(project.toLowerCase())) {
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
