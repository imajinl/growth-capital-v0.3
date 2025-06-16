function updateScoringMatchIndexEcosystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');
  const projectsSheet = ss.getSheetByName('ProjectsView');
  const investorSheet = ss.getSheetByName('InvestorView');
  const targetSheet = ss.getSheetByName('ScoringMatchIndexEcosystem');

  const selectedEcosystems = inputSheet.getRange('D14').getValue().split(',').map(e => e.trim()).filter(Boolean);
  if (selectedEcosystems.length === 0) return;

  const projectEcosystems = projectsSheet.getRange('L2:L').getValues().map(row => row[0] || "");
  const projectNames = projectsSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const investorProjects = investorSheet.getRange('AD2:AD').getValues().map(row => row[0] || "");
  const investorNames = investorSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const ecosystemData = {};

  selectedEcosystems.forEach(ecosystem => {
    const ecosystemCounts = {};
    for (let i = 0; i < projectEcosystems.length; i++) {
      if (!projectEcosystems[i].toLowerCase().includes(ecosystem.toLowerCase())) continue;
      const project = projectNames[i];
      if (!project) continue;

      for (let j = 0; j < investorProjects.length; j++) {
        if (investorProjects[j].toLowerCase().includes(project.toLowerCase())) {
          const investor = investorNames[j];
          if (!investor) continue;
          ecosystemCounts[investor] = (ecosystemCounts[investor] || 0) + 1;
        }
      }
    }

    const sorted = Object.entries(ecosystemCounts).sort((a, b) => b[1] - a[1]);
    ecosystemData[ecosystem] = sorted;
  });

  const headers = [];
  selectedEcosystems.forEach(e => {
    headers.push(`${e} - rank`);
    headers.push(`${e} - value`);
  });

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const maxRows = Math.max(...Object.values(ecosystemData).map(arr => arr.length));
  const output = Array.from({ length: maxRows }, () => Array(headers.length).fill(""));

  selectedEcosystems.forEach((ecosystem, ei) => {
    const colRank = ei * 2;
    const colValue = colRank + 1;
    ecosystemData[ecosystem].forEach((entry, ri) => {
      output[ri][colRank] = entry[0];
      output[ri][colValue] = entry[1];
    });
  });

  targetSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}
