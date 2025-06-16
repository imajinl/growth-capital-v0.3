function updateScoringMatchIndexTags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');
  const projectsSheet = ss.getSheetByName('ProjectsView');
  const investorSheet = ss.getSheetByName('InvestorView');
  const targetSheet = ss.getSheetByName('ScoringMatchIndexTags');

  const selectedTags = String(inputSheet.getRange('D12').getValue())
    .split(',')
    .map(t => t.trim())
    .filter(Boolean);
  if (selectedTags.length === 0) return;

  const projectTags = projectsSheet.getRange('H2:H').getValues().map(row => row[0]);
  const projectNames = projectsSheet.getRange('B2:B').getValues().map(row => row[0]);

  const investorProjects = investorSheet.getRange('AD2:AD').getValues().map(row => row[0]);
  const investorNames = investorSheet.getRange('B2:B').getValues().map(row => row[0]);

  const tagData = {};

  selectedTags.forEach(tag => {
    const tagCounts = {};

    for (let i = 0; i < projectTags.length; i++) {
      const tagCell = String(projectTags[i] || "").toLowerCase();
      if (!tagCell.includes(tag.toLowerCase())) continue;

      const project = String(projectNames[i] || "").trim();
      if (!project) continue;

      for (let j = 0; j < investorProjects.length; j++) {
        const invProjects = String(investorProjects[j] || "").toLowerCase();
        const investor = String(investorNames[j] || "").trim();
        if (!investor) continue;

        if (invProjects.includes(project.toLowerCase())) {
          tagCounts[investor] = (tagCounts[investor] || 0) + 1;
        }
      }
    }

    const sorted = Object.entries(tagCounts).sort((a, b) => b[1] - a[1]);
    tagData[tag] = sorted;
  });

  const headers = [];
  selectedTags.forEach(tag => {
    headers.push(`${tag} - rank`);
    headers.push(`${tag} - value`);
  });

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const maxRows = Math.max(...Object.values(tagData).map(arr => arr.length));
  const output = Array.from({ length: maxRows }, () => Array(headers.length).fill(""));

  selectedTags.forEach((tag, ti) => {
    const colRank = ti * 2;
    const colValue = colRank + 1;
    tagData[tag].forEach((entry, ri) => {
      output[ri][colRank] = entry[0];
      output[ri][colValue] = entry[1];
    });
  });

  if (output.length > 0) {
    targetSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
  }
}
