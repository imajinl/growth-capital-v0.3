function updateScoringMatchIndexArea() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Forgd - Growth Capital - Key Insights');
  const investorSheet = ss.getSheetByName('InvestorView');
  const targetSheet = ss.getSheetByName('ScoringMatchIndexArea');

  const selectedRegions = inputSheet.getRange('D13').getValue().split(',').map(r => r.trim()).filter(Boolean);
  if (selectedRegions.length === 0) return;

  const areaValues = investorSheet.getRange('C2:C').getValues().map(row => row[0] || "");
  const investorNames = investorSheet.getRange('B2:B').getValues().map(row => row[0] || "");

  const regionData = {};

  selectedRegions.forEach(region => {
    const regionCounts = {};
    for (let i = 0; i < areaValues.length; i++) {
      const areaList = areaValues[i].split(',').map(s => s.trim().toLowerCase());
      if (areaList.includes(region.toLowerCase())) {
        const investor = investorNames[i];
        if (!investor) continue;
        regionCounts[investor] = (regionCounts[investor] || 0) + 1;
      }
    }
    const sorted = Object.entries(regionCounts).sort((a, b) => b[1] - a[1]);
    regionData[region] = sorted;
  });

  const headers = [];
  selectedRegions.forEach(region => {
    headers.push(`${region} - rank`);
    headers.push(`${region} - value`);
  });

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const maxRows = Math.max(...Object.values(regionData).map(arr => arr.length));
  const output = Array.from({ length: maxRows }, () => Array(headers.length).fill(""));

  selectedRegions.forEach((region, ri) => {
    const colRank = ri * 2;
    const colValue = colRank + 1;
    regionData[region].forEach((entry, index) => {
      output[index][colRank] = entry[0];
      output[index][colValue] = entry[1];
    });
  });

  targetSheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}
