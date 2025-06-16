const API_KEY = PropertiesService.getScriptProperties().getProperty('ROOTDATA_API_KEY');
const BASE_URL_INVESTORS = 'https://api.rootdata.com/open/get_invest';

function fetchInvestorInfoBatch() {
  const sheetName = 'FetchInvestorInfo';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  sheet.clearContents();
  sheet.appendRow([
    'invest_id', 'invest_name', 'type', 'logo', 'area', 'last_fac_date',
    'last_invest_num', 'invest_num', 'invest_range', 'description',
    'invest_overview', 'investments', 'establishment_date',
    'invest_stics', 'team_members'
  ]);

  const payload = {
    page: 1,
    page_size: 50
  };

  const response = UrlFetchApp.fetch(BASE_URL_INVESTORS, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: {
      'apikey': API_KEY
    },
    muteHttpExceptions: true
  });

  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const results = (json && json.data && Array.isArray(json.data.items)) ? json.data.items : [];

  for (const p of results) {
    sheet.appendRow([
      p.invest_id,
      p.invest_name,
      p.type,
      p.logo,
      (p.area || []).join('; '),
      p.last_fac_date,
      p.last_invest_num,
      p.invest_num,
      JSON.stringify(p.invest_range || []),
      p.description,
      JSON.stringify(p.invest_overview || {}),
      (p.investments || []).map(i => i.name).join('; '),
      p.establishment_date,
      (p.invest_stics || []).map(s => `${s.track}:${s.invest_num}`).join('; '),
      (p.team_members || []).map(t => `${t.name} (${t.position})`).join('; ')
    ]);
  }
} 