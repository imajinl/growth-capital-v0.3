const API_KEY = PropertiesService.getScriptProperties().getProperty('ROOTDATA_API_KEY');
const BASE_URL_FUNDRAISING = 'https://api.rootdata.com/open/get_fac';

function fetchFundraisingRoundsBatch() {
  const sheetName = 'FetchFundraisingRounds';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  sheet.clearContents();
  sheet.appendRow([
    'project_id', 'name', 'logo', 'rounds', 'published_time',
    'amount', 'valuation', 'invests', 'data_status', 'source_url'
  ]);

  const payload = {
    page: 1,
    page_size: 50
  };

  const response = UrlFetchApp.fetch(BASE_URL_FUNDRAISING, {
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

  for (const r of results) {
    sheet.appendRow([
      r.project_id,
      r.name,
      r.logo,
      r.rounds,
      r.published_time,
      r.amount,
      r.valuation,
      (r.invests || []).map(i => i.name).join('; '),
      r.data_status,
      r.source_url
    ]);
  }
} 