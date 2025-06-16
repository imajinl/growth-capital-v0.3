const API_KEY = PropertiesService.getScriptProperties().getProperty('ROOTDATA_API_KEY');
const BASE_URL = 'https://api.rootdata.com/open/get_item';

const projectIds = Array.from({ length: 50 }, (_, i) => 8000 + i);

function fetchProjectsBatch() {
  const sheetName = 'FetchProjects';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  sheet.clearContents();
  sheet.appendRow([
    'project_id', 'project_name', 'logo', 'token_symbol', 'establishment_date',
    'one_liner', 'description', 'active', 'total_funding', 'tags',
    'rootdataurl', 'investors', 'social_media', 'similar_project', 'ecosystem'
  ]);

  for (const id of projectIds) {
    const payload = {
      project_id: id,
      include_team: false,
      include_investors: true
    };

    const response = UrlFetchApp.fetch(BASE_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        'apikey': API_KEY,
        'language': 'en'
      },
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    const p = json.data;

    if (!p || !p.project_id) continue;

    sheet.appendRow([
      p.project_id,
      p.project_name,
      p.logo,
      p.token_symbol,
      p.establishment_date,
      p.one_liner,
      p.description,
      p.active,
      p.total_funding,
      (p.tags || []).join('; '),
      p.rootdataurl,
      (p.investors || []).map(i => i.name).join('; '),
      Object.values(p.social_media || {}).join('; '),
      (p.similar_project || []).join('; '),
      (p.ecosystem || []).join('; ')
    ]);

    Utilities.sleep(200);
  }
} 