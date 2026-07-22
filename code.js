function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('FMS Client Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMasterDashboardPayload() {
  const PROJECT_SHEET_ID = "1_yoPWOS_1icAxyOdCNhJQ-seONvJRmVAxN3P8uVh3Io"; 
  const FMS_SHEET_ID = "1OUgNpHfAlQlOy8PgCvT2bcd-3po98MKI2HMlwGZMAps";

  let payload = { clients: [], fmsRawData: [], error: null };

  try {
    const ssProject = SpreadsheetApp.openById(PROJECT_SHEET_ID);
    let pSheet = ssProject.getSheetByName("Project Details") || ssProject.getSheets()[0];
    const pData = pSheet.getDataRange().getDisplayValues();
    
    if (pData.length > 1) {
      const headers = pData[0].map(h => h.trim().toLowerCase().replace(/[\s_-]/g, ""));
      const rows = pData.slice(1);
      
      const getVal = (rowArr, keys) => {
        for (let k of keys) {
          let cleanKey = k.toLowerCase().replace(/[\s_-]/g, "");
          let idx = headers.indexOf(cleanKey);
          if (idx !== -1 && rowArr[idx]) return rowArr[idx].trim();
        }
        return "";
      };

      payload.clients = rows.map((row, idx) => {
        return {
          id: "CLT-" + String(idx + 1).padStart(3, "0"),
          name: getVal(row, ["clientcontactperson", "clientname", "name", "client"]),
          project: getVal(row, ["projectname", "project", "title"]),
          email: getVal(row, ["email", "mailidclient", "emailid", "mail"]),
          phone: getVal(row, ["contactnoclient", "phonenumber", "phone", "contact"]),
          startDate: getVal(row, ["startdate", "start"]),
          expectedDate: getVal(row, ["expecteddate", "expectedenddate", "expectedcompletiondate", "enddate"]),
          poDate: getVal(row, ["podate"]),
          poNumber: getVal(row, ["ponumber", "pono"]),
          poAmount: getVal(row, ["poamount", "totalpoamount", "amount", "value"]),
          logo: getVal(row, ["logo", "avatar"]),
          docs: getVal(row, ["documents", "docs", "momupload", "files"])
        };
      }).filter(c => c.name || c.project);
    }

    const ssFms = SpreadsheetApp.openById(FMS_SHEET_ID);
    let fSheet = ssFms.getSheetByName("FMS") || ssFms.getSheetByName("fms") || ssFms.getSheets()[0];
    payload.fmsRawData = fSheet.getDataRange().getDisplayValues();

  } catch(e) {
    payload.error = e.message;
  }
  return payload;
}

function saveStorageBackend(payloadString) {
  try {
    PropertiesService.getUserProperties().setProperty('fms_dashboard_storage', payloadString || '');
    return true;
  } catch(e) { throw new Error(e.message); }
}

function loadStorageBackend() {
  try {
    return PropertiesService.getUserProperties().getProperty('fms_dashboard_storage') || '{"pw":{},"msgs":{},"stats":{"links":0,"proj":0,"clients":0}}';
  } catch(e) { return '{"pw":{},"msgs":{},"stats":{"links":0,"proj":0,"clients":0}}'; }
}
