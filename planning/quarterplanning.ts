let activeSheet = SpreadsheetApp.getActiveSpreadsheet();
let initiativeColumn = 4

//Project=Sheet
enum teamSheets {
  "POBS" = "POBS",
  "PBTD" = "PBTD",
  "PSRE" = "SRE",
  "PODM" = "AOPS",
  "PCP" = "PCP",
}

let jiraFetchArgs = {
  contentType: "application/json",
  headers: { Authorization: "Basic " + Utilities.base64Encode("lei.guo@hootsuite.com:aI8n1UXSGqwKRXb7XAJI69DF") },
  muteHttpExceptions: false
};

function onOpen() {
  let menuEntries = [{ name: "Update", functionName: "updateIssueForQuarter" }];
  activeSheet.addMenu("Commands", menuEntries);
}

//Retrieve issue key/summary/investment category
function updateIssueForQuarter() {

  for (let teamSheetName in teamSheets) {
    if (!isNaN(teamSheetName)) {
      let jiraSheet = activeSheet.getSheetByName(teamSheetName);
      let data = getData("https://hootsuite.atlassian.net/rest/api/3/search?jql=filter%3D21431%20and%20project%3D%22" + teamSheetName + "%22&fields=key,summary,customfield_14786", jiraFetchArgs)
      jiraSheet.getRange(2, 2, 1, 1).setValue("=HYPERLINK(\"https://hootsuite.atlassian.net/browse/" + data.key + data.fields.summary + "\")");
      SpreadsheetApp.flush();
    }
 }
}

function getData(endPoint: string, fetchArgs) {
  let httpResponse = UrlFetchApp.fetch(endPoint, fetchArgs);
  if (httpResponse) {
    let rspns = httpResponse.getResponseCode();
    if (rspns === 200) {
      return JSON.parse(httpResponse.getContentText());
    } else {
      SpreadsheetApp.getUi().alert(
        "Unable to make requests to " + endPoint + " with a response of:" + rspns
      );
    }
  } else {
    SpreadsheetApp.getUi().alert(
      "Unable to make requests to " + endPoint
      );
  }
}