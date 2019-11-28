let activeSheet = SpreadsheetApp.getActiveSpreadsheet();
let sourceSheet = activeSheet.getSheetByName("POD");
let jiraLink = '=HYPERLINK("https://hootsuite.atlassian.net/browse/';
enum JIRASTATUS {
  Ready="Not Started",
  Backlog="Not Started",
  Triage="Not Started",
  Open="Not Started",
  "In Progress"="On Track",
  Impeded="At Risk",
  Done="Done",
  Closed="Done",
}

const destinationSheetHeaderRows = 18;

function onOpen() {
  let menuEntries = [{ name: "Update POD Status", functionName: "updatePOD" }];
  activeSheet.addMenu("Commands", menuEntries);
}

function updatePOD() {
  if (sourceSheet == null) {
    SpreadsheetApp.getUi().alert("No source sheet present!");
    return;
  }

  let rawData: string[][] = sourceSheet.getDataRange().getValues();
  for (let i = destinationSheetHeaderRows; i < rawData.length; ++i) {
    let jiraKey = rawData[i][1].substr(
      5,
      rawData[i][1].length
    );
    if (jiraKey != undefined && jiraKey.length > 4) {
      let jiraStatus = getIssueStatus(jiraKey);
      updateStatus(i, jiraStatus);
    }
  }
}

function updateStatus(index: number, jiraStatus: string) {
  sourceSheet.getRange(index+1, 33, 1, 1).setValue(JIRASTATUS[jiraStatus]);
  SpreadsheetApp.flush();
}

function getIssueStatus(key) {
  let baseURL = "https://hootsuite.atlassian.net/rest/api/3/issue/";
  let username = "lei.guo@hootsuite.com";
  let password = "xxxxxxxxxxxxxxxxxxxxxx";
  let encCred = Utilities.base64Encode(username + ":" + password);

  let fetchArgs = {
    contentType: "application/json",
    headers: { Authorization: "Basic " + encCred },
    muteHttpExceptions: false
  };

  let jql = key + "?fields=status";
  let httpResponse = UrlFetchApp.fetch(baseURL + jql, fetchArgs);
  if (httpResponse) {
    let rspns = httpResponse.getResponseCode();
    if (rspns === 200) {
      let data = JSON.parse(httpResponse.getContentText());
      return data.fields.status.name;
    } else {
      SpreadsheetApp.getUi().alert(
        "Jira Error: Unable to make requests to Jira!" + rspns
      );
    }
  } else {
    SpreadsheetApp.getUi().alert(
      "Jira Error: Unable to make requests to Jira!"
    );
  }
}
