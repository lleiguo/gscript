let activeSheet = SpreadsheetApp.getActiveSpreadsheet();
let sourceSheet = activeSheet.getSheetByName("Team Sprints");
let baseURL = "https://hootsuite.atlassian.net/rest/greenhopper/latest/rapid/charts/sprintreport?rapidViewId=";
let sprintURL = "https://hootsuite.atlassian.net/rest/agile/1.0/board/"
let username = "lei.guo@hootsuite.com";
let password = "5vXCLN8IAEeVRupF5WuvC851";
let encCred = Utilities.base64Encode(username + ":" + password);

enum teamBoardIds {
  "PCRE" = 711,
  "PODM" = 813,
  "PBTD" = 702,
  "POBS" = 823
}

let fetchArgs = {
  contentType: "application/json",
  headers: { Authorization: "Basic " + encCred },
  muteHttpExceptions: true
};

function onOpen() {
  let menuEntries = [{ name: "Update Team Sprints", functionName: "update" }];
  activeSheet.addMenu("Commands", menuEntries);
}

function update() {
  if (sourceSheet == null) {
    SpreadsheetApp.getUi().alert("No source sheet present!");
    return;
  }

  activeSheet.deleteRows(2, activeSheet.getLastRow())
  for (let boardId in teamBoardIds) {
    updateSprints(boardId);
  }
}
function updateSprints(boardId: string) {
  let isLast = false
  let lastIndex = 0
  while (!isLast) {
    let httpResponse = UrlFetchApp.fetch(sprintURL+boardId+"/sprint?state=closed&maxResults=50&startAt="+lastIndex, fetchArgs);
    if (httpResponse) {
      let rspns = httpResponse.getResponseCode();
      if (rspns === 200) {
        let data = JSON.parse(httpResponse.getContentText());
        let sprints = data.values
        isLast = data.isLast
        lastIndex = lastIndex + sprints.length
        for (let i = 0; i < sprints.length; ++i) {
          let sprintId = sprints[i].id;
          let startDate = sprints[i].startDate
          
          if(Date.parse(startDate) > Date.parse("2019-01-01")) {
            updateSprint(activeSheet.getLastRow()+1, sprintId, boardId)
          }
        }
      } else {
        SpreadsheetApp.getUi().alert(
          "Jira Error: Unable to make requests to Jira!" + rspns
        );
      }
    }
  }
}


function updateSprint(row: number, sprintId: string, boardId: string) {
  let httpResponse = UrlFetchApp.fetch(baseURL+boardId+"&sprintId=" + sprintId, fetchArgs);
  if (httpResponse) {
    let rspns = httpResponse.getResponseCode();
    if (rspns === 200) {
      let data = JSON.parse(httpResponse.getContentText());
      sourceSheet.getRange(row, 1, 1, 1).setValue(boardId);
      sourceSheet.getRange(row, 2, 1, 1).setValue(teamBoardIds[boardId]);
      sourceSheet.getRange(row, 3, 1, 1).setValue(sprintId);
      sourceSheet.getRange(row, 4, 1, 1).setValue(data.sprint.name);
      sourceSheet.getRange(row, 5, 1, 1).setValue(data.contents.completedIssues.length);
      sourceSheet.getRange(row, 6, 1, 1).setValue(data.contents.issuesNotCompletedInCurrentSprint.length);
      sourceSheet.getRange(row, 7, 1, 1).setValue(Object.keys(data.contents.issueKeysAddedDuringSprint).length);
      sourceSheet.getRange(row, 8, 1, 1).setValue(data.contents.puntedIssues.length);
      sourceSheet.getRange(row, 9, 1, 1).setValue(data.contents.issuesCompletedInAnotherSprint.length);
      sourceSheet.getRange(row, 10, 1, 1).setValue(data.sprint.startDate);
      sourceSheet.getRange(row, 11, 1, 1).setValue(data.sprint.endDate);
    } else {
      SpreadsheetApp.getUi().alert(
        "Jira Error: Unable to make requests to Jira!" + rspns
      );
    }
  }
  SpreadsheetApp.flush();
}
