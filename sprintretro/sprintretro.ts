let sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sprint Predicability");
let baseURL = "https://hootsuite.atlassian.net/rest/greenhopper/latest/rapid/charts/sprintreport?rapidViewId=";
let sprintURL = "https://hootsuite.atlassian.net/rest/agile/1.0/board/"
let issueURL = "https://hootsuite.atlassian.net/rest/agile/1.0/sprint/"
let searchURL = "https://hootsuite.atlassian.net/rest/api/3/search?jql=filter=21389&type=epic&fields=key"
let encCred = Utilities.base64Encode("lei.guo@hootsuite.com:aI8n1UXSGqwKRXb7XAJI69DF");
let roadmap: Array[string] ;

enum teamBoardIds {
  "PCRE" = 711,
  "PODM" = 813,
  "PBTD" = 702,
  "POBS" = 820,
  "IDENTITY" = 605,
  "DODGES" = 643,
  "ENGAGE - Red" = 577,
  "ENGAGE - Phoenix" = 775,
  "MOBILE - Artemis" = 751,
  "MOBILE - Apollo" = 752,
  "P+C - Load Lifter" = 497,
  "P+C - Yoda" = 827,
  "PLATFORM - Backend" = 591,
  "ANALYTICS" = 327,
  "IMPACT" = 558,
  "PRODUCT GROWTH - Accquisition" = 818
}

let fetchArgs = {
  contentType: "application/json",
  headers: { Authorization: "Basic " + encCred },
  muteHttpExceptions: false
};

function onOpen() {
    let menu = SpreadsheetApp.getUi().createMenu("Commands");
    menu.addItem("Update Sprint Predicability Sheet", 'update');
    menu.addToUi()
}

function update() {
  if (sourceSheet == null) {
    SpreadsheetApp.getUi().alert("No source sheet present!");
        return;
      }

    if (sourceSheet.getLastRow() >= 2) {
      sourceSheet.deleteRows(2, getLastRowSpecial(sourceSheet.getRange("A:A").getValues())-1)
    }  

    getRoadmap()

    for (let boardId in teamBoardIds) {
      if (!isNaN(boardId)) {
        updateSprints(boardId);
      }
   }
}

function getLastRowSpecial(range){
  let rowNum = 0;
  let blank = false;
  for(let row = 0; row < range.length; row++){
    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;
    }else if(range[row][0] !== ""){
      blank = false;
    }
  }
  return rowNum;
}

function getRoadmap() {
  let httpResponse = UrlFetchApp.fetch(searchURL, fetchArgs);
  if (httpResponse) {
    let rspns = httpResponse.getResponseCode();
    if (rspns === 200) {
      let data = JSON.parse(httpResponse.getContentText());
      roadmap = data.issues
    } else {
      SpreadsheetApp.getUi().alert(
          "Jira Error: " + rspns
      );
    }
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
            updateSprint(sourceSheet.getLastRow()+1, sprintId, boardId)
          }
        }
      } else {
        SpreadsheetApp.getUi().alert(
          "Jira Error: " + rspns
        );
      }
    }
  }
}

function updateSprint(row: number, sprintId: string, boardId: string) {
  let httpResponse = UrlFetchApp.fetch(baseURL+boardId+"&sprintId=" + sprintId, fetchArgs);
  var data
  if (httpResponse) {
    let rspns = httpResponse.getResponseCode();
    if (rspns === 200) {
      data = JSON.parse(httpResponse.getContentText());
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
      sourceSheet.getRange(row, 15, 1, 1).setValue(data.contents.completedIssuesEstimateSum.value);
    } else {
      SpreadsheetApp.getUi().alert(
        "Jira Error: Unable to make requests to " + baseURL + ": " + rspns
      );
    }
  }

  //Update Issue counts
  httpResponse = UrlFetchApp.fetch(issueURL + sprintId + "/issue?jql=cf[11502] is not empty&fields=epic,resolution,resolutiondate", fetchArgs);
  if (httpResponse) {
    let rspns = httpResponse.getResponseCode();
    if (rspns === 200) {
      let issueData = JSON.parse(httpResponse.getContentText());
      let roadMapIssueCount = 0;
      let roadMapCompletion = 0;
      for (let i=0; i<issueData.issues.length; i++) {
        for (let r = 0;  r < roadmap.length; r++) {
          if (roadmap[r].key == issueData.issues[i].fields.epic.key ) {
            roadMapIssueCount++;
            if (issueData.issues[i].fields.resolution != null && issueData.issues[i].fields.resolution.name == "Done" ) {
              let resolutionDate = new Date(issueData.issues[i].fields.resolutiondate)
              let sprintStartDate = new Date(data.sprint.isoStartDate)
              let sprintEndDate = new Date(data.sprint.isoEndDate)
              if (resolutionDate >= sprintStartDate && resolutionDate <= sprintEndDate ) {
                roadMapCompletion++;
              }
            }
            break;
          }
        }
        sourceSheet.getRange(row, 12, 1, 1).setValue(issueData.total);
        sourceSheet.getRange(row, 13, 1, 1).setValue(roadMapIssueCount);
        sourceSheet.getRange(row, 14, 1, 1).setValue(roadMapCompletion);
      }
    } else {
      SpreadsheetApp.getUi().alert(
          "Jira Error: Unable to make requests to " + issueURL + ": " + rspns
      );
    }
  }

  SpreadsheetApp.flush();
}
