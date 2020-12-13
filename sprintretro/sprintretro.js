// Compiled using ts2gas 3.4.4 (TypeScript 3.7.2)
var exports = exports || {};
var module = module || { exports: exports };
var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sprint Predicability");
var baseURL = "https://hootsuite.atlassian.net/rest/greenhopper/latest/rapid/charts/sprintreport?rapidViewId=";
var sprintURL = "https://hootsuite.atlassian.net/rest/agile/1.0/board/";
var issueURL = "https://hootsuite.atlassian.net/rest/agile/1.0/sprint/";
var searchURL = "https://hootsuite.atlassian.net/rest/api/3/search?jql=filter=21017&type=epic&fields=key";
var username = "lei.guo@hootsuite.com";
var password = "aI8n1UXSGqwKRXb7XAJI69DF";
var encCred = Utilities.base64Encode(username + ":" + password);
var roadmap;
var teamBoardIds;
(function (teamBoardIds) {
    teamBoardIds[teamBoardIds["PCRE"] = 711] = "PCRE";
    teamBoardIds[teamBoardIds["PODM"] = 813] = "PODM";
    teamBoardIds[teamBoardIds["PBTD"] = 702] = "PBTD";
    teamBoardIds[teamBoardIds["POBS"] = 820] = "POBS";
})(teamBoardIds || (teamBoardIds = {}));
var fetchArgs = {
    contentType: "application/json",
    headers: { Authorization: "Basic " + encCred },
    muteHttpExceptions: false
};
function onOpen() {
    var menu = SpreadsheetApp.getUi().createMenu("Commands");
    menu.addItem("Update Sprint Predicability Sheet", 'update');
    menu.addToUi();
}
function update() {
    if (sourceSheet == null) {
        SpreadsheetApp.getUi().alert("No source sheet present!");
        return;
    }
    sourceSheet.deleteRows(2, sourceSheet.getLastRow()-1);
    getRoadmap();
    for (var boardId in teamBoardIds) {
        if (!isNaN(boardId)) {
            updateSprints(boardId);
        }
    }
}
function getRoadmap() {
    var httpResponse = UrlFetchApp.fetch(searchURL, fetchArgs);
    if (httpResponse) {
        var rspns = httpResponse.getResponseCode();
        if (rspns === 200) {
            var data = JSON.parse(httpResponse.getContentText());
            roadmap = data.issues;
        }
        else {
            SpreadsheetApp.getUi().alert("Jira Error: " + rspns);
        }
    }
}
function updateSprints(boardId) {
    var isLast = false;
    var lastIndex = 0;
    while (!isLast) {
        var httpResponse = UrlFetchApp.fetch(sprintURL + boardId + "/sprint?state=closed&maxResults=50&startAt=" + lastIndex, fetchArgs);
        if (httpResponse) {
            var rspns = httpResponse.getResponseCode();
            if (rspns === 200) {
                var data = JSON.parse(httpResponse.getContentText());
                var sprints = data.values;
                isLast = data.isLast;
                lastIndex = lastIndex + sprints.length;
                for (var i = 0; i < sprints.length; ++i) {
                    var sprintId = sprints[i].id;
                    var startDate = sprints[i].startDate;
                    if (Date.parse(startDate) > Date.parse("2019-01-01")) {
                        updateSprint(sourceSheet.getLastRow() + 1, sprintId, boardId);
                    }
                }
            }
            else {
                SpreadsheetApp.getUi().alert("Jira Error: " + rspns);
            }
        }
    }
}
function updateSprint(row, sprintId, boardId) {
    var httpResponse = UrlFetchApp.fetch(baseURL + boardId + "&sprintId=" + sprintId, fetchArgs);
    var data;
    if (httpResponse) {
        var rspns = httpResponse.getResponseCode();
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
        }
        else {
            SpreadsheetApp.getUi().alert("Jira Error: Unable to make requests to " + baseURL + ": " + rspns);
        }
    }
    //Update Issue counts
    httpResponse = UrlFetchApp.fetch(issueURL + sprintId + "/issue?jql=cf[11502] is not empty&fields=epic,resolution,resolutiondate", fetchArgs);
    if (httpResponse) {
        var rspns = httpResponse.getResponseCode();
        if (rspns === 200) {
            var issueData = JSON.parse(httpResponse.getContentText());
            var roadMapIssueCount = 0;
            var roadMapCompletion = 0;
            for (var i = 0; i < issueData.issues.length; i++) {
                for (var r = 0; r < roadmap.length; r++) {
                    if (roadmap[r].key == issueData.issues[i].fields.epic.key) {
                        roadMapIssueCount++;
                        if (issueData.issues[i].fields.resolution != null && issueData.issues[i].fields.resolution.name == "Done") {
                            var resolutionDate = new Date(issueData.issues[i].fields.resolutiondate);
                            var sprintStartDate = new Date(data.sprint.isoStartDate);
                            var sprintEndDate = new Date(data.sprint.isoEndDate);
                            if (resolutionDate >= sprintStartDate && resolutionDate <= sprintEndDate) {
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
        }
        else {
            SpreadsheetApp.getUi().alert("Jira Error: Unable to make requests to " + issueURL + ": " + rspns);
        }
    }
    SpreadsheetApp.flush();
}
