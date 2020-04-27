let activeSheet = SpreadsheetApp.getActiveSpreadsheet();
let jiraSheet = activeSheet.getSheetByName("JIRA");
let jiraBaseUrl = "https://hootsuite.atlassian.net/rest/api/3/"
let jiraLink = '=HYPERLINK("https://hootsuite.atlassian.net/browse/';
let jiraUser = "lei.guo@hootsuite.com";
let jiraToken = "aI8n1UXSGqwKRXb7XAJI69DF";
let githubUser = "lei-guo"
let githubToken = "5e4bbddd81eaadb663a7b796a029dc9fdf61e95b";
const pullRequestUrl = (repo, pullid) => `https://github.hootops.com/api/v3/repos/${repo}/pulls/${pullid}`
let issueKey = "PBTD-1495"

let jiraFetchArgs = {
  contentType: "application/json",
  headers: { Authorization: "Basic " + Utilities.base64Encode(jiraUser + ":" + jiraToken) },
  muteHttpExceptions: false
};

let githubFetchArgs = {
  contentType: "application/json",
  headers: { Authorization: "Basic " + Utilities.base64Encode(githubUser + ":" + githubToken) }
};

let headers = [
  "Issue Key",
  "Issue Id",
  "PR Id",
  "PR Status",
  "PR Destination",
  "PR URL",
  "Git Repo",
  "Jenkins URL"
];

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

function onOpen() {
  let menuEntries = [{ name: "Update", functionName: "update" }];
  activeSheet.addMenu("Commands", menuEntries);
}

function update(index: number, jiraStatus: string) {
  jiraSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  jiraSheet.getRange(2, 1, 1, 1).setValue(issueKey);
  // let issueId = getIssueIdFromKey(issueKey);
  // jiraSheet.getRange(2, 2, 1, 1).setValue(issueId);
  // let pullRequests = getIssuePullRequests(issueKey);
  // let prId = pullRequests[0].id.replace('#', '');
  // jiraSheet.getRange(2, 3, 1, 1).setValue(pullRequests[0].id);
  // jiraSheet.getRange(2, 4, 1, 1).setValue(pullRequests[0].status);
  // jiraSheet.getRange(2, 5, 1, 1).setValue(pullRequests[0].destination.branch);
  // jiraSheet.getRange(2, 6, 1, 1).setValue(pullRequests[0].url);
  // let prURL = pullRequests[0].url
  // let repo = prURL.substring("https://github.hootops.com/".length, prURL.indexOf("pull")-1)
  // jiraSheet.getRange(2, 7, 1, 1).setValue(repo);
  let gitPRs = getGitPullRequest("pod/jenkins-kubernetes", "369")
  jiraSheet.getRange(2, 8, 1, 1).setValue(gitPRs[0]._links.statuses.href);
  SpreadsheetApp.flush();
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
//Unfortunately, JIRA pull request API only take an issue id instead of issue key
function getIssueIdFromKey(key: string) {
  let issueEndPoint = "issue/" + key + "?fields=id";
  let data = getData(jiraBaseUrl + issueEndPoint, jiraFetchArgs)
  return data.id
}

//Retrieve pull requests associated with a JIRA issue
function getIssuePullRequests(key: string) {
  let prEndPoint = "https://hootsuite.atlassian.net/rest/dev-status/1.0/issue/detail?applicationType=githube&dataType=pullrequest&fields=pullRequests&issueId="
  let data = getData(prEndPoint + getIssueIdFromKey(key), jiraFetchArgs)

  //Look for pull request that is "status==MERGED" and "destination.branch==master"
  //url looks like this "https://github.hootops.com/pod/docker-build-images/pull/38"
  return data.detail[0].pullRequests
}

//Retrieve pull request data from Github
function getGitPullRequest(repo: string, pullId: string) {
   return getData(pullRequestUrl(repo, pullId), githubFetchArgs) 
}
  
