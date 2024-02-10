// Compiled using ts2gas 3.6.3 (TypeScript 3.9.7)
var Teams;
(function (Teams) {
    Teams[Teams["Backend Platform"] = 0] = "Backend Platform";
    Teams[Teams["Bluewater Mango"] = 1] = "Bluewater Mango";
    Teams[Teams["Frontend Platform"] = 2] = "Frontend Platform";
    Teams[Teams["Orange"] = 3] = "Orange";
    Teams[Teams["SRE"] = 4] = "SRE";
})(Teams || (Teams = {}));
var anniversaries = [];
var newHires = [];
var firstSyncDate = new Date("2024-01-11T08:00:00.000-08:00");
var daysSinceFirstSync = Math.ceil(new Date().getTime() - firstSyncDate.getTime()) / (1000 * 60 * 60 * 24);
var nextSyncDate = new Date(firstSyncDate.getTime() + (14 * (Math.floor(daysSinceFirstSync / 14) + 1) * (1000 * 60 * 60 * 24)));
var nextFacilitator = Teams[Teams.SRE];
var staffingSheet = SpreadsheetApp.openById("1HhHaRDxhngE5RrChY4vR03Zrea93oEDTXa3uixSfyp4").getSheetByName("People - Staffing");
var workerColIndex = 5;
var hireDateColIndex = 8;
function onOpen() {
    var menu = DocumentApp.getUi().createMenu("Commands");
    menu.addItem("Send Agenda to Doc", 'sendAgendaToGoogleDoc');
    menu.addToUi();
}
function getNextSyncDateAndFacilitator() {
    nextSyncDate.setHours(11, 0, 0);
    Logger.log(Utilities.formatDate(nextSyncDate, "GMT-5", "MMMM dd, yyyy'T'HH:mm:ss zzzz"));
    nextFacilitator = Teams[(Math.floor(daysSinceFirstSync / 14) + 1) % 5];
    Logger.log(nextFacilitator);
}
function getAnniversariesAndNewHires() {
    var staffingPositions = staffingSheet.getDataRange().getValues();
    var lastSyncDate = new Date(nextSyncDate.getTime() - (14 * 1000 * 60 * 60 * 24));
    Logger.log("last sync date: " + lastSyncDate);
    for (var s = 1; s < staffingPositions.length; ++s) {
        var hireDate = new Date(staffingPositions[s][hireDateColIndex]);
        var formattedHireDate = Utilities.formatDate(hireDate, "GMT+1", "MMMM dd, yyyy");
        Logger.log("hire date: " + hireDate);
        if (hireDate < nextSyncDate && hireDate >= lastSyncDate) {
            newHires.push(staffingPositions[s][workerColIndex] + " (" + formattedHireDate + ")");
        }
        else if (hireDate.getMonth() == nextSyncDate.getMonth() && nextSyncDate.getDate() < 15) {
            //Anniversaries is only announced at the first sync of the month
            anniversaries.push(staffingPositions[s][workerColIndex] + " (" + formattedHireDate + ")");
        }
    }
}
function syncAgenda() {
    getNextSyncDateAndFacilitator();
    getAnniversariesAndNewHires();
}
function sendAgendaToGoogleDoc() {
    syncAgenda();
    var doc = DocumentApp.openById("15WsCfgVlFL93pRhF1ehhc37PxX2SbxiKPm8tYHJfowA");
    var body = doc.getBody();
    var hr = body.findElement(DocumentApp.ElementType.TABLE);
    var date = Utilities.formatDate(nextSyncDate, "GMT-5", "MMMM dd, yyyy'T'HH:mm:ss zzzz");
    var dateElement = body.insertParagraph(body.getChildIndex(hr.getElement()) + 1, date + "\n");
    dateElement.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    var nextIndex = body.getChildIndex(dateElement) + 1;
    var recordingElement = body.insertParagraph(nextIndex, "Make sure you press record at the start of the meeting! \n").setHeading(DocumentApp.ParagraphHeading.NORMAL);
    var style = {};
    style[DocumentApp.Attribute.BOLD] = true;
    style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#FF0000";
    recordingElement.setAttributes(style);
    nextIndex = body.getChildIndex(recordingElement);
    var agendaElement = body.insertParagraph(nextIndex + 1, "Facilitator: " + nextFacilitator + "\nAgenda: \n").setHeading(DocumentApp.ParagraphHeading.NORMAL);
    nextIndex = body.getChildIndex(agendaElement);
    var record = body.insertListItem(nextIndex + 1, "Press record!");
    nextIndex = body.getChildIndex(record) + 1;
    if (newHires.length > 0) {
        var newHire = body.insertListItem(nextIndex, "New Hire Intro");
        newHires.forEach(function (value, index) {
            body.insertListItem(body.getChildIndex(newHire) + index + 1, value).setNestingLevel(1);
        });
        newHire.setNestingLevel(0);
        nextIndex = body.getChildIndex(newHire) + newHires.length + 1;
    }
    var tagdpg = body.insertListItem(nextIndex, "TAG/DPG Updates");
    tagdpg.setNestingLevel(0);
    nextIndex = body.getChildIndex(tagdpg) + 1;
    var announcement = body.insertListItem(nextIndex, "Announcement");
    announcement.setNestingLevel(0);
    nextIndex = body.getChildIndex(announcement) + 1;
    var shoutout = body.insertListItem(nextIndex, "Shout-out");
    shoutout.setNestingLevel(0);
    nextIndex = body.getChildIndex(shoutout) + 1;
    if (anniversaries.length > 0) {
        var anniversary = body.insertListItem(nextIndex, "Anniversaries");
        anniversaries.forEach(function (value, index) {
            body.insertListItem(body.getChildIndex(anniversary) + index + 1, value).setNestingLevel(1);
        });
        anniversary.setNestingLevel(0);
    }
}