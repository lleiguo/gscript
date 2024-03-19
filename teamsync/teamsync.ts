const SHEET_ID = "1HhHaRDxhngE5RrChY4vR03Zrea93oEDTXa3uixSfyp4";
const WORKER_COL_INDEX = 5;
const HIRE_DATE_COL_INDEX = 8;
const DOCUMENT_ID = "15WsCfgVlFL93pRhF1ehhc37PxX2SbxiKPm8tYHJfowA";

var Teams;
(function (Teams) {
    Teams[Teams["BackendPlatform"] = 0]= "Backend Platform";
    Teams[Teams["BluewaterMango"] = 1] = "Bluewater Mango";
    Teams[Teams["FrontendPlatform"] = 2] = "Frontend Platform";
    Teams[Teams["Orange"] = 3] = "Orange";
    Teams[Teams["SRE"] = 4] = "SRE";
})(Teams || (Teams = {}));

const firstSyncDate = new Date("2024-01-10T08:30:00.000-08:00");

const daysSinceFirstSync = Math.ceil(new Date().getTime() - firstSyncDate.getTime()) / (1000 * 60 * 60 * 24);
const nextSyncDate = new Date(firstSyncDate.getTime() + (14 * (Math.floor(daysSinceFirstSync / 14) + 1) * (1000 * 60 * 60 * 24)));
let nextFacilitator = Teams["SRE"];
const staffingSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("People - Staffing");

function onOpen() {
  var menu = DocumentApp.getUi().createMenu("Commands");
  menu.addItem("Send Agenda to Doc", 'sendAgendaToGoogleDoc');
  menu.addToUi();
}

function sendAgendaToGoogleDoc() {
  const syncAgenda = new SyncAgenda();
  syncAgenda.sync();

  const documentWriter = new DocumentWriter();
  documentWriter.writeAgenda(syncAgenda.getNextSyncDate(), syncAgenda.getNextFacilitator(), syncAgenda.getAnniversaries(), syncAgenda.getNewHires());
}

class SyncAgenda {
  constructor(private sheet: GoogleAppsScript.Spreadsheet = SpreadsheetApp.openById(SHEET_ID)) { }

  sync(): void {
    this.calculateNextSyncDateAndFacilitator();
    this.getAnniversariesAndNewHires();
  }

  getNextSyncDate(): Date {
    return this.calculateNextSyncDateAndFacilitator().nextSyncDate;
  }

  getNextFacilitator(): string {
    return this.calculateNextSyncDateAndFacilitator().nextFacilitator;
  }

  getAnniversaries(): string[] {
    return this.getAnniversariesAndNewHires().anniversaries;
  }

  getNewHires(): string[] {
    return this.getAnniversariesAndNewHires().newHires;
  }

  private calculateNextSyncDateAndFacilitator(): { nextSyncDate: Date; nextFacilitator: string } {
    const daysSinceFirstSync = Math.ceil((new Date().getTime() - firstSyncDate.getTime()) / (1000 * 60 * 60 * 24));
    const nextSyncDate = new Date(firstSyncDate.getTime() + (14 * (Math.floor(daysSinceFirstSync / 14) + 1) * (1000 * 60 * 60 * 24)));
    nextSyncDate.setHours(11, 15, 0);
    const nextFacilitator = Teams[(Math.floor(daysSinceFirstSync / 14) + 1) % 5];
    return { nextSyncDate, nextFacilitator };
  }

  private getAnniversariesAndNewHires(): { anniversaries: string[]; newHires: string[] } {
    const staffingPositions = this.sheet.getSheetByName("People - Staffing").getDataRange().getValues();
    const lastSyncDate = new Date(this.getNextSyncDate().getTime() - (14 * 1000 * 60 * 60 * 24));
    const anniversaries: string[] = [];
    const newHires: string[] = [];

    for (let s = 1; s < staffingPositions.length; s++) {
      const hireDate = new Date(staffingPositions[s][HIRE_DATE_COL_INDEX]);
      const formattedHireDate = Utilities.formatDate(hireDate, "GMT+1", "MMMM dd, yyyy");

      if (hireDate < this.getNextSyncDate() && hireDate >= lastSyncDate) {
        newHires.push(staffingPositions[s][WORKER_COL_INDEX] + " (" + formattedHireDate + ")");
      } else if (hireDate.getMonth() === this.getNextSyncDate().getMonth() && this.getNextSyncDate().getDate() < 15) {
        anniversaries.push(staffingPositions[s][WORKER_COL_INDEX] + " (" + formattedHireDate + ")");
      }
    }

    return { anniversaries, newHires };
  }
}

class DocumentWriter {
  private document: GoogleAppsScript.Document;

  constructor(documentId: string = DOCUMENT_ID) {
    this.document = DocumentApp.openById(documentId);
  }

  writeAgenda(nextSyncDate: Date, nextFacilitator: string, anniversaries: string[], newHires: string[]): void {
    this.insertDate(nextSyncDate);
    this.insertRecordingReminder();
    this.insertAgendaHeader(nextFacilitator);
    if (anniversaries.length > 0) {
      this.insertAgendaItems(anniversaries, "Anniversaries");
    }
    this.insertAgendaItems([], "Shoutouts");
    this.insertAgendaItems(["General Updates", "TAG/DPG Updates"], "Updates & Announcements");
    if (newHires.length > 0) {
      this.insertAgendaItems(newHires, "New Hire Intro");
    }
  }

  private insertDate(date: Date): void {
    const formattedDate = Utilities.formatDate(date, "GMT-5", "MMMM dd, yyyy'T'HH:mm:ss zzzz");
    const body = this.document.getBody();
    const dateElement = body.insertParagraph(body.getChildIndex(this.findTable(body)) + 1, formattedDate + "\n");
    dateElement.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  }

  private insertRecordingReminder(): void {
    const body = this.document.getBody();
    const recordingElement = body.insertParagraph(body.getChildIndex(this.findTable(body)) + 2, "Make sure you press record at the start of the meeting! \n").setHeading(DocumentApp.ParagraphHeading.NORMAL);
    const style = {};
    style[DocumentApp.Attribute.BOLD] = true;
    style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#FF0000";
    recordingElement.setAttributes(style);
  }

  private insertAgendaHeader(facilitator: string): void {
    const body = this.document.getBody();
    const agendaElement = body.insertParagraph(body.getChildIndex(this.findTable(body)) + 3, "Facilitator: " + facilitator + "\nAgenda: \n").setHeading(DocumentApp.ParagraphHeading.NORMAL);
  }

  private insertAgendaItems(items: string[], title: string): void {
    const body = this.document.getBody();
    const sectionElement = body.insertListItem(body.getChildIndex(this.findTable(body)) + 4, title);
    sectionElement.setNestingLevel(0);
    items.forEach((item) => {
      body.insertListItem(body.getChildIndex(sectionElement) + 1, item).setNestingLevel(1);
    });
  }

  private findTable(body: Body): GoogleAppsScript.Element {
    var hr = body.findElement(DocumentApp.ElementType.TABLE)
    return hr ? hr.getElement() : null;
  }
}

Announce who will be facilitating next