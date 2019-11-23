class Investment {
  name: string;
  id: string;
  portfolio: string;
  dependency: string;
  dependentPortfolio: string;
  dependencyType: string;
  category: string;
  productArea: string;

  //Construct investment from spreadsheet row
  constructor(investment: string[]) {
    this.name = investment[0];
    this.id = investment[1];
    this.category = investment[2];
    this.portfolio = investment[8];
    this.productArea = investment[9];
    this.dependencyType = investment[18];
    this.dependentPortfolio = investment[19];
    this.dependency = investment[20];
  }
}

const ss = SpreadsheetApp.getActiveSpreadsheet();
let dependencySheet: GoogleAppsScript.Spreadsheet.Sheet; //the sheet to write to
let portfolioName: string = "";

const destinationSheetHeaderRows: number = 12;
const portfolios: Array<String> = [
  "POD",
  "Product Growth",
  "Promote",
  "Platform",
  "Measure",
  "PIF",
  "Engage",
  "P+C",
  "InfoSec",
  "Software Dev",
  "Design",
  "Product Management",
  "Program Management",
  "Social Selling",
  "Ads"
];

enum investmentWeight {
  Unplanned = 0,
  "Stretch / H2 2020+" = 1,
  Target = 2,
  Committed = 3
}
enum investCategories {
  Unplanned = "Unplanned",
  "Stretch / H2 2020+" = "Stretch / H2 2020+",
  Target = "Target",
  Committed = "Committed"
}

enum dependencyTypes {
  HasDependency = "Has a Dependency",
  NoDependency = "No known dependency",
  Departmental = "Department Investment",
  IsDependened = "Is Depended On"
}

const severityColor = [
  { sev: 3, color: "#ff0000" },
  { sev: 2, color: "#ea9999" },
  { sev: 1, color: "#ffff00" },
  { sev: 0, color: "white" },
  { sev: -1, color: "white" },
  { sev: -2, color: "white" },
  { sev: -3, color: "white" }
];

let dependencyName: string,
  effort: string,
  dependencyType: string,
  investment: string,
  dependency: string,
  investmentId: string,
  category: string,
  portfolio: string,
  productArea: string;
let dependencySheetName = "Dependencies - Category Mismatch";
let allInvestments: Investment[];
let value: string | number[];

function onOpen() {
  // Create menu options
  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addSubMenu(
      ui.createMenu("Update Dependency").addItem("All", "updateDependency")
    )
    .addToUi();
}

function initialize() {
  //This will setup the proper column names
  let readmeSheet = ss
    .getSheets()[0]
    .getDataRange()
    .getValues();
  dependencyName = readmeSheet[44][0];
  effort = readmeSheet[42][0];
  dependencyType = readmeSheet[43][0];
  investment = readmeSheet[25][0];
  dependency = readmeSheet[45][0];
  investmentId = readmeSheet[26][0];
  category = readmeSheet[27][0];
  portfolio = readmeSheet[34][0];
  productArea = readmeSheet[35][0];

  let today = new Date();
  dependencySheetName =
    today.getMonth() + 1 + "/" + today.getDate() + "-" + dependencySheetName;
  dependencySheet = ss.getSheetByName(dependencySheetName); //the sheet to write to
}

function updateDependency() {
  initialize();
  if (dependencySheet != null) {
    ss.deleteSheet(dependencySheet);
  }
  dependencySheet = ss.insertSheet(dependencySheetName);

  let value = [
    portfolio,
    productArea,
    investmentId,
    investment,
    category,
    dependencyType,
    portfolio,
    dependency,
    investment,
    category,
    dependency,
    "Src",
    "Tgt",
    "Category Delta"
  ];
  dependencySheet.getRange(1, 1, 1, value.length).setValues([value]);
  dependencySheet.getRange(1, 1, 1, value.length).setFontWeight("bold");
  dependencySheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  portfolios.forEach(function(p: string) {
    let portfolioInvestments = ss
      .getSheetByName(p)
      .getDataRange()
      .getValues();
    for (
      let i = destinationSheetHeaderRows;
      i < portfolioInvestments.length;
      i++
    ) {
      if (
        portfolioInvestments[i][1] != undefined &&
        portfolioInvestments[i][1].length > 0
      ) {
        allInvestments[portfolioInvestments[i][1]] = new Investment(
          portfolioInvestments[i]
        );
      }
    }

    //Hide src/target column to avoid confusion
    dependencySheet.hideColumns(12, 2);
    SpreadsheetApp.flush();
  });

  let dependentInvestment = "",
    dependentCategory = "",
    dependentTldr = "";

  Object.keys(allInvestments).forEach(function(key) {
    //Any investment has a dependency
    let currentInvestment: Investment = allInvestments[key];
    if (
      currentInvestment.dependentPortfolio != undefined &&
      currentInvestment.dependentPortfolio.length > 0 &&
      currentInvestment.dependentPortfolio.toLowerCase() != "n/a" &&
      (currentInvestment.category == investCategories.Committed ||
        currentInvestment.category == investCategories.Target)
    ) {
      if (
        (portfolioName.length > 0 &&
          (currentInvestment.portfolio == portfolioName ||
            currentInvestment.dependentPortfolio == portfolioName)) ||
        portfolioName.length == 0
      ) {
        if (
          currentInvestment.dependency != undefined &&
          currentInvestment.dependency.split("\n").length > 1
        ) {
          currentInvestment.dependency
            .split("\n")
            .forEach(function(investmentid: string) {
              let patt1 = /[0-9]+/;
              if (investmentid.length > 1 && patt1.test(investmentid)) {
                writeInvestment(investmentid, currentInvestment);
              }
            });
        } else {
          writeInvestment(currentInvestment.dependency, currentInvestment);
        }
      }
    }
  });
}

function writeInvestment(investmentid: string, currentInvestment: Investment) {
  let dependentInvestment: string = "",
    dependentCategory: string = "",
    dependentTldr: string = "",
    dependentPortfolio: string = "";
  if (allInvestments[investmentid] != undefined) {
    dependentInvestment = allInvestments[investmentid].id;
    dependentCategory = allInvestments[investmentid].category;
    dependentTldr = allInvestments[investmentid].dependency;
    dependentPortfolio = allInvestments[investmentid].dependentPortfolio;
  }

  //Exclude matched dependencies and external dependencies
  if (
    currentInvestment.category == dependentCategory ||
    dependentPortfolio.indexOf("Data Central", 0) >= 0 ||
    dependentPortfolio.indexOf("Third Party", 0) >= 0 ||
    currentInvestment.dependentPortfolio.indexOf("Data Central", 0) >= 0 ||
    currentInvestment.dependentPortfolio.indexOf("Third Party", 0) >= 0
  ) {
    return;
  }
  let srcCategory: string = "",
    targetCategory: string = "";
  let delta: number = 0;
  let srcWeight: number = 0,
    targetWeight: number = 0;

  if (currentInvestment.dependencyType == dependencyTypes.HasDependency) {
    srcCategory = investCategories[currentInvestment.category];
    targetCategory = investCategories[dependentCategory];
  } else {
    srcCategory = investCategories[dependentCategory];
    targetCategory = investCategories[currentInvestment.category];
  }
  srcCategory = srcCategory != undefined ? srcCategory : "";
  targetCategory = targetCategory != undefined ? targetCategory : "";
  srcWeight =
    investmentWeight[srcCategory] == undefined
      ? 0
      : investmentWeight[srcCategory];
  targetWeight =
    investmentWeight[targetCategory] == undefined
      ? 0
      : investmentWeight[targetCategory];

  delta = srcWeight - targetWeight;
  if (
    delta > 0 ||
    (srcCategory.length == 0 && targetCategory.length == 0 && delta < 0)
  ) {
    let value = [
      currentInvestment.portfolio,
      currentInvestment.productArea,
      currentInvestment.id,
      currentInvestment.name,
      currentInvestment.category,
      currentInvestment.dependencyType,
      currentInvestment.dependentPortfolio,
      investmentid,
      dependentInvestment,
      dependentCategory,
      dependentTldr,
      srcCategory,
      targetCategory,
      delta
    ];
    dependencySheet
      .getRange(dependencySheet.getLastRow() + 1, 1, 1, value.length)
      .setValues([value]);
    let weiColor = severityColor.filter(function(wc) {
      return wc.sev == delta;
    });
    dependencySheet
      .getRange(dependencySheet.getLastRow(), value.length, 1, 1)
      .setBackground(weiColor[0].color);
    if (srcCategory == "") {
      dependencySheet
        .getRange(dependencySheet.getLastRow(), value.length - 2, 1, 1)
        .setBackground("red");
    }
    if (targetCategory == "") {
      dependencySheet
        .getRange(dependencySheet.getLastRow(), value.length - 1, 1, 1)
        .setBackground("red");
    }

    SpreadsheetApp.flush();
  }
}
