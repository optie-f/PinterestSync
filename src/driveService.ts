import Folder = GoogleAppsScript.Drive.Folder;
import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

export class SaveToDrive {
  static ss: Spreadsheet;
  static mainSheet: Sheet;
  static rootDir: Folder;

  static main(): void {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.mainSheet = this.ss.getSheetByName('main');
    const dirId = this.mainSheet.getRange(1, 2).getValue();
    this.rootDir = DriveApp.getFolderById(dirId);

    const sheets = this.ss.getSheets();

    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() == 'main') continue;
      this.saveImagesRecordedAtSheetOf(sheets[i]);
    }
  }

  static saveImagesRecordedAtSheetOf(sheet: Sheet): void {
    const dstFolder = this.createFolderIfNotYet(sheet);
    if (sheet.getLastRow() <= 2) return; // if no rows

    const data = sheet.getRange(3, 5, sheet.getLastRow() - 2, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      if (row[1] == '') {
        let res = UrlFetchApp.fetch(row[0]);
        dstFolder.createFile(res);
        sheet.getRange(3 + i, 6).setValue('1');
      }
    }
  }

  static createFolderIfNotYet(sheet: Sheet): Folder {
    let folderId = sheet.getRange(1, 6).getValue();
    let dstFolder: Folder;
    if (folderId == '') {
      dstFolder = this.rootDir.createFolder(sheet.getName());
      folderId = dstFolder.getId();
      sheet.getRange(1, 6).setValue(folderId);
    } else {
      dstFolder = DriveApp.getFolderById(folderId);
    }
    return dstFolder;
  }
}
