import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import Folder = GoogleAppsScript.Drive.Folder;
import HTTPResponse = GoogleAppsScript.URL_Fetch.HTTPResponse;

export class PinterestSync {
  static FIRSTROW = 3;
  static ss: Spreadsheet;
  static mainSheet: Sheet;
  static url_row_id: number;
  static urls: string[] = [];
  static rootDir: Folder;
  static cancelled: boolean;

  static main(): void {
    this.cancelled = false;
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.setUrlsFromSheet();

    this.urls.forEach(url => {
      let result_json = this.tryHttpGet(url);
      if (!result_json) return;
      let board_id: string = result_json['data'][0]['board']['id'];
      let sheet: Sheet | null = this.ss.getSheetByName(board_id);
      if (sheet === null) sheet = this.setUpSheet(board_id);

      do {
        let data: Array<JSON> = result_json['data'];
        let row_ptr = sheet.getRange(1, 4).getValue();
        let row_ptr_current = row_ptr;

        let rows = [];
        // data は新しい順に並んでいるはずなので, シートでも同様に記録しておく
        // そうするとシート先頭行とのだけの比較で差分がとれる
        for (let i = 0; i < data.length; i++) {
          let pin: JSON = data[i];
          let newest_pin_id = sheet.getRange(row_ptr, 1).getValue();
          if (pin['id'] == newest_pin_id) break;
          if (pin['image']['original']['width'] == 0) continue;
          let newRow = [
            pin['id'],
            pin['created_at'],
            pin['image']['original']['width'],
            pin['image']['original']['height'],
            pin['image']['original']['url'],
          ];
          row_ptr += 1;
          rows.push(newRow);
          this.addFileToDrive(pin['image']['original']['url']);
        }

        if (rows) {
          sheet.insertRows(row_ptr_current, rows.length);
          sheet.getRange(row_ptr_current, 1, rows.length, 5).setValues(rows);
        }

        let next: string | null = result_json['page']['next'];
        if (next !== null) {
          sheet.getRange(1, 2).setValue(next);
          sheet.getRange(1, 4).setValue(row_ptr);
          result_json = this.tryHttpGet(next);
        } else {
          sheet.getRange(1, 2).setValue('');
          sheet.getRange(1, 4).setValue(this.FIRSTROW);
          this.mainSheet.getRange(this.url_row_id, 3).setValue(1);
          this.url_row_id += 1;
          return;
        }
      } while (result_json);
    });
    if (!this.cancelled) {
      this.mainSheet
        .getRange(4, 3, this.mainSheet.getLastRow() - 3, 1)
        .setValue('');
    }
  }

  static setUrlsFromSheet(): void {
    this.mainSheet = this.ss.getSheetByName('main');

    const dirId = this.mainSheet.getRange(1, 2).getValue();
    this.rootDir = DriveApp.getFolderById(dirId);

    const token = this.mainSheet.getRange(2, 2).getValue();
    const data = this.mainSheet
      .getRange(4, 1, this.mainSheet.getLastRow() - 3, 3)
      .getValues();
    this.url_row_id = 4;

    for (let i = 0; i < data.length; i++) {
      let checked = data[i][2];
      if (checked == '') {
        let user = data[i][0];
        let boardname = data[i][1];
        let url = this.createUrl(user + '/' + boardname, token);
        this.urls.push(url);
      } else {
        this.url_row_id += 1;
      }
    }
  }

  static createUrl(boardname: string, access_token: string): string {
    let url = 'https://api.pinterest.com/';
    url += 'v1/boards/' + boardname + '/pins/';
    url += '?access_token=' + access_token;
    url += '&limit=100&fields=';
    url += 'board,id,created_at,image';
    return url;
  }

  static tryHttpGet(url: string): JSON | boolean {
    let res: HTTPResponse = UrlFetchApp.fetch(url);
    let text = res.getContentText();

    try {
      JSON.parse(text);
    } catch {
      Logger.log('JSON parse failed');
      return false;
    }

    let result_json: JSON = JSON.parse(text);
    Logger.log(result_json);
    if (res.getResponseCode() > 299 || !result_json['data']) {
      Logger.log(text);
      this.cancelled = true;
      return false;
    }
    return result_json;
  }

  static setUpSheet(name: string): Sheet {
    const sheet = this.ss.insertSheet(name);
    // 1 行目にはAPI制限による中断後に再開するための情報を記録
    sheet.getRange(1, 1).setValue('next:');
    sheet.getRange(1, 2).setValue('');
    sheet.getRange(1, 3).setValue('last modified row:');
    sheet.getRange(1, 4).setValue(this.FIRSTROW);
    // 2 行目は header
    sheet.getRange(2, 1).setValue('id');
    sheet.getRange(2, 2).setValue('created_at');
    sheet.getRange(2, 3).setValue('width');
    sheet.getRange(2, 4).setValue('height');
    sheet.getRange(2, 5).setValue('url');

    sheet.setColumnWidths(1, 2, 140);
    sheet
      .getRange(this.FIRSTROW, 5, sheet.getMaxRows() - this.FIRSTROW, 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.setFrozenRows(this.FIRSTROW - 1);
    sheet.getRange(2, 1, 1, 5).setFontWeight('bold');
    sheet
      .getRange(
        2,
        1,
        sheet.getMaxRows() - this.FIRSTROW + 1,
        sheet.getMaxColumns()
      )
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    return sheet;
  }

  static addFileToDrive(url: string) {
    const urlrow: string[][] = this.mainSheet
      .getRange(this.url_row_id, 1, 1, 2)
      .getValues();
    const folderName = urlrow[0][0] + '_' + urlrow[0][1];

    if (!this.rootDir.getFoldersByName(folderName).hasNext()) {
      this.rootDir.createFolder(folderName);
    }
    const folderItr = this.rootDir.getFoldersByName(folderName);
    const folder = folderItr.next();
    const res = UrlFetchApp.fetch(url);
    folder.createFile(res);
  }
}
