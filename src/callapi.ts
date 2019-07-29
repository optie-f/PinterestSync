import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import HTTPResponse = GoogleAppsScript.URL_Fetch.HTTPResponse;
import { TOKEN, USERBOARDS } from './config';

export class PinterestSync {
  static FIRSTROW = 3;
  static ss: Spreadsheet;
  static urls: string[] = [];

  static main(): void {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.setUrls();

    this.urls.forEach(url => {
      let result_json = this.tryHttpGet(url);
      if (!result_json) return;
      let board_id: string = result_json['data'][0]['board']['id'];
      let sheet: Sheet | null = this.ss.getSheetByName(board_id);
      if (sheet === null) sheet = this.setUpSheet(board_id);

      do {
        let data: Array<JSON> = result_json['data'];
        let row_ptr = sheet.getRange(1, 4).getValue();

        let rows = [];
        // data は新しい順に並んでいるはずなので, シートでも同様に記録しておく
        // そうするとシート先頭行とのだけの比較で差分がとれる
        for (let i = 0; i < data.length; i++) {
          let pin: JSON = data[i];
          let newest_pin_id = sheet.getRange(row_ptr, 1).getValue();
          if (pin['id'] == newest_pin_id) break;
          let newRow = [
            pin['id'],
            pin['created_at'],
            pin['image']['original']['width'],
            pin['image']['original']['height'],
            pin['image']['original']['url'],
          ];
          row_ptr = +1;
          rows.push(newRow);
        }

        if (rows) {
          sheet.insertRows(row_ptr, rows.length);
          sheet.getRange(row_ptr, 1, rows.length, 5).setValues(rows);
        }

        let next: string | null = result_json['page']['next'];
        if (next !== null) {
          sheet.getRange(1, 2).setValue(next);
          sheet.getRange(1, 4).setValue(row_ptr);
          result_json = this.tryHttpGet(next);
        } else {
          sheet.getRange(1, 2).setValue('');
          sheet.getRange(1, 4).setValue(this.FIRSTROW);
          return;
        }
      } while (result_json);
    });
  }

  static setUrls(): void {
    USERBOARDS.forEach(userBoard => {
      userBoard.boards.forEach(boardname => {
        let bname = userBoard.user + '/' + boardname;
        let url = this.createUrl(bname, TOKEN);
        this.urls.push(url);
      });
    });
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
    Logger.log(text);

    try {
      JSON.parse(text);
    } catch {
      Logger.log('JSON parse failed');
      return false;
    }

    let result_json: JSON = JSON.parse(text);
    Logger.log(result_json);
    if (res.getResponseCode() > 299 || !result_json['data']) return false;
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

    return sheet;
  }
}
