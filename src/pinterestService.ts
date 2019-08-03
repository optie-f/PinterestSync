import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import HTTPResponse = GoogleAppsScript.URL_Fetch.HTTPResponse;
import { parseSync } from '@babel/core';

export class RecordPinsData {
  static FIRSTROW = 3;
  static ss: Spreadsheet;
  static mainSheet: Sheet;
  static url_row_id: number;
  static urls: string[] = [];
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

      const prev = sheet.getRange(1, 2).getValue();
      if (prev != '') result_json = this.tryHttpGet(prev);

      let all_ids: string[];
      let id_set: Object = {};

      if (sheet.getLastRow() >= this.FIRSTROW) {
        all_ids = sheet
          .getRange(this.FIRSTROW, 1, sheet.getLastRow() - this.FIRSTROW + 1)
          .getValues()
          .map(x => x[0]);
        id_set = all_ids.reduce(function(result, item) {
          result[item] = item; // 疑似set
          return result;
        }, {});
      }

      do {
        let data: Array<JSON> = result_json['data'];

        let rows = [];
        // 新しく入るデータは常に 先頭行より新しい OR 最後行より古い のどちらか一方を仮定できる
        for (let i = 0; i < data.length; i++) {
          let pin: JSON = data[i];
          if (pin['image']['original']['width'] == 0) continue; // unavailableなやつ
          if (Object.prototype.hasOwnProperty.call(id_set, pin['id'])) continue;
          let newRow = [
            pin['id'],
            pin['created_at'],
            pin['image']['original']['width'],
            pin['image']['original']['height'],
            pin['image']['original']['url'],
          ];
          rows.push(newRow);
        }

        if (rows && rows.length >= 1) {
          if (
            sheet.getRange(this.FIRSTROW, 2).getValue() < data[0]['created_at']
          ) {
            sheet.insertRows(this.FIRSTROW, rows.length);
            sheet.getRange(this.FIRSTROW, 1, rows.length, 5).setValues(rows);
            Logger.log('top insertion: ' + rows.length + ' rows');
          } else {
            sheet.insertRows(sheet.getLastRow(), rows.length);
            sheet
              .getRange(sheet.getLastRow(), 1, rows.length, 5)
              .setValues(rows);
            Logger.log('bottom insertion: ' + rows.length + ' rows');
          }
        } else {
          Logger.log('no row to insert');
        }

        let next: string | null = result_json['page']['next'];
        if (next !== null) {
          sheet.getRange(1, 2).setValue(next);
          result_json = this.tryHttpGet(next);
        } else {
          sheet.getRange(1, 2).setValue('');
          sheet.getRange(1, 4).setValue(this.FIRSTROW);
          this.mainSheet.getRange(this.url_row_id, 3).setValue(1);
          this.url_row_id += 1;
          Logger.log('end scanning a board');
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
    Logger.log('try GET: ' + url);
    let res: HTTPResponse;
    let text: string;
    try {
      res = UrlFetchApp.fetch(url);
    } catch (e) {
      Logger.log(e.message);
    } finally {
      text = res.getContentText();
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
    sheet.getRange(1, 1).setValue('next url:');
    sheet.getRange(1, 2).setValue('');
    sheet.getRange(1, 4).setValue(this.FIRSTROW);
    sheet.getRange(1, 5).setValue('FolderID:');
    // 2 行目は header
    sheet.getRange(2, 1).setValue('id');
    sheet.getRange(2, 2).setValue('created_at');
    sheet.getRange(2, 3).setValue('width');
    sheet.getRange(2, 4).setValue('height');
    sheet.getRange(2, 5).setValue('url');
    sheet.getRange(2, 6).setValue('saved to Drive');

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
        sheet.getMaxColumns(),
      )
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    return sheet;
  }
}
