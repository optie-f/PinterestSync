# PinterestSync

> A Spreadsheet-bounded script to sync images of pins on a board with a specified folder in Google Drive

## Prerequisites

- [Node.js](https://nodejs.org/)
- [google/clasp](https://github.com/google/clasp)
- [Pinterest API](https://developers.pinterest.com/docs/api/overview/)

# How to Use

## 1. Set up spreadsheet

1. Create a spreadsheet to record pins.
2. Insert a sheet named `main`, and write data to some specific cells:
   - `B1`: Id of a directory in Google Drive to save folders( Images in each board are saved to the corresponding folder under that directory)
   - `B2`: Your access token for Pinterest API
3. Write names of boards to sync to `4th row and later` in the following manner:
   - column `A`: Username of the owner of the board to watch
   - column `B`: Name of the board
4. Open script editor and check `scriptId`.
   - What is scriptId ? https://github.com/google/clasp#scriptid-required

The `main` sheet should be like this:

|     | A           | B                          | ... |
| --- | ----------- | -------------------------- | --- |
| 1   | token:      | `<your_access_token>`      |     |
| 2   | dir_id:     | `<Id_of_parent_directory>` |     |
| 3   | username    | boardname                  |     |
| 4   | `user_hoge` | `fuga_board`               |     |
| 5   | ...         | ...                        |     |

notice: Don't decrease columns to less than 3. (3th column is used in this script)

## 2. Configuration

Create `PinterestSync/.clasp.json`, and write:

```json
{
  "scriptId": "<your_script_id>",
  "rootDir": "dist"
}
```

## 3. build

Install dependencies

```bash
yarn install # or 'npm install'
```

build:

```bash
yarn build  # or 'npm run build'
```

and push:

```
clasp push
```

## Other

This project is started with the help of [gas-clasp-starter](https://github.com/howdy39/gas-clasp-starter)
