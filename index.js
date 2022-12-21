// run `node index.js` in the terminal

const path = require('path');
const csv = require('csv-parser');
const fs = require('fs');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const assertString = (input) => {
  if (typeof input !== 'string') {
    throw Error(`Input was not string, but: "${typeof input}"`)
  }

  return input
}

// path to strong csv
// strong862973258176784790.csv
const INPUT_FILE = assertString(process.env.INPUT_FILE);
// follow instructions to get google sheet authorization
// https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication?id=service-account
const GOOGLE_PRIVATE_KEY = assertString(process.env.GOOGLE_PRIVATE_KEY).replace(/\\n/g, '\n')
const GOOGLE_SERVICE_ACCOUNT_EMAIL = assertString(process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL)
// doc ID is the long id in the sheets URL
const GOOGLE_SHEET_ID = assertString(process.env.GOOGLE_SHEET_ID)

const CSV_PATH = path.join(__dirname, INPUT_FILE);

/** @typedef {Object.<string, string>} Exercise */
/** @typedef {Exercise[]} Strong */
/** @typedef {Object.<string, string>} Row */
/** @typedef {import('google-spreadsheet').GoogleSpreadsheetWorksheet} Sheet */

/**
* @param {string} inputPath
* @returns {Promise<Strong>} 
*/
const load = async (inputPath) => {
  return new Promise((resolve) => {
    const results = [];

    fs.createReadStream(inputPath)
      .pipe(csv({ separator: ';' }))
      .on('data', (data) => results.push(data))
      .on('end', () => {
        resolve(results)
        // console.log(results);
        // [
        // {
        //    Date: '2022-10-28 13:42:11',
        //    'Workout Name': 'Week A Day 2',
        //    'Exercise Name': 'Deadlift (Barbell)',
        //    'Set Order': '1',
        //    Weight: '135',
        //    'Weight Unit': 'lbs',
        //    Reps: '4',
        //    RPE: '6',
        //    Distance: '',
        //    'Distance Unit': '',
        //    Seconds: '0',
        //    Notes: '',
        //    'Workout Notes': '',
        //    'Workout Duration': '1h 15m'
        //  }
        // ]
      });
  })
}

const doc = new GoogleSpreadsheet(GOOGLE_SHEET_ID);

const auth = async () => {
  await doc.useServiceAccountAuth({
    client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: GOOGLE_PRIVATE_KEY,
  });
}

const firstSheetTitle = 'Untitled'
const headerValues = ['Date', 'Set Order', 'Reps', 'Weight', 'RPE', 'Notes']

/** @param {Strong} strong */
const createSheets = async (strong) => {
  // get all unique 'Exercise Name' keys out of array
  /** @type {string[]} */
  const exercises = strong.reduce(
    (acc, current) => current['Exercise Name'] && acc.includes(current['Exercise Name']) ? acc : [...acc, current['Exercise Name']],
    []
  ).sort()

  for (title of exercises) {
    if (title === exercises[0]) {
      const firstSheet = doc.sheetsByTitle[firstSheetTitle]
      await firstSheet.updateProperties({ title })
      await firstSheet.setHeaderRow(headerValues)
    } else {
      await doc.addSheet({ title, headerValues })
    }
  }
}

/** @param {Strong} strong */
const resetSheets = async (strong) => {
  const existingSheetIds = Object.keys(doc.sheetsById)
  const lastSheetId = existingSheetIds[existingSheetIds.length - 1]
  const lastSheet = doc.sheetsById[lastSheetId]
  for (sheetId of existingSheetIds) {
    if (sheetId !== lastSheetId) {
      await doc.deleteSheet(sheetId)
    } else {
      await lastSheet.updateProperties({ title: firstSheetTitle })
      await lastSheet.clear()
    }
  }
}

/** 
* Reduce exercise payload down to keys with only specified header keys
* @param {Exercise} e
* @returns {Row}
*/
const row = (e) => headerValues.reduce(
  (acc, header) => e[header] ? { ...acc, [header]: e[header] } : acc, {}
)

/** @param {Strong} strong */
const fillSheets = async (strong) => {
  const sheets = Object.entries(doc.sheetsByTitle)
  for (const [exercise, sheet] of sheets) {
    const exercises = strong
      .filter((e) => e['Exercise Name'] === exercise)
      .map(row)

    await sheet.addRows(exercises)
  }
}

/** @param {Sheet} sheet */
const applyDateColumn = async (sheet) => {

  const limit = 200
  const cells = new Array(limit).fill(0)
  await sheet.loadCells(`A1:A${limit}`);

  cells.forEach((_, index) => {
    const cell = sheet.getCell(index, 0)
    const dateTimeFormat = 'ddd, m/d/yy' // https://developers.google.com/sheets/api/guides/formats
    cell.numberFormat = { type: 'DATE', pattern: dateTimeFormat }
  })
}

/** 
* Create letter from index (0 -> A, 1 -> B)
* @param {number} index  
*/
const l = (index) => String.fromCharCode(64 + index)

/** @param {Sheet} sheet */
const applyHeaderStyles = async (sheet) => {
  const limit = headerValues.length
  const cells = new Array(limit).fill(0)
  const L = l(limit)
  await sheet.loadCells(`A1:${L}1`);
  cells.forEach((_, index) => {
    const cell = sheet.getCell(0, index)
    cell.textFormat = { bold: true }
  })
}

const applyFormatting = async () => {
  const sheets = Object.entries(doc.sheetsByTitle)
  for (const [_title, sheet] of sheets) {
    await applyDateColumn(sheet)
    await applyHeaderStyles(sheet)
    await sheet.saveUpdatedCells();
  }
}

const run = async () => {
  console.log('Authenticating with Google Sheets API...')
  await auth()
  console.log('Loading doc...')
  await doc.loadInfo()

  const strong = await load(CSV_PATH)

  console.log('Resetting sheets...')
  await resetSheets(strong)
  console.log('Creating sheets...')
  await createSheets(strong)
  console.log('Filling sheets...')
  await fillSheets(strong)
  console.log('Format sheets...')
  await applyFormatting()
}

(async () => {
  try {
    await run();
  } catch (e) {
    console.error(e.message)
  }
})()
