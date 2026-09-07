/************************************************************
 * GLORY TO GLORY INTERNATIONAL SCHOOL
 * SCHOOL PAYMENT & ALLOCATION AUTOMATION
 *
 * FINANCE WORKBOOK
 * ----------------------------------------------------------
 * 1. Master Daily Inflow
 * 2. Parent/Student Financial Record
 * 3. Payment_Allocation        <- created automatically
 * 4. Allocation_Entry          <- created automatically
 * 5. _Payment_Baseline         <- created automatically
 *
 * STUDENT WORKBOOK
 * ----------------------------------------------------------
 * 1. Student Master
 * 2. Class sheets
 *
 ************************************************************/


/************************************************************
 * CONFIGURATION
 ************************************************************/

const CONFIG = {

  /*
   * Paste the ID of the Google Sheet that contains
   * Student Master and the class sheets.
   *
   * Example URL:
   * https://docs.google.com/spreadsheets/d/ABC123XYZ/edit
   *
   * The ID is:
   * ABC123XYZ
   */
  STUDENT_WORKBOOK_ID: '1KgaRmggfPFLh98SqXQDBPQFvZCB6o7P80bcZlbVLdK4',


  // -------------------------
  // SHEET NAMES
  // -------------------------

  PARENT_STUDENT_SHEET: 'Sch & Bus Fees 1st Term',

  INFLOW_SHEET: 'Daily Inflow 1st Term ',

  STUDENT_MASTER_SHEET: 'School FEE Control LIST ',

  ALLOCATION_LOG_SHEET: 'Payment_Allocation',

  ALLOCATION_ENTRY_SHEET: 'Allocation_Entry',

  BASELINE_SHEET: '_Payment_Baseline',

  SYNC_ISSUES_SHEET: 'Sync_Issues',


  // -------------------------
  // PARENT/STUDENT COLUMNS
  // -------------------------
  //
  // Based on the structure you supplied.
  //
  // Column numbers:
  // A = 1
  // B = 2
  // C = 3 etc.
  //

  PARENT: {

    FAMILY_SN: 1,

    PARENT_NAME: 2,

    PHONE: 3,

    PARENT_NO: 1,

    CHILD_SN: 5,

    STUDENT_NAME: 6,

    STUDENT_NO: 7,

    STUDENT_CLASS: 8,

    GENDER: 9,

    SCHOOL_FEE: 10,

    LESSON: 11,

    OTHER_PAYMENT: 12,

    OTHER_PAYMENT_TOTAL: 13,

    PRESENT_TERM_SCHOOL_FEES_TOTAL: 15,

    USES_BUS: 16,

    BUS_FEE: 17,

    BUS_FAMILY_TOTAL: 18,

    INDIVIDUAL_CURRENT_TERM_TOTAL: 19,

    FAMILY_CURRENT_TERM_TOTAL: 20,

    INDIVIDUAL_BROUGHT_FORWARD: 21,

    FAMILY_BROUGHT_FORWARD: 22,

    INDIVIDUAL_TOTAL_DUE: 24,

    FAMILY_TOTAL_DUE: 25,

    PAYMENT_1_INDIVIDUAL: 26,

    PAYMENT_1_FAMILY: 27,

    INDIVIDUAL_OLD_OUTSTANDING: 28,

    PAYMENT_1_DATE: 29,

    PAYMENT_1_RECEIPT: 30,

    BALANCE_1_FAMILY: 31,

    BALANCE_1_INDIVIDUAL: 32,

    PAYMENT_2_INDIVIDUAL: 33,

    PAYMENT_2_FAMILY: 34,

    PAYMENT_2_DATE: 35,

    PAYMENT_2_RECEIPT: 36,

    BALANCE_2_FAMILY: 37,

    BALANCE_2_INDIVIDUAL: 38,

    PAYMENT_3_INDIVIDUAL: 39,

    PAYMENT_3_FAMILY: 40,

    PAYMENT_3_DATE: 41,

    PAYMENT_3_RECEIPT: 42,

    BALANCE_3_FAMILY: 43,

    BALANCE_3_INDIVIDUAL: 44
  },


  // -------------------------
  // MASTER DAILY INFLOW
  // -------------------------
  //
  // These are your ORIGINAL columns.
  //
  // The script adds new automation columns
  // to the RIGHT of the existing sheet.
  //
  // Therefore it does NOT destroy your S/N.
  //

  INFLOW: {

    SN: 1,

    DATE: 2,

    PAYEE: 3,

    PARENT_NO: 4,

    AMOUNT_RECEIVED: 5,

    SCHOOL_FEES: 6,

    OUTSTANDING_SCHOOL_FEES: 7,

    REGISTRATION: 8,

    TEXTBOOKS: 9,

    // Column J = 10 is blank in your current layout

    GRADUATION: 11,

    UNIFORMS: 12,

    SPORTS_WEAR: 13,

    BUS_FEES: 14,

    LESSON: 15,

    COMMON_ENTRANCE: 16,

    JUNIOR_WAEC: 17,

    OTHERS: 18,

    TOTAL: 19,

    REMARK: 20
  },


  // -------------------------
  // STUDENT MASTER
  // -------------------------

  STUDENT: {

    STUDENT_NO: 1,       // A

    STUDENT_NAME: 2,     // B

    CLASS: 3,            // C

    ADMISSION_YEAR: 4,   // D

    SCHOOL_FEE: 5,       // E

    LESSON: 6,           // F

    TEXTBOOK: 7,         // G

    NOTEBOOK: 8,         // H

    UNIFORM: 9,          // I

    CLUBS: 10,           // J

    OUTSTANDING: 13,     // M

    TOTAL: 14,           // N

    AMOUNT_PAID: 15,     // O

    AMOUNT_OWING: 16,    // P

    COMMENTS: 17         // Q
  },


  TOLERANCE: 0.01
};



/************************************************************
 * MENU
 ************************************************************/

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('School Finance')
    .addItem(
      '1. Initial Setup',
      'initialSetup'
    )
    .addSeparator()
    .addItem(
      '2. Sync Student Numbers',
      'syncStudentNumbers'
    )
    .addItem(
      '3. Sync Fees to Student Master',
      'syncFeesToStudentMaster'
    )
    .addSeparator()
    .addItem(
      '4. Prepare Selected Payment',
      'prepareSelectedPayment'
    )
    .addItem(
      '5. Post Allocation',
      'postAllocation'
    )
    .addSeparator()
    .addItem(
      '6. Refresh Payment Balances',
      'refreshAllBalances'
    )
    .addToUi();

}



/************************************************************
 * INITIAL SETUP
 *
 * Run this ONCE before beginning automation.
 ************************************************************/

function initialSetup() {

  validateConfiguration();

  createAllocationLog();

  createAllocationEntrySheet();

  ensureInflowAutomationColumns();

  createPaymentBaseline();

  syncStudentNumbers();

  refreshAllBalances();

  SpreadsheetApp
    .getUi()
    .alert(
      'Initial setup completed successfully.\n\n' +
      'Your existing Amount Paid values have been saved ' +
      'as opening balances.'
    );

}



/************************************************************
 * VALIDATE CONFIG
 ************************************************************/

function validateConfiguration() {

  if (
    !CONFIG.STUDENT_WORKBOOK_ID ||
    CONFIG.STUDENT_WORKBOOK_ID ===
      'PASTE_STUDENT_WORKBOOK_ID_HERE'
  ) {

    throw new Error(
      'Please paste the Student Records Google Sheet ID ' +
      'into CONFIG.STUDENT_WORKBOOK_ID first.'
    );

  }

}



/************************************************************
 * NORMALIZE TEXT
 ************************************************************/

function normalize(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

}

/************************************************************
 * NORMALIZE CLASS
 *
 * Standardizes class names for comparison.
 *
 * Examples:
 * "Grade 1"   -> "grade 1"
 * " GRADE 1 " -> "grade 1"
 * "JSS   2"   -> "jss 2"
 ************************************************************/

function normalizeClass(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

}

function normalizeId(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return '';
  }

  return String(value)
    .trim();
}

/************************************************************
 * MONEY CONVERSION
 ************************************************************/

function money(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return 0;
  }

  if (typeof value === 'number') {
    return value;
  }

  const cleaned =
    String(value)
      .replace(/₦/g, '')
      .replace(/,/g, '')
      .trim();

  return Number(cleaned) || 0;

}



/************************************************************
 * STUDENT MATCH KEY
 *
 * Name + Class
 ************************************************************/

function studentMatchKey(name, studentClass) {

  return (
    normalize(name) +
    '|' +
    normalize(studentClass)
  );

}



/************************************************************
 * GET STUDENT WORKBOOK
 ************************************************************/

function getStudentWorkbook() {

  validateConfiguration();

  return SpreadsheetApp.openById(
    CONFIG.STUDENT_WORKBOOK_ID
  );

}



/************************************************************
 * CREATE PAYMENT ALLOCATION LOG
 ************************************************************/

function createAllocationLog() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.ALLOCATION_LOG_SHEET
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.ALLOCATION_LOG_SHEET
      );

  }

  const headers = [[

    'Allocation ID',

    'Payment ID',

    'Payment Date',

    'Parent No',

    'Parent / Payee',

    'Student No',

    'Student Name',

    'Class',

    'Category',

    'Amount',

    'Receipt No',

    'Remarks',

    'Entered By',

    'Logged At'

  ]];


  sheet
    .getRange(
      1,
      1,
      1,
      headers[0].length
    )
    .setValues(headers);


  sheet
    .getRange(
      1,
      1,
      1,
      headers[0].length
    )
    .setFontWeight('bold');


  sheet.setFrozenRows(1);

}



/************************************************************
 * CREATE ALLOCATION ENTRY SHEET
 *
 * This is your working screen.
 *
 * Payment_Allocation = permanent history.
 * Allocation_Entry   = temporary entry screen.
 ************************************************************/

function createAllocationEntrySheet() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.ALLOCATION_ENTRY_SHEET
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.ALLOCATION_ENTRY_SHEET
      );

  }


  sheet.clear();


  sheet
    .getRange('A1')
    .setValue(
      'PAYMENT ALLOCATION'
    )
    .setFontWeight('bold')
    .setFontSize(14);


  sheet.getRange('A3').setValue('Payment ID');

  sheet.getRange('A4').setValue('Payment Date');

  sheet.getRange('A5').setValue('Parent No');

  sheet.getRange('A6').setValue('Parent / Payee');

  sheet.getRange('A7').setValue('Amount Received');

  sheet.getRange('A8').setValue('Already Allocated');

  sheet.getRange('A9').setValue('Remaining To Allocate');


  const headers = [[

  'Student No',

  'Student Name',

  'Class',

  'Amount Due',

  'Category',

  'Category Available',

  'Amount To Allocate',

  'Remarks'

]];


  sheet
    .getRange(
      11,
      1,
      1,
      headers[0].length
    )
    .setValues(headers)
    .setFontWeight('bold');


  sheet.setFrozenRows(11);

}



/************************************************************
 * ADD AUTOMATION COLUMNS TO MASTER DAILY INFLOW
 *
 * They are appended to the RIGHT.
 *
 * Existing columns stay untouched.
 ************************************************************/

function ensureInflowAutomationColumns() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.INFLOW_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Master Daily Inflow sheet was not found.'
    );

  }


  const requiredHeaders = [

    'Payment ID',

    'Receipt No',

    'Payment Method',

    'Allocated Amount',

    'Unallocated Amount',

    'Allocation Status'

  ];


  const existingHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0];


  requiredHeaders.forEach(
    header => {

      if (
        !existingHeaders.includes(header)
      ) {

        const newColumn =
          sheet.getLastColumn() + 1;

        sheet
          .getRange(
            1,
            newColumn
          )
          .setValue(header);

        existingHeaders.push(header);

      }

    }
  );

}



/************************************************************
 * GET COLUMN BY HEADER
 ************************************************************/

function findHeaderColumn(
  sheet,
  header
) {

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0];


  const index =
    headers.indexOf(header);


  if (index === -1) {

    return 0;

  }


  return index + 1;

}



/************************************************************
 * GENERATE PAYMENT ID
 ************************************************************/

function makePaymentId(
  dateValue,
  rowNumber
) {

  const timezone =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSpreadsheetTimeZone();


  let datePart;


  if (
    dateValue instanceof Date
  ) {

    datePart =
      Utilities.formatDate(
        dateValue,
        timezone,
        'yyyyMMdd'
      );

  } else {

    datePart =
      Utilities.formatDate(
        new Date(),
        timezone,
        'yyyyMMdd'
      );

  }


  return (
    'PAY-' +
    datePart +
    '-' +
    String(rowNumber).padStart(4, '0')
  );

}



/************************************************************
 * ENSURE ONE PAYMENT HAS A PAYMENT ID
 ************************************************************/

function ensurePaymentId(
  sheet,
  row
) {

  ensureInflowAutomationColumns();


  const paymentIdCol =
    findHeaderColumn(
      sheet,
      'Payment ID'
    );


  let paymentId =
    sheet
      .getRange(
        row,
        paymentIdCol
      )
      .getDisplayValue()
      .trim();


  if (paymentId) {

    return paymentId;

  }


  const dateValue =
    sheet
      .getRange(
        row,
        CONFIG.INFLOW.DATE
      )
      .getValue();


  paymentId =
    makePaymentId(
      dateValue,
      row
    );


  sheet
    .getRange(
      row,
      paymentIdCol
    )
    .setValue(paymentId);


  return paymentId;

}



/************************************************************
 * SYNC STUDENT NUMBERS
 *
 * SOURCE:
 * Student Master
 *
 * MATCH:
 * Student Name + Class
 *
 * DESTINATION:
 * Parent/Student Financial Record
 ************************************************************/
function syncStudentNumbers() {

  validateConfiguration();

  const financeSS =
    SpreadsheetApp.getActiveSpreadsheet();

  const parentSheet =
    financeSS.getSheetByName(
      CONFIG.PARENT_STUDENT_SHEET
    );

  if (!parentSheet) {
    throw new Error(
      'Parent/Student sheet was not found: ' +
      CONFIG.PARENT_STUDENT_SHEET
    );
  }

  const studentSS =
    getStudentWorkbook();

  const studentSheet =
    studentSS.getSheetByName(
      CONFIG.STUDENT_MASTER_SHEET
    );

  if (!studentSheet) {
    throw new Error(
      'Student Master sheet was not found: ' +
      CONFIG.STUDENT_MASTER_SHEET
    );
  }

  const issueSheet =
    prepareSyncIssuesSheetBothWays();

  const studentData =
    studentSheet
      .getDataRange()
      .getValues();

  const parentData =
    parentSheet
      .getDataRange()
      .getValues();


  /********************************************************
   * BUILD STUDENT MASTER LOOKUP
   ********************************************************/

  const studentLookup = {};

  const duplicateStudentKeys =
    new Set();

  for (
    let i = 1;
    i < studentData.length;
    i++
  ) {

    const studentNo =
      studentData[i]
        [CONFIG.STUDENT.STUDENT_NO - 1];

    const studentName =
      studentData[i]
        [CONFIG.STUDENT.STUDENT_NAME - 1];

    const studentClass =
      studentData[i]
        [CONFIG.STUDENT.CLASS - 1];

    if (
      !studentName ||
      !studentClass
    ) {
      continue;
    }

    const key =
      studentMatchKey(
        studentName,
        studentClass
      );

    if (studentLookup[key]) {

      duplicateStudentKeys.add(key);

    } else {

      studentLookup[key] = {
        row: i + 1,
        studentNo: studentNo,
        studentName: studentName,
        studentClass: studentClass
      };

    }

  }


  /********************************************************
   * BUILD PARENT/STUDENT LOOKUP
   ********************************************************/

  const parentLookup = {};

  const duplicateParentKeys =
    new Set();

  for (
    let i = 3;
    i < parentData.length;
    i++
  ) {

    const studentName =
      parentData[i]
        [CONFIG.PARENT.STUDENT_NAME - 1];

    const studentClass =
      parentData[i]
        [CONFIG.PARENT.STUDENT_CLASS - 1];

    const existingStudentNo =
      parentData[i]
        [CONFIG.PARENT.STUDENT_NO - 1];

    if (
      !studentName ||
      !studentClass
    ) {
      continue;
    }

    const key =
      studentMatchKey(
        studentName,
        studentClass
      );

    if (parentLookup[key]) {

      duplicateParentKeys.add(key);

    } else {

      parentLookup[key] = {
        row: i + 1,
        studentNo: existingStudentNo,
        studentName: studentName,
        studentClass: studentClass
      };

    }

  }


  /********************************************************
   * SYNC STUDENT NUMBERS INTO PARENT/STUDENT
   ********************************************************/

  const studentNumberOutput = [];

  const issueRows = [];

  let matched = 0;

  let unmatchedParent = 0;

  let unmatchedMaster = 0;

  let duplicateCount = 0;


  for (
    let i = 3;
    i < parentData.length;
    i++
  ) {

    const studentName =
      parentData[i]
        [CONFIG.PARENT.STUDENT_NAME - 1];

    const studentClass =
      parentData[i]
        [CONFIG.PARENT.STUDENT_CLASS - 1];

    const existingStudentNo =
      parentData[i]
        [CONFIG.PARENT.STUDENT_NO - 1];

    if (
      !studentName ||
      !studentClass
    ) {

      studentNumberOutput.push([
        existingStudentNo
      ]);

      continue;

    }

    const key =
      studentMatchKey(
        studentName,
        studentClass
      );


    /******************************************************
     * DUPLICATE / AMBIGUOUS
     ******************************************************/

    if (
      duplicateStudentKeys.has(key) ||
      duplicateParentKeys.has(key)
    ) {

      studentNumberOutput.push([
        existingStudentNo || ''
      ]);

      issueRows.push([
        'AMBIGUOUS',
        i + 1,
        '',
        studentName,
        studentClass,
        existingStudentNo,
        'Duplicate Name + Class found',
        normalize(studentName),
        normalizeClass(studentClass)
      ]);

      duplicateCount++;

      continue;

    }


    /******************************************************
     * MATCH FOUND
     ******************************************************/

    if (studentLookup[key]) {

      studentNumberOutput.push([
        studentLookup[key].studentNo
      ]);

      matched++;

      continue;

    }


    /******************************************************
     * EXISTS ON PARENT/STUDENT BUT NOT STUDENT MASTER
     ******************************************************/

    studentNumberOutput.push([
      existingStudentNo || ''
    ]);

    issueRows.push([
      'PARENT/STUDENT ONLY',
      i + 1,
      '',
      studentName,
      studentClass,
      existingStudentNo,
      'Exists on Parent/Student but no match was found on Student Master',
      normalize(studentName),
      normalizeClass(studentClass)
    ]);

    unmatchedParent++;

  }


  /********************************************************
   * FIND STUDENTS THAT EXIST ON STUDENT MASTER
   * BUT NOT ON PARENT/STUDENT
   ********************************************************/

  for (
    let i = 1;
    i < studentData.length;
    i++
  ) {

    const studentNo =
      studentData[i]
        [CONFIG.STUDENT.STUDENT_NO - 1];

    const studentName =
      studentData[i]
        [CONFIG.STUDENT.STUDENT_NAME - 1];

    const studentClass =
      studentData[i]
        [CONFIG.STUDENT.CLASS - 1];

    if (
      !studentName ||
      !studentClass
    ) {
      continue;
    }

    const key =
      studentMatchKey(
        studentName,
        studentClass
      );

    if (
      duplicateStudentKeys.has(key) ||
      duplicateParentKeys.has(key)
    ) {
      continue;
    }

    if (!parentLookup[key]) {

      issueRows.push([
        'STUDENT MASTER ONLY',
        '',
        i + 1,
        studentName,
        studentClass,
        studentNo,
        'Exists on Student Master but no match was found on Parent/Student',
        normalize(studentName),
        normalizeClass(studentClass)
      ]);

      unmatchedMaster++;

    }

  }


  /********************************************************
   * WRITE STUDENT NUMBERS
   ********************************************************/

  if (
    studentNumberOutput.length
  ) {

    parentSheet
      .getRange(
        4,
        CONFIG.PARENT.STUDENT_NO,
        studentNumberOutput.length,
        1
      )
      .setValues(
        studentNumberOutput
      );

  }


  /********************************************************
   * WRITE SYNC ISSUES
   ********************************************************/

  if (
    issueRows.length
  ) {

    issueSheet
      .getRange(
        2,
        1,
        issueRows.length,
        issueRows[0].length
      )
      .setValues(
        issueRows
      );

  }


  /********************************************************
   * FINAL SUMMARY
   ********************************************************/

  const totalIssues =
  unmatchedParent +
  unmatchedMaster +
  duplicateCount;

if (totalIssues > 0) {

  SpreadsheetApp
    .getUi()
    .alert(
      'Student Number Sync Complete\n\n' +
      'Matched: ' +
      matched +
      '\n' +
      'Parent/Student only: ' +
      unmatchedParent +
      '\n' +
      'Student Master only: ' +
      unmatchedMaster +
      '\n' +
      'Duplicate / ambiguous: ' +
      duplicateCount +
      '\n\n' +
      'Please check the "' +
      CONFIG.SYNC_ISSUES_SHEET +
      '" sheet for details.'
    );

} else {

  SpreadsheetApp
    .getUi()
    .alert(
      'Student Number Sync Complete\n\n' +
      'Matched: ' +
      matched +
      '\n\n' +
      'All students matched successfully.'
    );

}

}



/************************************************************
 * SYNC PARENT/STUDENT FEE INFORMATION
 * TO STUDENT MASTER
 *
 * This removes your manual copying.
 *
 * At the moment we sync ONLY fields for which
 * your two supplied structures have a clear match:
 *
 * Parent School Fee -> Student Master School Fee
 * Parent Lesson     -> Student Master Lesson
 * Balance BF        -> Student Master Outstanding
 *
 * Other fields are deliberately NOT guessed.
 ************************************************************/

function syncFeesToStudentMaster() {

  const financeSS =
    SpreadsheetApp.getActiveSpreadsheet();


  const parentSheet =
    financeSS.getSheetByName(
      CONFIG.PARENT_STUDENT_SHEET
    );


  const studentSS =
    getStudentWorkbook();


  const studentSheet =
    studentSS.getSheetByName(
      CONFIG.STUDENT_MASTER_SHEET
    );


  const parentData =
    parentSheet
      .getDataRange()
      .getValues();


  const studentData =
    studentSheet
      .getDataRange()
      .getValues();


  /*
   * Student No -> Parent/Student financial row
   */

  const financeLookup = {};


  for (
    let i = 3;
    i < parentData.length;
    i++
  ) {

    const studentNo =
      normalizeId(
        parentData[i]
          [CONFIG.PARENT.STUDENT_NO - 1]
      );


    if (!studentNo) {
      continue;
    }


    financeLookup[studentNo] = {

      schoolFee:
        money(
          parentData[i]
            [CONFIG.PARENT.SCHOOL_FEE - 1]
        ),

      lesson:
        money(
          parentData[i]
            [CONFIG.PARENT.LESSON - 1]
        ),

      outstanding:
        money(
          parentData[i]
            [
              CONFIG.PARENT
                .INDIVIDUAL_BROUGHT_FORWARD - 1
            ]
        )

    };

  }


  for (
    let i = 1;
    i < studentData.length;
    i++
  ) {

    const studentNo =
      normalizeId(
        studentData[i]
          [CONFIG.STUDENT.STUDENT_NO - 1]
      );


    if (
      !studentNo ||
      !financeLookup[studentNo]
    ) {
      continue;
    }


    const record =
      financeLookup[studentNo];


    studentSheet
      .getRange(
        i + 1,
        CONFIG.STUDENT.SCHOOL_FEE
      )
      .setValue(
        record.schoolFee
      );


    studentSheet
      .getRange(
        i + 1,
        CONFIG.STUDENT.LESSON
      )
      .setValue(
        record.lesson
      );


    studentSheet
      .getRange(
        i + 1,
        CONFIG.STUDENT.OUTSTANDING
      )
      .setValue(
        record.outstanding
      );


    /*
     * TOTAL
     * E:M
     */

    studentSheet
      .getRange(
        i + 1,
        CONFIG.STUDENT.TOTAL
      )
      .setFormula(
        '=IFERROR(SUM(E' +
        (i + 1) +
        ':M' +
        (i + 1) +
        '),"")'
      );

  }


  SpreadsheetApp
    .getUi()
    .alert(
      'Fee information synced to Student Master.'
    );

}

function prepareSyncIssuesSheetBothWays() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.SYNC_ISSUES_SHEET
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.SYNC_ISSUES_SHEET
      );

  }

  sheet.clear();

  const headers = [[
    'Source',
    'Parent Sheet Row',
    'Student Master Row',
    'Student Name',
    'Class',
    'Student No',
    'Issue',
    'Normalized Name',
    'Normalized Class'
  ]];

  sheet
    .getRange(
      1,
      1,
      1,
      headers[0].length
    )
    .setValues(headers)
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  return sheet;

}

/************************************************************
 * CREATE OPENING PAYMENT BASELINE
 *
 * VERY IMPORTANT.
 *
 * This captures the Amount Paid already in
 * Student Master BEFORE automation.
 *
 * Existing payments are therefore not lost.
 ************************************************************/

function createPaymentBaseline() {

  const financeSS =
    SpreadsheetApp.getActiveSpreadsheet();


  let baselineSheet =
    financeSS.getSheetByName(
      CONFIG.BASELINE_SHEET
    );


  if (!baselineSheet) {

    baselineSheet =
      financeSS.insertSheet(
        CONFIG.BASELINE_SHEET
      );

  }


  /*
   * If baseline already contains data,
   * DO NOT recreate it.
   */

  if (
    baselineSheet.getLastRow() > 1
  ) {

    return;

  }


  baselineSheet.clear();


  baselineSheet
    .getRange(
      1,
      1,
      1,
      4
    )
    .setValues([[
      'Student No',
      'Student Name',
      'Opening Amount Paid',
      'Captured At'
    ]])
    .setFontWeight('bold');


  const studentSS =
    getStudentWorkbook();


  const studentSheet =
    studentSS.getSheetByName(
      CONFIG.STUDENT_MASTER_SHEET
    );


  const data =
    studentSheet
      .getDataRange()
      .getValues();


  const rows = [];


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const studentNo =
      data[i]
        [CONFIG.STUDENT.STUDENT_NO - 1];


    if (!studentNo) {
      continue;
    }


    rows.push([

      studentNo,

      data[i]
        [CONFIG.STUDENT.STUDENT_NAME - 1],

      money(
        data[i]
          [CONFIG.STUDENT.AMOUNT_PAID - 1]
      ),

      new Date()

    ]);

  }


  if (rows.length) {

    baselineSheet
      .getRange(
        2,
        1,
        rows.length,
        4
      )
      .setValues(rows);

  }


  /*
   * Hide the helper sheet.
   */

  baselineSheet.hideSheet();

}



/************************************************************
 * PAYMENT CATEGORY DEFINITIONS
 ************************************************************/

function getPaymentCategories() {

  return [

    {
      name: 'School Fees',
      column:
        CONFIG.INFLOW.SCHOOL_FEES
    },

    {
      name: 'Outstanding School Fees',
      column:
        CONFIG.INFLOW.OUTSTANDING_SCHOOL_FEES
    },

    {
      name: 'New Student Registration Fees',
      column:
        CONFIG.INFLOW.REGISTRATION
    },

    {
      name: 'Textbooks',
      column:
        CONFIG.INFLOW.TEXTBOOKS
    },

    {
      name: 'End of Session / Graduation',
      column:
        CONFIG.INFLOW.GRADUATION
    },

    {
      name: 'School Uniforms',
      column:
        CONFIG.INFLOW.UNIFORMS
    },

    {
      name: 'Sports Wear',
      column:
        CONFIG.INFLOW.SPORTS_WEAR
    },

    {
      name: 'Bus Fees',
      column:
        CONFIG.INFLOW.BUS_FEES
    },

    {
      name: 'Lesson',
      column:
        CONFIG.INFLOW.LESSON
    },

    {
      name: 'Common Entrance Registration',
      column:
        CONFIG.INFLOW.COMMON_ENTRANCE
    },

    {
      name: 'Junior WAEC Registration/Lesson',
      column:
        CONFIG.INFLOW.JUNIOR_WAEC
    },

    {
      name: 'Others',
      column:
        CONFIG.INFLOW.OTHERS
    }

  ];

}



/************************************************************
 * GET EXISTING ALLOCATIONS FOR PAYMENT
 ************************************************************/

function getExistingAllocations(
  paymentId
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CONFIG.ALLOCATION_LOG_SHEET
    );


  const result = {

    total: 0,

    byCategory: {}

  };


  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    return result;

  }


  const data =
    sheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      normalize(
        data[i][1]
      ) !==
      normalize(paymentId)
    ) {
      continue;
    }


    const category =
      String(
        data[i][8] || ''
      ).trim();


    const amount =
      money(
        data[i][9]
      );


    result.total += amount;


    if (
      !result.byCategory[category]
    ) {

      result.byCategory[category] = 0;

    }


    result.byCategory[category] +=
      amount;

  }


  return result;

}



/************************************************************
 * GET CHILDREN FOR PARENT
 *
 * Handles parent numbers that appear only
 * once at the start of a family block.
 ************************************************************/

function getChildrenForParent(
  parentNo
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CONFIG.PARENT_STUDENT_SHEET
    );


  const data =
    sheet
      .getDataRange()
      .getValues();


  const target =
    normalizeId(parentNo);


  const children = [];


  let currentParentNo = '';

  let currentParentName = '';


  const seenStudents =
    new Set();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const rawParentNo =
      data[i]
        [CONFIG.PARENT.PARENT_NO - 1];


    const rawParentName =
      data[i]
        [CONFIG.PARENT.PARENT_NAME - 1];


    if (rawParentNo) {

      currentParentNo =
        rawParentNo;

    }


    if (rawParentName) {

      currentParentName =
        rawParentName;

    }


    if (
      normalizeId(
        currentParentNo
      ) !== target
    ) {
      continue;
    }


    const studentNo =
      data[i]
        [CONFIG.PARENT.STUDENT_NO - 1];


    const studentName =
      data[i]
        [CONFIG.PARENT.STUDENT_NAME - 1];


    const studentClass =
      data[i]
        [CONFIG.PARENT.STUDENT_CLASS - 1];

    const amountDue =
        money(
      data[i]
        [CONFIG.PARENT.INDIVIDUAL_TOTAL_DUE - 1]
  );

    


    if (
      !studentName ||
      !studentClass
    ) {
      continue;
    }


    const uniqueKey =
      normalizeId(studentNo) ||
      studentMatchKey(
        studentName,
        studentClass
      );


    if (
      seenStudents.has(uniqueKey)
    ) {
      continue;
    }


    seenStudents.add(
      uniqueKey
    );


  children.push({

  studentNo:
    studentNo,

  studentName:
    studentName,

  studentClass:
    studentClass,

  amountDue:
    amountDue,

  parentName:
    currentParentName

});

  }


  return children;

}



/************************************************************
 * PREPARE SELECTED PAYMENT
 *
 * HOW TO USE:
 *
 * 1. Go to Master Daily Inflow.
 * 2. Click ANY CELL in the payment row.
 * 3. School Finance ->
 *    Prepare Selected Payment.
 *
 * Allocation_Entry is then prepared automatically.
 ************************************************************/

function prepareSelectedPayment() {

  //syncStudentNumbers();

  ensureInflowAutomationColumns();

  createAllocationLog();

  createAllocationEntrySheet();


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const inflowSheet =
    ss.getSheetByName(
      CONFIG.INFLOW_SHEET
    );


  const activeSheet =
    ss.getActiveSheet();


  if (
    activeSheet.getName() !==
    CONFIG.INFLOW_SHEET
  ) {

    throw new Error(
      'Please select a payment row in Master Daily Inflow first.'
    );

  }


  const row =
    activeSheet
      .getActiveCell()
      .getRow();


  if (row <= 1) {

    throw new Error(
      'Please select an actual payment row.'
    );

  }


  const amountReceived =
    money(
      inflowSheet
        .getRange(
          row,
          CONFIG.INFLOW.AMOUNT_RECEIVED
        )
        .getValue()
    );


  if (
    amountReceived <= 0
  ) {

    throw new Error(
      'The selected row does not contain a valid Amount Received.'
    );

  }


  const paymentId =
    ensurePaymentId(
      inflowSheet,
      row
    );


  const date =
    inflowSheet
      .getRange(
        row,
        CONFIG.INFLOW.DATE
      )
      .getValue();


  const payee =
    inflowSheet
      .getRange(
        row,
        CONFIG.INFLOW.PAYEE
      )
      .getDisplayValue();


  const parentNo =
    inflowSheet
      .getRange(
        row,
        CONFIG.INFLOW.PARENT_NO
      )
      .getDisplayValue();


  if (!parentNo) {

    throw new Error(
      'The selected payment has no Parent No.'
    );

  }


  const children =
    getChildrenForParent(
      parentNo
    );


  if (!children.length) {

    throw new Error(
      'No children were found for Parent No: ' +
      parentNo +
      '.\n\n' +
      'Check the Parent No and run Sync Student Numbers.'
    );

  }


  const existing =
    getExistingAllocations(
      paymentId
    );


  const remainingPayment =
    amountReceived -
    existing.total;


  const categoryRows = [];


  const categories =
    getPaymentCategories();


  categories.forEach(
    category => {

      const originalAmount =
        money(
          inflowSheet
            .getRange(
              row,
              category.column
            )
            .getValue()
        );


      if (
        originalAmount <= 0
      ) {
        return;
      }


      const alreadyAllocated =
        existing.byCategory[
          category.name
        ] || 0;


      const available =
        originalAmount -
        alreadyAllocated;


      if (
        available >
        CONFIG.TOLERANCE
      ) {

        categoryRows.push({

          name:
            category.name,

          available:
            available

        });

      }

    }
  );


  /*
   * If no category was supplied in
   * Master Daily Inflow, allow
   * "Unspecified".
   */

  if (
    !categoryRows.length &&
    remainingPayment >
      CONFIG.TOLERANCE
  ) {

    categoryRows.push({

      name:
        'Unspecified',

      available:
        remainingPayment

    });

  }


  const entrySheet =
    ss.getSheetByName(
      CONFIG.ALLOCATION_ENTRY_SHEET
    );


  /*
   * Clear old transaction.
   */

  entrySheet
    .getRange(
      3,
      2,
      7,
      6
    )
    .clearContent();


  const lastRow =
    Math.max(
      entrySheet.getLastRow(),
      12
    );


  if (lastRow >= 12) {

    entrySheet
      .getRange(
        12,
        1,
        lastRow - 11,
        8
      )
      .clearContent();

  }


  entrySheet
    .getRange('B3')
    .setValue(paymentId);


  entrySheet
    .getRange('B4')
    .setValue(date);


  entrySheet
    .getRange('B5')
    .setValue(parentNo);


  entrySheet
    .getRange('B6')
    .setValue(payee);


  entrySheet
    .getRange('B7')
    .setValue(amountReceived);


  entrySheet
    .getRange('B8')
    .setValue(existing.total);


  entrySheet
    .getRange('B9')
    .setValue(remainingPayment);


  /*
   * Create one candidate row for
   * every Child x Category.
   *
   * You enter only the amounts
   * that apply.
   */

  const rows = [];


  children.forEach(
    child => {

      categoryRows.forEach(
        category => {

          rows.push([

            child.studentNo,

            child.studentName,

            child.studentClass,

            child.amountDue,

            category.name,

            category.available,

            '',

            ''

          ]);

      }
    );

  }
);


  if (rows.length) {

    entrySheet
      .getRange(
        12,
        1,
        rows.length,
        8
      )
      .setValues(rows);

  }

  entrySheet.activate();



}



/************************************************************
 * POST ALLOCATION
 *
 * Reads Allocation_Entry and permanently logs
 * the records into Payment_Allocation.
 ************************************************************/

function postAllocation() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const entrySheet =
    ss.getSheetByName(
      CONFIG.ALLOCATION_ENTRY_SHEET
    );


  const logSheet =
    ss.getSheetByName(
      CONFIG.ALLOCATION_LOG_SHEET
    );


  const inflowSheet =
    ss.getSheetByName(
      CONFIG.INFLOW_SHEET
    );


  const paymentId =
    entrySheet
      .getRange('B3')
      .getDisplayValue()
      .trim();


  const paymentDate =
    entrySheet
      .getRange('B4')
      .getValue();


  const parentNo =
    entrySheet
      .getRange('B5')
      .getDisplayValue();


  const payee =
    entrySheet
      .getRange('B6')
      .getDisplayValue();


  const remainingBefore =
    money(
      entrySheet
        .getRange('B9')
        .getValue()
    );


  if (!paymentId) {

    throw new Error(
      'No payment is currently loaded in Allocation_Entry.'
    );

  }


  if (
    remainingBefore <=
    CONFIG.TOLERANCE
  ) {

    throw new Error(
      'This payment is already fully allocated.'
    );

  }


  const lastRow =
    entrySheet.getLastRow();


  if (lastRow < 12) {

    throw new Error(
      'There are no allocation rows.'
    );

  }


  const data =
    entrySheet
      .getRange(
        12,
        1,
        lastRow - 11,
        8
      )
      .getValues();


  const allocations = [];


  const categoryTotals = {};


  let allocationTotal = 0;


  data.forEach(
    row => {

      const studentNo =
          row[0];

        const studentName =
          row[1];

        const studentClass =
          row[2];

        const amountDue =
          money(row[3]);

        const category =
          row[4];

        const categoryAvailable =
          money(row[5]);

        const amount =
          money(row[6]);

        const remarks =
          row[7];


      if (
        amount <= 0
      ) {
        return;
      }


      if (!studentNo) {

        throw new Error(
          'An allocation amount was entered for ' +
          studentName +
          ' but the student has no Student Number.'
        );

      }


      if (
        amount >
        categoryAvailable +
        CONFIG.TOLERANCE
      ) {

        throw new Error(
          studentName +
          ' has an allocation greater than the ' +
          'available amount for ' +
          category +
          '.'
        );

      }


      allocations.push({

        studentNo:
          studentNo,

        studentName:
          studentName,

        studentClass:
          studentClass,

        category:
          category,

        amount:
          amount,

        remarks:
          remarks

      });


      allocationTotal +=
        amount;


      if (
        !categoryTotals[category]
      ) {

        categoryTotals[category] = 0;

      }


      categoryTotals[category] +=
        amount;

    }
  );


  if (!allocations.length) {

    throw new Error(
      'Enter at least one allocation amount.'
    );

  }


  if (
    allocationTotal >
    remainingBefore +
    CONFIG.TOLERANCE
  ) {

    throw new Error(
      'You are attempting to allocate ₦' +
      allocationTotal.toLocaleString() +
      ' but only ₦' +
      remainingBefore.toLocaleString() +
      ' remains on this payment.'
    );

  }


  /*
   * Ensure category total does not exceed
   * actual category amount remaining.
   */

  const availableByCategory = {};


  data.forEach(
    row => {

    const category =
      row[4];

    const available =
      money(row[5]);


      if (
        !category ||
        available <= 0
      ) {
        return;
      }


      availableByCategory[category] =
        available;

    }
  );


  Object.keys(
    categoryTotals
  ).forEach(
    category => {

      if (
        categoryTotals[category] >
        availableByCategory[category] +
        CONFIG.TOLERANCE
      ) {

        throw new Error(
          'Total allocation for ' +
          category +
          ' exceeds the amount available.'
        );

      }

    }
  );


  /*
   * Get receipt number from inflow.
   */

  const paymentIdCol =
    findHeaderColumn(
      inflowSheet,
      'Payment ID'
    );


  const receiptCol =
    findHeaderColumn(
      inflowSheet,
      'Receipt No'
    );


  let receiptNo = '';


  const inflowData =
    inflowSheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < inflowData.length;
    i++
  ) {

    if (
      normalize(
        inflowData[i]
          [paymentIdCol - 1]
      ) ===
      normalize(paymentId)
    ) {

      receiptNo =
        inflowData[i]
          [receiptCol - 1];

      break;

    }

  }


  const enteredBy =
    Session
      .getActiveUser()
      .getEmail() || 'Unknown User';


  const logRows =
    allocations.map(
      allocation => [

        'AL-' +
        Utilities
          .getUuid()
          .substring(0, 10)
          .toUpperCase(),

        paymentId,

        paymentDate,

        parentNo,

        payee,

        allocation.studentNo,

        allocation.studentName,

        allocation.studentClass,

        allocation.category,

        allocation.amount,

        receiptNo,

        allocation.remarks,

        enteredBy,

        new Date()

      ]
    );


  logSheet
    .getRange(
      logSheet.getLastRow() + 1,
      1,
      logRows.length,
      logRows[0].length
    )
    .setValues(logRows);


  /*
   * Update balances everywhere.
   */

  refreshAllBalances();


  /*
   * Refresh this same payment so that
   * only the remaining amount appears.
   */

  const paymentRow =
    findPaymentRow(
      paymentId
    );


  if (paymentRow) {

    inflowSheet
      .getRange(
        paymentRow,
        1
      )
      .activate();


    prepareSelectedPayment();

  }


  SpreadsheetApp
    .getUi()
    .alert(
      'Allocation posted successfully.\n\n' +
      'Allocated now: ₦' +
      allocationTotal.toLocaleString()
    );

}



/************************************************************
 * FIND PAYMENT ROW
 ************************************************************/

function findPaymentRow(
  paymentId
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      CONFIG.INFLOW_SHEET
    );


  const paymentIdCol =
    findHeaderColumn(
      sheet,
      'Payment ID'
    );


  if (!paymentIdCol) {
    return 0;
  }


  const data =
    sheet
      .getRange(
        2,
        paymentIdCol,
        Math.max(
          sheet.getLastRow() - 1,
          1
        ),
        1
      )
      .getDisplayValues();


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    if (
      normalize(data[i][0]) ===
      normalize(paymentId)
    ) {

      return i + 2;

    }

  }


  return 0;

}



/************************************************************
 * REFRESH EVERYTHING
 ************************************************************/

function refreshAllBalances() {

  ensureInflowAutomationColumns();

  createAllocationLog();

  updateInflowAllocationStatus();

  updateStudentMasterPayments();

}



/************************************************************
 * UPDATE MASTER DAILY INFLOW STATUS
 ************************************************************/

function updateInflowAllocationStatus() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const inflow =
    ss.getSheetByName(
      CONFIG.INFLOW_SHEET
    );


  const logSheet =
    ss.getSheetByName(
      CONFIG.ALLOCATION_LOG_SHEET
    );


  const paymentIdCol =
    findHeaderColumn(
      inflow,
      'Payment ID'
    );


  const allocatedCol =
    findHeaderColumn(
      inflow,
      'Allocated Amount'
    );


  const unallocatedCol =
    findHeaderColumn(
      inflow,
      'Unallocated Amount'
    );


  const statusCol =
    findHeaderColumn(
      inflow,
      'Allocation Status'
    );


  /*
   * Payment ID -> Allocated Total
   */

  const allocatedLookup = {};


  if (
    logSheet.getLastRow() >= 2
  ) {

    const allocationData =
      logSheet
        .getDataRange()
        .getValues();


    for (
      let i = 1;
      i < allocationData.length;
      i++
    ) {

      const paymentId =
        normalize(
          allocationData[i][1]
        );


      const amount =
        money(
          allocationData[i][9]
        );


      if (!paymentId) {
        continue;
      }


      allocatedLookup[paymentId] =
        (
          allocatedLookup[paymentId] ||
          0
        ) + amount;

    }

  }


  const lastRow =
    inflow.getLastRow();


  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {

    const amountReceived =
      money(
        inflow
          .getRange(
            row,
            CONFIG.INFLOW.AMOUNT_RECEIVED
          )
          .getValue()
      );


    if (
      amountReceived <= 0
    ) {
      continue;
    }


    const paymentId =
      ensurePaymentId(
        inflow,
        row
      );


    const allocated =
      allocatedLookup[
        normalize(paymentId)
      ] || 0;


    const remaining =
      amountReceived -
      allocated;


    let status =
      'UNALLOCATED';


    if (
      Math.abs(remaining) <=
      CONFIG.TOLERANCE
    ) {

      status =
        'FULLY ALLOCATED';

    } else if (
      allocated > 0
    ) {

      status =
        'PARTIALLY ALLOCATED';

    }


    inflow
      .getRange(
        row,
        allocatedCol
      )
      .setValue(
        allocated
      );


    inflow
      .getRange(
        row,
        unallocatedCol
      )
      .setValue(
        remaining
      );


    inflow
      .getRange(
        row,
        statusCol
      )
      .setValue(
        status
      );

  }

}



/************************************************************
 * UPDATE STUDENT MASTER
 *
 * Amount Paid =
 *
 * Existing Opening Paid Balance
 * +
 * New Payment Allocations
 *
 * Amount Owing =
 *
 * Total - Amount Paid
 ************************************************************/

function updateStudentMasterPayments() {

  const financeSS =
    SpreadsheetApp.getActiveSpreadsheet();


  let baselineSheet =
    financeSS.getSheetByName(
      CONFIG.BASELINE_SHEET
    );


  if (
    !baselineSheet ||
    baselineSheet.getLastRow() < 2
  ) {

    createPaymentBaseline();

    baselineSheet =
      financeSS.getSheetByName(
        CONFIG.BASELINE_SHEET
      );

  }


  const allocationSheet =
    financeSS.getSheetByName(
      CONFIG.ALLOCATION_LOG_SHEET
    );


  const baselineData =
    baselineSheet
      .getDataRange()
      .getValues();


  const openingPaid = {};


  for (
    let i = 1;
    i < baselineData.length;
    i++
  ) {

    const studentNo =
      normalizeId(
        baselineData[i][0]
      );


    if (!studentNo) {
      continue;
    }


    openingPaid[studentNo] =
      money(
        baselineData[i][2]
      );

  }


  const newPayments = {};


  if (
    allocationSheet.getLastRow() >= 2
  ) {

    const allocationData =
      allocationSheet
        .getDataRange()
        .getValues();


    for (
      let i = 1;
      i < allocationData.length;
      i++
    ) {

      const studentNo =
        normalizeId(
          allocationData[i][5]
        );


      const amount =
        money(
          allocationData[i][9]
        );


      if (!studentNo) {
        continue;
      }


      newPayments[studentNo] =
        (
          newPayments[studentNo] ||
          0
        ) + amount;

    }

  }


  const studentSS =
    getStudentWorkbook();


  const studentSheet =
    studentSS.getSheetByName(
      CONFIG.STUDENT_MASTER_SHEET
    );


  const studentData =
    studentSheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < studentData.length;
    i++
  ) {

    const row =
      i + 1;


    const studentNo =
      normalizeId(
        studentData[i]
          [CONFIG.STUDENT.STUDENT_NO - 1]
      );


    if (!studentNo) {
      continue;
    }


    /*
     * New students not present when
     * baseline was captured begin at zero.
     */

    const opening =
      openingPaid[studentNo] || 0;


    const allocated =
      newPayments[studentNo] || 0;


    const amountPaid =
      opening +
      allocated;


    studentSheet
      .getRange(
        row,
        CONFIG.STUDENT.AMOUNT_PAID
      )
      .setValue(
        amountPaid
      );


    /*
     * Keep your Amount Owing formula.
     */

    studentSheet
      .getRange(
        row,
        CONFIG.STUDENT.AMOUNT_OWING
      )
      .setFormula(
        '=IFERROR(N' +
        row +
        '-O' +
        row +
        ',"")'
      );

  }

}
