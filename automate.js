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
    FAMILY_ID: 4,       // D - permanent family automation key
    PARENT_NO: 1,       // display only

    CHILD_SN: 5,
    STUDENT_NAME: 6,
    STUDENT_NO: 7,      // display only
    STUDENT_ID: 8,      // permanent student automation key
    STUDENT_CLASS: 9,
    GENDER: 10,
    SCHOOL_FEE: 11,
    LESSON: 12,
    OTHER_PAYMENT: 13,
    OTHER_PAYMENT_TOTAL: 14,
    PRESENT_TERM_SCHOOL_FEES_TOTAL: 16,
    USES_BUS: 17,
    BUS_FEE: 18,
    BUS_FAMILY_TOTAL: 19,
    INDIVIDUAL_CURRENT_TERM_TOTAL: 20,
    FAMILY_CURRENT_TERM_TOTAL: 21,
    INDIVIDUAL_BROUGHT_FORWARD: 22,
    FAMILY_BROUGHT_FORWARD: 23,
    INDIVIDUAL_TOTAL_DUE: 25,
    FAMILY_TOTAL_DUE: 26,
    PAYMENT_1_INDIVIDUAL: 27,
    PAYMENT_1_FAMILY: 28,
    INDIVIDUAL_OLD_OUTSTANDING: 29,
    PAYMENT_1_DATE: 30,
    PAYMENT_1_RECEIPT: 31,
    BALANCE_1_FAMILY: 32,
    BALANCE_1_INDIVIDUAL: 33,
    PAYMENT_2_INDIVIDUAL: 34,
    PAYMENT_2_FAMILY: 35,
    PAYMENT_2_DATE: 36,
    PAYMENT_2_RECEIPT: 37,
    BALANCE_2_FAMILY: 38,
    BALANCE_2_INDIVIDUAL: 39,
    PAYMENT_3_INDIVIDUAL: 40,
    PAYMENT_3_FAMILY: 41,
    PAYMENT_3_DATE: 42,
    PAYMENT_3_RECEIPT: 43,
    BALANCE_3_FAMILY: 44,
    BALANCE_3_INDIVIDUAL: 45
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

  FAMILY_ID: 5,

  AMOUNT_RECEIVED: 6,

  SCHOOL_FEES: 7,

  OUTSTANDING_SCHOOL_FEES: 8,

  REGISTRATION: 9,

  TEXTBOOKS: 10,

  // Column K = 11
  // Name and Class of Pupils/Student

  GRADUATION: 12,

  UNIFORMS: 13,

  SPORTS_WEAR: 14,

  FRIDAY_WEAR: 15,

  BUS_FEES: 16,

  LESSON: 17,

  COMMON_ENTRANCE: 18,

  JUNIOR_WAEC: 19,

  OTHERS: 20,

  TOTAL: 21,

  REMARK: 22
},


  // -------------------------
  // STUDENT MASTER
  // -------------------------

  STUDENT: {

  STUDENT_NO: 1,       // A - Nos
  STUDENT_NAME: 2,     // B - Name
  CLASS: 3,            // C - Class
  ADMISSION_YEAR: 4,   // D - Admission Year
  STUDENT_ID: 5,       // E - Student ID

  SCHOOL_FEE: 6,       // F
  LESSON: 7,           // G
  TEXTBOOK: 8,         // H
  NOTEBOOK: 9,         // I
  UNIFORM: 10,         // J
  SPORTS_WEAR: 11,     // K
  FRIDAY_WEAR: 12,     // L
  CLUBS: 13,           // M
  BUS: 14,             // N
  OTHERS: 15,          // O
  OUTSTANDING: 16,     // P
  TOTAL: 17,           // Q
  AMOUNT_PAID: 18,     // R
  AMOUNT_OWING: 19,    // S
  COMMENTS: 20         // T
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
      '2. Sync Student IDs & Numbers',
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

  migrateLegacyHistoryToPermanentIds();

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.ALLOCATION_LOG_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ALLOCATION_LOG_SHEET);

  // Keep the original 14 columns in place so existing allocation history is not shifted.
  const headers = [[
    'Allocation ID', 'Payment ID', 'Payment Date', 'Parent No',
    'Parent / Payee', 'Student No', 'Student Name', 'Class',
    'Category', 'Amount', 'Receipt No', 'Remarks', 'Entered By', 'Logged At',
    'Family ID', 'Student ID'
  ]];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.ALLOCATION_ENTRY_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ALLOCATION_ENTRY_SHEET);
  sheet.clear();

  sheet.getRange('A1').setValue('PAYMENT ALLOCATION').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setValue('Payment ID');
  sheet.getRange('A4').setValue('Payment Date');
  sheet.getRange('A5').setValue('Parent No');
  sheet.getRange('A6').setValue('Family ID');
  sheet.getRange('A7').setValue('Parent / Payee');
  sheet.getRange('A8').setValue('Amount Received');
  sheet.getRange('A9').setValue('Already Allocated');
  sheet.getRange('A10').setValue('Remaining To Allocate');

  const headers = [[
    'Student No', 'Student Name', 'Class', 'Amount Due',
    'Category', 'Category Available', 'Amount To Allocate', 'Remarks', 'Student ID'
  ]];
  sheet.getRange(12, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  sheet.setFrozenRows(12);
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

function findHeaderColumnInRows(
  sheet,
  header,
  startRow,
  endRow
) {

  const numRows =
    endRow - startRow + 1;


  const values =
    sheet
      .getRange(
        startRow,
        1,
        numRows,
        sheet.getLastColumn()
      )
      .getDisplayValues();


  for (
    let r = 0;
    r < values.length;
    r++
  ) {

    for (
      let c = 0;
      c < values[r].length;
      c++
    ) {

      if (
        String(values[r][c])
          .trim() === header
      ) {

        return c + 1;

      }

    }

  }


  return 0;

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


  if (!parentSheet) {

    throw new Error(
      'Parent/Student Financial Record sheet was not found.'
    );

  }


  if (!studentSheet) {

    throw new Error(
      'Student Master sheet was not found.'
    );

  }


  /************************************************************
   * FIND STUDENT ID COLUMNS
   ************************************************************/

 const parentStudentIdCol =
  CONFIG.PARENT.STUDENT_ID;


  const studentMasterIdCol =
    findHeaderColumn(
      studentSheet,
      'Student ID'
    );




  if (!studentMasterIdCol) {

    throw new Error(
      'Student ID column was not found on Student Master.'
    );

  }


  /************************************************************
   * READ DATA
   ************************************************************/

  const parentData =
    parentSheet
      .getDataRange()
      .getValues();


  const studentData =
    studentSheet
      .getDataRange()
      .getValues();


  /************************************************************
   * BUILD STUDENT MASTER LOOKUP
   *
   * Student ID -> Student information
   ************************************************************/

  const studentLookup = {};

  const duplicateStudentIds = {};


  for (
    let i = 1;
    i < studentData.length;
    i++
  ) {

    const studentId =
      normalizeId(
        studentData[i]
          [studentMasterIdCol - 1]
      );


    if (!studentId) {
      continue;
    }


    const studentNo =
      normalizeId(
        studentData[i]
          [CONFIG.STUDENT.STUDENT_NO - 1]
      );


    const studentName =
      studentData[i]
        [CONFIG.STUDENT.STUDENT_NAME - 1];


    const studentClass =
      studentData[i]
        [CONFIG.STUDENT.CLASS - 1];


    if (
      studentLookup[studentId]
    ) {

      duplicateStudentIds[studentId] = true;

      continue;

    }


    studentLookup[studentId] = {

      studentNo:
        studentNo,

      studentName:
        studentName,

      studentClass:
        studentClass,

      row:
        i + 1

    };

  }


  /************************************************************
   * PREPARE SYNC ISSUES SHEET
   ************************************************************/

  let issuesSheet =
    financeSS.getSheetByName(
      CONFIG.SYNC_ISSUES_SHEET
    );


  if (!issuesSheet) {

    issuesSheet =
      financeSS.insertSheet(
        CONFIG.SYNC_ISSUES_SHEET
      );

  }


  issuesSheet.clearContents();


  issuesSheet
    .getRange(
      1,
      1,
      1,
      7
    )
    .setValues([
      [
        'Issue Type',
        'Student ID',
        'Student Name',
        'Class',
        'Current Student No',
        'Master Student No',
        'Source Row'
      ]
    ]);


  const issues = [];


  let matched = 0;

  let unmatchedParent = 0;

  let missingStudentId = 0;

  let duplicateCount = 0;


  /************************************************************
   * UPDATE PARENT MASTER
   *
   * Rows 1-3 are headers
   * Data starts row 4
   ************************************************************/

  for (
    let i = 3;
    i < parentData.length;
    i++
  ) {

    const studentName =
      parentData[i]
        [CONFIG.PARENT.STUDENT_NAME - 1];


    if (
      !normalize(studentName)
    ) {
      continue;
    }


    const studentId =
      normalizeId(
        parentData[i]
          [parentStudentIdCol - 1]
      );


    const currentStudentNo =
      normalizeId(
        parentData[i]
          [CONFIG.PARENT.STUDENT_NO - 1]
      );


    const studentClass =
      parentData[i]
        [CONFIG.PARENT.STUDENT_CLASS - 1];


    if (!studentId) {

      missingStudentId++;


      issues.push([
        'MISSING STUDENT ID',
        '',
        studentName,
        studentClass,
        currentStudentNo,
        '',
        i + 1
      ]);


      continue;

    }


    if (
      duplicateStudentIds[studentId]
    ) {

      duplicateCount++;


      issues.push([
        'DUPLICATE STUDENT ID IN STUDENT MASTER',
        studentId,
        studentName,
        studentClass,
        currentStudentNo,
        '',
        i + 1
      ]);


      continue;

    }


    const match =
      studentLookup[studentId];


    if (!match) {

      unmatchedParent++;


      issues.push([
        'STUDENT ID NOT FOUND IN STUDENT MASTER',
        studentId,
        studentName,
        studentClass,
        currentStudentNo,
        '',
        i + 1
      ]);


      continue;

    }


    parentSheet
      .getRange(
        i + 1,
        CONFIG.PARENT.STUDENT_NO
      )
      .setValue(
        match.studentNo
      );


    matched++;

  }


  /************************************************************
   * REPORT DUPLICATE STUDENT IDS
   ************************************************************/

  Object.keys(
    duplicateStudentIds
  ).forEach(
    studentId => {

      const match =
        studentLookup[studentId];


      issues.push([
        'DUPLICATE STUDENT ID IN STUDENT MASTER',
        studentId,
        match
          ? match.studentName
          : '',
        match
          ? match.studentClass
          : '',
        '',
        match
          ? match.studentNo
          : '',
        match
          ? match.row
          : ''
      ]);

    }
  );


  /************************************************************
   * WRITE ISSUES
   ************************************************************/

  if (
    issues.length > 0
  ) {

    issuesSheet
      .getRange(
        2,
        1,
        issues.length,
        7
      )
      .setValues(
        issues
      );

  }


  /************************************************************
   * FINAL MESSAGE
   ************************************************************/

  const totalIssues =
    unmatchedParent +
    missingStudentId +
    duplicateCount;


  if (
    totalIssues > 0
  ) {

    SpreadsheetApp
      .getUi()
      .alert(
        'Student Number Sync Complete\n\n' +

        'Matched / Updated: ' +
        matched +
        '\n' +

        'Missing Student ID: ' +
        missingStudentId +
        '\n' +

        'Student ID Not Found: ' +
        unmatchedParent +
        '\n' +

        'Duplicate Student IDs: ' +
        duplicateCount +
        '\n\n' +

        'Please check the Sync_Issues sheet.'
      );

  } else {

    SpreadsheetApp
      .getUi()
      .alert(
        'Student Number Sync Complete\n\n' +

        'Matched / Updated: ' +
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
  const financeSS = SpreadsheetApp.getActiveSpreadsheet();
  const parentSheet = financeSS.getSheetByName(CONFIG.PARENT_STUDENT_SHEET);
  const studentSheet = getStudentWorkbook().getSheetByName(CONFIG.STUDENT_MASTER_SHEET);
  const parentData = parentSheet.getDataRange().getValues();
  const studentData = studentSheet.getDataRange().getValues();
  const financeLookup = {};

  for (let i = 3; i < parentData.length; i++) {
    const studentId = normalizeId(parentData[i][CONFIG.PARENT.STUDENT_ID - 1]);
    if (!studentId) continue;
    financeLookup[studentId] = {
      schoolFee: money(parentData[i][CONFIG.PARENT.SCHOOL_FEE - 1]),
      lesson: money(parentData[i][CONFIG.PARENT.LESSON - 1]),
      outstanding: money(parentData[i][CONFIG.PARENT.INDIVIDUAL_BROUGHT_FORWARD - 1])
    };
  }

  for (let i = 1; i < studentData.length; i++) {
    const studentId = normalizeId(studentData[i][CONFIG.STUDENT.STUDENT_ID - 1]);
    if (!studentId || !financeLookup[studentId]) continue;
    const row = i + 1;
    const record = financeLookup[studentId];
    studentSheet.getRange(row, CONFIG.STUDENT.SCHOOL_FEE).setValue(record.schoolFee);
    studentSheet.getRange(row, CONFIG.STUDENT.LESSON).setValue(record.lesson);
    studentSheet.getRange(row, CONFIG.STUDENT.OUTSTANDING).setValue(record.outstanding);
    studentSheet.getRange(row, CONFIG.STUDENT.TOTAL).setFormula('=IFERROR(SUM(F' + row + ':P' + row + '),"")');
  }
  SpreadsheetApp.getUi().alert('Fee information synced to Student Master using permanent Student ID.');
}

function prepareSyncIssuesSheetBothWays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SYNC_ISSUES_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SYNC_ISSUES_SHEET);
  sheet.clear();
  const headers = [[
    'Source', 'Parent Sheet Row', 'Student Master Row', 'Student Name', 'Class',
    'Student No', 'Student ID', 'Issue', 'Family ID'
  ]];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
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
  const financeSS = SpreadsheetApp.getActiveSpreadsheet();
  let baselineSheet = financeSS.getSheetByName(CONFIG.BASELINE_SHEET);
  if (!baselineSheet) baselineSheet = financeSS.insertSheet(CONFIG.BASELINE_SHEET);
  if (baselineSheet.getLastRow() > 1) return;
  baselineSheet.clear();
  baselineSheet.getRange(1, 1, 1, 5).setValues([[
    'Student ID', 'Student No', 'Student Name', 'Opening Amount Paid', 'Captured At'
  ]]).setFontWeight('bold');

  const studentSheet = getStudentWorkbook().getSheetByName(CONFIG.STUDENT_MASTER_SHEET);
  const data = studentSheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const studentId = normalizeId(data[i][CONFIG.STUDENT.STUDENT_ID - 1]);
    if (!studentId) continue;
    rows.push([
      studentId,
      data[i][CONFIG.STUDENT.STUDENT_NO - 1],
      data[i][CONFIG.STUDENT.STUDENT_NAME - 1],
      money(data[i][CONFIG.STUDENT.AMOUNT_PAID - 1]),
      new Date()
    ]);
  }
  if (rows.length) baselineSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  baselineSheet.hideSheet();
}



/************************************************************
 * MIGRATE LEGACY HISTORY TO PERMANENT IDS
 *
 * Runs safely more than once. It backfills permanent Student ID
 * and Family ID into existing baseline/allocation history while
 * the current display numbers are still available for matching.
 ************************************************************/
function migrateLegacyHistoryToPermanentIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parentSheet = ss.getSheetByName(CONFIG.PARENT_STUDENT_SHEET);
  const parentData = parentSheet.getDataRange().getValues();
  const familyIdCol =
  findHeaderColumnInRows(
    parentSheet,
    'Family ID',
    1,
    3
  );
  if (!familyIdCol) throw new Error('Family ID column was not found on the Parent/Student Financial Record.');

  const byStudentNo = {};
  let currentFamilyId = '';
  for (let i = 3; i < parentData.length; i++) {
    if (parentData[i][familyIdCol - 1]) currentFamilyId = normalizeId(parentData[i][familyIdCol - 1]);
    const studentNo = normalizeId(parentData[i][CONFIG.PARENT.STUDENT_NO - 1]);
    const studentId = normalizeId(parentData[i][CONFIG.PARENT.STUDENT_ID - 1]);
    if (studentNo && studentId) byStudentNo[studentNo] = { studentId, familyId: currentFamilyId };
  }

  const baseline = ss.getSheetByName(CONFIG.BASELINE_SHEET);
  if (baseline && baseline.getLastRow() > 1 && !findHeaderColumn(baseline, 'Student ID')) {
    baseline.insertColumnBefore(1);
    baseline.getRange(1, 1).setValue('Student ID');
    const data = baseline.getDataRange().getValues();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const oldStudentNo = normalizeId(data[i][1]);
      out.push([byStudentNo[oldStudentNo] ? byStudentNo[oldStudentNo].studentId : '']);
    }
    if (out.length) baseline.getRange(2, 1, out.length, 1).setValues(out);
  }

  createAllocationLog();
  const log = ss.getSheetByName(CONFIG.ALLOCATION_LOG_SHEET);
  if (log.getLastRow() > 1) {
    const data = log.getDataRange().getValues();
    const studentNoCol = findHeaderColumn(log, 'Student No');
    const familyIdLogCol = findHeaderColumn(log, 'Family ID');
    const studentIdLogCol = findHeaderColumn(log, 'Student ID');
    const familyOut = [], studentOut = [];
    for (let i = 1; i < data.length; i++) {
      const existingFamilyId = normalizeId(data[i][familyIdLogCol - 1]);
      const existingStudentId = normalizeId(data[i][studentIdLogCol - 1]);
      const studentNo = normalizeId(data[i][studentNoCol - 1]);
      const match = byStudentNo[studentNo];
      familyOut.push([existingFamilyId || (match ? match.familyId : '')]);
      studentOut.push([existingStudentId || (match ? match.studentId : '')]);
    }
    log.getRange(2, familyIdLogCol, familyOut.length, 1).setValues(familyOut);
    log.getRange(2, studentIdLogCol, studentOut.length, 1).setValues(studentOut);
  }
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
  name: 'Friday Wear',
  column:
    CONFIG.INFLOW.FRIDAY_WEAR
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

function getChildrenForParent(familyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PARENT_STUDENT_SHEET);
  const data = sheet.getDataRange().getValues();
  const target = normalizeId(familyId);
  const children = [];
  const seenStudents = new Set();
  let currentFamilyId = '';
  let currentParentName = '';
  let currentParentNo = '';

  for (let i = 3; i < data.length; i++) {
    const rawFamilyId = data[i][CONFIG.PARENT.FAMILY_ID - 1];
    const rawParentName = data[i][CONFIG.PARENT.PARENT_NAME - 1];
    const rawParentNo = data[i][CONFIG.PARENT.PARENT_NO - 1];
    if (rawFamilyId) currentFamilyId = rawFamilyId;
    if (rawParentName) currentParentName = rawParentName;
    if (rawParentNo) currentParentNo = rawParentNo;
    if (normalizeId(currentFamilyId) !== target) continue;

    const studentId = normalizeId(data[i][CONFIG.PARENT.STUDENT_ID - 1]);
    const studentNo = data[i][CONFIG.PARENT.STUDENT_NO - 1];
    const studentName = data[i][CONFIG.PARENT.STUDENT_NAME - 1];
    const studentClass = data[i][CONFIG.PARENT.STUDENT_CLASS - 1];
    const amountDue = money(data[i][CONFIG.PARENT.INDIVIDUAL_TOTAL_DUE - 1]);
    if (!studentId || !studentName || !studentClass) continue;
    if (seenStudents.has(studentId)) continue;
    seenStudents.add(studentId);
    children.push({ studentId, studentNo, studentName, studentClass, amountDue, parentName: currentParentName, parentNo: currentParentNo });
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


  const parentNo = inflowSheet.getRange(row, CONFIG.INFLOW.PARENT_NO).getDisplayValue();

  const familyId = normalizeId(
    inflowSheet.getRange(row, CONFIG.INFLOW.FAMILY_ID).getDisplayValue()
  );


  if (!familyId) {
    throw new Error('The selected payment has no permanent Family ID.');
  }


  const children =
    getChildrenForParent(
      familyId
    );


  if (!children.length) {

    throw new Error(
      'No children were found for Family ID: ' + familyId + '.\n\n' +
      'Check the Family ID on Daily Inflow and Parent/Student Financial Record.'
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
      8,
      6
    )
    .clearContent();


  const lastRow =
    Math.max(
      entrySheet.getLastRow(),
      12
    );


  if (lastRow >= 13) {

    entrySheet
      .getRange(
        13,
        1,
        lastRow - 12,
        9
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


  entrySheet.getRange('B6').setValue(familyId);
  entrySheet.getRange('B7').setValue(payee);
  entrySheet.getRange('B8').setValue(amountReceived);
  entrySheet.getRange('B9').setValue(existing.total);
  entrySheet.getRange('B10').setValue(remainingPayment);


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

            '',

            child.studentId

          ]);

      }
    );

  }
);


  if (rows.length) {

    entrySheet
      .getRange(
        13,
        1,
        rows.length,
        9
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrySheet = ss.getSheetByName(CONFIG.ALLOCATION_ENTRY_SHEET);
  const logSheet = ss.getSheetByName(CONFIG.ALLOCATION_LOG_SHEET);
  const inflowSheet = ss.getSheetByName(CONFIG.INFLOW_SHEET);

  const paymentId = entrySheet.getRange('B3').getDisplayValue().trim();
  const paymentDate = entrySheet.getRange('B4').getValue();
  const parentNo = entrySheet.getRange('B5').getDisplayValue();
  const familyId = normalizeId(entrySheet.getRange('B6').getDisplayValue());
  const payee = entrySheet.getRange('B7').getDisplayValue();
  const remainingBefore = money(entrySheet.getRange('B10').getValue());

  if (!paymentId) throw new Error('No payment is currently loaded in Allocation_Entry.');
  if (!familyId) throw new Error('This payment has no permanent Family ID.');
  if (remainingBefore <= CONFIG.TOLERANCE) throw new Error('This payment is already fully allocated.');

  const lastRow = entrySheet.getLastRow();
  if (lastRow < 13) throw new Error('There are no allocation rows.');
  const data = entrySheet.getRange(13, 1, lastRow - 12, 9).getValues();
  const allocations = [];
  const categoryTotals = {};
  let allocationTotal = 0;

  data.forEach(row => {
    const studentNo = row[0];
    const studentName = row[1];
    const studentClass = row[2];
    const category = row[4];
    const categoryAvailable = money(row[5]);
    const amount = money(row[6]);
    const remarks = row[7];
    const studentId = normalizeId(row[8]);
    if (amount <= 0) return;
    if (!studentId) throw new Error('An allocation amount was entered for ' + studentName + ' but the student has no permanent Student ID.');
    if (amount > categoryAvailable + CONFIG.TOLERANCE) throw new Error(studentName + ' has an allocation greater than the available amount for ' + category + '.');

    allocations.push({ studentId, studentNo, studentName, studentClass, category, amount, remarks });
    allocationTotal += amount;
    categoryTotals[category] = (categoryTotals[category] || 0) + amount;
  });

  if (!allocations.length) throw new Error('Enter at least one allocation amount.');
  if (allocationTotal > remainingBefore + CONFIG.TOLERANCE) {
    throw new Error('You are attempting to allocate ₦' + allocationTotal.toLocaleString() + ' but only ₦' + remainingBefore.toLocaleString() + ' remains on this payment.');
  }

  const availableByCategory = {};
  data.forEach(row => {
    const category = row[4];
    const available = money(row[5]);
    if (category && available > 0) availableByCategory[category] = available;
  });
  Object.keys(categoryTotals).forEach(category => {
    if (categoryTotals[category] > availableByCategory[category] + CONFIG.TOLERANCE)
      throw new Error('Total allocation for ' + category + ' exceeds the amount available.');
  });

  const paymentIdCol = findHeaderColumn(inflowSheet, 'Payment ID');
  const receiptCol = findHeaderColumn(inflowSheet, 'Receipt No');
  let receiptNo = '';
  const inflowData = inflowSheet.getDataRange().getValues();
  for (let i = 1; i < inflowData.length; i++) {
    if (normalize(inflowData[i][paymentIdCol - 1]) === normalize(paymentId)) {
      receiptNo = inflowData[i][receiptCol - 1];
      break;
    }
  }

  const enteredBy = Session.getActiveUser().getEmail() || 'Unknown User';
  const logRows = allocations.map(a => [
    'AL-' + Utilities.getUuid().substring(0, 10).toUpperCase(),
    paymentId, paymentDate, parentNo, payee, a.studentNo, a.studentName, a.studentClass,
    a.category, a.amount, receiptNo, a.remarks, enteredBy, new Date(), familyId, a.studentId
  ]);
  logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 16).setValues(logRows);

  refreshAllBalances();
  const paymentRow = findPaymentRow(paymentId);
  if (paymentRow) {
    inflowSheet.getRange(paymentRow, 1).activate();
    prepareSelectedPayment();
  }
  SpreadsheetApp.getUi().alert('Allocation posted successfully.\n\nAllocated now: ₦' + allocationTotal.toLocaleString());
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

  migrateLegacyHistoryToPermanentIds();

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
   * Find Payment_Allocation columns by header.
   * This avoids problems when Family ID,
   * Student ID or other columns are inserted.
   */

  const logPaymentIdCol =
    findHeaderColumn(
      logSheet,
      'Payment ID'
    );


  const logAmountCol =
    findHeaderColumn(
      logSheet,
      'Amount'
    );


  if (
    !logPaymentIdCol ||
    !logAmountCol
  ) {

    throw new Error(
      'Payment_Allocation must contain Payment ID and Amount columns.'
    );

  }


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
          allocationData[i]
            [logPaymentIdCol - 1]
        );


      const amount =
        money(
          allocationData[i]
            [logAmountCol - 1]
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
  const financeSS = SpreadsheetApp.getActiveSpreadsheet();
  let baselineSheet = financeSS.getSheetByName(CONFIG.BASELINE_SHEET);
  if (!baselineSheet || baselineSheet.getLastRow() < 2) {
    createPaymentBaseline();
    baselineSheet = financeSS.getSheetByName(CONFIG.BASELINE_SHEET);
  }

  const allocationSheet = financeSS.getSheetByName(CONFIG.ALLOCATION_LOG_SHEET);
  const baselineData = baselineSheet.getDataRange().getValues();
  const openingPaid = {};
  const baselineStudentIdCol = findHeaderColumn(baselineSheet, 'Student ID');
  const baselineOpeningCol = findHeaderColumn(baselineSheet, 'Opening Amount Paid');
  if (!baselineStudentIdCol || !baselineOpeningCol) {
    throw new Error('The payment baseline has not yet been migrated to permanent Student IDs. Run Initial Setup once.');
  }
  for (let i = 1; i < baselineData.length; i++) {
    const studentId = normalizeId(baselineData[i][baselineStudentIdCol - 1]);
    if (studentId) openingPaid[studentId] = money(baselineData[i][baselineOpeningCol - 1]);
  }

  const newPayments = {};
  if (allocationSheet.getLastRow() >= 2) {
    const allocationData = allocationSheet.getDataRange().getValues();
    const logStudentIdCol = findHeaderColumn(allocationSheet, 'Student ID');
    const logAmountCol = findHeaderColumn(allocationSheet, 'Amount');
    for (let i = 1; i < allocationData.length; i++) {
      const studentId = normalizeId(allocationData[i][logStudentIdCol - 1]);
      if (!studentId) continue;
      newPayments[studentId] = (newPayments[studentId] || 0) + money(allocationData[i][logAmountCol - 1]);
    }
  }

  const studentSheet = getStudentWorkbook().getSheetByName(CONFIG.STUDENT_MASTER_SHEET);
  const studentData = studentSheet.getDataRange().getValues();
  for (let i = 1; i < studentData.length; i++) {
    const row = i + 1;
    const studentId = normalizeId(studentData[i][CONFIG.STUDENT.STUDENT_ID - 1]);
    if (!studentId) continue;
    const amountPaid = (openingPaid[studentId] || 0) + (newPayments[studentId] || 0);
    studentSheet.getRange(row, CONFIG.STUDENT.AMOUNT_PAID).setValue(amountPaid);
    studentSheet.getRange(row, CONFIG.STUDENT.AMOUNT_OWING).setFormula('=IFERROR(Q' + row + '-R' + row + ',"")');
  }
}
