require('dotenv').config();
const express = require('express');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const exceljs = require('exceljs');
const Sequelize = require('sequelize');

const app = express();
const port = 3000;

const uploadDir = path.join(__dirname, 'public/uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'public/uploads')));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Ensure the database directory exists
const dbDir = path.join(__dirname, 'db');
if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir);
}

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const sequelize = new Sequelize({
  dialect: 'sqlite',
  storage: path.join(dbDir, 'database.sqlite')
});

// Define your models here (User, Memo, Feedback)
const User = sequelize.define('User', {
  firstName: {
    type: Sequelize.STRING,
    allowNull: false
  },
  lastName: {
    type: Sequelize.STRING,
    allowNull: false
  },
  email: {
    type: Sequelize.STRING,
    allowNull: false,
    unique: true
  },
  password: {
    type: Sequelize.STRING,
    allowNull: false
  },
  isVerified: {
    type: Sequelize.BOOLEAN,
    defaultValue: false
  },
  securityQuestion: {
    type: Sequelize.STRING,
    allowNull: false
  },
  securityAnswer: {
    type: Sequelize.STRING,
    allowNull: false
  },
  isApproved: {
    type: Sequelize.BOOLEAN,
    defaultValue: false
  },
  resetToken: {
    type: Sequelize.STRING,
    allowNull: true
  },
  resetTokenExpiration: {
    type: Sequelize.DATE,
    allowNull: true
  }
}, {
  timestamps: true
});

const Memo = sequelize.define('Memo', {
  content: {
    type: Sequelize.TEXT,
    allowNull: false
  },
  classification: {
    type: Sequelize.STRING,
    allowNull: false
  },
  status: {
    type: Sequelize.STRING,
    allowNull: false
  },
  assigned_department: {
    type: Sequelize.STRING,
    allowNull: true
  },
  created_at: {
    type: Sequelize.DATE,
    allowNull: false
  }
}, {
  timestamps: true
});

const Feedback = sequelize.define('Feedback', {
  memo_id: {
    type: Sequelize.INTEGER,
    references: {
      model: Memo,
      key: 'id'
    }
  },
  user_classification: {
    type: Sequelize.STRING,
    allowNull: false
  },
  correct_classification: {
    type: Sequelize.STRING,
    allowNull: true
  },
  created_at: {
    type: Sequelize.DATE,
    allowNull: false
  }
}, {
  timestamps: true
});

let loggedErrors = new Set();

function resetLoggedErrors() {
  loggedErrors = new Set();
  console.log('Reset loggedErrors set.');
}

function highlightCell(worksheet, column, row_id, error, color) {
  const cell = worksheet.getCell(`${column}${row_id}`);
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: color }
  };
  if (cell.note) {
    cell.note += ' | ' + error;
  } else {
    cell.note = error;
  }
}

function logAndAddErrorUnique(errors, column, error, row_id, universeSheet, worksheet, color) {
  let errorIdentifier = `${column}-${row_id}-${error}`;
  if (!loggedErrors.has(errorIdentifier)) {
    errors.push({ column, row_id, error });
    loggedErrors.add(errorIdentifier);
    console.log(`Logging error for Row ${row_id}, Column ${column}: ${error}`);
    if (universeSheet) {
      universeSheet.addRow([column, row_id.toString(), error]);
    } else {
      console.error('Universe sheet is undefined!');
    }
    if (column) {
      highlightCell(worksheet, column, row_id, error, color);
    }
  } else {
    console.log(`Duplicate error detected for Row ${row_id}, Column ${column}: ${error}`);
  }
}

function isValidDate(dateString) {
  const regex = /^\d{4}\/\d{2}\/\d{2}$/;
  if (!regex.test(dateString)) return false;
  const [year, month, day] = dateString.split('/').map(Number);
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() + 1 === month && date.getDate() === day;
}

function isValidTimeFormat(timeString) {
  const regex = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]$/;
  return regex.test(timeString);
}

function calculateDateDifference(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const diffTime = Math.abs(end - start);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return diffDays;
}

function validateFirstName(firstName, rowNumber, errors, universeSheet, worksheet) {
  if (!firstName || firstName.trim() === '') {
    logAndAddErrorUnique(errors, 'A', 'Invalid value. Please ensure the Enrollee First Name field is not left blank.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    if (/\d/.test(firstName)) {
      logAndAddErrorUnique(errors, 'A', 'Invalid value. The Enrollee First Name field cannot contain numeric values. Please enter a valid first name.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
    if (firstName.length > 50) {
      logAndAddErrorUnique(errors, 'A', 'Invalid. Exceeds Character Length Limit', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validateLastName(lastName, rowNumber, errors, universeSheet, worksheet) {
  if (!lastName || lastName.trim() === '') {
    logAndAddErrorUnique(errors, 'B', 'Invalid value. Please ensure the Enrollee Last Name field is not left blank.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    if (/\d/.test(lastName)) {
      logAndAddErrorUnique(errors, 'B', 'Invalid value. The Enrollee Last Name field cannot contain numeric values. Please enter a valid last name.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
    if (lastName.length > 50) {
      logAndAddErrorUnique(errors, 'B', 'Invalid. Exceeds Character Length Limit', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validateEnrolleeID(enrolleeID, rowNumber, errors, universeSheet, worksheet) {
  if (!enrolleeID) {
    logAndAddErrorUnique(errors, 'C', 'Invalid value. The Medicare Beneficiary Identifier field cannot be left blank. Please enter a valid identifier.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (enrolleeID.length !== 11) {
    logAndAddErrorUnique(errors, 'C', 'Invalid value. The Medicare Beneficiary Identifier (MBI) must not exceed 11 characters. Please review and ensure the MBI is correctly entered.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (!/^[A-Z0-9]{11}$/.test(enrolleeID)) {
    logAndAddErrorUnique(errors, 'C', 'Invalid value. Please ensure the Medicare Beneficiary Identifier (MBI) field contains the correct format and does not exceed 11 characters.', rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateContractID(contractID, rowNumber, errors, universeSheet, worksheet) {
  if (!contractID) {
    logAndAddErrorUnique(errors, 'D', 'Invalid value. The Contract ID field cannot be left blank. Please enter a valid Contract ID', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (!/^[A-Z][0-9]{4}$/.test(contractID)) {
    if (contractID.length !== 5) {
      logAndAddErrorUnique(errors, 'D', 'Character length issue detected. Please ensure the input meets the specified length requirements.', rowNumber, universeSheet, worksheet, 'FFFF00');
    } else {
      logAndAddErrorUnique(errors, 'D', 'Invalid Contract ID. Contract IDs must begin with a letter followed by three numbers. Please correct the entry accordingly.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validatePBP(pbp, rowNumber, errors, universeSheet, worksheet) {
  if (!pbp) {
    logAndAddErrorUnique(errors, 'E', 'Invalid value. The Plan Benefit Package field cannot be left blank. Please enter a valid Plan Benefit Package ID.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    if (pbp.length > 3 || pbp.trim().length !== 3) {
      logAndAddErrorUnique(errors, 'E', 'Invalid value. The Plan Benefit Package field cannot exceed 3 characters. Please review and ensure the entry meets the specified requirements.', rowNumber, universeSheet, worksheet, 'FFFF00');
    } else if (!/^\d+$/.test(pbp)) {
      logAndAddErrorUnique(errors, 'E', 'Invalid value. Please ensure the Plan Benefit Package field contains the correct format and meets the specified requirements', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validateFirstTierEntity(firstTierEntity, rowNumber, errors, universeSheet, worksheet) {
  if (!firstTierEntity || firstTierEntity.trim() === '') {
    logAndAddErrorUnique(errors, 'F', 'Invalid value. The First Tier, Downstream, and Related Entity field cannot be left blank. Please enter a valid value for this field.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (firstTierEntity.trim().length > 70) {
    logAndAddErrorUnique(errors, 'F', 'Invalid value. The First Tier, Downstream, and Related Entity field cannot exceed 70 characters. Please ensure the entry is within the specified limit.', rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateAuthOrClaimNumber(authOrClaimNumber, rowNumber, errors, universeSheet, worksheet) {
  if (!authOrClaimNumber) {
    logAndAddErrorUnique(errors, 'G', 'Invalid value. The Authorization or Claim Number field cannot be left blank. Please provide a valid authorization or claim number.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    if (authOrClaimNumber.length > 40) {
      logAndAddErrorUnique(errors, 'G', 'Invalid value. The Authorization or Claim Number cannot exceed 40 characters. Please ensure the entry is within the specified character limit.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
    if (!/^[\w\s]+$/.test(authOrClaimNumber)) {
      logAndAddErrorUnique(errors, 'G', 'Invalid value. Please ensure the Authorization or Claim Number field contains the correct format and valid information.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validateDateReceived(dateReceived, rowNumber, errors, universeSheet, worksheet) {
  if (!dateReceived) {
    logAndAddErrorUnique(errors, 'H', 'Invalid value. The Received Date field cannot be left blank. Please enter the date the request was received', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    if (dateReceived.length !== 10) {
      logAndAddErrorUnique(errors, 'H', 'Invalid value. The field should contain exactly 10 characters. Please ensure the input length is correct.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
    if (!/^\d{4}\/\d{2}\/\d{2}$/.test(dateReceived)) {
      logAndAddErrorUnique(errors, 'H', 'Invalid date format. Please ensure the date is entered in the correct format', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
    if (dateReceived === 'None') {
      logAndAddErrorUnique(errors, 'H', 'Invalid value. Please enter the correct date when the request was received.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validateTimeReceived(timeReceived, columnN, partBDrugRequest, rowNumber, errors, universeSheet, worksheet) {
  if (columnN === 'S' && partBDrugRequest === 'N') {
    if (timeReceived.trim() !== 'None') {
      logAndAddErrorUnique(errors, 'I', 'Invalid Value. Time is not required for Standard non Part B Cases.', rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  } else if (columnN === 'S' && partBDrugRequest === 'Y') {
    if (timeReceived.trim() === 'None' || timeReceived.trim().length !== 8) {
      logAndAddErrorUnique(errors, 'I', 'Invalid Value. Time is required for Standard Part B Request.', rowNumber, universeSheet, worksheet, 'FFFF00');
    } else {
      if (!/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]$/.test(timeReceived)) {
        logAndAddErrorUnique(errors, 'I', 'Invalid Value. The time format for Standard Part B Requests is incorrect. Please ensure the time is entered in the correct format.', rowNumber, universeSheet, worksheet, 'FFFF00');
      }
      if (timeReceived.trim().length > 8) {
        logAndAddErrorUnique(errors, 'I', 'Invalid value. Please ensure the character length meets the specified requirements.', rowNumber, universeSheet, worksheet, 'FFFF00');
      }
    }
  } else if (columnN === 'E') {
    if (timeReceived.trim() === 'None' && partBDrugRequest === 'N') {
      logAndAddErrorUnique(errors, 'I', 'Invalid Value. None is not a correct response for expedited requests.', rowNumber, universeSheet, worksheet, 'FFFF00');
    } else if (timeReceived.trim() === 'None' || timeReceived.trim().length !== 8) {
      logAndAddErrorUnique(errors, 'I', 'Invalid value. Time is required for Expedited Requests. Please ensure the time is provided.', rowNumber, universeSheet, worksheet, 'FFFF00');
    } else {
      if (!/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(timeReceived)) {
        logAndAddErrorUnique(errors, 'I', 'Invalid value. The time format for Expedited Part B Requests is incorrect. Please ensure the time is entered in the correct format.', rowNumber, universeSheet, worksheet, 'FFFF00');
      }
      if (timeReceived.trim().length > 8) {
        logAndAddErrorUnique(errors, 'I', 'Invalid value. Please ensure the character length meets the specified requirements.', rowNumber, universeSheet, worksheet, 'FFFF00');
      }
    }
  }
}

function validatePartBDrugRequest(partBDrugRequest, rowNumber, errors, universeSheet, worksheet) {
  if (partBDrugRequest.trim() === '') {
    logAndAddErrorUnique(errors, 'J', 'Invalid Value. The field cannot be left blank. Please enter Y for Yes or N for No.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    let errorMessages = [];

    if (partBDrugRequest.trim() !== 'Y' && partBDrugRequest.trim() !== 'N') {
      errorMessages.push('Invalid value. The acceptable response for this field is either Y for Yes or N for No. Please provide a valid response.');
    }

    if (partBDrugRequest.trim().length !== 1) {
      errorMessages.push('Invalid value. The input exceeds the allowed character length.');
    }

    if (errorMessages.length > 0) {
      logAndAddErrorUnique(errors, 'J', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validateAORDate(columnK, columnV, columnM, rowNumber, errors, universeSheet, worksheet) {
  if (!columnK.trim()) {
    logAndAddErrorUnique(errors, 'K', 'Invalid value. The AOR Receipt Date field is blank. Please review and ensure it is filled with the appropriate date', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    let errorMessages = [];

    if (['E', 'CP', 'NCP'].includes(columnV.trim())) {
      if (columnK.trim() !== "None") {
        errorMessages.push('Invalid value. The field should display None if the request is received from the Enrollee, Contracted Provider, or Non-Contracted Provider.');
      }
    } else if (columnV.trim() === 'ER' && columnM.trim() === 'Denied' && columnK.trim() === 'None') {
      errorMessages.push('Invalid value. A valid AOR Date is required when the request is approved or denied and received from the ER.');
    } else {
      if (!isValidDate(columnK.trim())) {
        errorMessages.push('Invalid date format. Please ensure the date is entered in the correct format.');
      }
    }

    if (columnM.trim() === 'Dismissed' && columnK.trim() !== "None") {
      errorMessages.push('Invalid value. The field should display None when the case is dismissed.');
    }

    if (columnK.trim().length > 10) {
      errorMessages.push('Invalid value. The field exceeds the allowed character length.');
    }

    if (errorMessages.length > 0) {
      logAndAddErrorUnique(errors, 'K', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
    }
  }
}

function validateAORReceiptTime(columnL, columnV, columnN, columnM, columnJ, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!columnL.trim()) {
    logAndAddErrorUnique(errors, 'L', 'Invalid value. The Time of AOR Receipt field is blank. Please review and ensure it is filled with the appropriate time.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else {
    if (['E', 'CP', 'NCP'].includes(columnV.trim())) {
      if (columnL.trim() !== 'None') {
        errorMessages.push('Invalid Value. Case Received from E, CP or NCP, field should Display None');
      }
    } else if (columnV === 'ER' && columnN === 'E' && ['Approved', 'Denied'].includes(columnM)) {
      if (!isValidTimeFormat(columnL.trim()) || columnL.trim() === 'None' || /[!@#$%^&*(),]/.test(columnL.trim())) {
        errorMessages.push('Invalid Value. Expedited Case Received from the Enrollee Representative, Missing Time the AOR was received');
      }
    } else if (columnV === 'ER' && columnN === 'S' && columnJ === 'Y' && ['Approved', 'Denied'].includes(columnM)) {
      if (!isValidTimeFormat(columnL.trim()) || columnL.trim() === 'None' || /[!@#$%^&*(),]/.test(columnL.trim())) {
        errorMessages.push('Invalid Value. Expedited Case Received from the Enrollee Representative, Missing Time the AOR was received');
      }
    } else if (columnV === 'ER' && columnN === 'S' && columnJ === 'N') {
      if (columnL.trim() !== 'None') {
        errorMessages.push('Invalid Value. For Standard Non Part B Case, Time the AOR was received should Display None');
      }
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'L', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateRequestDetermination(columnM, rowNumber, errors, universeSheet, worksheet) {
  if (!columnM.trim()) {
    logAndAddErrorUnique(errors, 'M', 'Invalid value. The Request Determination field cannot be left blank.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (!['Approved', 'Denied', 'Dismissed'].includes(columnM.trim())) {
    logAndAddErrorUnique(errors, 'M', 'Invalid value. The Request Determination field should be either Approved, Denied, or Dismissed.', rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateRequestProcessing(columnN, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!columnN.trim()) {
    errorMessages.push('Invalid. Acceptable options are "S" for Standard or "E" for Expedited');
  } else {
    if (columnN.trim() !== 'S' && columnN.trim() !== 'E') {
      errorMessages.push("Invalid value. Acceptable options are 'S' for Standard or 'E' for Expedited.");
    }
    if (columnN.trim().length > 1) {
      errorMessages.push('Invalid. Exceeds character limit of 1');
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'N', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateTimeframeExtension(columnO, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!columnO.trim()) {
    errorMessages.push('Invalid. Acceptable options are "Y" for Yes and "N" for No.');
  } else if (columnO.trim() !== 'Y' && columnO.trim() !== 'N') {
    errorMessages.push('Invalid. Acceptable options are "Y" for Yes and "N" for No.');
  } else if (columnO.trim().length > 1) {
    errorMessages.push('Invalid. Exceeds character limit.');
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'O', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateDateOfDetermination(columnP, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!columnP.trim()) {
    errorMessages.push('Invalid value. The Date of Determination field cannot be left blank.');
  } else {
    if (!isValidDate(columnP.trim())) {
      errorMessages.push('Invalid date format. Use CCYY/MM/DD format (e.g., 2020/01/01).');
    }
    if (columnP.trim().length !== 10) {
      errorMessages.push('Invalid Character Length. The provided value does not match the permitted length of 10 characters. Please ensure that the entry contains exactly 10 characters.');
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'P', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateTimeOfDetermination(columnQ, columnN, columnJ, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!columnQ.trim()) {
    errorMessages.push('Invalid value. The Time of Determination field cannot be left blank.');
  } else {
    if (columnN.trim() === 'S' && columnJ.trim() === 'N') {
      if (columnQ.trim() !== 'None') {
        errorMessages.push('Invalid value. For Standard non Part B Cases, time of determination is not required.');
      }
      if (columnQ.trim().length !== 4) {
        errorMessages.push('Error with Character Length');
      }
    } else if (columnN.trim() === 'S' && columnJ.trim() === 'Y') {
      if (columnQ.trim() === 'None') {
        errorMessages.push('Invalid value. For Standard Part B Requests, time of determination is required.');
      }
      if (!isValidTimeFormat(columnQ.trim())) {
        errorMessages.push('Invalid time format. Use HH:MM:SS format (e.g., 13:45:00).');
      }
      if (columnQ.trim().length !== 8) {
        errorMessages.push('Invalid Value. Error with Character Length');
      }
    } else if (columnN.trim() === 'E') {
      if (columnQ.trim() === 'None') {
        errorMessages.push('Invalid value. For Expedited Requests, time of determination is required.');
      }
      if (!isValidTimeFormat(columnQ.trim())) {
        errorMessages.push('Invalid time format. Use HH:MM:SS format (e.g., 13:45:00).');
      }
      if (columnQ.trim().length !== 8) {
        errorMessages.push('Invalid Value. Error with Character Length');
      }
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'Q', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateDateOralNotification(dateOralNotification, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!dateOralNotification.trim()) {
    errorMessages.push('Invalid value. The Date of Oral Notification field cannot be left blank.');
  } else {
    if (dateOralNotification.trim() === 'None') {
      if (dateOralNotification.trim().length !== 4) {
        errorMessages.push('Invalid Character Length. The provided value does not match the permitted length of 4 characters. Please ensure that the entry contains exactly 4 characters.');
      }
    } else {
      if (!isValidDate(dateOralNotification.trim())) {
        errorMessages.push('Invalid date format. Use CCYY/MM/DD format (e.g., 2020/01/01).');
      }
      if (dateOralNotification.trim().length !== 10) {
        errorMessages.push('Invalid Character Length. The provided value does not match the permitted length of 10 characters. Please ensure that the entry contains exactly 10 characters.');
      }
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'R', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateTimeOfOralNotification(rowData, rowNumber, errors, universeSheet, worksheet) {
  const timeOfOralNotification = rowData.getCell(19).text.trim();
  const processingType = rowData.getCell(14).text.trim();
  const partBDrugRequest = rowData.getCell(10).text.trim();

  if (!timeOfOralNotification.trim()) {
    logAndAddErrorUnique(errors, 'S', "The 'Time of Oral Notification' field cannot be empty.", rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (timeOfOralNotification.trim().toLowerCase() === "none") {
    console.log(`Row ${rowNumber}: 'None' is accepted for 'Time of Oral Notification'.`);
  } else {
    if (!isValidTimeFormat(timeOfOralNotification.trim())) {
      logAndAddErrorUnique(errors, 'S', "The 'Time of Oral Notification' must be in the correct format HH:MM:SS in military time.", rowNumber, universeSheet, worksheet, 'FFFF00');
    }

    if (timeOfOralNotification.trim().length !== 8) {
      logAndAddErrorUnique(errors, 'S', "The 'Time of Oral Notification' must be exactly 8 characters long in the format HH:MM:SS.", rowNumber, universeSheet, worksheet, 'FFFF00');
    }

    if (processingType === "S" && partBDrugRequest === "N") {
      if (timeOfOralNotification.trim().toLowerCase() !== "none") {
        logAndAddErrorUnique(errors, 'S', "For Standard non Part B Cases, 'Time of Oral Notification' should be 'None'.", rowNumber, universeSheet, worksheet, 'FFFF00');
      }
    }
  }
}

function validateNotificationMethod(notificationMethod, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!notificationMethod.trim()) {
    errorMessages.push("The 'Notification Method' field cannot be blank.");
  } else if (notificationMethod.trim().length > 3) {
    errorMessages.push("Invalid Character Length. The value exceeds the allowed limit of 3 characters.");
  } else if (!['Oral', 'Written'].includes(notificationMethod.trim())) {
    errorMessages.push("Invalid value. The 'Notification Method' field should be either 'Oral' or 'Written'.");
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'T', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateNotificationDate(notificationDate, rowNumber, errors, universeSheet, worksheet) {
  if (!notificationDate.trim()) {
    logAndAddErrorUnique(errors, 'U', 'Invalid value. The Notification Date field is blank. Please ensure it is filled with the correct date.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (!isValidDate(notificationDate.trim())) {
    logAndAddErrorUnique(errors, 'U', 'Invalid value. The Notification Date field should be in the format CCYY/MM/DD. Please correct the entry accordingly.', rowNumber, universeSheet, worksheet, 'FFFF00');
  } else if (notificationDate.trim().length !== 10) {
    logAndAddErrorUnique(errors, 'U', 'Invalid Character Length. The provided value does not match the permitted length of 10 characters. Please ensure that the entry contains exactly 10 characters.', rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateRequestReceivedFrom(requestReceivedFrom, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!requestReceivedFrom.trim()) {
    errorMessages.push("The 'Request Received From' field cannot be blank.");
  } else {
    const validValues = ["ER", "E", "CP", "NCP"];
    if (!validValues.includes(requestReceivedFrom.trim())) {
      errorMessages.push("Invalid value. The 'Request Received From' field should be either 'ER', 'E', 'CP', or 'NCP'.");
    }

    if (requestReceivedFrom.trim().length !== 3) {
      errorMessages.push("Invalid Character Length. The value exceeds the allowed limit of 3 characters.");
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'V', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateDescriptionOfExtension(descriptionOfExtension, columnO, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!descriptionOfExtension.trim()) {
    if (columnO.trim() === "Y") {
      errorMessages.push("Invalid Value. The 'Description of the Extension' field cannot be blank if 'Timeframe Extension' is 'Y'.");
    }
  } else {
    if (columnO.trim() === "Y" && descriptionOfExtension.trim() === "None") {
      errorMessages.push("Invalid value. The 'Description of the Extension' field cannot be 'None' if 'Timeframe Extension' is 'Y'.");
    }

    if (descriptionOfExtension.trim().length > 500) {
      errorMessages.push("Invalid Character Length. The value exceeds the allowed limit of 500 characters.");
    }

    if (columnO.trim() === "N" && descriptionOfExtension.trim() !== "None") {
      errorMessages.push("Invalid value. The 'Description of the Extension' field should be 'None' if 'Timeframe Extension' is 'N'.");
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'W', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateSupportingDocumentation(supportingDocumentation, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!supportingDocumentation.trim()) {
    errorMessages.push("The 'Supporting Documentation' field cannot be blank.");
  } else if (supportingDocumentation.trim().length > 5) {
    errorMessages.push("Invalid Character Length. The value exceeds the allowed limit of 5 characters.");
  } else if (!['Yes', 'No'].includes(supportingDocumentation.trim())) {
    errorMessages.push("Invalid value. The 'Supporting Documentation' field should be either 'Yes' or 'No'.");
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'X', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function validateProcessNote(processNote, rowNumber, errors, universeSheet, worksheet) {
  let errorMessages = [];

  if (!processNote.trim()) {
    errorMessages.push("The 'Process Note' field cannot be blank.");
  } else if (processNote.trim().length > 2000) {
    errorMessages.push("Invalid Character Length. The value exceeds the allowed limit of 2000 characters.");
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'Y', errorMessages.join(' | '), rowNumber, universeSheet, worksheet, 'FFFF00');
  }
}

function calculateTimeliness(rowNumber, requestProcessing, dateReceived, timeReceived, dateOfDetermination, timeOfDetermination, timelinessSheet, worksheet) {
  let errors = [];
  let color = 'FFFF00';

  let errorMessages = [];

  if (!requestProcessing.trim()) {
    errorMessages.push('Invalid value. The Processing Type field cannot be left blank.');
  } else if (requestProcessing.trim() === 'S') {
    if (!isValidDate(dateReceived.trim())) {
      errorMessages.push('Invalid Received Date. Please ensure the date is entered in the correct format.');
    }

    if (timeReceived.trim() !== 'None' && !isValidTimeFormat(timeReceived.trim())) {
      errorMessages.push('Invalid Time Received. Please ensure the time is entered in the correct format.');
    }

    if (!isValidDate(dateOfDetermination.trim())) {
      errorMessages.push('Invalid Date of Determination. Please ensure the date is entered in the correct format.');
    }

    if (timeOfDetermination.trim() !== 'None' && !isValidTimeFormat(timeOfDetermination.trim())) {
      errorMessages.push('Invalid Time of Determination. Please ensure the time is entered in the correct format.');
    }

    if (errorMessages.length === 0) {
      const receivedDate = new Date(dateReceived.trim());
      const determinationDate = new Date(dateOfDetermination.trim());

      const receivedTime = timeReceived.trim() !== 'None' ? timeReceived.trim().split(':').map(Number) : null;
      const determinationTime = timeOfDetermination.trim() !== 'None' ? timeOfDetermination.trim().split(':').map(Number) : null;

      let receivedDateTime = receivedDate;
      let determinationDateTime = determinationDate;

      if (receivedTime) {
        receivedDateTime.setHours(receivedTime[0], receivedTime[1], receivedTime[2]);
      }

      if (determinationTime) {
        determinationDateTime.setHours(determinationTime[0], determinationTime[1], determinationTime[2]);
      }

      const diffMs = determinationDateTime - receivedDateTime;
      const diffDays = diffMs / (1000 * 60 * 60 * 24);

      if (diffDays <= 14) {
        errorMessages.push('Request processed within the required timeframe.');
      } else {
        errorMessages.push(`Request processing exceeded the required timeframe by ${diffDays - 14} days.`);
      }
    }
  } else if (requestProcessing.trim() === 'E') {
    if (!isValidDate(dateReceived.trim())) {
      errorMessages.push('Invalid Received Date. Please ensure the date is entered in the correct format.');
    }

    if (!isValidTimeFormat(timeReceived.trim())) {
      errorMessages.push('Invalid Time Received. Please ensure the time is entered in the correct format.');
    }

    if (!isValidDate(dateOfDetermination.trim())) {
      errorMessages.push('Invalid Date of Determination. Please ensure the date is entered in the correct format.');
    }

    if (!isValidTimeFormat(timeOfDetermination.trim())) {
      errorMessages.push('Invalid Time of Determination. Please ensure the time is entered in the correct format.');
    }

    if (errorMessages.length === 0) {
      const receivedDate = new Date(dateReceived.trim());
      const determinationDate = new Date(dateOfDetermination.trim());

      const receivedTime = timeReceived.trim().split(':').map(Number);
      const determinationTime = timeOfDetermination.trim().split(':').map(Number);

      let receivedDateTime = receivedDate;
      let determinationDateTime = determinationDate;

      receivedDateTime.setHours(receivedTime[0], receivedTime[1], receivedTime[2]);
      determinationDateTime.setHours(determinationTime[0], determinationTime[1], determinationTime[2]);

      const diffMs = determinationDateTime - receivedDateTime;
      const diffHours = diffMs / (1000 * 60 * 60);

      if (diffHours <= 72) {
        errorMessages.push('Request processed within the required timeframe.');
      } else {
        errorMessages.push(`Request processing exceeded the required timeframe by ${diffHours - 72} hours.`);
      }
    }
  }

  if (errorMessages.length > 0) {
    logAndAddErrorUnique(errors, 'Z', errorMessages.join(' | '), rowNumber, timelinessSheet, worksheet, color);
  }
}

// Forgot password form
app.get('/forgot-password', (req, res) => {
  res.render('forgot-password');
});

// Forgot Password Route
app.post('/forgot-password', async (req, res) => {
  const { email, securityQuestion, securityAnswer } = req.body;
  const user = await User.findOne({ where: { email, securityQuestion, securityAnswer } });
  if (user) {
    // Generate a reset token and expiration
    const resetToken = crypto.randomBytes(32).toString('hex');
    const resetTokenExpiration = Date.now() + 3600000; // 1 hour

    user.resetToken = resetToken;
    user.resetTokenExpiration = resetTokenExpiration;
    await user.save();

    // Send email with reset link
    const transporter = nodemailer.createTransport({
      service: 'Gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
      }
    });

    const mailOptions = {
      to: email,
      from: process.env.EMAIL_USER,
      subject: 'Password Reset',
      text: `You are receiving this because you (or someone else) have requested the reset of the password for your account.\n\n
      Please click on the following link, or paste this into your browser to complete the process:\n\n
      http://${req.headers.host}/reset/${resetToken}\n\n
      If you did not request this, please ignore this email and your password will remain unchanged.\n`
    };

    transporter.sendMail(mailOptions, (err) => {
      if (err) {
        console.error('Error sending email:', err);
        req.flash('error', 'Error sending password reset email. Please try again.');
        return res.redirect('/forgot-password');
      } else {
        console.log('Password reset email sent');
        req.flash('success', 'An email has been sent to your email address with further instructions.');
        res.redirect('/login');
      }
    });
  } else {
    req.flash('error', 'No account with that email address or security question answer exists.');
    res.redirect('/forgot-password');
  }
});

// Routes for user registration and authentication
app.get('/login', (req, res) => {
  res.render('login', { messages: req.flash('error') });
});

app.get('/register', (req, res) => {
  res.render('register', { messages: req.flash('error') });
});

app.post('/login', async (req, res) => {
  const { email, password } = req.body;
  const user = await User.findOne({ where: { email } });

  if (user && user.isApproved && await bcrypt.compare(password, user.password)) {
    req.session.userId = user.id;
    res.redirect('/dashboard'); // Redirect to the dashboard
  } else {
    req.flash('error', 'Invalid email, password, or account not approved yet');
    res.redirect('/login');
  }
});

app.get('/dashboard', (req, res) => {
  res.render('dashboard', { user: 'Guest' });
});

app.post('/register', async (req, res) => {
  const { firstName, lastName, email, password, confirmPassword, securityQuestion, securityAnswer } = req.body;
  if (password !== confirmPassword) {
    req.flash('error', 'Passwords do not match');
    return res.redirect('/register');
  }
  try {
    const hashedPassword = await bcrypt.hash(password, 10);
    const user = await User.create({
      firstName,
      lastName,
      email,
      password: hashedPassword,
      securityQuestion,
      securityAnswer
    });

    // Send approval request email to admin
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
      }
    });

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: process.env.ADMIN_EMAIL,
      subject: 'New User Registration Approval Needed',
      html: `
        <p>A new user has registered. Please review the details below:</p>
        <p>Name: ${firstName} ${lastName}</p>
        <p>Email: ${email}</p>
        <p><a href="${req.protocol}://${req.get('host')}/approve/${user.id}">Approve</a> | <a href="${req.protocol}://${req.get('host')}/deny/${user.id}">Deny</a></p>
      `
    };

    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        console.error('Error sending email:', error);
        req.flash('error', 'Error sending approval email. Please try again.');
        return res.redirect('/register');
      } else {
        console.log('Email sent:', info.response);
        req.flash('success', 'Registration successful. Your request is being reviewed for approval.');
        return res.redirect('/register');
      }
    });
  } catch (error) {
    console.error('Error registering user:', error);
    req.flash('error', 'Email already registered');
    res.redirect('/register');
  }
});

app.get('/approve/:userId', async (req, res) => {
  const { userId } = req.params;
  const user = await User.findByPk(userId);
  if (user) {
    user.isApproved = true;
    await user.save();

    // Send approval confirmation email to user
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
      }
    });

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: user.email,
      subject: 'Account Approved',
      text: 'Your account has been approved. You can now log in and access the scrubber tool and memo functionality.'
    };

    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        console.error('Error sending email:', error);
      } else {
        console.log('Email sent:', info.response);
      }
      res.send('User approved successfully');
    });
  } else {
    res.send('Invalid approval link');
  }
});

app.get('/deny/:userId', async (req, res) => {
  const { userId } = req.params;
  const user = await User.findByPk(userId);
  if (user) {
    await user.destroy();

    // Send denial notification email to user
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
      }
    });

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: user.email,
      subject: 'Registration Denied',
      text: 'Your registration has been denied. Please contact support@example.com if you have any questions.'
    };

    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        console.error('Error sending email:', error);
      } else {
        console.log('Email sent:', info.response);
      }
      res.send('User denied successfully');
    });
  } else {
    res.send('Invalid denial link');
  }
});

app.get('/logout', (req, res) => {
  res.redirect('/');
});

// Main route
app.get('/', async (req, res) => {
  try {
    res.render('index', { user: 'Guest' });
  } catch (error) {
    console.error('Error fetching user:', error);
    res.redirect('/');
  }
});

// Route for rendering the Regulatory Memo Tracking setup page
app.get('/dashboard/regulatory-memo-tracking', (req, res) => {
  console.log('GET /dashboard/regulatory-memo-tracking hit');
  res.render('regulatory-memo-tracking-setup', { title: 'Regulatory Memo Tracking Setup' });
});

// Route for handling form submission from the Regulatory Memo Tracking setup page
app.post('/dashboard/regulatory-memo-tracking', async (req, res) => {
  console.log('POST /dashboard/regulatory-memo-tracking hit');
  const { departments, leaderName, leaderEmail } = req.body;

  res.render('regulatory-memo-tracking-success');
});

// Route for rendering the ODAG upload page
app.get('/odag', (req, res) => {
  res.render('odag', { title: 'ODAG Scrubber Tool' });
});

app.post('/odag/upload', upload.single('file'), async (req, res) => {
  const { startDate, endDate } = req.body;
  console.log('Received file upload request');  // Debug log
  console.log('Request body:', req.body);

  try {
    const workbook = new exceljs.Workbook();
    console.log('Loading workbook from uploaded file buffer');
    await workbook.xlsx.load(req.file.buffer);

    const worksheet = workbook.getWorksheet('Sheet1');
    console.log('Worksheet loaded:', worksheet ? 'Sheet1 found' : 'Sheet1 not found');
    if (!worksheet) {
      throw new Error('Worksheet "Sheet1" not found in the uploaded file.');
    }

    const universeSheet = workbook.addWorksheet('Universe Accuracy');
    const timelinessSheet = workbook.addWorksheet('Timeliness Calculation');
    console.log('Added new worksheets for universe accuracy and timeliness calculation');

    universeSheet.addRow(['Column', 'Row', 'Error']);
    timelinessSheet.addRow(['Column', 'Row', 'Timeliness Error']);

    let errors = [];
    let timelinessErrors = [];
    console.log('Initialized error arrays');

    resetLoggedErrors();
    console.log('Reset logged errors');

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // Skip the header row
        console.log(`Processing row number ${rowNumber}`);
        const firstName = row.getCell('A').text.trim();
        const lastName = row.getCell('B').text.trim();
        const enrolleeID = row.getCell('C').text.trim();
        const contractID = row.getCell('D').text.trim();
        const pbp = row.getCell('E').text.trim();
        const firstTierEntity = row.getCell('F').text.trim();
        const authOrClaimNumber = row.getCell('G').text.trim();
        const dateReceived = row.getCell('H').text.trim();
        const timeReceived = row.getCell('I').text.trim();
        const partBDrugRequest = row.getCell('J').text.trim();
        const aorDate = row.getCell('K').text.trim();
        const aorReceiptTime = row.getCell('L').text.trim();
        const requestDetermination = row.getCell('M').text.trim();
        const requestProcessing = row.getCell('N').text.trim();
        const timeframeExtension = row.getCell('O').text.trim();
        const dateOfDetermination = row.getCell('P').text.trim();
        const timeOfDetermination = row.getCell('Q').text.trim();
        const dateOralNotification = row.getCell('R').text.trim();
        const timeOfOralNotification = row.getCell('S').text.trim();
        const notificationMethod = row.getCell('T').text.trim();
        const notificationDate = row.getCell('U').text.trim();
        const requestReceivedFrom = row.getCell('V').text.trim();
        const descriptionOfExtension = row.getCell('W').text.trim();
        const supportingDocumentation = row.getCell('X').text.trim();
        const processNote = row.getCell('Y').text.trim();

        validateFirstName(firstName, rowNumber, errors, universeSheet, worksheet);
        validateLastName(lastName, rowNumber, errors, universeSheet, worksheet);
        validateEnrolleeID(enrolleeID, rowNumber, errors, universeSheet, worksheet);
        validateContractID(contractID, rowNumber, errors, universeSheet, worksheet);
        validatePBP(pbp, rowNumber, errors, universeSheet, worksheet);
        validateFirstTierEntity(firstTierEntity, rowNumber, errors, universeSheet, worksheet);
        validateAuthOrClaimNumber(authOrClaimNumber, rowNumber, errors, universeSheet, worksheet);
        validateDateReceived(dateReceived, rowNumber, errors, universeSheet, worksheet);
        validateTimeReceived(timeReceived, requestProcessing, partBDrugRequest, rowNumber, errors, universeSheet, worksheet);
        validatePartBDrugRequest(partBDrugRequest, rowNumber, errors, universeSheet, worksheet);
        validateAORDate(aorDate, requestReceivedFrom, requestDetermination, rowNumber, errors, universeSheet, worksheet);
        validateAORReceiptTime(aorReceiptTime, requestReceivedFrom, requestProcessing, requestDetermination, partBDrugRequest, rowNumber, errors, universeSheet, worksheet);
        validateRequestDetermination(requestDetermination, rowNumber, errors, universeSheet, worksheet);
        validateRequestProcessing(requestProcessing, rowNumber, errors, universeSheet, worksheet);
        validateTimeframeExtension(timeframeExtension, rowNumber, errors, universeSheet, worksheet);
        validateDateOfDetermination(dateOfDetermination, rowNumber, errors, universeSheet, worksheet);
        validateTimeOfDetermination(timeOfDetermination, requestProcessing, partBDrugRequest, rowNumber, errors, universeSheet, worksheet);
        validateDateOralNotification(dateOralNotification, rowNumber, errors, universeSheet, worksheet);
        validateTimeOfOralNotification(row, rowNumber, errors, universeSheet, worksheet);
        validateNotificationMethod(notificationMethod, rowNumber, errors, universeSheet, worksheet);
        validateNotificationDate(notificationDate, rowNumber, errors, universeSheet, worksheet);
        validateRequestReceivedFrom(requestReceivedFrom, rowNumber, errors, universeSheet, worksheet);
        validateDescriptionOfExtension(descriptionOfExtension, timeframeExtension, rowNumber, errors, universeSheet, worksheet);
        validateSupportingDocumentation(supportingDocumentation, rowNumber, errors, universeSheet, worksheet);
        validateProcessNote(processNote, rowNumber, errors, universeSheet, worksheet);
        calculateTimeliness(rowNumber, requestProcessing, dateReceived, timeReceived, dateOfDetermination, timeOfDetermination, timelinessSheet, worksheet);
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = `validated_${Date.now()}.xlsx`;
    console.log('Writing validated workbook to buffer');

    fs.writeFileSync(path.join(uploadDir, fileName), buffer);
    console.log('File written to disk:', fileName);

    res.download(path.join(uploadDir, fileName), fileName, (err) => {
      if (err) {
        console.error('Error downloading the file:', err);
        res.status(500).send('Error downloading the file');
      } else {
        console.log('File download initiated:', fileName);
      }
    });
  } catch (error) {
    console.error('Error processing the file:', error.message);
    res.status(500).send(`Error processing the file: ${error.message}`);
  }
});

// Database synchronization and server start
sequelize.sync().then(() => {
  console.log('Database synced!');
  app.listen(port, () => {
    console.log(`Server running on port ${port}`);
  });
});
