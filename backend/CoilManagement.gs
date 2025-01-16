const SPREADSHEET_ID = "1ErvO0zu9RIi22w-HcffVf6GAzdM1HVhbGQDftiO7Rlk";
const COL_TIMESTAMP = 1;
const COL_LOT_NUMBER = 2;
const COL_GAUGE = 3;
const COL_COLOR = 4;
const COL_WIDTH = 5;
const COL_WEIGHT = 6;
const COL_LENGTH = 7;
const COL_LOCATION = 8;
const COL_EMAIL = 9;
const COL_SQUAREFEET = 10;
const COL_BALANCE_WEIGHT = 11;
const COL_BALANCE_LENGTH = 12;
const COL_BALANCE_SQUAREFEET = 13;
const COL_TIMESTAMP_SECONDARY = 14;
const COL_EMAIL_SECONDARY = 15;


function doGet() {
  return HtmlService.createHtmlOutputFromFile('ReceiptFrm').setTitle('Coil Management');
}


// Function to get user email
function getEmail() {
  return Session.getActiveUser().getEmail();
}


// Retrieve unique color options from the "Data" sheet, column B
function getColorOptions() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Data');
  const colorColumn = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues(); // Column B
  const uniqueColors = [...new Set(colorColumn.flat())].filter(color => color);
  return uniqueColors;
}


// Check if lot number exists in the Inventory sheet
function checkLotNumberExists(lotNumber) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // No data below headers
  const data = sheet.getRange(2, COL_LOT_NUMBER, lastRow - 1, 1).getValues();
  return data.some(row => row[0] === lotNumber);
}


// Retrieve data for a specific lot number
function getCoilData(lotNumber) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_LOT_NUMBER - 1] === lotNumber) {
      return {
        gauge: data[i][COL_GAUGE - 1],
        color: data[i][COL_COLOR - 1],
        width: data[i][COL_WIDTH - 1],
        length: data[i][COL_LENGTH - 1],
        balanceLength: data[i][COL_BALANCE_LENGTH - 1],
        balanceWeight: data[i][COL_BALANCE_WEIGHT - 1] // Correct column for balance weight
      };
    }
  }
  return null; // Return null if the lot number is not found
}


function createCoil(coilData) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();

  // Find the last generated lot number
  const prefix = "A";
  let maxNumber = 110; // Start from 'A111'
  data.forEach(row => {
    if (row[COL_LOT_NUMBER - 1] && row[COL_LOT_NUMBER - 1].startsWith(prefix)) {
      const number = parseInt(row[COL_LOT_NUMBER - 1].substring(1));
      if (!isNaN(number)) {
        maxNumber = Math.max(maxNumber, number);
      }
    }
  });

  // Generate the next lot number
  const newLotNumber = `${prefix}${maxNumber + 1}`;

  // Append the new coil data
  sheet.appendRow([
    new Date(),                // COL_TIMESTAMP
    newLotNumber,              // COL_LOT_NUMBER
    coilData.gauge,            // COL_GAUGE
    coilData.color,            // COL_COLOR
    coilData.width,            // COL_WIDTH
    coilData.weight || '',     // COL_WEIGHT
    coilData.length,           // COL_LENGTH
    coilData.location,         // COL_LOCATION
    coilData.email,            // COL_EMAIL
    '', '',                    // Empty: COL_J, COL_K,
    coilData.length,           // COL_BALANCE_LENGTH (set to Length)
    '',                        // Empty: COL_
    new Date(),                // COL_TIMESTAMP_SECONDARY
    coilData.email             // COL_EMAIL_SECONDARY
  ]);

  // Return data for the created coil
  return {
    lotNumber: newLotNumber,
    gauge: coilData.gauge,
    color: coilData.color,
    width: coilData.width,
    length: coilData.length,
    balanceLength: coilData.length // Include balance length for printing or UI feedback
  };
}




function submitReceive(coilData) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  
  Logger.log("Submitting new coil data: " + JSON.stringify(coilData));
  
  sheet.appendRow([
    new Date(),                 // COL_TIMESTAMP
    coilData.lotNumber,         // COL_LOT_NUMBER
    coilData.gauge,             // COL_GAUGE
    coilData.color,             // COL_COLOR
    coilData.width,             // COL_WIDTH
    coilData.weight,            // COL_WEIGHT
    coilData.length,            // COL_LENGTH
    coilData.location,          // COL_LOCATION
    coilData.email,             // COL_EMAIL
    '',                         // COL_J: Formula will be added
    '',                         // COL_K: Formula will be added
    coilData.length,            // COL_BALANCE_LENGTH, same as COL_LENGTH
    '',                         // COL_M: Formula will be added
    new Date(),                 // COL_TIMESTAMP_SECONDARY
    coilData.email              // COL_EMAIL_SECONDARY
  ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL_BALANCE_WEIGHT).setFormula(`=F${lastRow}/G${lastRow}*L${lastRow}`); // Formula for Balance Weight
  sheet.getRange(lastRow, COL_SQUAREFEET ).setFormula(`=E${lastRow}/12*G${lastRow}`); // Formula for column J (e.g., Volume)
  sheet.getRange(lastRow, COL_BALANCE_SQUAREFEET).setFormula(`=E${lastRow}/12*L${lastRow}`); // Formula for column M (e.g., Adjusted Volume)

  return true; // Indicate success
}


// Submit Reduce function
function submitReduce(formData) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_LOT_NUMBER - 1] === formData.lotNumber) {
      const currentBalance = parseFloat(data[i][COL_BALANCE_LENGTH - 1]);
      const reduction = parseFloat(formData.linearFeetReduced);

      if (isNaN(currentBalance) || isNaN(reduction)) {
        throw new Error("Invalid balance length or reduction amount.");
      }

      const newBalanceL = currentBalance - reduction;
      if (newBalanceL < 0) {
        throw new Error("Insufficient balance length to reduce.");
      }

      // Update the new balance length
      sheet.getRange(i + 1, COL_BALANCE_LENGTH).setValue(newBalanceL);

      // Correctly calculate balance weight
      const originalWeight = parseFloat(data[i][COL_WEIGHT - 1]); // Original weight of the coil
      const originalLength = parseFloat(data[i][COL_LENGTH - 1]); // Original length of the coil

      if (isNaN(originalWeight) || isNaN(originalLength) || originalLength === 0) {
        throw new Error("Invalid original weight or length.");
      }

      const balanceWeight = (originalWeight / originalLength) * newBalanceL;

      // Update balance weight
      sheet.getRange(i + 1, COL_BALANCE_WEIGHT).setValue(balanceWeight);

      // Update secondary fields
      sheet.getRange(i + 1, COL_TIMESTAMP_SECONDARY).setValue(new Date());
      sheet.getRange(i + 1, COL_EMAIL_SECONDARY).setValue(formData.email);

      return { balanceLength: newBalanceL, balanceWeight }; // Return both balance length and weight
    }
  }
  throw new Error(`Lot number ${formData.lotNumber} not found.`);
}




function submitSlit(coilData) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();
  let masterRow;
  const createdCoils = [];
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  // Find the current batch suffix for this lot number
  const existingBatchSuffixes = data
    .filter(row => row[COL_LOT_NUMBER - 1].startsWith(`${coilData.lotNumber}.`))
    .map(row => row[COL_LOT_NUMBER - 1].split('.')[1]);

  let nextBatchSuffix = "A";
  if (existingBatchSuffixes.length > 0) {
    const lastBatch = existingBatchSuffixes.sort().pop(); // Get the last batch suffix
    const lastBatchIndex = alphabet.indexOf(lastBatch);
    if (lastBatchIndex >= 0 && lastBatchIndex < alphabet.length - 1) {
      nextBatchSuffix = alphabet[lastBatchIndex + 1]; // Increment to the next letter
    } else if (lastBatchIndex === alphabet.length - 1) {
      throw new Error("Exceeded the number of available batch suffixes (A-Z).");
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_LOT_NUMBER - 1] === coilData.lotNumber) {
      masterRow = data[i];
      const currentBalance = masterRow[COL_BALANCE_LENGTH - 1];
      const totalSlitLength = coilData.slitLength;
      const newBalanceL = currentBalance - totalSlitLength;

      if (newBalanceL < 0) {
        throw new Error("Insufficient balance length to slit.");
      }

      // Original coil dimensions and weight
      const originalWeight = parseFloat(masterRow[COL_WEIGHT - 1]);
      const originalLength = parseFloat(masterRow[COL_LENGTH - 1]);
      const originalWidth = parseFloat(masterRow[COL_WIDTH - 1]);

      if (isNaN(originalWeight) || isNaN(originalLength) || originalLength === 0) {
        throw new Error("Invalid original weight or length.");
      }

      const weightPerSqFt = originalWeight / (originalLength * (originalWidth / 12)); // Weight per square foot

      // Update balance weight for the master coil
      const balanceWeight = weightPerSqFt * newBalanceL * (originalWidth / 12);
      sheet.getRange(i + 1, COL_BALANCE_LENGTH).setValue(newBalanceL);
      sheet.getRange(i + 1, COL_BALANCE_WEIGHT).setValue(balanceWeight);
      sheet.getRange(i + 1, COL_TIMESTAMP_SECONDARY).setValue(new Date());
      sheet.getRange(i + 1, COL_EMAIL_SECONDARY).setValue(coilData.email);

      // Add master coil data to the array for label printing
      createdCoils.push({
        lotNumber: coilData.lotNumber,
        gauge: masterRow[COL_GAUGE - 1],
        color: masterRow[COL_COLOR - 1],
        width: masterRow[COL_WIDTH - 1],
        balanceLength: newBalanceL,
        balanceWeight: balanceWeight
      });

      // Generate data for each slit coil
      for (let j = 0; j < coilData.qty; j++) {
        const slitWeight = weightPerSqFt * coilData.slitLength * (coilData.slitWidth / 12); // Slit coil weight
        const newLotNumber = `${coilData.lotNumber}.${nextBatchSuffix}.${j + 1}`;

        // Append slit coil to the sheet
        sheet.appendRow([
          new Date(),
          newLotNumber,
          coilData.gauge,
          coilData.color,
          coilData.slitWidth,
          slitWeight,
          coilData.slitLength,
          coilData.location,
          coilData.email,
          '', 
          '', 
          coilData.slitLength,
          '', 
          new Date(),
          coilData.email
          ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL_BALANCE_WEIGHT).setFormula(`=F${lastRow}/G${lastRow}*L${lastRow}`); // Formula for Balance Weight
  sheet.getRange(lastRow, COL_SQUAREFEET ).setFormula(`=E${lastRow}/12*G${lastRow}`); // Formula for column J (e.g., Volume)
  sheet.getRange(lastRow, COL_BALANCE_SQUAREFEET).setFormula(`=E${lastRow}/12*L${lastRow}`); // Formula for column M (e.g., Adjusted Volume)


        // Add slit coil data to the array for label printing
        createdCoils.push({
          lotNumber: newLotNumber,
          gauge: coilData.gauge,
          color: coilData.color,
          width: coilData.slitWidth,
          balanceLength: coilData.slitLength,
          balanceWeight: slitWeight
        });
      }
      return createdCoils;
    }
  }
  throw new Error(`Lot number ${coilData.lotNumber} not found.`);
}




// Submit Move function
function submitMove(coilData) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_LOT_NUMBER - 1] === coilData.lotNumber) {
      sheet.getRange(i + 1, COL_LOCATION).setValue(coilData.newLocation);
      sheet.getRange(i + 1, COL_TIMESTAMP_SECONDARY).setValue(new Date());
      sheet.getRange(i + 1, COL_EMAIL_SECONDARY).setValue(coilData.email);
      return true;
    }
  }
  throw new Error(`Lot number ${coilData.lotNumber} not found.`);
}

function getCoilData(lotNumber) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_LOT_NUMBER - 1] === lotNumber) {
      return {
        gauge: data[i][COL_GAUGE - 1],
        color: data[i][COL_COLOR - 1],
        width: data[i][COL_WIDTH - 1],
        balanceLength: data[i][COL_BALANCE_LENGTH - 1],
        balanceWeight: data[i][COL_BALANCE_WEIGHT - 1]
      };
    }
  }
  return null; // Return null if the lot number is not found
}









