/** @OnlyCurrentDoc */


function Transformconsoledata() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const rowBegin = 2;
  const lastRow = sheet.getLastRow();

  // Keep only time for L, 12 column
  const orderExecutionTimeRange = sheet.getRange(rowBegin, 12, lastRow - rowBegin + 1);
  const orderExecutionTimes = orderExecutionTimeRange.getValues();
  const timestamps = [];
  for (let row in orderExecutionTimes) {
    for (let col in orderExecutionTimes[row]) {
      timestamps.push([orderExecutionTimes[row][col].toString().split('T')[1]]);
    }
  }
  orderExecutionTimeRange.setValues(timestamps);

  // Copy oet from L(M) after adjustment to C
  const orderExecutionTime = sheet.getRange('M:M');
  sheet.insertColumnAfter(3);
  orderExecutionTime.copyTo(sheet.getRange(1, 4));

  // // Delete unwanted cols.
  sheet.deleteColumn(13);
  sheet.deleteColumn(12);
  sheet.deleteColumn(11);
  sheet.deleteColumn(7);
  sheet.deleteColumn(6);
  sheet.deleteColumn(5);
  sheet.deleteColumn(2);

  const lastCol = sheet.getLastColumn();

  // Group same orders within an interval of 3 min,
  
  for(let i = rowBegin; i <= lastRow; i++) {
    const currRowRange = sheet.getRange(i, 1, 1, lastCol);
    const currRow = currRowRange.getValues()[0];

    // get all same rows within 3 min. Store in list.
    const sameOrderIndicies = []; // index is excel row number: 1, 2, 3 ...
    const currSymbol =  currRow[0];
    const currDate = currRow[1];
    const currTime = currRow[2];
    const currType = currRow[3];
    for(let j = i + 1; j <= lastRow; j++) {
      const row = sheet.getRange(j, 1, 1, lastCol).getValues()[0];
      Logger.log(row);
      const symbol = row[0];
      const date = row[1];
      const time = row[2];
      const type = row[3];
      if(currDate === date && isWithinInterval(time, currTime, 3)) {
        if(currSymbol === symbol && currType === type) {
          sameOrderIndicies[j];
        }
      }
      else {
        break;
      }
    }

    // combine the list with current entry. weighted average price.
    const numer = 0, denom = 0;
    for (let j = 0 ; j < sameOrderIndicies.length; j++) {
      const row = sheet.getRange(j, 1, 1, lastCol).getValues()[0];
      const currQty = row[4];
      const currPrice = row[5];
      numer += currQty * currPrice;
      denom += currQty;
    }

    // Set new qty
    sheet.getRange(i, 5).setValue(denom);

    // Set new price
    sheet.getRange(i, 6).setValue(numer);

    // delete those rows.
    sameOrderIndicies.reverse().forEach(index => sheet.deleteRow(index));
  }

  // Add col for exit date and time
  sheet.insertColumnAfter(3);
  sheet.insertColumnAfter(3);
  sheet.getRange('D1').setValue('Exit date');
  sheet.getRange('E1').setValue('Exit time');

  // Populate exit date and time

};

function isWithinInterval(t1, t2, interval) {
  const [hhT1, mmT1, ssT1] = t1.toString().split(':');
  const [hhT2, mmT2, ssT2] = t2.toString().split(':');

  // TODO: add hour shift support
  if (hhT1 !== hhT2) {
    return false;
  }

  if (parseInt(mmT1) - parseInt(mmT2) <= interval) {
    return true;
  }

  return false;
}
