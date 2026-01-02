function buildPrintout() {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ss.getSheetByName('Donations_Log'); // change if needed
  if (!dataSheet) throw new Error("Sheet 'Donations_Log' not found.");

  // Create/clear Printout sheet
  let rpt = ss.getSheetByName('Printout');
  if (!rpt) rpt = ss.insertSheet('Printout');
  rpt.clear();
  rpt.setHiddenGridlines(true);

  // Basic layout
  rpt.setColumnWidths(1, 1, 28);   // A
  rpt.setColumnWidths(2, 1, 36);   // B
  rpt.setColumnWidths(3, 1, 65);   // C (descriptions)
  rpt.setColumnWidths(4, 1, 16);   // D (amount)

  const year = new Date().getFullYear(); // adjust if you want manual year
  const name = ss.getOwner() ? ss.getOwner().getEmail() : 'YOUR NAME';

  // Title block (like ItsDeductible)
  rpt.getRange('A1:D1').merge().setValue(`${year} Tax Year`).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  rpt.getRange('A2:D2').merge().setValue('YTD Charitable Deductions').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  rpt.getRange('A3:D3').merge().setValue('YOUR NAME').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');

  let row = 5;

  // Read data
  const values = dataSheet.getDataRange().getValues();
  const headers = values.shift();
  const idx = Object.fromEntries(headers.map((h,i)=>[h,i]));

  function isCash(r){ return String(r[idx['IRS Donation Type Classification']] || '').toLowerCase().includes('cash'); }
  function isNonCash(r){ return String(r[idx['IRS Donation Type Classification']] || '').toLowerCase().includes('non'); }

  function fmtDate(d){
    if (!(d instanceof Date)) return d;
    return Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }

  function writeSection(title, filterFn) {
    rpt.getRange(row,1,1,4).merge().setValue(title).setFontWeight('bold');
    row++;

    // Header row
    rpt.getRange(row,1).setValue('Charity Name / Donated Date').setFontWeight('bold');
    rpt.getRange(row,2).setValue('Charity Address').setFontWeight('bold');
    rpt.getRange(row,3).setValue('Donation Description').setFontWeight('bold');
    rpt.getRange(row,4).setValue('Donation Amount').setFontWeight('bold').setHorizontalAlignment('right');
    row++;

    const rows = values.filter(filterFn).sort((a,b)=>{
      const ca = String(a[idx['Charity']]||'');
      const cb = String(b[idx['Charity']]||'');
      if (ca !== cb) return ca.localeCompare(cb);
      return new Date(a[idx['Date']]) - new Date(b[idx['Date']]);
    });

    let total = 0;

    rows.forEach(r=>{
      const charity = r[idx['Charity']] || '';
      const addr = r[idx['Charity Address']] || '';
      const date = fmtDate(r[idx['Date']]);
      const desc = r[idx['Description']] || '';
      const amt = Number(r[idx['Donation Value in $']] || 0);

      rpt.getRange(row,1).setValue(`${charity}\n${date}`).setWrap(true);
      rpt.getRange(row,2).setValue(addr).setWrap(true);
      rpt.getRange(row,3).setValue(desc).setWrap(true);
      rpt.getRange(row,4).setValue(amt).setNumberFormat('$#,##0.00').setHorizontalAlignment('right');

      total += amt;
      row++;
    });

    rpt.getRange(row,3).setValue('Subtotal :').setFontWeight('bold').setHorizontalAlignment('right');
    rpt.getRange(row,4).setValue(total).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');
    row++;

    return total;
  }

  const nonCashTotal = writeSection('Non-Cash Donations', isNonCash);
  rpt.getRange(row,3).setValue('Total Non-Cash Donations:').setFontWeight('bold').setHorizontalAlignment('right');
  rpt.getRange(row,4).setValue(nonCashTotal).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');
  row += 2;

  const cashTotal = writeSection('Cash Donations', isCash);
  rpt.getRange(row,3).setValue('Total Cash Donations:').setFontWeight('bold').setHorizontalAlignment('right');
  rpt.getRange(row,4).setValue(cashTotal).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');
  row += 2;

  rpt.getRange(row,3).setValue('Grand Total:').setFontWeight('bold').setHorizontalAlignment('right');
  rpt.getRange(row,4).setValue(nonCashTotal + cashTotal).setNumberFormat('$#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');

  // Cosmetic borders
  rpt.getRange(6,1,row,4).setVerticalAlignment('top');
  rpt.getDataRange().setFontFamily('Arial').setFontSize(10);
}
