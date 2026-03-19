const ExcelJS = require("exceljs");

async function createUmraTracker() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Umra Tracker";
  workbook.created = new Date();

  const people = [
    "Osoba 1", "Osoba 2", "Osoba 3", "Osoba 4", "Osoba 5",
    "Osoba 6", "Osoba 7", "Osoba 8", "Osoba 9", "Osoba 10",
    "Osoba 11", "Osoba 12", "Osoba 13", "Osoba 14", "Osoba 15",
  ];

  const months = [
    "Januar", "Februar", "Mart", "April", "Maj", "Juni",
    "Juli", "August", "Septembar", "Oktobar", "Novembar", "Decembar",
  ];

  const years = [2026, 2027, 2028, 2029, 2030];

  // ── Color palette ──
  const darkGreen = "1B5E20";
  const medGreen = "2E7D32";
  const lightGreen = "E8F5E9";
  const gold = "F9A825";
  const darkGold = "F57F17";
  const white = "FFFFFF";
  const lightGray = "F5F5F5";
  const medGray = "E0E0E0";
  const darkText = "212121";
  const paidGreen = "C8E6C9";
  const unpaidRed = "FFCDD2";

  const thinBorder = {
    top: { style: "thin", color: { argb: "BDBDBD" } },
    left: { style: "thin", color: { argb: "BDBDBD" } },
    bottom: { style: "thin", color: { argb: "BDBDBD" } },
    right: { style: "thin", color: { argb: "BDBDBD" } },
  };

  const medBorder = {
    top: { style: "medium", color: { argb: "9E9E9E" } },
    left: { style: "medium", color: { argb: "9E9E9E" } },
    bottom: { style: "medium", color: { argb: "9E9E9E" } },
    right: { style: "medium", color: { argb: "9E9E9E" } },
  };

  // ══════════════════════════════════════════════════════════════
  //  YEARLY SHEETS (one per year)
  // ══════════════════════════════════════════════════════════════
  for (const year of years) {
    const ws = workbook.addWorksheet(`${year}`, {
      properties: { defaultColWidth: 14 },
      views: [{ state: "frozen", xSplit: 2, ySplit: 4 }],
    });

    // Row 1: Title
    ws.mergeCells("A1:O1");
    const titleCell = ws.getCell("A1");
    titleCell.value = `UMRA TRACKER - ${year}`;
    titleCell.font = { name: "Calibri", size: 22, bold: true, color: { argb: white } };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: darkGreen } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(1).height = 50;

    // Row 2: Subtitle
    ws.mergeCells("A2:O2");
    const subCell = ws.getCell("A2");
    subCell.value = `Mjesečna uplata: 10 KM po osobi  |  Cilj: Zajednička Umra`;
    subCell.font = { name: "Calibri", size: 11, italic: true, color: { argb: white } };
    subCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGreen } };
    subCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(2).height = 28;

    // Row 3: spacer
    ws.getRow(3).height = 8;
    for (let c = 1; c <= 15; c++) {
      ws.getRow(3).getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: white } };
    }

    // Row 4: Header row — "#", "Ime i prezime", then 12 months, then "UKUPNO (year)"
    const headerRow = ws.getRow(4);
    headerRow.height = 32;

    ws.getColumn(1).width = 5;
    ws.getColumn(2).width = 22;
    for (let c = 3; c <= 14; c++) ws.getColumn(c).width = 14;
    ws.getColumn(15).width = 18;

    const headers = ["#", "Ime i prezime", ...months, `UKUPNO ${year}`];
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: white } };
      cell.fill = {
        type: "pattern", pattern: "solid",
        fgColor: { argb: i === headers.length - 1 ? darkGold : darkGreen },
      };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = thinBorder;
    });

    // Data rows (rows 5–19 for 15 people)
    people.forEach((name, idx) => {
      const rowNum = 5 + idx;
      const row = ws.getRow(rowNum);
      row.height = 26;

      // # column
      const numCell = row.getCell(1);
      numCell.value = idx + 1;
      numCell.font = { name: "Calibri", size: 10, bold: true, color: { argb: darkText } };
      numCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? lightGreen : white } };
      numCell.alignment = { horizontal: "center", vertical: "middle" };
      numCell.border = thinBorder;

      // Name column
      const nameCell = row.getCell(2);
      nameCell.value = name;
      nameCell.font = { name: "Calibri", size: 11, bold: true, color: { argb: darkText } };
      nameCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? lightGreen : white } };
      nameCell.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      nameCell.border = thinBorder;

      // Month cells (columns 3–14): user enters amount paid
      for (let m = 3; m <= 14; m++) {
        const mCell = row.getCell(m);
        mCell.value = null;
        mCell.numFmt = '#,##0.00 "KM"';
        mCell.font = { name: "Calibri", size: 11, color: { argb: darkText } };
        mCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? lightGreen : white } };
        mCell.alignment = { horizontal: "center", vertical: "middle" };
        mCell.border = thinBorder;

        // Data validation: only 0 or 10 KM (or blank)
        mCell.dataValidation = {
          type: "list",
          allowBlank: true,
          formulae: ['"0,10"'],
          showErrorMessage: true,
          errorTitle: "Pogrešan unos",
          error: "Unesite 0 ili 10 (KM)",
        };
      }

      // TOTAL per person for this year (column 15)
      const colLetter = (c) => String.fromCharCode(64 + c);
      const totalCell = row.getCell(15);
      totalCell.value = { formula: `SUM(C${rowNum}:N${rowNum})` };
      totalCell.numFmt = '#,##0.00 "KM"';
      totalCell.font = { name: "Calibri", size: 12, bold: true, color: { argb: darkText } };
      totalCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9C4" } };
      totalCell.alignment = { horizontal: "center", vertical: "middle" };
      totalCell.border = medBorder;
    });

    // Row 20: spacer
    const spacerRow = 5 + people.length;

    // Row 21: Monthly totals
    const totalRowNum = spacerRow + 1;
    const totalRow = ws.getRow(totalRowNum);
    totalRow.height = 30;

    const tLabelCell = totalRow.getCell(1);
    ws.mergeCells(`A${totalRowNum}:B${totalRowNum}`);
    tLabelCell.value = "UKUPNO PO MJESECU";
    tLabelCell.font = { name: "Calibri", size: 11, bold: true, color: { argb: white } };
    tLabelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: darkGreen } };
    tLabelCell.alignment = { horizontal: "center", vertical: "middle" };
    tLabelCell.border = medBorder;

    for (let m = 3; m <= 14; m++) {
      const cl = String.fromCharCode(64 + m);
      const cell = totalRow.getCell(m);
      cell.value = { formula: `SUM(${cl}5:${cl}${5 + people.length - 1})` };
      cell.numFmt = '#,##0.00 "KM"';
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: white } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGreen } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = medBorder;
    }

    // Grand total for the year
    const grandCell = totalRow.getCell(15);
    grandCell.value = { formula: `SUM(O5:O${5 + people.length - 1})` };
    grandCell.numFmt = '#,##0.00 "KM"';
    grandCell.font = { name: "Calibri", size: 14, bold: true, color: { argb: white } };
    grandCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: darkGold } };
    grandCell.alignment = { horizontal: "center", vertical: "middle" };
    grandCell.border = medBorder;

    // Row 22: Expected vs Actual
    const expRowNum = totalRowNum + 1;
    const expRow = ws.getRow(expRowNum);
    expRow.height = 26;
    ws.mergeCells(`A${expRowNum}:B${expRowNum}`);
    const expLabel = expRow.getCell(1);
    expLabel.value = "OČEKIVANO PO MJESECU";
    expLabel.font = { name: "Calibri", size: 10, bold: true, color: { argb: darkText } };
    expLabel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGray } };
    expLabel.alignment = { horizontal: "center", vertical: "middle" };
    expLabel.border = thinBorder;

    for (let m = 3; m <= 14; m++) {
      const cell = expRow.getCell(m);
      cell.value = people.length * 10;
      cell.numFmt = '#,##0.00 "KM"';
      cell.font = { name: "Calibri", size: 10, color: { argb: darkText } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGray } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = thinBorder;
    }

    const expTotal = expRow.getCell(15);
    expTotal.value = people.length * 10 * 12;
    expTotal.numFmt = '#,##0.00 "KM"';
    expTotal.font = { name: "Calibri", size: 11, bold: true, color: { argb: darkText } };
    expTotal.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGray } };
    expTotal.alignment = { horizontal: "center", vertical: "middle" };
    expTotal.border = thinBorder;

    // Conditional formatting for payment cells
    ws.addConditionalFormatting({
      ref: `C5:N${5 + people.length - 1}`,
      rules: [
        {
          type: "cellIs",
          operator: "equal",
          formulae: [10],
          style: {
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: paidGreen } },
            font: { bold: true, color: { argb: "1B5E20" } },
          },
          priority: 1,
        },
        {
          type: "cellIs",
          operator: "equal",
          formulae: [0],
          style: {
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: unpaidRed } },
            font: { bold: true, color: { argb: "B71C1C" } },
          },
          priority: 2,
        },
      ],
    });

    // Print setup
    ws.pageSetup = {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
    };
  }

  // ══════════════════════════════════════════════════════════════
  //  SUMMARY SHEET ("Pregled")
  // ══════════════════════════════════════════════════════════════
  const summary = workbook.addWorksheet("PREGLED", {
    properties: { defaultColWidth: 16 },
    views: [{ state: "frozen", xSplit: 2, ySplit: 4 }],
  });

  // Move summary to front
  summary.orderNo = 0;

  // Title
  const sumCols = 2 + years.length + 1; // #, Name, year cols, Grand Total
  summary.mergeCells(1, 1, 1, sumCols);
  const sTitle = summary.getCell("A1");
  sTitle.value = "UMRA TRACKER - UKUPNI PREGLED";
  sTitle.font = { name: "Calibri", size: 22, bold: true, color: { argb: white } };
  sTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: darkGreen } };
  sTitle.alignment = { horizontal: "center", vertical: "middle" };
  summary.getRow(1).height = 50;

  // Subtitle
  summary.mergeCells(2, 1, 2, sumCols);
  const sSub = summary.getCell("A2");
  sSub.value = `Pregled uplata svih ${people.length} osoba kroz sve godine  |  Cilj: Zajednička Umra`;
  sSub.font = { name: "Calibri", size: 11, italic: true, color: { argb: white } };
  sSub.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGreen } };
  sSub.alignment = { horizontal: "center", vertical: "middle" };
  summary.getRow(2).height = 28;

  // Spacer
  summary.getRow(3).height = 8;

  // Headers
  summary.getColumn(1).width = 5;
  summary.getColumn(2).width = 22;
  for (let i = 3; i <= 2 + years.length; i++) summary.getColumn(i).width = 18;
  summary.getColumn(sumCols).width = 22;

  const sHeaderRow = summary.getRow(4);
  sHeaderRow.height = 32;
  const sHeaders = ["#", "Ime i prezime", ...years.map(y => `${y}`), "SVEUKUPNO"];
  sHeaders.forEach((h, i) => {
    const cell = sHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: white } };
    cell.fill = {
      type: "pattern", pattern: "solid",
      fgColor: { argb: i === sHeaders.length - 1 ? darkGold : darkGreen },
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = thinBorder;
  });

  // Data rows
  people.forEach((name, idx) => {
    const rowNum = 5 + idx;
    const row = summary.getRow(rowNum);
    row.height = 28;

    // #
    const numCell = row.getCell(1);
    numCell.value = idx + 1;
    numCell.font = { name: "Calibri", size: 10, bold: true };
    numCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? lightGreen : white } };
    numCell.alignment = { horizontal: "center", vertical: "middle" };
    numCell.border = thinBorder;

    // Name
    const nameCell = row.getCell(2);
    nameCell.value = name;
    nameCell.font = { name: "Calibri", size: 11, bold: true };
    nameCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? lightGreen : white } };
    nameCell.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
    nameCell.border = thinBorder;

    // Per-year totals (reference each year sheet's O column)
    years.forEach((year, yi) => {
      const cell = row.getCell(3 + yi);
      cell.value = { formula: `'${year}'!O${rowNum}` };
      cell.numFmt = '#,##0.00 "KM"';
      cell.font = { name: "Calibri", size: 11 };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: idx % 2 === 0 ? lightGreen : white } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = thinBorder;
    });

    // Grand total across all years
    const colStart = String.fromCharCode(67); // C
    const colEnd = String.fromCharCode(67 + years.length - 1);
    const grandCell = row.getCell(sumCols);
    grandCell.value = { formula: `SUM(${colStart}${rowNum}:${colEnd}${rowNum})` };
    grandCell.numFmt = '#,##0.00 "KM"';
    grandCell.font = { name: "Calibri", size: 12, bold: true, color: { argb: darkText } };
    grandCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9C4" } };
    grandCell.alignment = { horizontal: "center", vertical: "middle" };
    grandCell.border = medBorder;
  });

  // Totals row on summary
  const sTotalRowNum = 5 + people.length + 1;
  const sTotalRow = summary.getRow(sTotalRowNum);
  sTotalRow.height = 32;

  summary.mergeCells(`A${sTotalRowNum}:B${sTotalRowNum}`);
  const stLabel = sTotalRow.getCell(1);
  stLabel.value = "UKUPNO PO GODINI";
  stLabel.font = { name: "Calibri", size: 11, bold: true, color: { argb: white } };
  stLabel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: darkGreen } };
  stLabel.alignment = { horizontal: "center", vertical: "middle" };
  stLabel.border = medBorder;

  years.forEach((year, yi) => {
    const cell = sTotalRow.getCell(3 + yi);
    const cl = String.fromCharCode(67 + yi);
    cell.value = { formula: `SUM(${cl}5:${cl}${5 + people.length - 1})` };
    cell.numFmt = '#,##0.00 "KM"';
    cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: white } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGreen } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = medBorder;
  });

  // Grand total of everything
  const colStart = String.fromCharCode(67);
  const colEnd = String.fromCharCode(67 + years.length - 1);
  const gGrand = sTotalRow.getCell(sumCols);
  gGrand.value = { formula: `SUM(${colStart}${sTotalRowNum}:${colEnd}${sTotalRowNum})` };
  gGrand.numFmt = '#,##0.00 "KM"';
  gGrand.font = { name: "Calibri", size: 16, bold: true, color: { argb: white } };
  gGrand.fill = { type: "pattern", pattern: "solid", fgColor: { argb: darkGold } };
  gGrand.alignment = { horizontal: "center", vertical: "middle" };
  gGrand.border = medBorder;

  // Expected row
  const sExpRowNum = sTotalRowNum + 1;
  const sExpRow = summary.getRow(sExpRowNum);
  sExpRow.height = 26;
  summary.mergeCells(`A${sExpRowNum}:B${sExpRowNum}`);
  const seLabel = sExpRow.getCell(1);
  seLabel.value = "OČEKIVANO PO GODINI";
  seLabel.font = { name: "Calibri", size: 10, bold: true };
  seLabel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGray } };
  seLabel.alignment = { horizontal: "center", vertical: "middle" };
  seLabel.border = thinBorder;

  years.forEach((year, yi) => {
    const cell = sExpRow.getCell(3 + yi);
    cell.value = people.length * 10 * 12;
    cell.numFmt = '#,##0.00 "KM"';
    cell.font = { name: "Calibri", size: 10 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGray } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = thinBorder;
  });

  const seGrand = sExpRow.getCell(sumCols);
  seGrand.value = people.length * 10 * 12 * years.length;
  seGrand.numFmt = '#,##0.00 "KM"';
  seGrand.font = { name: "Calibri", size: 11, bold: true };
  seGrand.fill = { type: "pattern", pattern: "solid", fgColor: { argb: medGray } };
  seGrand.alignment = { horizontal: "center", vertical: "middle" };
  seGrand.border = thinBorder;

  // Conditional formatting on summary year columns
  summary.addConditionalFormatting({
    ref: `C5:${String.fromCharCode(67 + years.length - 1)}${5 + people.length - 1}`,
    rules: [
      {
        type: "cellIs",
        operator: "equal",
        formulae: [120],
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: paidGreen } },
          font: { bold: true, color: { argb: "1B5E20" } },
        },
        priority: 1,
      },
      {
        type: "cellIs",
        operator: "lessThan",
        formulae: [120],
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF9C4" } },
        },
        priority: 2,
      },
    ],
  });

  summary.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    paperSize: 9,
  };

  // ── Save ──
  const filePath = "Umra.xlsx";
  await workbook.xlsx.writeFile(filePath);
  console.log(`Umra Tracker created: ${filePath}`);
  console.log(`  - PREGLED sheet (summary across all years)`);
  console.log(`  - ${years.length} yearly sheets: ${years.join(", ")}`);
  console.log(`  - ${people.length} people tracked`);
  console.log(`  - Monthly payment: 10 KM`);
}

createUmraTracker().catch(console.error);
