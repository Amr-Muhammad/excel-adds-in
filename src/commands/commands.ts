/// <reference types="office-js" />

Office.onReady(() => {
  // Office is ready
});

/**
 * Shows a notification in Excel
 */
function showNotification(message: string) {
  Office.context.ui.displayDialogAsync(
    `data:text/html,<html><body><h2>${message}</h2></body></html>`,
    { height: 30, width: 20 }
  );
}

/**
 * Creates a Balance Sheet
 * @param event The event object from the ribbon button
 */
async function createBalanceSheet(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getUsedRange().clear();

      // Company header
      sheet.getRange("A1").values = [["ABC COMPANY"]];
      sheet.getRange("A1").format.font.bold = true;
      sheet.getRange("A1").format.font.size = 16;

      sheet.getRange("A2").values = [["Balance Sheet"]];
      sheet.getRange("A2").format.font.size = 14;

      sheet.getRange("A3").values = [["As of December 31, 2024"]];
      sheet.getRange("A3").format.font.italic = true;

      // ASSETS SECTION
      sheet.getRange("A5").values = [["ASSETS"]];
      sheet.getRange("A5").format.font.bold = true;
      sheet.getRange("A5").format.font.size = 12;
      sheet.getRange("A5").format.fill.color = "#4472C4";
      sheet.getRange("A5:B5").merge();
      sheet.getRange("A5").format.font.color = "white";

      // Current Assets
      sheet.getRange("A6").values = [["Current Assets"]];
      sheet.getRange("A6").format.font.bold = true;
      sheet.getRange("A6").format.font.underline = "Single";

      const currentAssets = [
        ["Cash and Cash Equivalents", 50000],
        ["Accounts Receivable", 35000],
        ["Inventory", 45000],
        ["Prepaid Expenses", 5000],
        ["Total Current Assets", "=SUM(B7:B10)"],
      ];
      sheet.getRangeByIndexes(6, 0, currentAssets.length, 2).values = currentAssets;
      sheet.getRange("A11").format.font.bold = true;
      sheet.getRange("B11").format.font.bold = true;

      // Non-Current Assets
      sheet.getRange("A13").values = [["Non-Current Assets"]];
      sheet.getRange("A13").format.font.bold = true;
      sheet.getRange("A13").format.font.underline = "Single";

      const nonCurrentAssets = [
        ["Property, Plant & Equipment", 200000],
        ["Less: Accumulated Depreciation", -50000],
        ["Intangible Assets", 30000],
        ["Long-term Investments", 25000],
        ["Total Non-Current Assets", "=SUM(B14:B17)"],
      ];
      sheet.getRangeByIndexes(13, 0, nonCurrentAssets.length, 2).values = nonCurrentAssets;
      sheet.getRange("A18").format.font.bold = true;
      sheet.getRange("B18").format.font.bold = true;

      // Total Assets
      sheet.getRange("A20").values = [["TOTAL ASSETS"]];
      sheet.getRange("B20").formulas = [["=B11+B18"]];
      sheet.getRange("A20:B20").format.font.bold = true;
      sheet.getRange("A20:B20").format.font.size = 12;
      sheet.getRange("A20:B20").format.fill.color = "#D9E1F2";
      sheet.getRange("A20:B20").format.borders.getItem("EdgeTop").style = "Double";
      sheet.getRange("A20:B20").format.borders.getItem("EdgeBottom").style = "Double";

      // LIABILITIES SECTION
      sheet.getRange("A22").values = [["LIABILITIES"]];
      sheet.getRange("A22").format.font.bold = true;
      sheet.getRange("A22").format.font.size = 12;
      sheet.getRange("A22").format.fill.color = "#4472C4";
      sheet.getRange("A22:B22").merge();
      sheet.getRange("A22").format.font.color = "white";

      // Current Liabilities
      sheet.getRange("A23").values = [["Current Liabilities"]];
      sheet.getRange("A23").format.font.bold = true;
      sheet.getRange("A23").format.font.underline = "Single";

      const currentLiabilities = [
        ["Accounts Payable", 28000],
        ["Short-term Debt", 15000],
        ["Accrued Expenses", 12000],
        ["Income Tax Payable", 8000],
        ["Total Current Liabilities", "=SUM(B24:B27)"],
      ];
      sheet.getRangeByIndexes(23, 0, currentLiabilities.length, 2).values = currentLiabilities;
      sheet.getRange("A28").format.font.bold = true;
      sheet.getRange("B28").format.font.bold = true;

      // Non-Current Liabilities
      sheet.getRange("A30").values = [["Non-Current Liabilities"]];
      sheet.getRange("A30").format.font.bold = true;
      sheet.getRange("A30").format.font.underline = "Single";

      const nonCurrentLiabilities = [
        ["Long-term Debt", 100000],
        ["Deferred Tax Liabilities", 15000],
        ["Total Non-Current Liabilities", "=SUM(B31:B32)"],
      ];
      sheet.getRangeByIndexes(30, 0, nonCurrentLiabilities.length, 2).values =
        nonCurrentLiabilities;
      sheet.getRange("A33").format.font.bold = true;
      sheet.getRange("B33").format.font.bold = true;

      // Total Liabilities
      sheet.getRange("A35").values = [["TOTAL LIABILITIES"]];
      sheet.getRange("B35").formulas = [["=B28+B33"]];
      sheet.getRange("A35:B35").format.font.bold = true;
      sheet.getRange("A35:B35").format.fill.color = "#E7E6E6";

      // EQUITY SECTION
      sheet.getRange("A37").values = [["SHAREHOLDERS' EQUITY"]];
      sheet.getRange("A37").format.font.bold = true;
      sheet.getRange("A37").format.font.size = 12;
      sheet.getRange("A37").format.fill.color = "#4472C4";
      sheet.getRange("A37:B37").merge();
      sheet.getRange("A37").format.font.color = "white";

      const equity = [
        ["Common Stock", 50000],
        ["Retained Earnings", 119000],
        ["Additional Paid-in Capital", 20000],
        ["Total Shareholders' Equity", "=SUM(B38:B40)"],
      ];
      sheet.getRangeByIndexes(37, 0, equity.length, 2).values = equity;
      sheet.getRange("A41").format.font.bold = true;
      sheet.getRange("B41").format.font.bold = true;

      // TOTAL LIABILITIES AND EQUITY
      sheet.getRange("A43").values = [["TOTAL LIABILITIES AND EQUITY"]];
      sheet.getRange("B43").formulas = [["=B35+B41"]];
      sheet.getRange("A43:B43").format.font.bold = true;
      sheet.getRange("A43:B43").format.font.size = 12;
      sheet.getRange("A43:B43").format.fill.color = "#D9E1F2";
      sheet.getRange("A43:B43").format.borders.getItem("EdgeTop").style = "Double";
      sheet.getRange("A43:B43").format.borders.getItem("EdgeBottom").style = "Double";

      // Format currency
      const currencyRanges = [
        "B7:B11",
        "B14:B18",
        "B20",
        "B24:B28",
        "B31:B33",
        "B35",
        "B38:B41",
        "B43",
      ];
      currencyRanges.forEach((range) => {
        sheet.getRange(range).numberFormat = [["$#,##0"]];
      });

      // Set column widths
      sheet.getRange("A:A").format.columnWidth = 280;
      sheet.getRange("B:B").format.columnWidth = 120;
      sheet.getRange("B:B").format.horizontalAlignment = "Right";

      // Borders
      sheet.getRange("A5:B20").format.borders.getItem("EdgeLeft").style = "Continuous";
      sheet.getRange("A5:B20").format.borders.getItem("EdgeRight").style = "Continuous";
      sheet.getRange("A22:B35").format.borders.getItem("EdgeLeft").style = "Continuous";
      sheet.getRange("A22:B35").format.borders.getItem("EdgeRight").style = "Continuous";
      sheet.getRange("A37:B43").format.borders.getItem("EdgeLeft").style = "Continuous";
      sheet.getRange("A37:B43").format.borders.getItem("EdgeRight").style = "Continuous";

      await context.sync();
    });

    event.completed();
  } catch (error) {
    console.error(error);
    event.completed();
  }
}

/**
 * Creates a Cash Flow Statement
 * @param event The event object from the ribbon button
 */
async function createCashFlowStatement(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getUsedRange().clear();

      // Company header
      sheet.getRange("A1").values = [["ABC COMPANY"]];
      sheet.getRange("A1").format.font.bold = true;
      sheet.getRange("A1").format.font.size = 16;

      sheet.getRange("A2").values = [["Statement of Cash Flows"]];
      sheet.getRange("A2").format.font.size = 14;

      sheet.getRange("A3").values = [["For the Year Ended December 31, 2024"]];
      sheet.getRange("A3").format.font.italic = true;

      // OPERATING ACTIVITIES
      sheet.getRange("A5").values = [["CASH FLOWS FROM OPERATING ACTIVITIES"]];
      sheet.getRange("A5").format.font.bold = true;
      sheet.getRange("A5").format.font.size = 12;
      sheet.getRange("A5").format.fill.color = "#4472C4";
      sheet.getRange("A5:B5").merge();
      sheet.getRange("A5").format.font.color = "white";

      const operatingActivities = [
        ["Net Income", 75000],
        ["Adjustments to reconcile net income to net cash:", ""],
        ["Depreciation and Amortization", 25000],
        ["Loss on Sale of Equipment", 3000],
        ["Changes in Operating Assets and Liabilities:", ""],
        ["Accounts Receivable", -8000],
        ["Inventory", -12000],
        ["Prepaid Expenses", 2000],
        ["Accounts Payable", 6000],
        ["Accrued Expenses", 4000],
        ["Income Tax Payable", 3000],
        ["Net Cash Provided by Operating Activities", "=B6+B8+B9+B11+B12+B13+B14+B15+B16"],
      ];
      sheet.getRangeByIndexes(5, 0, operatingActivities.length, 2).values = operatingActivities;
      sheet.getRange("A7").format.font.italic = true;
      sheet.getRange("A10").format.font.italic = true;
      sheet.getRange("A17").format.font.bold = true;
      sheet.getRange("B17").format.font.bold = true;
      sheet.getRange("A17:B17").format.fill.color = "#E7E6E6";
      sheet.getRange("A17:B17").format.borders.getItem("EdgeTop").style = "Continuous";

      // INVESTING ACTIVITIES
      sheet.getRange("A19").values = [["CASH FLOWS FROM INVESTING ACTIVITIES"]];
      sheet.getRange("A19").format.font.bold = true;
      sheet.getRange("A19").format.font.size = 12;
      sheet.getRange("A19").format.fill.color = "#4472C4";
      sheet.getRange("A19:B19").merge();
      sheet.getRange("A19").format.font.color = "white";

      const investingActivities = [
        ["Purchase of Property, Plant & Equipment", -45000],
        ["Purchase of Investments", -15000],
        ["Proceeds from Sale of Equipment", 8000],
        ["Purchase of Intangible Assets", -10000],
        ["Net Cash Used in Investing Activities", "=SUM(B20:B23)"],
      ];
      sheet.getRangeByIndexes(19, 0, investingActivities.length, 2).values = investingActivities;
      sheet.getRange("A24").format.font.bold = true;
      sheet.getRange("B24").format.font.bold = true;
      sheet.getRange("A24:B24").format.fill.color = "#E7E6E6";
      sheet.getRange("A24:B24").format.borders.getItem("EdgeTop").style = "Continuous";

      // FINANCING ACTIVITIES
      sheet.getRange("A26").values = [["CASH FLOWS FROM FINANCING ACTIVITIES"]];
      sheet.getRange("A26").format.font.bold = true;
      sheet.getRange("A26").format.font.size = 12;
      sheet.getRange("A26").format.fill.color = "#4472C4";
      sheet.getRange("A26:B26").merge();
      sheet.getRange("A26").format.font.color = "white";

      const financingActivities = [
        ["Proceeds from Issuance of Common Stock", 20000],
        ["Proceeds from Long-term Debt", 50000],
        ["Repayment of Short-term Debt", -10000],
        ["Payment of Dividends", -15000],
        ["Repurchase of Common Stock", -5000],
        ["Net Cash Provided by Financing Activities", "=SUM(B27:B31)"],
      ];
      sheet.getRangeByIndexes(26, 0, financingActivities.length, 2).values = financingActivities;
      sheet.getRange("A32").format.font.bold = true;
      sheet.getRange("B32").format.font.bold = true;
      sheet.getRange("A32:B32").format.fill.color = "#E7E6E6";
      sheet.getRange("A32:B32").format.borders.getItem("EdgeTop").style = "Continuous";

      // NET INCREASE/DECREASE
      sheet.getRange("A34").values = [["NET INCREASE (DECREASE) IN CASH"]];
      sheet.getRange("B34").formulas = [["=B17+B24+B32"]];
      sheet.getRange("A34:B34").format.font.bold = true;
      sheet.getRange("A34:B34").format.font.size = 11;
      sheet.getRange("A34:B34").format.fill.color = "#D9E1F2";
      sheet.getRange("A34:B34").format.borders.getItem("EdgeTop").style = "Continuous";

      // CASH BALANCES
      sheet.getRange("A36").values = [["Cash and Cash Equivalents at Beginning of Year", 25000]];
      sheet.getRange("A37").values = [["Cash and Cash Equivalents at End of Year"]];
      sheet.getRange("B37").formulas = [["=B36+B34"]];
      sheet.getRange("A37:B37").format.font.bold = true;
      sheet.getRange("A37:B37").format.font.size = 11;
      sheet.getRange("A37:B37").format.fill.color = "#D9E1F2";
      sheet.getRange("A37:B37").format.borders.getItem("EdgeTop").style = "Double";
      sheet.getRange("A37:B37").format.borders.getItem("EdgeBottom").style = "Double";

      // Format currency
      const currencyRanges = ["B6:B17", "B20:B24", "B27:B32", "B34", "B36:B37"];
      currencyRanges.forEach((range) => {
        sheet.getRange(range).numberFormat = [["$#,##0"]];
      });

      // Set column widths
      sheet.getRange("A:A").format.columnWidth = 350;
      sheet.getRange("B:B").format.columnWidth = 130;
      sheet.getRange("B:B").format.horizontalAlignment = "Right";

      await context.sync();
    });

    event.completed();
  } catch (error) {
    console.error(error);
    event.completed();
  }
}

/**
 * Creates an Income Statement
 * @param event The event object from the ribbon button
 */
async function createIncomeStatement(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getUsedRange().clear();

      // Company header
      sheet.getRange("A1").values = [["ABC COMPANY"]];
      sheet.getRange("A1").format.font.bold = true;
      sheet.getRange("A1").format.font.size = 16;

      sheet.getRange("A2").values = [["Income Statement"]];
      sheet.getRange("A2").format.font.size = 14;

      sheet.getRange("A3").values = [["For the Year Ended December 31, 2024"]];
      sheet.getRange("A3").format.font.italic = true;

      // REVENUE
      sheet.getRange("A5").values = [["REVENUE"]];
      sheet.getRange("A5").format.font.bold = true;
      sheet.getRange("A5").format.font.size = 12;
      sheet.getRange("A5").format.fill.color = "#4472C4";
      sheet.getRange("A5:B5").merge();
      sheet.getRange("A5").format.font.color = "white";

      const revenue = [
        ["Sales Revenue", 500000],
        ["Service Revenue", 150000],
        ["Total Revenue", "=SUM(B6:B7)"],
      ];
      sheet.getRangeByIndexes(5, 0, revenue.length, 2).values = revenue;
      sheet.getRange("A8").format.font.bold = true;
      sheet.getRange("B8").format.font.bold = true;
      sheet.getRange("A8:B8").format.fill.color = "#E7E6E6";

      // COST OF GOODS SOLD
      sheet.getRange("A10").values = [["COST OF GOODS SOLD"]];
      sheet.getRange("A10").format.font.bold = true;
      sheet.getRange("A10").format.font.size = 12;
      sheet.getRange("A10").format.fill.color = "#4472C4";
      sheet.getRange("A10:B10").merge();
      sheet.getRange("A10").format.font.color = "white";

      const cogs = [
        ["Cost of Goods Sold", 300000],
        ["GROSS PROFIT", "=B8-B11"],
      ];
      sheet.getRangeByIndexes(10, 0, cogs.length, 2).values = cogs;
      sheet.getRange("A12").format.font.bold = true;
      sheet.getRange("B12").format.font.bold = true;
      sheet.getRange("A12:B12").format.fill.color = "#D9E1F2";

      // OPERATING EXPENSES
      sheet.getRange("A14").values = [["OPERATING EXPENSES"]];
      sheet.getRange("A14").format.font.bold = true;
      sheet.getRange("A14").format.font.size = 12;
      sheet.getRange("A14").format.fill.color = "#4472C4";
      sheet.getRange("A14:B14").merge();
      sheet.getRange("A14").format.font.color = "white";

      const expenses = [
        ["Selling Expenses", 50000],
        ["Administrative Expenses", 75000],
        ["Research & Development", 30000],
        ["Depreciation & Amortization", 25000],
        ["Total Operating Expenses", "=SUM(B15:B18)"],
        ["OPERATING INCOME", "=B12-B19"],
      ];
      sheet.getRangeByIndexes(14, 0, expenses.length, 2).values = expenses;
      sheet.getRange("A19").format.font.bold = true;
      sheet.getRange("B19").format.font.bold = true;
      sheet.getRange("A20").format.font.bold = true;
      sheet.getRange("B20").format.font.bold = true;
      sheet.getRange("A20:B20").format.fill.color = "#D9E1F2";

      // OTHER INCOME/EXPENSES
      sheet.getRange("A22").values = [["OTHER INCOME (EXPENSES)"]];
      sheet.getRange("A22").format.font.bold = true;
      sheet.getRange("A22").format.font.size = 12;
      sheet.getRange("A22").format.fill.color = "#4472C4";
      sheet.getRange("A22:B22").merge();
      sheet.getRange("A22").format.font.color = "white";

      const other = [
        ["Interest Income", 5000],
        ["Interest Expense", -8000],
        ["Gain on Sale of Assets", 3000],
        ["Total Other Income (Expenses)", "=SUM(B23:B25)"],
        ["INCOME BEFORE TAXES", "=B20+B26"],
      ];
      sheet.getRangeByIndexes(22, 0, other.length, 2).values = other;
      sheet.getRange("A26").format.font.bold = true;
      sheet.getRange("B26").format.font.bold = true;
      sheet.getRange("A27").format.font.bold = true;
      sheet.getRange("B27").format.font.bold = true;
      sheet.getRange("A27:B27").format.fill.color = "#E7E6E6";

      // INCOME TAX
      const tax = [
        ["Income Tax Expense", "=B27*0.25"],
        ["NET INCOME", "=B27-B29"],
      ];
      sheet.getRangeByIndexes(28, 0, tax.length, 2).values = tax;
      sheet.getRange("A30").format.font.bold = true;
      sheet.getRange("B30").format.font.bold = true;
      sheet.getRange("A30:B30").format.font.size = 12;
      sheet.getRange("A30:B30").format.fill.color = "#D9E1F2";
      sheet.getRange("A30:B30").format.borders.getItem("EdgeTop").style = "Double";
      sheet.getRange("A30:B30").format.borders.getItem("EdgeBottom").style = "Double";

      // Format currency
      const currencyRanges = ["B6:B8", "B11:B12", "B15:B20", "B23:B27", "B29:B30"];
      currencyRanges.forEach((range) => {
        sheet.getRange(range).numberFormat = [["$#,##0"]];
      });

      // Set column widths
      sheet.getRange("A:A").format.columnWidth = 280;
      sheet.getRange("B:B").format.columnWidth = 120;
      sheet.getRange("B:B").format.horizontalAlignment = "Right";

      await context.sync();
    });

    event.completed();
  } catch (error) {
    console.error(error);
    event.completed();
  }
}

// Register functions
Office.actions.associate("createBalanceSheet", createBalanceSheet);
Office.actions.associate("createCashFlowStatement", createCashFlowStatement);
Office.actions.associate("createIncomeStatement", createIncomeStatement);
