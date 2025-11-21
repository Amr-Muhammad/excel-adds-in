/// <reference types="office-js" />

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("generateBtn").onclick = generateBalanceSheet;

    // Set today's date as default
    const today = new Date().toISOString().split("T")[0];
    (document.getElementById("asOfDate") as HTMLInputElement).value = today;
  }
});

async function generateBalanceSheet() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getUsedRange().clear();

      // Get values from form
      const companyName = (document.getElementById("companyName") as HTMLInputElement).value;
      const asOfDate = (document.getElementById("asOfDate") as HTMLInputElement).value;
      const cash = parseInt((document.getElementById("cash") as HTMLInputElement).value);
      const accountsReceivable = parseInt(
        (document.getElementById("accountsReceivable") as HTMLInputElement).value
      );
      const inventory = parseInt((document.getElementById("inventory") as HTMLInputElement).value);
      const prepaidExpenses = parseInt(
        (document.getElementById("prepaidExpenses") as HTMLInputElement).value
      );
      const ppe = parseInt((document.getElementById("ppe") as HTMLInputElement).value);
      const depreciation = parseInt(
        (document.getElementById("depreciation") as HTMLInputElement).value
      );

      // Format date
      const dateObj = new Date(asOfDate);
      const formattedDate = dateObj.toLocaleDateString("en-US", {
        year: "numeric",
        month: "long",
        day: "numeric",
      });

      // Company header
      sheet.getRange("A1").values = [[companyName]];
      sheet.getRange("A1").format.font.bold = true;
      sheet.getRange("A1").format.font.size = 16;

      sheet.getRange("A2").values = [["Balance Sheet"]];
      sheet.getRange("A2").format.font.size = 14;

      sheet.getRange("A3").values = [[`As of ${formattedDate}`]];
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
        ["  Cash and Cash Equivalents", cash],
        ["  Accounts Receivable", accountsReceivable],
        ["  Inventory", inventory],
        ["  Prepaid Expenses", prepaidExpenses],
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
        ["  Property, Plant & Equipment", ppe],
        ["  Less: Accumulated Depreciation", -depreciation],
        ["  Intangible Assets", 30000],
        ["  Long-term Investments", 25000],
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

      // LIABILITIES (using fixed values for now)
      sheet.getRange("A22").values = [["LIABILITIES"]];
      sheet.getRange("A22").format.font.bold = true;
      sheet.getRange("A22").format.font.size = 12;
      sheet.getRange("A22").format.fill.color = "#4472C4";
      sheet.getRange("A22:B22").merge();
      sheet.getRange("A22").format.font.color = "white";

      const currentLiabilities = [
        ["Current Liabilities"],
        ["  Accounts Payable", 28000],
        ["  Short-term Debt", 15000],
        ["  Accrued Expenses", 12000],
        ["  Income Tax Payable", 8000],
        ["Total Current Liabilities", "=SUM(B24:B27)"],
      ];
      sheet.getRangeByIndexes(22, 0, currentLiabilities.length, 2).values = currentLiabilities;
      sheet.getRange("A23").format.font.bold = true;
      sheet.getRange("A23").format.font.underline = "Single";
      sheet.getRange("A28").format.font.bold = true;
      sheet.getRange("B28").format.font.bold = true;

      const nonCurrentLiabilities = [
        ["Non-Current Liabilities"],
        ["  Long-term Debt", 100000],
        ["  Deferred Tax Liabilities", 15000],
        ["Total Non-Current Liabilities", "=SUM(B31:B32)"],
      ];
      sheet.getRangeByIndexes(29, 0, nonCurrentLiabilities.length, 2).values =
        nonCurrentLiabilities;
      sheet.getRange("A30").format.font.bold = true;
      sheet.getRange("A30").format.font.underline = "Single";
      sheet.getRange("A33").format.font.bold = true;
      sheet.getRange("B33").format.font.bold = true;

      sheet.getRange("A35").values = [["TOTAL LIABILITIES"]];
      sheet.getRange("B35").formulas = [["=B28+B33"]];
      sheet.getRange("A35:B35").format.font.bold = true;
      sheet.getRange("A35:B35").format.fill.color = "#E7E6E6";

      // EQUITY
      sheet.getRange("A37").values = [["SHAREHOLDERS' EQUITY"]];
      sheet.getRange("A37").format.font.bold = true;
      sheet.getRange("A37").format.font.size = 12;
      sheet.getRange("A37").format.fill.color = "#4472C4";
      sheet.getRange("A37:B37").merge();
      sheet.getRange("A37").format.font.color = "white";

      const equity = [
        ["  Common Stock", 50000],
        ["  Retained Earnings", 119000],
        ["  Additional Paid-in Capital", 20000],
        ["Total Shareholders' Equity", "=SUM(B38:B40)"],
      ];
      sheet.getRangeByIndexes(37, 0, equity.length, 2).values = equity;
      sheet.getRange("A41").format.font.bold = true;
      sheet.getRange("B41").format.font.bold = true;

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

      // Column widths
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

      alert("Balance Sheet generated successfully!");
    });
  } catch (error) {
    console.error(error);
    alert("Error generating balance sheet: " + error);
  }
}
