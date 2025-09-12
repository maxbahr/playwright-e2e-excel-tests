import { expect, test } from "@playwright/test";
import moment from "moment";
import { LaunchExcelPage } from "../pages/excel-launch.page";
import { ExcelWorkbookPage } from "../pages/excel-workbook.page";
import { LoginPage } from "../pages/login.page";
import { CellNumberFormattingOptions } from "../types/cell-formatting-options.type";
import { CellParameters } from "../types/cell-parameters.type";

test.describe("Excel Online", () => {
  test("function: today", async ({ page }) => {
    const dateFormat = "YYYY-MM-DD";
    const cellParameters: CellParameters = {
      address: "A2",
      info: "Contains Formula",
      formula: "=TODAY()",
      value: moment().format(dateFormat),
    };
    const cellNumberFormatOptions: CellNumberFormattingOptions = {
      category: "Date",
      locale: "English (United States)",
      type: moment("2012-03-14").format(dateFormat),
    };

    let excelWorkbookPage: ExcelWorkbookPage;
    await test.step("Login and create new workbook", async () => {
      const loginPage: LoginPage = new LoginPage(page);
      await loginPage.login(process.env.USER!, process.env.PASSWORD!);

      const launchExcelPage: LaunchExcelPage = new LaunchExcelPage(page);
      excelWorkbookPage = await launchExcelPage.createNewBlankWorkbook();
    });

    await test.step("Enter formula and verify cell value", async () => {
      await excelWorkbookPage.verifyTitle();
      await excelWorkbookPage.selectCell(cellParameters.address);
      await excelWorkbookPage.enterFormula(cellParameters.formula!);
      await excelWorkbookPage.selectCell(cellParameters.address);
      await excelWorkbookPage.openFormatCellPane();
      await excelWorkbookPage.formatCell(cellNumberFormatOptions);

      await expect
        .poll(
          async () => {
            console.log("Checking cell parameters...");
            return await excelWorkbookPage.getCellParams();
          },
          {
            message: "Waiting for cell parameters to match expected values",
            timeout: 2000,
          }
        )
        .toEqual(cellParameters);
    });
  });
});
