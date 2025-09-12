import { Page } from "@playwright/test";
import { ExcelWorkbookPage } from "./excel-workbook.page";

export class LaunchExcelPage {
  constructor(private page: Page) {}

  async createNewBlankWorkbook(): Promise<ExcelWorkbookPage> {
    const [workbookPage] = await Promise.all([
      this.page.waitForEvent("popup"),
      this.page.locator('[role=link][aria-label="Create a new blank workbook"]').click(),
    ]);
    return new ExcelWorkbookPage([workbookPage][0]);
  }
}
