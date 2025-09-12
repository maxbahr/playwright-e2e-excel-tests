import { expect, FrameLocator, Locator, Page } from "@playwright/test";
import { CellNumberFormattingOptions } from "../types/cell-formatting-options.type";
import { CellParameters } from "../types/cell-parameters.type";

export class ExcelWorkbookPage {
  iframe: FrameLocator;
  cellReader: Locator;
  formulaBar: Locator;

  constructor(private page: Page) {
    this.iframe = this.page.locator('iframe[name="WacFrame_Excel_0"]').contentFrame();
    this.cellReader = this.iframe.locator("#m_excelWebRenderer_ewaCtl_readoutElement1");
    this.formulaBar = this.iframe.getByRole("textbox", { name: "formula bar" });
  }

  async verifyTitle(): Promise<void> {
    await expect(this.page).toHaveTitle(/Book( \d+)?\.xlsx/);
  }

  async selectCell(cell: string): Promise<void> {
    await expect
      .poll(
        async () => {
          const cellAddressBox = this.iframe.getByRole("combobox", { name: "Name Box" });
          await cellAddressBox.click();
          await cellAddressBox.clear();
          await cellAddressBox.fill("");
          await cellAddressBox.pressSequentially(cell, { delay: 50 });
          await cellAddressBox.press("Enter");

          await expect(this.cellReader).toHaveAttribute("aria-label", new RegExp(`\\b${cell}\\b`));

          return cellAddressBox.inputValue();
        },
        { timeout: 30000 }
      )
      .toBe(cell);
  }

  async openFormatCellPane(): Promise<void> {
    await expect
      .poll(
        async () => {
          await this.iframe.getByRole("button", { name: "More Options", exact: true }).click();
          await this.iframe.getByRole("menuitem", { name: "Format", exact: true }).click();
          await this.iframe.getByRole("menuitem", { name: "Format Cells" }).click();
          await this.iframe.getByTestId("sidebar").getByRole("group", { name: "Number" }).click();
          return this.iframe.getByRole("heading", { name: "All formatting options" }).isVisible();
        },
        { timeout: 5_000 }
      )
      .toBeTruthy();
  }

  async formatCell(options: CellNumberFormattingOptions): Promise<void> {
    await this.iframe.getByRole("tab", { name: "Number" }).click();
    await this.iframe.getByRole("combobox", { name: "Choose Category" }).locator("div").click();
    await this.iframe.getByRole("option", { name: options.category }).click();

    if (options.locale) {
      const comboboxLocale = this.iframe.getByRole("combobox", { name: "Locale (location)" });
      await comboboxLocale.click();
      const listbox = this.iframe.locator("#LocaleDropdown-list");
      await listbox.getByRole("option", { name: options.locale }).first().click();
      await expect(comboboxLocale).toHaveValue(options.locale);
    }

    if (options.type) {
      const listType = this.iframe
        .locator(".ms-DetailsList")
        .getByRole("gridcell", { name: options.type, exact: true });
      await listType.click();
      await expect(listType.locator("..").locator("..")).toHaveClass(/is-selected/);
    }

    await this.iframe.getByRole("button", { name: "Close" }).click();
  }

  async enterFormula(formula: string): Promise<void> {
    await this.formulaBar.click();
    await this.formulaBar.clear();
    await this.formulaBar.pressSequentially(formula, { delay: 50 });
    await this.formulaBar.press("Enter");
  }

  async getCellParams(): Promise<CellParameters> {
    const label = (await this.cellReader.getAttribute("aria-label")) || undefined;

    if (!label) {
      throw new Error("The cell reader aria-label is empty or undefined.");
    }

    const labelValues: string[] = label.split(" . ").filter((s) => s.length > 0);
    console.log("Cell:", ...labelValues);

    let cellProps: CellParameters;

    switch (labelValues.length) {
      case 3:
        //formula in cell: value + address + info
        cellProps = {
          value: labelValues[0],
          address: labelValues[1],
          info: labelValues[2],
          formula: (await this.formulaBar.textContent()) || undefined,
        };
        break;
      case 2:
        //text in cell: value + address
        cellProps = {
          value: labelValues[0],
          address: labelValues[1],
        };
        break;
      case 1:
        //empty cell: address
        cellProps = {
          address: labelValues[0],
        };
        break;
      default:
        throw new Error("Unexpected label format");
    }

    return cellProps;
  }
}
