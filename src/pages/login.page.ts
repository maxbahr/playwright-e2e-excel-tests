import { expect, Page } from "@playwright/test";

export class LoginPage {
  constructor(private page: Page) {}

  async login(email: string, password: string): Promise<void> {
    if (!email || !password) {
      throw new Error("Email and password must be provided for login.");
    }

    await this.page.goto("/");
    await this.page.getByLabel("Email, phone, or Skype").fill(email);
    await this.page.getByRole("button", { name: "Next" }).click();
    await this.page.getByRole("button", { name: "Use your password" }).click();
    await this.page.getByTestId("passwordEntry").getByText("Password").click();
    await this.page.getByRole("textbox", { name: "Password" }).fill(password);
    await this.page.getByTestId("primaryButton").click();
    await this.page.getByTestId("secondaryButton").click();
    await expect(this.page).toHaveTitle("Excel | M365 Copilot");
  }
}
