import { expect, type Locator, type Page } from '@playwright/test';
import { BasePage } from './BasePage';

export class ElementsPage extends BasePage {
  readonly textBoxOutput: Locator;
  readonly checkBoxResult: Locator;
  readonly radioResult: Locator;
  readonly doubleClickMessage: Locator;
  readonly rightClickMessage: Locator;

  constructor(page: Page) {
    super(page);
    this.textBoxOutput = page.locator('#output');
    this.checkBoxResult = page.locator('#result');
    this.radioResult = page.locator('.text-success');
    this.doubleClickMessage = page.locator('#doubleClickMessage');
    this.rightClickMessage = page.locator('#rightClickMessage');
  }

  async openTextBox(): Promise<void> {
    await this.goto('/text-box');
  }

  async submitTextBoxForm(user?: {
    fullName?: string;
    email?: string;
    currentAddress?: string;
    permanentAddress?: string;
  }): Promise<void> {
    if (user?.fullName) {
      await this.fillInput(this.page.getByPlaceholder('Full Name'), user.fullName);
    }
    if (user?.email) {
      await this.fillInput(this.page.getByPlaceholder('name@example.com'), user.email);
    }
    if (user?.currentAddress) {
      await this.fillInput(this.page.getByPlaceholder('Current Address'), user.currentAddress);
    }
    if (user?.permanentAddress) {
      await this.fillInput(this.page.locator('#permanentAddress'), user.permanentAddress);
    }

    await this.safeClick(this.page.getByRole('button', { name: 'Submit' }));
  }

  async openCheckBox(): Promise<void> {
    await this.goto('/checkbox');
  }

  async selectCheckBox(): Promise<void> {
    await this.safeClick(this.page.getByRole('checkbox', { name: 'Select Home' }));
  }

  async openRadioButton(): Promise<void> {
    await this.goto('/radio-button');
  }

  async selectRadioButton(option = 'Yes'): Promise<void> {
    await this.safeClick(this.page.getByText(option, { exact: true }));
  }

  async openButtons(): Promise<void> {
    await this.goto('/buttons');
  }

  async doubleClickButton(): Promise<void> {
    await this.page.getByRole('button', { name: 'Double Click Me' }).dblclick();
  }

  async rightClickButton(): Promise<void> {
    await this.page.getByRole('button', { name: 'Right Click Me' }).click({ button: 'right' });
  }

  async expectNoTextBoxOutput(): Promise<void> {
    await expect(this.textBoxOutput).toBeHidden();
  }
}
