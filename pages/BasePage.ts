import { type Locator, type Page } from '@playwright/test';

export class BasePage {
  protected readonly page: Page;

  constructor(page: Page) {
    this.page = page;
  }

  async goto(path: string): Promise<void> {
    await this.page.goto(path);
  }

  async safeClick(locator: Locator): Promise<void> {
    await locator.scrollIntoViewIfNeeded();
    await locator.click();
  }

  async fillInput(locator: Locator, value: string): Promise<void> {
    await locator.scrollIntoViewIfNeeded();
    await locator.fill(value);
  }
}
