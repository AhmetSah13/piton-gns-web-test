import { expect, type Locator, type Page } from '@playwright/test';
import { BasePage } from './BasePage';

export class WidgetsPage extends BasePage {
  readonly dateInput: Locator;
  readonly sliderInput: Locator;
  readonly sliderValue: Locator;
  readonly progressBar: Locator;

  constructor(page: Page) {
    super(page);
    this.dateInput = page.locator('#datePickerMonthYearInput');
    this.sliderInput = page.locator('input[type="range"]');
    this.sliderValue = page.locator('#sliderValue');
    this.progressBar = page.getByRole('progressbar');
  }

  async openAccordian(): Promise<void> {
    await this.goto('/accordian');
  }

  async openDatePicker(): Promise<void> {
    await this.goto('/date-picker');
  }

  async selectDate(day: string): Promise<void> {
    await this.safeClick(this.dateInput);
    await this.safeClick(this.page.getByRole('gridcell', { name: new RegExp(`^Choose .+ May ${day}(st|nd|rd|th), 2026$`) }));
  }

  async openSlider(): Promise<void> {
    await this.goto('/slider');
  }

  async moveSlider(value = 70): Promise<void> {
    await this.sliderInput.focus();

    const currentValue = Number(await this.sliderInput.inputValue());
    const direction = value >= currentValue ? 'ArrowRight' : 'ArrowLeft';

    for (let step = 0; step < Math.abs(value - currentValue); step += 1) {
      await this.sliderInput.press(direction);
    }
  }

  async openProgressBar(): Promise<void> {
    await this.goto('/progress-bar');
  }

  async startProgressBar(): Promise<void> {
    await this.safeClick(this.page.getByRole('button', { name: 'Start' }));
  }

  async openTabs(): Promise<void> {
    await this.goto('/tabs');
  }

  async switchTab(tabName: string): Promise<void> {
    await this.safeClick(this.page.getByRole('tab', { name: tabName }));
  }

  async expectAccordianSectionOpen(sectionHeading: string, contentLocator: Locator): Promise<void> {
    await this.safeClick(this.page.getByText(sectionHeading, { exact: true }));
    await expect(contentLocator).toBeVisible();
  }
}
