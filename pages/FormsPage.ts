import { expect, type Locator, type Page } from '@playwright/test';
import { BasePage } from './BasePage';

export class FormsPage extends BasePage {
  readonly modal: Locator;

  constructor(page: Page) {
    super(page);
    this.modal = page.getByRole('dialog');
  }

  async openPracticeForm(): Promise<void> {
    await this.goto('/automation-practice-form');
  }

  async fillRequiredStudentFormFields(student: {
    firstName: string;
    lastName: string;
    email?: string;
    gender?: string;
    mobile: string;
  }): Promise<void> {
    await this.fillInput(this.page.getByPlaceholder('First Name'), student.firstName);
    await this.fillInput(this.page.getByPlaceholder('Last Name'), student.lastName);
    if (student.email) {
      await this.fillInput(this.page.getByPlaceholder('name@example.com'), student.email);
    }
    if (student.gender) {
      await this.selectGender(student.gender);
    }
    await this.fillInput(this.page.getByPlaceholder('Mobile Number'), student.mobile);
  }

  async fillInvalidEmail(email: string): Promise<void> {
    await this.fillInput(this.page.getByPlaceholder('name@example.com'), email);
  }

  async selectGender(gender: string): Promise<void> {
    await this.safeClick(this.page.getByText(gender, { exact: true }));
  }

  async submitForm(): Promise<void> {
    await this.safeClick(this.page.getByRole('button', { name: 'Submit' }));
  }

  async verifyFormSubmissionModal(expectedValues: string[]): Promise<void> {
    await expect(this.modal).toBeVisible();
    for (const value of expectedValues) {
      await expect(this.modal).toContainText(value);
    }
  }

  getEmailInput(): Locator {
    return this.page.getByPlaceholder('name@example.com');
  }

  getGenderLabel(gender: string): Locator {
    return this.page.locator(`label[for="gender-radio-${this.genderValue(gender)}"]`);
  }

  getMobileInput(): Locator {
    return this.page.getByPlaceholder('Mobile Number');
  }

  private genderValue(gender: string): number {
    const genderMap: Record<string, number> = {
      Male: 1,
      Female: 2,
      Other: 3,
    };

    return genderMap[gender];
  }
}
