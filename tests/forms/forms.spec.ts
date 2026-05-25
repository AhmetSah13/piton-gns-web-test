import { expect, test } from '@playwright/test';
import { FormsPage } from '../../pages/FormsPage';
import { invalidStudent, validStudent } from '../../test-data/demoqaTestData';

test.describe('DemoQA Forms module', () => {
  test('FRM-001 Practice Form should submit with valid required data', async ({ page }) => {
    const formsPage = new FormsPage(page);

    await formsPage.openPracticeForm();
    await formsPage.fillRequiredStudentFormFields(validStudent);
    await formsPage.submitForm();

    await formsPage.verifyFormSubmissionModal([
      `${validStudent.firstName} ${validStudent.lastName}`,
      validStudent.gender,
      validStudent.mobile,
    ]);
  });

  test('FRM-002 Practice Form should reject invalid email format visually', async ({ page }) => {
    const formsPage = new FormsPage(page);

    await formsPage.openPracticeForm();
    await formsPage.fillRequiredStudentFormFields({
      firstName: invalidStudent.firstName,
      lastName: invalidStudent.lastName,
      gender: 'Male',
      mobile: invalidStudent.mobile,
    });
    await formsPage.fillInvalidEmail(invalidStudent.invalidEmail);
    await formsPage.submitForm();

    await expect(formsPage.modal).toBeHidden();
    await expect(formsPage.getEmailInput()).toHaveCSS('border-color', 'rgb(220, 53, 69)');
  });

  test('FRM-003 Practice Form should require gender selection', async ({ page }) => {
    const formsPage = new FormsPage(page);

    await formsPage.openPracticeForm();
    await formsPage.fillRequiredStudentFormFields({
      firstName: validStudent.firstName,
      lastName: validStudent.lastName,
      email: validStudent.email,
      mobile: validStudent.mobile,
    });
    await formsPage.submitForm();

    await expect(formsPage.modal).toBeHidden();
    await expect(formsPage.getGenderLabel('Male')).toHaveCSS('color', 'rgb(220, 53, 69)');
  });

  test('FRM-004 Practice Form should accept mobile number with 10 digits', async ({ page }) => {
    const formsPage = new FormsPage(page);

    await formsPage.openPracticeForm();
    await formsPage.fillRequiredStudentFormFields(validStudent);
    await formsPage.submitForm();

    await formsPage.verifyFormSubmissionModal([validStudent.mobile]);
  });

  test('FRM-005 Practice Form should display submitted values in confirmation modal', async ({ page }) => {
    const formsPage = new FormsPage(page);

    await formsPage.openPracticeForm();
    await formsPage.fillRequiredStudentFormFields(validStudent);
    await formsPage.submitForm();

    await formsPage.verifyFormSubmissionModal([
      `${validStudent.firstName} ${validStudent.lastName}`,
      validStudent.email,
      validStudent.gender,
      validStudent.mobile,
    ]);
  });
});
