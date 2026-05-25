import { expect, test } from '@playwright/test';
import { ElementsPage } from '../../pages/ElementsPage';
import { textBoxUser } from '../../test-data/demoqaTestData';

test.describe('DemoQA Elements module', () => {
  test('ELM-001 Text Box form should submit with valid data', async ({ page }) => {
    const elementsPage = new ElementsPage(page);

    await elementsPage.openTextBox();
    await elementsPage.submitTextBoxForm(textBoxUser);

    await expect(elementsPage.textBoxOutput).toContainText(textBoxUser.fullName);
    await expect(elementsPage.textBoxOutput).toContainText(textBoxUser.email);
    await expect(elementsPage.textBoxOutput).toContainText(textBoxUser.currentAddress);
    await expect(elementsPage.textBoxOutput).toContainText(textBoxUser.permanentAddress);
  });

  test('ELM-002 Text Box should not display output for empty submission', async ({ page }) => {
    const elementsPage = new ElementsPage(page);

    await elementsPage.openTextBox();
    await elementsPage.submitTextBoxForm();

    await elementsPage.expectNoTextBoxOutput();
  });

  test('ELM-003 Check Box should allow selecting Home checkbox', async ({ page }) => {
    const elementsPage = new ElementsPage(page);

    await elementsPage.openCheckBox();
    await elementsPage.selectCheckBox();

    await expect(elementsPage.checkBoxResult).toContainText('home');
  });

  test('ELM-004 Radio Button should allow selecting Yes option', async ({ page }) => {
    const elementsPage = new ElementsPage(page);

    await elementsPage.openRadioButton();
    await elementsPage.selectRadioButton('Yes');

    await expect(elementsPage.radioResult).toHaveText('Yes');
  });

  test('ELM-005 Buttons page should verify double click and right click actions', async ({ page }) => {
    const elementsPage = new ElementsPage(page);

    await elementsPage.openButtons();
    await elementsPage.doubleClickButton();
    await elementsPage.rightClickButton();

    await expect(elementsPage.doubleClickMessage).toHaveText('You have done a double click');
    await expect(elementsPage.rightClickMessage).toHaveText('You have done a right click');
  });
});
