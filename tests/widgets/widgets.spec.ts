import { expect, test } from '@playwright/test';
import { WidgetsPage } from '../../pages/WidgetsPage';
import { datePickerData } from '../../test-data/demoqaTestData';

test.describe('DemoQA Widgets module', () => {
  test('WGT-001 Accordian sections should expand and collapse', async ({ page }) => {
    const widgetsPage = new WidgetsPage(page);
    const firstSection = page.getByRole('button', { name: 'What is Lorem Ipsum?' });
    const secondSection = page.getByRole('button', { name: 'Where does it come from?' });

    await widgetsPage.openAccordian();
    await expect(firstSection).toHaveAttribute('aria-expanded', 'true');
    await secondSection.click();

    await expect(secondSection).toHaveAttribute('aria-expanded', 'true');
    await expect(firstSection).toHaveAttribute('aria-expanded', 'false');
  });

  test('WGT-002 Date Picker should allow selecting a date', async ({ page }) => {
    const widgetsPage = new WidgetsPage(page);

    await widgetsPage.openDatePicker();
    await widgetsPage.selectDate(datePickerData.date);

    await expect(widgetsPage.dateInput).toHaveValue(datePickerData.expectedDate);
  });

  test('WGT-003 Slider value should change after interaction', async ({ page }) => {
    const widgetsPage = new WidgetsPage(page);

    await widgetsPage.openSlider();
    await widgetsPage.moveSlider(70);

    await expect(widgetsPage.sliderValue).toHaveValue('70');
  });

  test('WGT-004 Progress Bar should start and show progress', async ({ page }) => {
    const widgetsPage = new WidgetsPage(page);

    await widgetsPage.openProgressBar();
    await widgetsPage.startProgressBar();

    await expect(widgetsPage.progressBar).not.toHaveAttribute('aria-valuenow', '0');
  });

  test('WGT-005 Tabs should switch between tab contents', async ({ page }) => {
    const widgetsPage = new WidgetsPage(page);

    await widgetsPage.openTabs();
    await widgetsPage.switchTab('Origin');

    await expect(page.getByRole('tabpanel')).toContainText('Contrary to popular belief');
  });
});
