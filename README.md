# PITON GNS Web Test

## Study Case Summary

This repository is a Playwright + TypeScript test automation study case for the PITON Technology internship evaluation. It automates selected DemoQA scenarios and prepares a documentation structure for manual exploratory testing reports.

Target automation site: https://demoqa.com/

Automated modules:

- Elements
- Forms
- Widgets

Each module currently contains exactly 5 meaningful test cases, including positive and negative coverage where it fits the DemoQA behavior.

## Tech Stack

- Playwright Test
- TypeScript
- Node.js
- GitHub Actions

## Folder Structure

```text
piton-gns-web-test/
  README.md
  package.json
  playwright.config.ts
  .gitignore
  pages/
    BasePage.ts
    ElementsPage.ts
    FormsPage.ts
    WidgetsPage.ts
  tests/
    elements/
      elements.spec.ts
    forms/
      forms.spec.ts
    widgets/
      widgets.spec.ts
  test-data/
    demoqaTestData.ts
  scripts/
    generate-docs.js
  docs/
    test-scenarios.xlsx
    exploratory-test-report.xlsx
    test-scenarios-draft.md
    exploratory-test-report-draft.md
    screenshots/
  .github/
    workflows/
      playwright.yml
```

## Installation

```bash
npm install
npx playwright install
```

## Running All Tests

```bash
npm test
```

Equivalent command:

```bash
npx playwright test
```

## Running Tests by Module

```bash
npm run test:elements
npm run test:forms
npm run test:widgets
```

## Viewing Playwright Report

```bash
npm run report
```

The HTML report is generated in `playwright-report/` after a test run.

Equivalent command:

```bash
npx playwright show-report
```

## Generating Excel Deliverables

```bash
npm run docs:generate
```

This command creates:

- `docs/test-scenarios.xlsx`
- `docs/exploratory-test-report.xlsx`

## Deliverables

- Automated Playwright tests for DemoQA Elements, Forms, and Widgets modules
- Test scenario Excel file: `docs/test-scenarios.xlsx`
- Exploratory test report Excel file: `docs/exploratory-test-report.xlsx`
- Playwright HTML report: `playwright-report/`
- Screenshot evidence folder: `docs/screenshots/`
- Short video walkthrough for study-case delivery

## Architecture

The automation uses a simple Page Object Model structure:

- `pages/BasePage.ts` stores the shared Playwright `Page` instance and common helper methods.
- `pages/ElementsPage.ts` contains reusable actions for the DemoQA Elements module.
- `pages/FormsPage.ts` contains reusable actions for the DemoQA Practice Form.
- `pages/WidgetsPage.ts` contains reusable actions for the DemoQA Widgets module.
- `test-data/demoqaTestData.ts` keeps reusable test data separate from test logic.
- `tests/` is organized by module so each area can be run independently.

Locators prefer Playwright user-facing selectors such as roles, text, and placeholders where possible. CSS locators are used only when DemoQA controls do not expose stable accessible selectors.

## CI/CD

GitHub Actions workflow is located at `.github/workflows/playwright.yml`.

The workflow runs on pushes and pull requests to `main` or `master`, installs dependencies with `npm ci`, installs the Chromium browser, runs `npm test`, and uploads the Playwright HTML report as a workflow artifact.

## Documentation Files

- `docs/test-scenarios-draft.md` contains the 15 drafted automation scenarios with expected results.
- `docs/test-scenarios.xlsx` contains the 15 automated DemoQA scenarios with local execution results.
- `docs/exploratory-test-report-draft.md` is a manual exploratory testing report template for:
  - https://www.gnsmetal.com/home
  - https://piton.com.tr/
- `docs/exploratory-test-report.xlsx` contains a structured manual exploratory report workbook with summary, findings, browser compatibility, and responsive checks sheets.
- `docs/screenshots/` is reserved for exploratory testing evidence and defect screenshots.

## Manual Exploratory Test Results

- DemoQA automation: 15 Playwright tests passed.
- PITON manual test: 2 low severity UI/responsive issues found.
- GNS Metal manual test: no critical issue observed in the tested scope.
- Browsers checked: Chrome and Edge.
- Viewports checked: desktop, mobile, and Galaxy Z Fold 5 responsive view.

Evidence screenshots for the open PITON findings are stored under `docs/screenshots/`.
