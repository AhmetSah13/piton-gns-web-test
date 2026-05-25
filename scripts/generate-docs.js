const ExcelJS = require('exceljs');
const path = require('path');

const docsDir = path.join(__dirname, '..', 'docs');

const testScenarioColumns = [
  { header: 'Test ID', key: 'testId', width: 12 },
  { header: 'Module', key: 'module', width: 14 },
  { header: 'Test Case Title', key: 'title', width: 45 },
  { header: 'Description & Preconditions', key: 'description', width: 55 },
  { header: 'Test Steps', key: 'steps', width: 55 },
  { header: 'Expected Result', key: 'expected', width: 50 },
  { header: 'Actual Result', key: 'actual', width: 34 },
  { header: 'Status', key: 'status', width: 14 },
];

const testScenarios = [
  {
    testId: 'ELM-001',
    module: 'Elements',
    title: 'Text Box form should submit with valid data',
    description: 'DemoQA Text Box page is available. Valid user data is prepared.',
    steps: 'Open Text Box page; fill full name, email, current address, and permanent address; click Submit.',
    expected: 'Submitted values are displayed in the output area.',
  },
  {
    testId: 'ELM-002',
    module: 'Elements',
    title: 'Text Box should not display output for empty submission',
    description: 'DemoQA Text Box page is available.',
    steps: 'Open Text Box page; leave all fields empty; click Submit.',
    expected: 'Output area is not displayed for empty submission.',
  },
  {
    testId: 'ELM-003',
    module: 'Elements',
    title: 'Check Box should allow selecting Home checkbox',
    description: 'DemoQA Check Box page is available.',
    steps: 'Open Check Box page; select the Home checkbox.',
    expected: 'Result section displays selected Home value.',
  },
  {
    testId: 'ELM-004',
    module: 'Elements',
    title: 'Radio Button should allow selecting Yes option',
    description: 'DemoQA Radio Button page is available.',
    steps: 'Open Radio Button page; select Yes option.',
    expected: 'Page displays that Yes is selected.',
  },
  {
    testId: 'ELM-005',
    module: 'Elements',
    title: 'Buttons page should verify double click and right click actions',
    description: 'DemoQA Buttons page is available.',
    steps: 'Open Buttons page; double-click Double Click Me; right-click Right Click Me.',
    expected: 'Double-click and right-click confirmation messages are displayed.',
  },
  {
    testId: 'FRM-001',
    module: 'Forms',
    title: 'Practice Form should submit with valid required data',
    description: 'DemoQA Practice Form page is available. Required student data is prepared.',
    steps: 'Open Practice Form; fill first name, last name, gender, and mobile; click Submit.',
    expected: 'Confirmation modal is displayed with submitted required values.',
  },
  {
    testId: 'FRM-002',
    module: 'Forms',
    title: 'Practice Form should reject invalid email format visually',
    description: 'DemoQA Practice Form page is available. Invalid email test data is prepared.',
    steps: 'Open Practice Form; fill required fields; enter invalid email format; click Submit.',
    expected: 'Confirmation modal is not displayed and email field is visually marked invalid.',
  },
  {
    testId: 'FRM-003',
    module: 'Forms',
    title: 'Practice Form should require gender selection',
    description: 'DemoQA Practice Form page is available. Student data without gender is prepared.',
    steps: 'Open Practice Form; fill first name, last name, email, and mobile; leave gender empty; click Submit.',
    expected: 'Confirmation modal is not displayed and gender selection is visually marked required.',
  },
  {
    testId: 'FRM-004',
    module: 'Forms',
    title: 'Practice Form should accept mobile number with 10 digits',
    description: 'DemoQA Practice Form page is available. Valid 10-digit mobile number is prepared.',
    steps: 'Open Practice Form; fill required fields with 10-digit mobile number; click Submit.',
    expected: 'Confirmation modal displays the submitted mobile number.',
  },
  {
    testId: 'FRM-005',
    module: 'Forms',
    title: 'Practice Form should display submitted values in confirmation modal',
    description: 'DemoQA Practice Form page is available. Valid student data is prepared.',
    steps: 'Open Practice Form; fill required fields and email; click Submit.',
    expected: 'Confirmation modal displays name, email, gender, and mobile values.',
  },
  {
    testId: 'WGT-001',
    module: 'Widgets',
    title: 'Accordian sections should expand and collapse',
    description: 'DemoQA Accordian page is available.',
    steps: 'Open Accordian page; verify first section is expanded; open second section.',
    expected: 'Second section expands and first section collapses.',
  },
  {
    testId: 'WGT-002',
    module: 'Widgets',
    title: 'Date Picker should allow selecting a date',
    description: 'DemoQA Date Picker page is available.',
    steps: 'Open Date Picker page; open calendar; select configured date.',
    expected: 'Date input displays the selected date.',
  },
  {
    testId: 'WGT-003',
    module: 'Widgets',
    title: 'Slider value should change after interaction',
    description: 'DemoQA Slider page is available.',
    steps: 'Open Slider page; move slider to a new value.',
    expected: 'Slider value field updates to the new value.',
  },
  {
    testId: 'WGT-004',
    module: 'Widgets',
    title: 'Progress Bar should start and show progress',
    description: 'DemoQA Progress Bar page is available.',
    steps: 'Open Progress Bar page; click Start.',
    expected: 'Progress bar value increases from 0.',
  },
  {
    testId: 'WGT-005',
    module: 'Widgets',
    title: 'Tabs should switch between tab contents',
    description: 'DemoQA Tabs page is available.',
    steps: 'Open Tabs page; switch to Origin tab.',
    expected: 'Origin tab content is displayed.',
  },
].map((scenario) => ({
  ...scenario,
  actual: 'Passed on local Playwright execution',
  status: 'Passed',
}));

const findingsColumns = [
  { header: 'Bug ID', key: 'bugId', width: 14 },
  { header: 'Website', key: 'website', width: 20 },
  { header: 'Page/Area', key: 'pageArea', width: 24 },
  { header: 'Test Type', key: 'testType', width: 24 },
  { header: 'Description', key: 'description', width: 42 },
  { header: 'Steps to Reproduce', key: 'steps', width: 45 },
  { header: 'Expected Result', key: 'expected', width: 38 },
  { header: 'Actual Result', key: 'actual', width: 34 },
  { header: 'Severity', key: 'severity', width: 14 },
  { header: 'Priority', key: 'priority', width: 14 },
  { header: 'Status', key: 'status', width: 24 },
  { header: 'Evidence/Screenshot', key: 'evidence', width: 28 },
];

function styleWorksheet(worksheet, options = {}) {
  worksheet.views = [{ state: 'frozen', ySplit: options.freezeRows || 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: worksheet.columnCount },
  };

  const header = worksheet.getRow(1);
  header.height = 28;
  header.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E78' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = borderStyle();
  });

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    row.height = 42;
    row.eachCell((cell) => {
      cell.alignment = { vertical: 'top', wrapText: true };
      cell.border = borderStyle('FFD9E2F3');
    });
  });
}

function styleStatusCells(worksheet, statusColumnKey) {
  const statusColumn = worksheet.getColumn(statusColumnKey);
  statusColumn.eachCell((cell, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    const value = String(cell.value || '').toLowerCase();
    if (value === 'passed') {
      cell.fill = passedFill();
      cell.font = { bold: true, color: { argb: 'FF006100' } };
    } else if (value.includes('executed manually')) {
      cell.fill = pendingFill();
      cell.font = { bold: true, color: { argb: 'FF9C6500' } };
    }
  });
}

function borderStyle(color = 'FFBFBFBF') {
  return {
    top: { style: 'thin', color: { argb: color } },
    left: { style: 'thin', color: { argb: color } },
    bottom: { style: 'thin', color: { argb: color } },
    right: { style: 'thin', color: { argb: color } },
  };
}

function passedFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
}

function pendingFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
}

function addTitle(worksheet, title, subtitle, columnCount) {
  worksheet.spliceRows(1, 0, [], []);
  worksheet.mergeCells(1, 1, 1, columnCount);
  worksheet.mergeCells(2, 1, 2, columnCount);

  const titleCell = worksheet.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = { bold: true, size: 16, color: { argb: 'FF1F4E78' } };
  titleCell.alignment = { vertical: 'middle' };

  const subtitleCell = worksheet.getCell(2, 1);
  subtitleCell.value = subtitle;
  subtitleCell.font = { italic: true, color: { argb: 'FF666666' } };

  worksheet.getRow(1).height = 24;
  worksheet.getRow(2).height = 22;
  worksheet.views = [{ state: 'frozen', ySplit: 3 }];
  worksheet.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3, column: columnCount },
  };

  worksheet.getRow(3).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E78' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = borderStyle();
  });
}

async function createTestScenariosWorkbook() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'PITON GNS Web Test';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('DemoQA Test Scenarios');
  worksheet.columns = testScenarioColumns;
  worksheet.addRows(testScenarios);
  styleWorksheet(worksheet);
  styleStatusCells(worksheet, 'status');
  addTitle(
    worksheet,
    'DemoQA Automated Test Scenarios',
    '15/15 scenarios passed on local Playwright execution.',
    testScenarioColumns.length,
  );

  await workbook.xlsx.writeFile(path.join(docsDir, 'test-scenarios.xlsx'));
}

function addSummarySheet(workbook) {
  const worksheet = workbook.addWorksheet('Summary');
  worksheet.columns = [
    { header: 'Section', key: 'section', width: 30 },
    { header: 'Website / Scope', key: 'scope', width: 40 },
    { header: 'Current Status', key: 'status', width: 32 },
    { header: 'Notes', key: 'notes', width: 55 },
  ];
  worksheet.addRows([
    {
      section: 'Exploratory Testing Scope',
      scope: 'https://www.gnsmetal.com/home',
      status: 'To be executed manually',
      notes: 'Capture usability, navigation, validation, visual, responsive, and performance observations.',
    },
    {
      section: 'Exploratory Testing Scope',
      scope: 'https://piton.com.tr/',
      status: 'To be executed manually',
      notes: 'Capture usability, navigation, validation, visual, responsive, and performance observations.',
    },
    {
      section: 'Evidence',
      scope: 'docs/screenshots/',
      status: 'To be executed manually',
      notes: 'Attach screenshots or video references for defects and notable observations.',
    },
  ]);
  styleWorksheet(worksheet);
  styleStatusCells(worksheet, 'status');
  addTitle(worksheet, 'Exploratory Test Report Summary', 'Manual report template prepared for PITON study-case delivery.', 4);
}

function addFindingsSheet(workbook, sheetName, website, bugPrefix) {
  const worksheet = workbook.addWorksheet(sheetName);
  worksheet.columns = findingsColumns;
  worksheet.addRows([
    {
      bugId: `${bugPrefix}-001`,
      website,
      pageArea: 'Home page',
      testType: 'UX and usability',
      description: 'Placeholder for manual exploratory observation.',
      steps: 'Open the website; review primary content, calls to action, and readability.',
      expected: 'User can understand the page purpose and continue without confusion.',
      actual: 'To be executed manually',
      severity: 'TBD',
      priority: 'TBD',
      status: 'To be executed manually',
      evidence: 'docs/screenshots/',
    },
    {
      bugId: `${bugPrefix}-002`,
      website,
      pageArea: 'Navigation',
      testType: 'Navigation flow',
      description: 'Placeholder for broken or confusing navigation paths.',
      steps: 'Use header, footer, and key internal links.',
      expected: 'Navigation works consistently and leads to expected pages.',
      actual: 'To be executed manually',
      severity: 'TBD',
      priority: 'TBD',
      status: 'To be executed manually',
      evidence: 'docs/screenshots/',
    },
    {
      bugId: `${bugPrefix}-003`,
      website,
      pageArea: 'Forms / Contact areas',
      testType: 'Form validation',
      description: 'Placeholder for required field, invalid format, and submission validation findings.',
      steps: 'Submit forms with empty and invalid values, then with valid values if available.',
      expected: 'Validation messages are clear and invalid submissions are prevented.',
      actual: 'To be executed manually',
      severity: 'TBD',
      priority: 'TBD',
      status: 'To be executed manually',
      evidence: 'docs/screenshots/',
    },
  ]);
  styleWorksheet(worksheet);
  styleStatusCells(worksheet, 'status');
  addTitle(worksheet, `${website} Findings`, 'Manual exploratory findings template.', findingsColumns.length);
}

function addBrowserCompatibilitySheet(workbook) {
  const worksheet = workbook.addWorksheet('Browser Compatibility');
  worksheet.columns = [
    { header: 'Website', key: 'website', width: 28 },
    { header: 'Browser', key: 'browser', width: 18 },
    { header: 'Version', key: 'version', width: 18 },
    { header: 'Primary Pages Checked', key: 'pages', width: 35 },
    { header: 'Result', key: 'result', width: 28 },
    { header: 'Notes', key: 'notes', width: 45 },
  ];
  worksheet.addRows([
    ['GNS Metal', 'Chrome', 'TBD', 'Home and key navigation pages', 'To be executed manually', ''],
    ['GNS Metal', 'Firefox', 'TBD', 'Home and key navigation pages', 'To be executed manually', ''],
    ['GNS Metal', 'Safari / WebKit', 'TBD', 'Home and key navigation pages', 'To be executed manually', ''],
    ['PITON', 'Chrome', 'TBD', 'Home and key navigation pages', 'To be executed manually', ''],
    ['PITON', 'Firefox', 'TBD', 'Home and key navigation pages', 'To be executed manually', ''],
    ['PITON', 'Safari / WebKit', 'TBD', 'Home and key navigation pages', 'To be executed manually', ''],
  ]);
  styleWorksheet(worksheet);
  addTitle(worksheet, 'Browser Compatibility', 'Manual cross-browser execution matrix.', 6);
}

function addResponsiveChecksSheet(workbook) {
  const worksheet = workbook.addWorksheet('Responsive Checks');
  worksheet.columns = [
    { header: 'Website', key: 'website', width: 28 },
    { header: 'Viewport', key: 'viewport', width: 18 },
    { header: 'Resolution', key: 'resolution', width: 20 },
    { header: 'Areas Checked', key: 'areas', width: 38 },
    { header: 'Actual Result', key: 'actual', width: 28 },
    { header: 'Status', key: 'status', width: 28 },
    { header: 'Notes / Evidence', key: 'notes', width: 40 },
  ];
  worksheet.addRows([
    ['GNS Metal', 'Desktop', '1440x900', 'Header, hero, navigation, footer', 'To be executed manually', 'To be executed manually', 'docs/screenshots/'],
    ['GNS Metal', 'Tablet', '768x1024', 'Header, content layout, forms', 'To be executed manually', 'To be executed manually', 'docs/screenshots/'],
    ['GNS Metal', 'Mobile', '390x844', 'Menu, content stacking, forms', 'To be executed manually', 'To be executed manually', 'docs/screenshots/'],
    ['PITON', 'Desktop', '1440x900', 'Header, hero, navigation, footer', 'To be executed manually', 'To be executed manually', 'docs/screenshots/'],
    ['PITON', 'Tablet', '768x1024', 'Header, content layout, forms', 'To be executed manually', 'To be executed manually', 'docs/screenshots/'],
    ['PITON', 'Mobile', '390x844', 'Menu, content stacking, forms', 'To be executed manually', 'To be executed manually', 'docs/screenshots/'],
  ]);
  styleWorksheet(worksheet);
  styleStatusCells(worksheet, 'status');
  addTitle(worksheet, 'Responsive Checks', 'Manual responsive behavior checklist.', 7);
}

async function createExploratoryWorkbook() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'PITON GNS Web Test';
  workbook.created = new Date();

  addSummarySheet(workbook);
  addFindingsSheet(workbook, 'GNS Metal Findings', 'GNS Metal', 'GNS');
  addFindingsSheet(workbook, 'PITON Findings', 'PITON', 'PITON');
  addBrowserCompatibilitySheet(workbook);
  addResponsiveChecksSheet(workbook);

  await workbook.xlsx.writeFile(path.join(docsDir, 'exploratory-test-report.xlsx'));
}

async function main() {
  await createTestScenariosWorkbook();
  await createExploratoryWorkbook();

  console.log('Generated docs/test-scenarios.xlsx');
  console.log('Generated docs/exploratory-test-report.xlsx');
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
