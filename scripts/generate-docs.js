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
  { header: 'Bug/Check ID', key: 'bugId', width: 18 },
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
    } else if (value === 'observation') {
      cell.fill = pendingFill();
      cell.font = { bold: true, color: { argb: 'FF9C6500' } };
    } else if (value === 'open') {
      cell.fill = openFill();
      cell.font = { bold: true, color: { argb: 'FF9C0006' } };
    } else if (value === 'not applicable') {
      cell.fill = neutralFill();
      cell.font = { bold: true, color: { argb: 'FF666666' } };
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

function openFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
}

function neutralFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
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
    { header: 'Category', key: 'category', width: 30 },
    { header: 'Item', key: 'item', width: 42 },
    { header: 'Status / Value', key: 'status', width: 32 },
    { header: 'Notes', key: 'notes', width: 62 },
  ];
  worksheet.addRows([
    {
      category: 'Project name',
      item: 'PITON & GNS Web Test',
      status: 'Completed',
      notes: 'Manual exploratory testing deliverable for corporate websites, alongside DemoQA automation.',
    },
    {
      category: 'Test type',
      item: 'Manual Exploratory Testing',
      status: 'Completed',
      notes: 'Corporate websites were tested manually, not automated.',
    },
    {
      category: 'Website tested',
      item: 'https://www.gnsmetal.com/home',
      status: 'Completed',
      notes: 'GNS Metal website tested in the documented scope.',
    },
    {
      category: 'Website tested',
      item: 'https://piton.com.tr/',
      status: 'Completed',
      notes: 'PITON website tested in the documented scope.',
    },
    {
      category: 'Browsers checked',
      item: 'Chrome, Edge',
      status: 'Passed',
      notes: 'No major browser-specific rendering difference was observed in the tested scope.',
    },
    {
      category: 'Viewports checked',
      item: 'Desktop, Tablet, Mobile, Galaxy Z Fold 5 responsive view',
      status: 'Open',
      notes: 'PITON Galaxy Z Fold 5 responsive view has a low severity hero text clipping issue.',
    },
    {
      category: 'Automation note',
      item: 'Corporate websites',
      status: 'Not Applicable',
      notes: 'No automation was written for corporate websites, in accordance with the study-case requirement.',
    },
    {
      category: 'DemoQA automation',
      item: 'https://demoqa.com/',
      status: 'Passed',
      notes: '15 Playwright tests passed for Elements, Forms, and Widgets modules.',
    },
    {
      category: 'PITON summary counts',
      item: 'Passed checks: 11; Open issues: 2',
      status: 'Open',
      notes: 'PITON has two low-severity UI/responsive findings.',
    },
    {
      category: 'GNS Metal summary counts',
      item: 'Passed checks: 9; Not Applicable checks: 1; Observations: 1; Open issues: 0',
      status: 'Observation',
      notes: 'GNS Metal had one low-severity page loading observation and no critical functional issue observed in the tested scope.',
    },
    {
      category: 'Evidence',
      item: 'docs/screenshots/',
      status: 'Passed',
      notes: 'Evidence screenshots for open PITON issues are stored under docs/screenshots/.',
    },
    {
      category: 'Brief conclusion',
      item: 'Manual exploratory result',
      status: 'Completed',
      notes: 'PITON has two low-severity UI/responsive findings. GNS Metal has one low-severity page loading observation and no critical functional issue observed in the tested scope.',
    },
  ]);
  styleWorksheet(worksheet);
  styleStatusCells(worksheet, 'status');
  addTitle(worksheet, 'Exploratory Test Report Summary', 'Manual exploratory observations based on real execution results.', 4);
}

function addFindingsSheet(workbook, sheetName, website, bugPrefix) {
  const worksheet = workbook.addWorksheet(sheetName);
  worksheet.columns = findingsColumns;
  worksheet.addRows(getFindingsRows(website, bugPrefix));
  styleWorksheet(worksheet);
  styleStatusCells(worksheet, 'status');
  addTitle(worksheet, `${website} Findings`, 'Manual exploratory findings from tested scope.', findingsColumns.length);
}

function getFindingsRows(website, bugPrefix) {
  if (bugPrefix === 'PITON') {
    return [
      {
        bugId: 'PITON-UX-001',
        website,
        pageArea: 'Homepage',
        testType: 'UX and usability',
        description: 'UX and homepage usability checked.',
        steps: 'Open PITON homepage and review main content, visual hierarchy, readability, and primary navigation entry points.',
        expected: 'Homepage is understandable and usable for a first-time visitor.',
        actual: 'Homepage was generally usable in the tested scope.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-NAV-001',
        website,
        pageArea: 'Header menu',
        testType: 'Navigation flow',
        description: 'Header/menu navigation links checked.',
        steps: 'Open PITON website and click header/menu navigation links.',
        expected: 'Navigation links open the expected pages or sections.',
        actual: 'Header/menu navigation links work correctly.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-FOOT-001',
        website,
        pageArea: 'Footer',
        testType: 'Navigation flow',
        description: 'Footer links checked.',
        steps: 'Open PITON website and click footer links.',
        expected: 'Footer links open the expected destinations.',
        actual: 'Footer links work correctly.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-FORM-001',
        website,
        pageArea: 'Contact form',
        testType: 'Form validation',
        description: 'Empty submit validation checked.',
        steps: 'Open contact form and submit without filling required fields.',
        expected: 'Required field warnings are displayed and empty submission is prevented.',
        actual: 'Contact form warns the user when required fields are empty.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-FORM-002',
        website,
        pageArea: 'Contact form email field',
        testType: 'Form validation',
        description: 'Invalid email validation checked.',
        steps: 'Enter an invalid email value without @ and submit the contact form.',
        expected: 'Email field displays a clear invalid email warning.',
        actual: 'Email field warns the user that an @ character is required.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-UI-001',
        website,
        pageArea: 'Contact form subject dropdown',
        testType: 'Visual/UI consistency',
        description: 'Contact form subject dropdown visual contrast/theme inconsistency.',
        steps: 'Open the contact form and expand the Konu dropdown.',
        expected: 'Dropdown option area follows the site dark theme and remains readable.',
        actual: 'Dropdown option area appears as a large white/blank area, creating visual consistency and contrast problems.',
        severity: 'Low',
        priority: 'Medium',
        status: 'Open',
        evidence: 'docs/screenshots/piton-contact-subject-dropdown-contrast-issue.png',
      },
      {
        bugId: 'PITON-UI-002',
        website,
        pageArea: 'General pages',
        testType: 'Visual/UI consistency',
        description: 'Visual/UI consistency general check.',
        steps: 'Review tested PITON pages for obvious layout, spacing, typography, and color consistency issues.',
        expected: 'No major visual consistency problem is observed outside documented defects.',
        actual: 'General visual/UI consistency looked acceptable in the tested scope, except PITON-UI-001.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-RESP-002',
        website,
        pageArea: 'Desktop layout',
        testType: 'Responsive behavior',
        description: 'Responsive desktop check.',
        steps: 'Open PITON website in desktop viewport and inspect main sections.',
        expected: 'Layout is readable and usable on desktop.',
        actual: 'Desktop view worked correctly in the tested scope.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-RESP-003',
        website,
        pageArea: 'Mobile layout',
        testType: 'Responsive behavior',
        description: 'Responsive mobile check.',
        steps: 'Open PITON website in mobile viewport and inspect navigation, content stacking, and contact form areas.',
        expected: 'Mobile layout is readable and usable.',
        actual: 'Mobile view generally works correctly.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-RESP-001',
        website,
        pageArea: 'Homepage hero',
        testType: 'Responsive behavior',
        description: 'Homepage hero text clipping on Galaxy Z Fold 5 responsive view.',
        steps: 'Open PITON homepage in Galaxy Z Fold 5 responsive viewport and inspect hero text.',
        expected: 'Hero text should be fully visible without clipping.',
        actual: 'A small part of the homepage hero text is clipped/not fully visible.',
        severity: 'Low',
        priority: 'Low',
        status: 'Open',
        evidence: 'docs/screenshots/piton-home-galaxy-z-fold-text-clipping.png',
      },
      {
        bugId: 'PITON-COMPAT-CHROME-001',
        website,
        pageArea: 'General visual layout',
        testType: 'Browser compatibility',
        description: 'Chrome compatibility checked.',
        steps: 'Open PITON website in Chrome and review tested pages.',
        expected: 'No major Chrome-specific rendering issue is observed.',
        actual: 'Chrome visual check looks consistent.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-COMPAT-EDGE-001',
        website,
        pageArea: 'General visual layout',
        testType: 'Browser compatibility',
        description: 'Edge compatibility checked.',
        steps: 'Open PITON website in Edge and review tested pages.',
        expected: 'No major Edge-specific rendering issue is observed.',
        actual: 'Edge visual check looks consistent.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
      {
        bugId: 'PITON-PERF-001',
        website,
        pageArea: 'General page loading',
        testType: 'Performance observation',
        description: 'Basic page loading observation.',
        steps: 'Open PITON website and observe whether tested pages load without obvious blocking or broken content.',
        expected: 'Pages load sufficiently for manual exploration without obvious blocking issues.',
        actual: 'No major page loading issue was observed during manual checks.',
        severity: 'N/A',
        priority: 'N/A',
        status: 'Passed',
        evidence: '',
      },
    ];
  }

  return [
    {
      bugId: `${bugPrefix}-UX-001`,
      website,
      pageArea: 'Homepage',
      testType: 'UX and usability',
      description: 'UX and homepage usability checked.',
      steps: 'Open GNS Metal homepage and review main content, visual hierarchy, readability, and navigation entry points.',
      expected: 'Homepage is understandable and usable for a first-time visitor.',
      actual: 'No major UX/usability issue was observed in the tested scope.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-NAV-001`,
      website,
      pageArea: 'Header menu',
      testType: 'Navigation flow',
      description: 'Header/menu navigation links checked.',
      steps: 'Open GNS Metal website and click header/menu navigation links.',
      expected: 'Navigation links open the expected pages or sections.',
      actual: 'Header/menu navigation links work correctly.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-FOOT-001`,
      website,
      pageArea: 'Footer',
      testType: 'Navigation flow',
      description: 'Footer links checked.',
      steps: 'Open GNS Metal website and click footer links.',
      expected: 'Footer links open the expected destinations.',
      actual: 'Footer links work correctly.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-LINK-001`,
      website,
      pageArea: 'Tested navigation scope',
      testType: 'Broken links',
      description: 'Broken link observation in tested navigation scope.',
      steps: 'Use header/menu and footer links during exploratory navigation.',
      expected: 'Checked links should navigate successfully without obvious broken-link behavior.',
      actual: 'No broken link behavior was observed in the tested navigation scope.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-FORM-001`,
      website,
      pageArea: 'Contact form',
      testType: 'Form validation',
      description: 'Contact form validation testing.',
      steps: 'Search tested scope for a standard user-fillable contact form.',
      expected: 'If a user-fillable contact form exists, required field and invalid input validation can be tested.',
      actual: 'Not Applicable because no standard user-fillable contact form was identified in the tested scope.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Not Applicable',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-UI-001`,
      website,
      pageArea: 'General pages',
      testType: 'Visual/UI consistency',
      description: 'Visual/UI consistency checked.',
      steps: 'Review tested GNS Metal pages for obvious layout, spacing, typography, and color consistency issues.',
      expected: 'No major visual consistency issue is observed.',
      actual: 'No major visual/UI consistency issue was observed in the tested scope.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-RESP-001`,
      website,
      pageArea: 'Desktop layout',
      testType: 'Responsive behavior',
      description: 'Responsive desktop check.',
      steps: 'Open GNS Metal website in desktop viewport and inspect main sections.',
      expected: 'Layout is readable and usable on desktop.',
      actual: 'Desktop responsive layout was checked and no major visual issue was observed.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-RESP-002`,
      website,
      pageArea: 'Tablet/mobile layout',
      testType: 'Responsive behavior',
      description: 'Responsive tablet/mobile check.',
      steps: 'Open GNS Metal website in tablet and mobile viewports and inspect content stacking and navigation.',
      expected: 'Layout adapts without major visual or usability issues.',
      actual: 'Tablet/mobile responsive layout was checked and no major visual issue was observed.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-COMPAT-CHROME-001`,
      website,
      pageArea: 'General visual layout',
      testType: 'Browser compatibility',
      description: 'Chrome compatibility checked.',
      steps: 'Open GNS Metal website in Chrome and review tested pages.',
      expected: 'No major Chrome-specific rendering issue is observed.',
      actual: 'Chrome visual check looks consistent.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-COMPAT-EDGE-001`,
      website,
      pageArea: 'General visual layout',
      testType: 'Browser compatibility',
      description: 'Edge compatibility checked.',
      steps: 'Open GNS Metal website in Edge and review tested pages.',
      expected: 'No major Edge-specific rendering issue is observed.',
      actual: 'Edge visual check looks consistent.',
      severity: 'N/A',
      priority: 'N/A',
      status: 'Passed',
      evidence: '',
    },
    {
      bugId: `${bugPrefix}-PERF-001`,
      website,
      pageArea: 'Ürünlerimiz page',
      testType: 'Performance / Page Loading Observation',
      description: 'The Ürünlerimiz page was observed to load slower than other checked pages during manual exploratory testing. The page contains many product images and detailed content, which may contribute to the perceived loading delay.',
      steps: '1. Open the GNS Metal website.\n2. Click the Ürünlerimiz navigation item.\n3. Compare the loading behavior with other main navigation pages.',
      expected: 'The page should load within a reasonable time and provide a smooth navigation experience.',
      actual: 'The Ürünlerimiz page loads slower compared to other checked pages.',
      severity: 'Low',
      priority: 'Low',
      status: 'Observation',
      evidence: 'docs/screenshots/gns-products-page-loading-observation.png',
    },
  ];
}

function addBrowserCompatibilitySheet(workbook) {
  const worksheet = workbook.addWorksheet('Browser Compatibility');
  worksheet.columns = [
    { header: 'Website', key: 'website', width: 22 },
    { header: 'Browser', key: 'browser', width: 18 },
    { header: 'Viewport', key: 'viewport', width: 18 },
    { header: 'Primary Pages Checked', key: 'pages', width: 35 },
    { header: 'Result', key: 'result', width: 18 },
    { header: 'Notes', key: 'notes', width: 58 },
  ];
  worksheet.addRows([
    ['PITON', 'Chrome', 'Desktop', 'Home, navigation, footer, contact form', 'Passed', 'No major browser-specific rendering difference was observed.'],
    ['PITON', 'Edge', 'Desktop', 'Home, navigation, footer, contact form', 'Passed', 'No major browser-specific rendering difference was observed.'],
    ['GNS Metal', 'Chrome', 'Desktop', 'Home and key navigation pages', 'Passed', 'No major browser-specific rendering difference was observed.'],
    ['GNS Metal', 'Edge', 'Desktop', 'Home and key navigation pages', 'Passed', 'No major browser-specific rendering difference was observed.'],
  ]);
  styleWorksheet(worksheet);
  styleStatusCells(worksheet, 'result');
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
    ['PITON', 'Desktop', 'Manual check', 'Header, hero, navigation, footer, contact form', 'Desktop view worked correctly in the tested scope.', 'Passed', ''],
    ['PITON', 'Tablet', 'Manual check', 'Navigation, content stacking, contact form', 'Tablet view was checked with no major visual issue observed.', 'Passed', ''],
    ['PITON', 'Mobile', 'Manual check', 'Menu, content stacking, contact form', 'Mobile view generally works correctly.', 'Passed', ''],
    ['PITON', 'Galaxy Z Fold 5', 'Responsive viewport', 'Homepage hero', 'A small part of hero text is clipped/not fully visible.', 'Open', 'PITON-RESP-001; docs/screenshots/piton-home-galaxy-z-fold-text-clipping.png'],
    ['GNS Metal', 'Desktop', 'Manual check', 'Header, hero, navigation, footer', 'No major visual issue observed.', 'Passed', ''],
    ['GNS Metal', 'Tablet', 'Manual check', 'Header, content layout, navigation, footer', 'No major visual issue observed.', 'Passed', ''],
    ['GNS Metal', 'Mobile', 'Manual check', 'Menu, content stacking, key pages', 'No major visual issue observed.', 'Passed', ''],
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
