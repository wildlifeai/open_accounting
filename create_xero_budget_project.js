// Configuration object to store constants
const CONFIG = {
  XERO_SHEET_ID: "1VHtZsZRzJJ29tt3SRDXnIP6ebKChoDtlHqyV5Ua-3Yg",
  ACCOUNT_COL_INDEX: 9,
  EXCLUDED_ACCOUNTS: [
    "Accounts Payable (800)",
    "Accounts Receivable (610)",
    "ANZ Term Deposit (605)",
    "Computer Equipment (720)",
    "GST (820)",
    "Historical Adjustment (840)",
    "Income Tax (830)",
    "Inventory (630)",
    "Less Accumulated Depreciation on Computer Equipment (721)",
    "Less Accumulated Depreciation on Office Equipment (711)",
    "less Provision for Doubtful Debts (611)",
    "Loan (900)",
    "Office Equipment (710)",
    "Owner A Drawings (980)",
    "Owner A Funds Introduced (970)",
    "PAYE Payable (825)",
    "Payroll Accrual (834)",
    "Prepayments (620)",
    "Retained Earnings (960)",
    "Suspense (850)",
    "Tracking Transfers (877)",
    "Unpaid Expense Claims (801)",
    "Visa Prezzy Card (622)",
    "Wages Deductions Payable (816)",
    "Wages Payable - Payroll (814)",
    "WILDLIFE.AI TRUST (600)",
    "Withholding tax paid (625)",
  ],
  MONTH_NAMES: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  REQUIRED_COLUMNS: ['*Account', 'Start', 'End', 'Amount']
};

// Date handling utility class
class DateUtil {
  static parseDate(dateValue) {
    if (dateValue instanceof Date) {
      return new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
    }

    if (typeof dateValue === 'string') {
      // Check if it's in DD/MMM/YY format
      const slashRegex = /^(\d{2})\/([A-Za-z]{3})\/(\d{2})$/;
      const slashMatch = dateValue.match(slashRegex);
      
      if (slashMatch) {
        const [, day, monthStr, yearStr] = slashMatch;
        const month = CONFIG.MONTH_NAMES.indexOf(monthStr);
        if (month === -1) return null;
        
        const year = 2000 + parseInt(yearStr);
        return new Date(year, month, parseInt(day));
      }
      
      // If not in DD/MMM/YY format, try standard date parsing
      const standardDate = new Date(dateValue);
      if (!isNaN(standardDate)) {
        return new Date(
          standardDate.getFullYear(), 
          standardDate.getMonth(), 
          standardDate.getDate()
        );
      }
    }

    throw new Error(`Unable to parse date: ${dateValue}`);
  }

  static parseStandardDate(str) {
    const date = new Date(str);
    return isNaN(date) ? null : new Date(date.getFullYear(), date.getMonth(), dateValue.getDay());
  }

  static parseSlashDate(str) {
    const parts = str.split('/');
    if (parts.length !== 3) return null;

    const day = parseInt(parts[0]);
    const month = parseInt(parts[1]) - 1;
    let year = parts[2];
    if (year.length === 2) year = parseInt(year) + 2000;
    
    return new Date(year, month, day);
  }

  static formatMonthYear(date) {
    const month = CONFIG.MONTH_NAMES[date.getMonth()];
    const year = date.getFullYear().toString().slice(-2);
    return `${month}-${year}`;
  }

  static getAllMonthsBetween(startDate, endDate) {
    const months = [];
    let currentDate = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    const lastDate = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

    while (currentDate <= lastDate) {
      months.push(new Date(currentDate));
      currentDate.setMonth(currentDate.getMonth() + 1);
    }

    return months;
  }
}

// Sheet management class
class SheetManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  cleanExistingSheets() {
    const sheets = this.spreadsheet.getSheets();
    
    // Delete all sheets except the first one
    sheets.forEach((sheet, index) => {
      if (index > 0) {
        try {
          this.spreadsheet.deleteSheet(sheet);
        } catch (e) {
          Logger.log(`Could not delete sheet ${sheet.getName()}: ${e.message}`);
        }
      }
    });

    // Clear first sheet
    if (sheets[0]) {
      sheets[0].clear();
    }
  }

  createSheet(sheetName, data, monthColumns) {
    let sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet(sheetName);
    }
    sheet.clearContents();

    const range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);

    if (monthColumns) {
      this.formatDateColumns(sheet, data[0].length);
    }

    return sheet;
  }

  formatDateColumns(sheet, numColumns) {
    if (numColumns <= 1) return;

    for (let i = 2; i <= numColumns; i++) {
      const range = sheet.getRange(1, i, 1, 1);
      range.setNumberFormat("mmm-yy");
    }
  }
}

// Budget data processor class
class BudgetProcessor {
  constructor(validAccounts) {
    this.validAccounts = validAccounts;
    this.globalStartMonth = null;
    this.globalEndMonth = null;
    this.allPossibleMonths = [];
    this.incomeTracking = {};
    this.quarterlyExpenses = {};
    this.deferredRevenue = {};
  }

  async processSecuredFolder(securedFolder) {
    await this.determineGlobalDateRange(securedFolder);
    this.allPossibleMonths = DateUtil.getAllMonthsBetween(
      this.globalStartMonth, 
      this.globalEndMonth
    );
    
    return this.processFiles(securedFolder);
  }

  async determineGlobalDateRange(folder) {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    while (files.hasNext()) {
      const file = files.next();
      const sheet = SpreadsheetApp.openById(file.getId())
                                 .getSheetByName("Budget");
      
      if (!sheet) continue;
      
      const data = sheet.getDataRange().getValues();
      const headerRow = data[0];
      
      const columns = this.getColumnIndices(headerRow);
      if (!columns.startCol || !columns.endCol) continue;

      this.updateGlobalDateRange(data, columns);
    }
  }

  getColumnIndices(headerRow) {
    return {
      accountCol: headerRow.indexOf("*Account"),
      startCol: headerRow.indexOf("Start"),
      endCol: headerRow.indexOf("End"),
      amountCol: headerRow.indexOf("Amount")
    };
  }

  updateGlobalDateRange(data, columns) {
    for (let i = 1; i < data.length; i++) {
      try {
        const startMonth = DateUtil.parseDate(data[i][columns.startCol]);
        const endMonth = DateUtil.parseDate(data[i][columns.endCol]);
        
        if (!startMonth || !endMonth) continue;

        if (!this.globalStartMonth || startMonth < this.globalStartMonth) {
          this.globalStartMonth = startMonth;
        }
        if (!this.globalEndMonth || endMonth > this.globalEndMonth) {
          this.globalEndMonth = endMonth;
        }
      } catch (error) {
        continue;
      }
    }
  }

  async processFiles(folder) {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const reshapedData = {};

    while (files.hasNext()) {
      const file = files.next();
      const fileData = await this.processFile(file);
      if (fileData) {
        // Subtract original revenue entries
        await this.adjustRevenueCategories(fileData, file);
        reshapedData[file.getName()] = fileData;
      }
    }

    return reshapedData;
  }

  async adjustRevenueCategories(fileData, file) {
    try {
      // Read the original file again
      const spreadsheet = SpreadsheetApp.openById(file.getId());
      const sheet = spreadsheet.getSheetByName("Budget");
      
      if (!sheet) {
        Logger.log(`No Budget sheet found in ${file.getName()}`);
        return;
      }

      const data = sheet.getDataRange().getValues();
      const headerRow = data[0];
      const columns = this.getColumnIndices(headerRow);

      // Process only revenue accounts from original data
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const account = row[columns.accountCol];
        
        if (!this.isRevenueAccount(account)) continue;

        const amount = parseFloat(String(row[columns.amountCol]).replace(/[^0-9.-]+/g, "")) || 0;
        if (isNaN(amount) || amount === 0) continue;

        try {
          const startMonth = DateUtil.parseDate(row[columns.startCol]);
          const endMonth = DateUtil.parseDate(row[columns.endCol]);
          
          if (!startMonth || !endMonth) continue;

          // Calculate days in the period
          const totalDays = (endMonth - startMonth) / (24 * 60 * 60 * 1000) + 1;
          
          // Get all months in the period
          const monthsInPeriod = DateUtil.getAllMonthsBetween(startMonth, endMonth);
          
          // Subtract original revenue amounts
          monthsInPeriod.forEach(month => {
            const monthStart = new Date(month.getFullYear(), month.getMonth(), 1);
            const monthEnd = new Date(month.getFullYear(), month.getMonth() + 1, 0);
            
            const daysInThisMonth = this.getOverlapDays(startMonth, endMonth, monthStart, monthEnd);
            const monthlyAmount = (amount * daysInThisMonth) / totalDays;

            if (fileData[account] && fileData[account][month]) {
              fileData[account][month] -= monthlyAmount;
              Logger.log(`Subtracted original ${account} amount ${monthlyAmount.toFixed(2)} for ${month.toDateString()}`);
            }
          });
        } catch (error) {
          Logger.log(`Error processing revenue row ${i}: ${error.message}`);
          continue;
        }
      }
    } catch (error) {
      Logger.log(`Error adjusting revenue categories: ${error.message}`);
    }
  }

  isRevenueAccount(account) {
    return account === "Grants (102)" || account === "Project Contract Income (181)";
  }  
  
  async processFile(file) {
    const fileName = file.getName();
    Logger.log(`\n=== Processing Funding Source: ${fileName} ===`);
    
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const sheet = spreadsheet.getSheetByName("Budget");
    
    if (!sheet) {
      Logger.log(`No Budget sheet found in ${fileName}`);
      return null;
    }
  
    const data = sheet.getDataRange().getValues();
    const columns = this.getColumnIndices(data[0]);
    
    if (!this.validateColumns(columns)) {
      Logger.log(`Invalid column structure in ${fileName}`);
      return null;
    }
  
    // Initialize tracking for this file
    this.incomeTracking[fileName] = [];
    this.quarterlyExpenses[fileName] = {};
    this.deferredRevenue[fileName] = {};
  
    // First pass: collect all income entries
    this.collectIncomeEntries(fileName, data, columns);
  
    // Log detailed breakdown of expenses by quarter - passing fileName as first parameter
    Logger.log(`\n=== ORIGINAL EXPENSE BREAKDOWN BEFORE PROCESSING ===`);
    this.logQuarterlyExpenseBreakdown(fileName, data, columns);
  
    // Split overheads into quarters before processing the data
    const modifiedData = this.splitOverheadsIntoQuarters(data, columns);
  
    // Log detailed breakdown after overhead splitting - passing fileName as first parameter
    Logger.log(`\n=== EXPENSE BREAKDOWN AFTER OVERHEAD SPLITTING ===`);
    this.logQuarterlyExpenseBreakdown(fileName, modifiedData, columns);
  
    // Process expenses by quarter and calculate released/deferred revenue
    this.processExpensesAndRevenue(fileName, modifiedData, columns);
  
    // Add deferred and released revenue rows to the data
    const dataWithRevenue = this.addRevenueRows(fileName, modifiedData, columns);
  
    // Log the final breakdown including revenue - passing fileName as first parameter
    Logger.log(`\n=== FINAL DATA BREAKDOWN INCLUDING REVENUE ===`);
    this.logQuarterlyExpenseBreakdown(fileName, dataWithRevenue, columns);
  
    return this.processFileData(dataWithRevenue, columns);
  }

  collectIncomeEntries(fileName, data, columns) {
    data.forEach((row, rowIndex) => {
      if (rowIndex === 0) return; // Skip header

      const account = row[columns.accountCol];
      if (account !== "Grants (102)" && account !== "Project Contract Income (181)") return;

      const amount = parseFloat(String(row[columns.amountCol]).replace(/[^0-9.-]+/g, ""));
      if (isNaN(amount) || amount === 0) return;

      this.incomeTracking[fileName].push({
        account,
        amount,
        remainingAmount: amount,
        date: DateUtil.parseDate(row[columns.startCol])
      });
    });

    // Sort income entries by date
    this.incomeTracking[fileName].sort((a, b) => a.date - b.date);
  }

  processExpensesAndRevenue(fileName, data, columns) {
    const quarters = this.getQuartersBetween(this.globalStartMonth, this.globalEndMonth);
    
    // Initialize quarterly tracking
    quarters.forEach(quarter => {
      this.quarterlyExpenses[fileName][quarter.label] = {
        total: 0,
        released: {
          "Grants (102)": 0,
          "Project Contract Income (181)": 0
        }
      };
      this.deferredRevenue[fileName][quarter.label] = {
        "Grants (102)": 0,
        "Project Contract Income (181)": 0
      };
    });
  
    // Log starting income balances
    Logger.log(`\n=== INCOME ENTRIES FOR ${fileName} ===`);
    this.incomeTracking[fileName].forEach((income, index) => {
      Logger.log(`Income Entry #${index+1}: ${income.account} - $${income.amount.toFixed(2)} on ${income.date.toDateString()}`);
    });
  
    // Process expenses quarter by quarter
    quarters.forEach(quarter => {
      Logger.log(`\n=== PROCESSING QUARTER: ${quarter.label} ===`);
      let quarterlyExpense = 0;
      const accountExpenses = {};
  
      // Calculate total expenses for the quarter
      data.forEach((row, rowIndex) => {
        if (rowIndex === 0) return; // Skip header
  
        const account = row[columns.accountCol];
        // Skip revenue and deferred revenue accounts
        if (account === "Grants (102)" || 
            account === "Project Contract Income (181)" || 
            account === "Unused Donations and Grants with Conditions (835)") {
          return;
        }
  
        const startDate = DateUtil.parseDate(row[columns.startCol]);
        const endDate = DateUtil.parseDate(row[columns.endCol]);
        const amount = parseFloat(String(row[columns.amountCol]).replace(/[^0-9.-]+/g, ""));
        
        if (isNaN(amount)) return;
  
        const overlap = this.getOverlapDays(startDate, endDate, quarter.start, quarter.end);
        if (overlap > 0) {
          const days = (endDate - startDate) / (24 * 60 * 60 * 1000) + 1;
          const expenseForThisQuarter = (amount / days) * overlap;
          quarterlyExpense += expenseForThisQuarter;
          
          // Track by account for detailed logging
          if (!accountExpenses[account]) accountExpenses[account] = 0;
          accountExpenses[account] += expenseForThisQuarter;
        }
      });
  
      this.quarterlyExpenses[fileName][quarter.label].total = quarterlyExpense;
  
      // Log expenses by account
      Logger.log(`Expenses for ${quarter.label}:`);
      Object.entries(accountExpenses)
        .sort((a, b) => b[1] - a[1]) // Sort by amount (highest first)
        .forEach(([account, amount]) => {
          const percentage = (amount / quarterlyExpense) * 100;
          Logger.log(`  ${account}: $${amount.toFixed(2)} (${percentage.toFixed(1)}%)`);
        });
      Logger.log(`  TOTAL: $${quarterlyExpense.toFixed(2)}`);
  
      // Process the expenses against available income
      if (quarterlyExpense > 0) {
        let remainingExpense = quarterlyExpense;
        
        Logger.log(`\nAllocating expenses to income sources:`);
        
        // Try to cover expenses with available income
        this.incomeTracking[fileName].forEach((income, index) => {
          if (remainingExpense <= 0 || income.remainingAmount <= 0) return;
          
          // Only use income that's available before or during this quarter
          if (income.date > quarter.end) {
            Logger.log(`  Income #${index+1} (${income.account}): Not available yet for this quarter`);
            return;
          }
  
          const amountBefore = income.remainingAmount;
          const amountToUse = Math.min(remainingExpense, income.remainingAmount);
          income.remainingAmount -= amountToUse;
          remainingExpense -= amountToUse;
          
          // Record released revenue
          this.quarterlyExpenses[fileName][quarter.label].released[income.account] += amountToUse;
          
          Logger.log(`  Income #${index+1} (${income.account}): $${amountBefore.toFixed(2)} → Used $${amountToUse.toFixed(2)} → Remaining $${income.remainingAmount.toFixed(2)}`);
        });
        
        if (remainingExpense > 0) {
          Logger.log(`  WARNING: Uncovered expenses: $${remainingExpense.toFixed(2)}`);
        }
      }
  
      // Calculate deferred revenue for this quarter
      Logger.log(`\nCalculating deferred revenue for ${quarter.label}:`);
      this.incomeTracking[fileName].forEach((income, index) => {
        if (income.date <= quarter.end) {
          this.deferredRevenue[fileName][quarter.label][income.account] += income.remainingAmount;
          Logger.log(`  Income #${index+1} (${income.account}): Deferring $${income.remainingAmount.toFixed(2)}`);
        } else {
          Logger.log(`  Income #${index+1} (${income.account}): Not received yet in this quarter`);
        }
      });
  
      // Summary for the quarter
      Logger.log(`\nQUARTER SUMMARY - ${fileName} - ${quarter.label}:`);
      Logger.log(`  Total Expenses: $${this.quarterlyExpenses[fileName][quarter.label].total.toFixed(2)}`);
      Logger.log(`  Released Revenue:`)
      Logger.log(`    Grants (102): $${this.quarterlyExpenses[fileName][quarter.label].released["Grants (102)"].toFixed(2)}`);
      Logger.log(`    Project Contract Income (181): $${this.quarterlyExpenses[fileName][quarter.label].released["Project Contract Income (181)"].toFixed(2)}`);
      Logger.log(`  Deferred Revenue:`)
      Logger.log(`    Grants (102): $${this.deferredRevenue[fileName][quarter.label]["Grants (102)"].toFixed(2)}`);
      Logger.log(`    Project Contract Income (181): $${this.deferredRevenue[fileName][quarter.label]["Project Contract Income (181)"].toFixed(2)}`);
    });
  }

  addRevenueRows(fileName, data, columns) {
    const newData = [...data];
    const quarters = this.getQuartersBetween(this.globalStartMonth, this.globalEndMonth);
    
    quarters.forEach(quarter => {
      const quarterDate = quarter.end; // Use the end date directly from the quarter object
      
      // Add released revenue rows
      ["Grants (102)", "Project Contract Income (181)"].forEach(account => {
        const releasedAmount = this.quarterlyExpenses[fileName][quarter.label].released[account];
        if (releasedAmount > 0) {
          const releasedRow = Array(data[0].length).fill('');
          releasedRow[columns.accountCol] = account;
          releasedRow[columns.startCol] = quarterDate;
          releasedRow[columns.endCol] = quarterDate;
          releasedRow[columns.amountCol] = releasedAmount;
          newData.push(releasedRow);
          
          Logger.log(`Added released revenue row for ${quarter.label}: ${account} = $${releasedAmount.toFixed(2)}`);
        }
  
        // Add deferred revenue row
        const deferredAmount = this.deferredRevenue[fileName][quarter.label][account];
        if (deferredAmount > 0) {
          const deferredRow = Array(data[0].length).fill('');
          deferredRow[columns.accountCol] = "Unused Donations and Grants with Conditions (835)";
          deferredRow[columns.startCol] = quarterDate;
          deferredRow[columns.endCol] = quarterDate;
          deferredRow[columns.amountCol] = deferredAmount;
          newData.push(deferredRow);
          
          Logger.log(`Added deferred revenue row for ${quarter.label}: ${account} = $${deferredAmount.toFixed(2)}`);
        }
      });
    });
    
    return newData;
  }

  getQuarterEndDate(quarterLabel) {
    const [year, quarter] = quarterLabel.split(' ');
    const month = (parseInt(quarter.slice(1)) * 3) - 1;
    return new Date(parseInt(year), month, new Date(parseInt(year), month + 1, 0).getDate());
  }

  splitOverheadsIntoQuarters(data, columns) {
    // Create a copy of the data array
    const newData = [...data];
    
    // Find the overhead row
    const overheadRowIndex = newData.findIndex(row => 
      row[columns.accountCol] === "Overhead Allocation (500)"
    );
    
    if (overheadRowIndex === -1) {
      Logger.log("No overhead row found");
      return newData;
    }

    // Extract the overhead row
    const overheadRow = newData[overheadRowIndex];
    const startDate = DateUtil.parseDate(overheadRow[columns.startCol]);
    const endDate = DateUtil.parseDate(overheadRow[columns.endCol]);
    const totalAmount = parseFloat(String(overheadRow[columns.amountCol]).replace(/[^0-9.-]+/g, ""));

    Logger.log(`Processing overhead split from ${startDate} to ${endDate} for amount ${totalAmount}`);

    // Get quarters between the dates
    const quarters = this.getQuartersBetween(startDate, endDate);
    
    // Calculate expenses per quarter (excluding income and overheads)
    const quarterlyExpenses = this.calculateQuarterlyExpenses(newData, quarters, columns);
    
    // Calculate total expenses across all quarters
    const totalExpenses = Object.values(quarterlyExpenses).reduce((a, b) => a + b, 0);
    
    // Create new rows for each quarter with weighted amounts
    const quarterlyRows = quarters.map(quarter => {
      const row = [...overheadRow]; // Clone the original row
      
      // Calculate weighted amount based on quarter's proportion of total expenses
      const quarterExpense = quarterlyExpenses[quarter.label] || 0;
      const weightedAmount = totalExpenses > 0 
        ? (quarterExpense / totalExpenses) * totalAmount 
        : totalAmount / quarters.length; // Fallback to even distribution if no expenses
      
      // Set the dates for this quarter
      row[columns.startCol] = quarter.end; // Last day of the quarter
      row[columns.endCol] = quarter.end; // Last day of the quarter
        
      row[columns.amountCol] = weightedAmount;
      row[columns.accountCol] = `Overhead Allocation (500)`;
      
      Logger.log(`Quarter ${quarter.label}: Expenses $${quarterExpense.toFixed(2)}, ` +
                `Allocation $${weightedAmount.toFixed(2)} (${((quarterExpense / totalExpenses) * 100).toFixed(1)}%)`);
      
      return row;
    });

    // Remove the original overhead row and add the new quarterly rows
    newData.splice(overheadRowIndex, 1, ...quarterlyRows);

    return newData;
  }

  calculateQuarterlyExpenses(data, quarters, columns) {
    const quarterlyExpenses = {};
    quarters.forEach(quarter => {
      quarterlyExpenses[quarter.label] = 0;
    });

    // Process each row
    data.forEach((row, rowIndex) => {
      if (rowIndex === 0 || 
          row[columns.accountCol] === "Overhead Allocation (500)" ||
          row[columns.accountCol] === "Grants (102)" ||
          row[columns.accountCol] === "Project Contract Income (181)" ||
          row[columns.accountCol] === "Unused Donations and Grants with Conditions (835)") {
        return;
      }

      try {
        const startDate = DateUtil.parseDate(row[columns.startCol]);
        const endDate = DateUtil.parseDate(row[columns.endCol]);
        const amount = parseFloat(String(row[columns.amountCol]).replace(/[^0-9.-]+/g, ""));
        
        if (isNaN(amount)) return;

        // Calculate total days in the expense period
        const totalDays = (endDate - startDate) / (24 * 60 * 60 * 1000) + 1;
        
        quarters.forEach(quarter => {
          // Calculate days that fall within this quarter
          const quarterStart = quarter.start;
          const quarterEnd = quarter.end;
          
          const overlapDays = this.getOverlapDays(startDate, endDate, quarterStart, quarterEnd);
          
          if (overlapDays > 0) {
            // Calculate the proportion of expense that belongs to this quarter
            const quarterAmount = (amount * overlapDays) / totalDays;
            quarterlyExpenses[quarter.label] += quarterAmount;
            
            Logger.log(`Expense distribution for ${row[columns.accountCol]}:`);
            Logger.log(`  Period: ${startDate.toDateString()} to ${endDate.toDateString()}`);
            Logger.log(`  Quarter: ${quarter.label} (${quarterStart.toDateString()} to ${quarterEnd.toDateString()})`);
            Logger.log(`  Total Amount: $${amount.toFixed(2)}`);
            Logger.log(`  Days in Period: ${totalDays}`);
            Logger.log(`  Days in Quarter: ${overlapDays}`);
            Logger.log(`  Allocated Amount: $${quarterAmount.toFixed(2)}`);
          }
        });
      } catch (error) {
        Logger.log(`Error processing row ${rowIndex}: ${error.message}`);
      }
    });

    return quarterlyExpenses;
  }

  getOverlapDays(startA, endA, startB, endB) {
    const start = Math.max(startA.getTime(), startB.getTime());
    const end = Math.min(endA.getTime(), endB.getTime());
    return Math.max(0, (end - start) / (24 * 60 * 60 * 1000) + 1);
  }

  getQuartersBetween(startDate, endDate) {
    Logger.log(`Calculating quarters between ${startDate} and ${endDate}`);
    
    const quarters = [];
    const startYear = startDate.getFullYear();
    const endYear = endDate.getFullYear();
    
    for (let year = startYear; year <= endYear; year++) {
      for (let q = 1; q <= 4; q++) {
        const quarterStart = new Date(year, (q - 1) * 3, 1);
        const quarterEnd = new Date(year, q * 3, 0); // Last day of the quarter
        
        if (quarterStart > endDate) break;
        if (quarterEnd >= startDate) {
          quarters.push({
            label: `${year} Q${q}`,
            start: quarterStart,
            end: quarterEnd,
            index: quarters.length // Add index for easier next quarter reference
          });
        }
      }
    }
    
    Logger.log(`Found ${quarters.length} quarters`);
    return quarters;
  }

  validateColumns(columns) {
    return !Object.values(columns).includes(-1);
  }

  processFileData(data, columns) {
    const fileCategories = this.initializeCategories();
  
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const account = row[columns.accountCol];
      const amount = parseFloat(String(row[columns.amountCol]).replace(/[^0-9.-]+/g, "")) || 0;
  
      if (!this.validAccounts.includes(account) || isNaN(amount) || amount === 0) {
        continue;
      }

      try {
        const startMonth = DateUtil.parseDate(row[columns.startCol]);
        const endMonth = DateUtil.parseDate(row[columns.endCol]);
        
        if (!startMonth || !endMonth) {
          Logger.log(`Invalid date formats for row ${i}`);
          continue;
        }
  
        const totalDays = (endMonth - startMonth) / (24 * 60 * 60 * 1000) + 1;
        const monthsInPeriod = DateUtil.getAllMonthsBetween(startMonth, endMonth);
        
        monthsInPeriod.forEach(month => {
          const monthStart = new Date(month.getFullYear(), month.getMonth(), 1);
          const monthEnd = new Date(month.getFullYear(), month.getMonth() + 1, 0);
          
          const daysInThisMonth = this.getOverlapDays(startMonth, endMonth, monthStart, monthEnd);
          const monthlyAmount = (amount * daysInThisMonth) / totalDays;
          
          if (!fileCategories[account]) {
            fileCategories[account] = {};
          }
          
          if (!fileCategories[account][month]) {
            fileCategories[account][month] = 0;
          }
          fileCategories[account][month] += monthlyAmount;
          
          Logger.log(`Monthly distribution for ${account}:`);
          Logger.log(`  Month: ${month.toDateString()}`);
          Logger.log(`  Days in month: ${daysInThisMonth}`);
          Logger.log(`  Allocated amount: $${monthlyAmount.toFixed(2)}`);
        });
      } catch (error) {
        Logger.log(`Error processing row ${i}: ${error.message}`);
        continue;
      }
    }    
  
    return fileCategories;
  }
  
  initializeCategories() {
    const categories = {};
    this.validAccounts.forEach(account => {
      categories[account] = {};
      this.allPossibleMonths.forEach(month => {
        categories[account][month] = 0;
      });
    });
    return categories;
  }

  distributeAmount(categories, account, months, monthlyAmount) {
    months.forEach(month => {
      categories[account][month] += monthlyAmount;
    });
  }

  generateSheetData(categories) {
    const header = ["*Account", ...this.allPossibleMonths];
    const rows = [header];

    this.validAccounts.forEach(account => {
      const row = [account];
      this.allPossibleMonths.forEach(month => {
        const value = categories[account][month] || 0;
        row.push(Number(value.toFixed(2)));
      });
      rows.push(row);
    });

    return rows;
  }
  
  logQuarterlyExpenseBreakdown(fileName, data, columns) {
    const quarters = this.getQuartersBetween(this.globalStartMonth, this.globalEndMonth);
    
    // Initialize tracking structure for each account in each quarter
    const quarterlyExpenseBreakdown = {};
    quarters.forEach(quarter => {
      quarterlyExpenseBreakdown[quarter.label] = {
        byAccount: {},
        total: 0
      };
    });
    
    // Process each row to categorize expenses by account and quarter
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const account = row[columns.accountCol];
      
      // Skip revenue accounts
      if (account === "Grants (102)" || 
          account === "Project Contract Income (181)" || 
          account === "Unused Donations and Grants with Conditions (835)") {
        continue;
      }
      
      try {
        const startDate = DateUtil.parseDate(row[columns.startCol]);
        const endDate = DateUtil.parseDate(row[columns.endCol]);
        const amount = parseFloat(String(row[columns.amountCol]).replace(/[^0-9.-]+/g, ""));
        
        if (isNaN(amount) || amount === 0) continue;
        
        // Calculate daily amount
        const days = (endDate - startDate) / (24 * 60 * 60 * 1000) + 1;
        const dailyAmount = amount / days;
        
        // Distribute to quarters
        quarters.forEach(quarter => {
          const overlap = this.getOverlapDays(startDate, endDate, quarter.start, quarter.end);
          if (overlap > 0) {
            const quarterAmount = dailyAmount * overlap;
            
            // Initialize account if not exists
            if (!quarterlyExpenseBreakdown[quarter.label].byAccount[account]) {
              quarterlyExpenseBreakdown[quarter.label].byAccount[account] = 0;
            }
            
            // Add to account total
            quarterlyExpenseBreakdown[quarter.label].byAccount[account] += quarterAmount;
            quarterlyExpenseBreakdown[quarter.label].total += quarterAmount;
          }
        });
      } catch (error) {
        Logger.log(`Error processing row ${i}: ${error.message}`);
        continue;
      }
    }
    
    // Log the breakdown for each quarter
    quarters.forEach(quarter => {
      Logger.log(`\n=== DETAILED EXPENSE BREAKDOWN FOR ${quarter.label} ===`);
      Logger.log(`Quarter period: ${quarter.start.toDateString()} to ${quarter.end.toDateString()}`);
      Logger.log(`TOTAL EXPENSES: $${quarterlyExpenseBreakdown[quarter.label].total.toFixed(2)}`);
      
      // Sort accounts by expense amount (highest first)
      const sortedAccounts = Object.entries(quarterlyExpenseBreakdown[quarter.label].byAccount)
        .sort((a, b) => b[1] - a[1]);
      
      if (sortedAccounts.length > 0) {
        Logger.log('\nBREAKDOWN BY ACCOUNT:');
        sortedAccounts.forEach(([account, amount]) => {
          const percentage = (amount / quarterlyExpenseBreakdown[quarter.label].total) * 100;
          Logger.log(`  ${account}: $${amount.toFixed(2)} (${percentage.toFixed(1)}%)`);
        });
      } else {
        Logger.log('No expenses found in this quarter.');
      }
      
      // Check if the quarterly expenses for this file and quarter exist
      if (this.quarterlyExpenses[fileName] && this.quarterlyExpenses[fileName][quarter.label]) {
        const processedTotal = this.quarterlyExpenses[fileName][quarter.label].total || 0;
        const difference = Math.abs(processedTotal - quarterlyExpenseBreakdown[quarter.label].total);
        
        Logger.log(`\nRECONCILIATION CHECK:`);
        Logger.log(`  Calculated total: $${quarterlyExpenseBreakdown[quarter.label].total.toFixed(2)}`);
        Logger.log(`  Processed total : $${processedTotal.toFixed(2)}`);
        Logger.log(`  Difference      : $${difference.toFixed(2)} (${difference > 0.01 ? '❌ MISMATCH' : '✓ MATCH'})`);
      } else {
        Logger.log(`\nRECONCILIATION CHECK: Cannot perform - quarterly expenses not yet initialized for ${quarter.label}`);
      }
    });
    
    return quarterlyExpenseBreakdown;
  }
}

// Main function
async function create_xero_budget_project() {
  try {
    Logger.log("Starting process...");
    
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetManager = new SheetManager(activeSpreadsheet);
    
    // Clean existing sheets
    sheetManager.cleanExistingSheets();

    // Get valid accounts
    const xeroSheet = SpreadsheetApp.openById(CONFIG.XERO_SHEET_ID).getActiveSheet();
    const xeroData = xeroSheet.getDataRange().getValues();
    const validAccounts = xeroData
      .map(row => row[CONFIG.ACCOUNT_COL_INDEX])
      .filter(account => account && 
              account !== "*Account" && 
              !CONFIG.EXCLUDED_ACCOUNTS.includes(account));

    // Process secured folder
    const parentFolder = DriveApp.getFileById(activeSpreadsheet.getId())
                                .getParents().next();
    const securedFolder = parentFolder.getFoldersByName("secured").next();
    
    const budgetProcessor = new BudgetProcessor(validAccounts);
    const reshapedData = await budgetProcessor.processSecuredFolder(securedFolder);

    // Create individual sheets
    Object.entries(reshapedData).forEach(([fileName, fileData]) => {
      const sheetData = budgetProcessor.generateSheetData(fileData);
      sheetManager.createSheet(fileName, sheetData, true);
    });

    // Create combined sheet
    const combinedCategories = budgetProcessor.initializeCategories();
    Object.values(reshapedData).forEach(fileData => {
      Object.entries(fileData).forEach(([account, monthData]) => {
        Object.entries(monthData).forEach(([month, value]) => {
          combinedCategories[account][month] += value;
        });
      });
    });

    const combinedData = budgetProcessor.generateSheetData(combinedCategories);
    sheetManager.createSheet(parentFolder.getName(), combinedData, true);

    Logger.log("Project and funding source specific budgets created successfully");
  } catch (error) {
    Logger.log(`Error in budget project creation: ${error.message}`);
    throw error;
  }
}