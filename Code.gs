function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Salary and Credit Management System')
    .setFaviconUrl('https://www.google.com/images/icons/product/sheets-32.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

function checkLogin(username, password) {
  var sheet = getSheet('Login');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return true;
    }
  }
  return false;
}

function getEmployees() {
  var sheet = getSheet('Employees');
  var data = sheet.getDataRange().getValues();
  var employeeList = [];

  for (var i = 1; i < data.length; i++) {
    employeeList.push(data[i]);
  }
  return employeeList;
}

function addEmployee(empId, name, department, designation, contact) {
  var sheet = getSheet('Employees');
  sheet.appendRow([empId, name, department, designation, contact]);
  return "Employee added successfully";
}

function updateEmployee(empId, name, department, designation, contact) {
  var sheet = getSheet('Employees');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == empId) {
      sheet.getRange(i + 1, 2).setValue(name);
      sheet.getRange(i + 1, 3).setValue(department);
      sheet.getRange(i + 1, 4).setValue(designation);
      sheet.getRange(i + 1, 5).setValue(contact);
      return "Employee updated successfully";
    }
  }
  return "Employee not found";
}

function deleteEmployee(empId) {
  var sheet = getSheet('Employees');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == empId) {
      sheet.deleteRow(i + 1);
      return "Employee deleted successfully";
    }
  }
  return "Employee not found";
}


function getSalaryData() {
  var sheet = getSheet('Salary');
  var data = sheet.getDataRange().getValues();
  var salaryList = [];

  for (var i = 1; i < data.length; i++) {
    salaryList.push(data[i]);
  }
  return salaryList;
}

function addSalary(empId, name, basic, bonus, deduction) {
  var sheet = getSheet('Salary');

  var netSalary = Number(basic) + Number(bonus) - Number(deduction);

  sheet.appendRow([empId, name, basic, bonus, deduction, netSalary]);
  return "Salary record added. Net Salary: " + netSalary;
}


function getCreditData() {
  var sheet = getSheet('Credit');
  var data = sheet.getDataRange().getValues();
  var creditList = [];

  for (var i = 1; i < data.length; i++) {
    creditList.push(data[i]);
  }
  return creditList;
}

function addCredit(creditId, empId, name, amount) {
  var sheet = getSheet('Credit');

  var balance = Number(amount);
  var status = "Pending";

  sheet.appendRow([creditId, empId, name, amount, balance, status]);
  return "Credit added successfully";
}

function updateCreditBalance(creditId, paidAmount) {
  var sheet = getSheet('Credit');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == creditId) {
      var currentBalance = Number(data[i][4]);
      var newBalance = currentBalance - Number(paidAmount);

      if (newBalance < 0) {
        newBalance = 0;
      }

      var status = newBalance == 0 ? "Paid" : "Pending";

      sheet.getRange(i + 1, 5).setValue(newBalance);
      sheet.getRange(i + 1, 6).setValue(status);

      return "Balance updated. Remaining: " + newBalance;
    }
  }
  return "Credit record not found";
}


