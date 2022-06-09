function timeCardBuilder() {

  let ss = SpreadsheetApp.getActiveSpreadsheet(); 
  let sheet = ss.getSheetByName('Employee Information'); 
  let lastRow = sheet.getLastRow(); 

  console.log(lastRow); 

  for (let i = 0; i <= lastRow; i++) {
    let emp_name = sheet.getRange(2+i,1).getValue(); 
    let emp_id = sheet.getRange(2+i, 2).getValue(); 

    console.log(i + ": " + emp_name + " | " + emp_id); 

    let period1 = "Period beginning 6/6/2022"; 
    let period2 = "Period Beginninig 6/20/2022";
    let timecard_name = emp_name + " - " + emp_id + " - " + period1; 
    let timecard_name2 = emp_name + " - " + emp_id + " - " + period2; 

    // CREATE TIME CARD 1 FROM TEMPLATE 
    console.log("Making FIRST time card for ", emp_name); 
    let destination_folder = DriveApp.getFolderById('1mrdpUB8a-Mmxa5xTNQ8gLIojD76X7WgF'); 
    let new_timecard = DriveApp.getFileById('11FFxTXEbS-0qMYtvuE3zKzTcBc89Ft5q0T3Neaw8YwA').makeCopy(timecard_name, destination_folder).getId(); 
    console.log(new_timecard); 

    // GET NEW TIME CARD 
    let ss = SpreadsheetApp.openById(new_timecard); 
    let timecard = ss.getSheetByName("Sheet1"); 

    // EMPLOYEE HEADER INFORMATION 
    timecard.getRange("A3").setValue("Employee ID: " + emp_id); 
    timecard.getRange(4,1).setValue("Name: " + emp_name); 
    timecard.getRange(5,1).setValue("Department/Location: "); 
    timecard.getRange(6,1).setValue("Supervisor: "); 
    timecard.getRange(4,6).setValue("Week of: 6/6/2022 - 6/12/2022"); 
    timecard.getRange(5,6).setValue("Check Date: 6/24/2022"); 
    timecard.getRange("A27").setValue("Employee ID: " + emp_id); 
    timecard.getRange("A28").setValue("Name: " + emp_name); 
    timecard.getRange("A29").setValue("Department/Location: "); 
    timecard.getRange("A30").setValue("Supervisor: "); 
    timecard.getRange("F28").setValue("Week of: 6/13/2022 - 6/19/2022"); 
    timecard.getRange("F29").setValue("Check Date: 6/24/2022"); 

    // DATE INFORMATION 
    timecard.getRange(9,2).setValue("6/6/2022"); 
    timecard.getRange(10,2).setValue("6/7/2022"); 
    timecard.getRange(11,2).setValue("6/8/2022"); 
    timecard.getRange(12,2).setValue("6/9/2022"); 
    timecard.getRange(13,2).setValue("6/10/2022"); 
    timecard.getRange(14,2).setValue("6/11/2022"); 
    timecard.getRange(15,2).setValue("6/12/2022");
    timecard.getRange("B33").setValue("6/13/2022"); 
    timecard.getRange("B34").setValue("6/14/2022"); 
    timecard.getRange("B35").setValue("6/15/2022"); 
    timecard.getRange("B36").setValue("6/16/2022"); 
    timecard.getRange("B37").setValue("6/17/2022"); 
    timecard.getRange("B38").setValue("6/18/2022"); 
    timecard.getRange("B39").setValue("6/19/2022"); 

    // CREATE TIME CARD 2 FROM TEMPLATE 
    console.log("Making SECOND time card for ", emp_name); 
    let new_timecard2 = DriveApp.getFileById('11FFxTXEbS-0qMYtvuE3zKzTcBc89Ft5q0T3Neaw8YwA').makeCopy(timecard_name2, destination_folder).getId(); 
    console.log(new_timecard2); 


    // GET NEW TIME CARD 2
    let ss2 = SpreadsheetApp.openById(new_timecard2);
    let timecard2 = ss2.getSheetByName("Sheet1");

    // EMPLOYEE HEADER INFORMATION 
    timecard2.getRange("A3").setValue("Employee ID: " + emp_id); 
    timecard2.getRange(4,1).setValue("Name: " + emp_name); 
    timecard2.getRange(5,1).setValue("Department/Location: "); 
    timecard2.getRange(6,1).setValue("Supervisor: "); 
    timecard2.getRange(4,6).setValue("Week of: 6/20/2022 - 7/3/2022"); 
    timecard2.getRange(5,6).setValue("Check Date: 7/8/2022"); 
    timecard2.getRange("A27").setValue("Employee ID: " + emp_id); 
    timecard2.getRange("A28").setValue("Name: " + emp_name); 
    timecard2.getRange("A29").setValue("Department/Location: "); 
    timecard2.getRange("A30").setValue("Supervisor: "); 
    timecard2.getRange("F28").setValue("Week of: 6/20/2022 - 7/3/2022"); 
    timecard2.getRange("F29").setValue("Check Date: 7/8/2022"); 

    // DATE INFORMATION 
    timecard2.getRange(9,2).setValue("6/20/2022"); 
    timecard2.getRange(10,2).setValue("6/21/2022"); 
    timecard2.getRange(11,2).setValue("6/22/2022"); 
    timecard2.getRange(12,2).setValue("6/23/2022"); 
    timecard2.getRange(13,2).setValue("6/24/2022"); 
    timecard2.getRange(14,2).setValue("6/25/2022"); 
    timecard2.getRange(15,2).setValue("6/26/2022"); 
    timecard2.getRange("B33").setValue("6/27/2022"); 
    timecard2.getRange("B34").setValue("6/28/2022"); 
    timecard2.getRange("B35").setValue("6/29/2022"); 
    timecard2.getRange("B36").setValue("6/30/2022"); 
    timecard2.getRange("B37").setValue("7/1/2022"); 
    timecard2.getRange("B38").setValue("7/2/2022"); 
    timecard2.getRange("B39").setValue("7/3/2022");     
  }
}
