// @ts-nocheck
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Google Classroom Utilities')
      .addSubMenu(
        ui.createMenu('Import from classroom')
        .addItem('Import Google classrooms', 'importCourses')
        .addItem('Import Classroom topics', 'importTopics')
      )
      .addSubMenu(
        ui.createMenu('Post to classroom')
        .addItem('Batch create topics', 'createTopics')
        .addItem('Batch create assignments', 'createAssignments')
      )
      .addToUi();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName('Classrooms');
  const topicSheet = ss.getSheetByName('Topics');
  const assignSheet = ss.getSheetByName('Assignments')

// Check to see if sheets are created already
  if (classSheet && topicSheet && assignSheet) {
    Logger.log('All sheets are created');
  } else {
    ss.toast('Creating sheets. Please wait.', 'Status');
    if (!classSheet) {
      Logger.log("Class sheet does not exist");
      const sheet1 = ss.getSheets()[0];
      sheet1.setName('Classrooms');
      const headerArr = [['Classroom','Link','Enrollment Code','Drive Folder','Classroom ID']];
      var headerRow = ss.getSheetByName('Classrooms').getRange(1,1,1,5);
      headerRow.setValues(headerArr);

      if (headerRow.isBlank()) {
        console.log("Range is empty");
        headerRow.setValues(headerArr);
      }
    createClassNamedRange();
    }

    if (!topicSheet) {
      Logger.log("Topics sheet does not exist");
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.insertSheet('Topics', 1);
      const headerArr = [['Class','courseId','Topics','topicId','Status']];
      var headerRow = ss.getSheetByName('Topics').getRange(1,1,1,5);
      headerRow.setValues(headerArr);
      if (headerRow.isBlank()) {
        console.log("Range is empty");
        headerRow.setValues(headerArr);
      }      
    }

    if (!assignSheet) {
      Logger.log("Assignments sheet does not exist");
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.insertSheet('Assignments', 2);
      const headerArr = [['Class','courseId','Topics','topicId', 'Title', 'Points', 'Due Date (M/DD/YYYY)', 'Due Time Hour', 'Due Time Minute', 'Status']];
      var headerRow = ss.getSheetByName('Assignments').getRange(1,1,1,10);
      headerRow.setValues(headerArr);
      if (headerRow.isBlank()) {
        console.log("Range is empty");
        headerRow.setValues(headerArr);
      }

    assignSheetSetup();
    }
  
    //pop-up to confirm sheets creation
    sheetCompleteAlert();
  }


}

function createClassNamedRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName('Classrooms');

  var classAvals = classSheet.getRange("A2:A").getValues();
  var lastRow = classAvals.filter(String).length;
  const startRow = 2;
  const startCol = 1;

  var classData = classSheet.getRange(startRow,startCol,lastRow,5).activate();
  ss.setNamedRange("ClassData",classData);

}

function assignSheetSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName('Classrooms');
  const assignmentsSheet = ss.getSheetByName('Assignments');
  var classAvals = classSheet.getRange("A2:A").getValues();
  var lastRow = classAvals.filter(String).length;
  var classMenu = [];


  var classMenuRange = ss.getRangeByName("ClassData");

  for (i = 0; i<classMenuRange.getValues().length; i++){
    classMenu.push(classMenuRange.getValues()[i][0]);
  }
  
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(classMenu,true);
  assignmentsSheet.getRange("A2:A").activate().setDataValidation(rule);
  assignmentsSheet.getRange("B2").setFormula('=ArrayFormula(if(A2:A="","",VLOOKUP(A2:A,Classrooms!$B:$F,5,FALSE)))');
  assignmentsSheet.getRange("D2").setFormula('=if(A2:A="","",dget(Topics!$A$1:D,"TopicID",{"Class","courseId","Topics";A2,B2,C2}))');

  assignTopics = assignmentsSheet.getRange(3,4,lastRow,1);
  assignmentsSheet.getRange("D2").copyTo(assignTopics);

}


//Format courseIds and topicIds as plain text
function classFormatColumns() {
  const ss = SpreadsheetApp.getActive();
  const classSheet = ss.getSheetByName('Classrooms');
  classSheet.getRange("A2:E").setNumberFormat("@");
}

function assignFormatColumns() {
  const ss = SpreadsheetApp.getActive();
  const assignmentsSheet = ss.getSheetByName('Assignments');
  assignmentsSheet.getRange("B2:I").setNumberFormat("@");
}

function topicsFormatColumns() {
  const ss = SpreadsheetApp.getActive();
  const topicSheet = ss.getSheetByName('Topics');
  topicSheet.getRange("B2:D").setNumberFormat("@");
}

function sheetCompleteAlert() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Setup complete','You are ready to go!',ui.ButtonSet.OK);
}

function doGet() {
  var html = HtmlService
      .createTemplateFromFile('page')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'My custom dialog');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getData() {

}

// https://richardaanderson.org/centrally-manage-your-google-classrooms-from-a-google-sheet

/**
 * Lists 10 course names and ids.
 */
function importCourses() {
  /**  here pass pageSize Query parameter as argument to get maximum number of result
   * @see https://developers.google.com/classroom/reference/rest/v1/courses/list
   */
  const optionalArgs = {
    pageSize: 10
    // Use other parameter here if needed
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName('Classrooms');
  const startRow = 2;
  const startCol = 1;

  try {
    // call courses.list() method to list the courses in classroom
    const response = Classroom.Courses.list(optionalArgs);
    const courses = response.courses;
    if (!courses || courses.length === 0) {
      Logger.log('No courses found.');
      return;
    }
    // Print the course names and IDs of the courses
    const DATA = courses.map(c=>{
      return [c.name, c.alternateLink, c.enrollmentCode, c.teacherFolder.alternateLink,c.id]
      // Course name, course url, course enrollment code, course Drive folder URL, course id
    });
    classSheet.getRange(startRow,startCol,DATA.length,DATA[0].length).setValues(DATA);
    classSheet.getRange(startRow,startCol,DATA.length,DATA[0].length).setDataValidation(null);


  } catch (err) {
    // TODO (developer)- Handle Courses.list() exception from Classroom API
    // get errors like PERMISSION_DENIED/INVALID_ARGUMENT/NOT_FOUND
    Logger.log('Failed with error %s', err.message);
  }
  classFormatColumns(); //Format class IDs as plain text
}


function importTopics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName('Classrooms');
  const topicSheet = ss.getSheetByName('Topics');
  const startRow = 2;
  const startCol = 2;
  const classStartCol = 6;
  var Avals = classSheet.getRange("A1:A").getValues();
  var lastRow = Avals.filter(String).length;
  var Fvalues = classSheet.getRange("F2:F").getValues();
  var lastRowCourseId = Fvalues.filter(String).length;

  var classrooms = classSheet.getRange(startRow,1,lastRow, 6).getValues();
  var data = [];
  for (var i = startRow; i <= classrooms.length; i++) {
    var classNames = classSheet.getRange(i,2,1,1).getValue();
    var courseIds = classSheet.getRange(i,6,1,1).getValue().toString();
    
    // Logger.log(i);
     if (courseIds != "") {
      try {
        // call courses..topics.list() method to list the courses in classroom
        const response = Classroom.Courses.Topics.list(courseIds);
        const topics = response.topic;
        if (!topics || topics.length === 0) {
          Logger.log('No topics found.');
          return;
        }
        for (var j=0; j <= topics.length; j++){
          // Materials that are already posted without topics return a null value;
          // This shortcircuits the problem.
          if(topics[j] != null) {
            // Push info to data arr
            data.push([topics[j].courseId.toString(),topics[j].name.toString(),topics[j].topicId.toString()]);
          }

        }
      } catch (err) {
        // TODO (developer)- Handle Courses.list() exception from Classroom API
        // get errors like PERMISSION_DENIED/INVALID_ARGUMENT/NOT_FOUND
        Logger.log('Failed with error %s', err.message);
      }
    }
  }
  topicSheet.getRange(startRow,startCol,data.length,3).setValues(data);
  topicSheet.getRange(startRow,5,data.length,1).setValue("IMPORTED");
  topicsFormatColumns(); //Format B2:D as plain text
  var Dvalues = classSheet.getRange("D2:D").getValues();
  var topicsLastRow = Dvalues.filter(String).length;
  
}

//Creates Topics in the specified class.  Must run from the Google Classroom Admin/Co-Teacher account.  Super admin does not work.
//All courses must first be accepted before any topics can be created.
function createTopics() {
  var user = Session.getActiveUser();
  //change to your own Google Classroom admin account
    var ss= SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName("Topics");
    var Avals = sheet.getRange("A1:A").getValues();
    var lastRow = Avals.filter(String).length;
    // Logger.log(lastRow);
    var Dvals = sheet.getRange("D1:D").getValues();
    var lastRowStatus = Dvals.filter(String).length;
    var data = sheet.getRange(2, 1, lastRow, 4).getValues();
    // Logger.log(data);

    for (var i = lastRowStatus+1; i <= data.length; i++){
      if (sheet.getRange(i,2,1,1).getValue() != ""){ 
        var topics = Classroom.Courses.Topics.create(
          {
            "name": sheet.getRange(i,3,1,1).getValue(),
          },
          sheet.getRange(i,2,1,1).getValue(), //courseId
        );
        sheet.getRange(i,4,1,1).setValue(topics.topicId)
        sheet.getRange(i,5,1,1).setValue("CREATED");
      }
      else {
        sheet.getRange(i,3,1,1).setValue("NO TOPIC FOUND");
      }
    }
  topicsFormatColumns(); //Format B2:D as plain text
}

function createAssignments() {
  var user = Session.getActiveUser();
  var ss= SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Assignments");
  var Avals = sheet.getRange("A1:A").getValues();
  var lastRow = Avals.filter(String).length;
  var Dvals = sheet.getRange("J1:J").getValues();
  var lastRowStatus = Dvals.filter(String).length;
  var data = sheet.getRange(2, 1, lastRow, 9).getValues();

  for (var i = lastRowStatus+1; i <= data.length; i++){
    if (sheet.getRange(i,4,1,1).getValue() != "" && sheet.getRange(i,10,1,1).getValue() != "CREATED"){ 
      var courseId = sheet.getRange(i,2,1,1).getValue();
      var title = sheet.getRange(i,5,1,1).getValue();
      var topicId = sheet.getRange(i,4,1,1).getValue();
      var maxPoints = sheet.getRange(i,6,1,1).getValue();
      var dueDate = sheet.getRange(i,7,1,1).getValue().toString();
      var dueHour = sheet.getRange(i,8,1,1).getValue();
      var dueMin = sheet.getRange(i,9,1,1).getValue();

      if (dueDate != ""){
        var courseWork = {
          "title": title,
          "topicId": topicId,
          "maxPoints": maxPoints,
          "workType": "ASSIGNMENT",
          "state": "PUBLISHED",
          "dueDate": {
            "year": dueDate.split("/")[2],
            "month": dueDate.split("/")[0],
            "day": dueDate.split("/")[1]
          },
          "dueTime": {
            "hours": dueHour,
            "minutes": dueMin
          }
        };
      } else {
        var courseWork = {
          "title": title,
          "topicId": topicId,
          "maxPoints": maxPoints,
          "workType": "ASSIGNMENT",
          "state": "PUBLISHED"
        };
      }

      var assignments = Classroom.Courses.CourseWork.create(courseWork,courseId);
      sheet.getRange(i,10,1,1).setValue("CREATED");
    }
    else {
      sheet.getRange(i,10,1,1).setValue("NOT CREATED");
    }
  }
  assignFormatColumns(); //Format B2:I as plain text
}



// Historical code; not used
// function getRosters() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const classSheet = ss.getSheetByName('Classrooms');

//   const startRow = 2;
//   const startCol = 1;
//   const endRow = 10;
//   const endCol = 6;

//   const rosterStartRow = 2;
//   const rosterStartCol = 1;

//   var classRange = classSheet.getRange(startRow,startCol,endRow,endCol);

//   var sheetData = classRange.getValues();
  
//   Logger.log(classRange.getValues());

//   sheetData.forEach(function(row) {
//     var rosterSheet = ss.getSheetByName(row[1]);

//     try {
//       if (row[0] != true) {
//         Logger.log(row[1] + ' not checked.');
//         return;
//       } else {
//         Logger.log(row[1] + ' checked.');
        
//       }

//       if (row[0] && !rosterSheet) {
//         ss.insertSheet(row[1]);
//         Utilities.sleep(5000);

//         const headerArr = [['First Name','Last Name','Email','Student ID']];
//         // var headerRow = rosterSheet.getRange(1,1,1,4);

//         var headerRow = ss.getSheetByName(row[1]).getRange(1,1,1,4);
//         if (headerRow.isBlank()) {
//           Logger.log("Header is blank");
//           headerRow.setValues(headerArr);

//         }
        

//         if (rosterSheet) {        
//           Logger.log('Sheet already exists.')
//         }

        
//         Logger.log('Sheet created.');
//       }

//       // const headerArr = [['First Name','Last Name','Email','Student ID']];
//       // var headerRow = rosterSheet.getRange(1,1,1,4);
//       // headerRow.setValues(headerArr);

//       const students = Classroom.Courses.Students.list(row[5]).students;  

//       const DATA = students.map(s=>{
//         return [s.profile.name.givenName,s.profile.name.familyName,s.profile.emailAddress,'="000"&MID('+'"'+s.profile.emailAddress+'"'+',4,6)']
//       });

//       ss.getSheetByName(row[1]).getRange(rosterStartRow,rosterStartCol,DATA.length,DATA[0].length).setValues(DATA);

//     } catch (err) {
//       // TODO (developer)- Handle Courses.list() exception from Classroom API
//       // get errors like PERMISSION_DENIED/INVALID_ARGUMENT/NOT_FOUND
//       Logger.log('Failed with error %s', err.message);
//     }
//   });
// }


