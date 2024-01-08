function onOpen() {
    var ui = SpreadsheetApp.getUi();
  
    // Create a custom menu in the spreadsheet
    ui.createMenu('Google Tasks')
      .addItem('Update Tasks', 'updateTaskDetails')
      .addItem('Migrate Tasks', 'confirmSorting')
      .addToUi();
  
    SpreadsheetApp.getActive().toast("Update your tasks using Google Tasks button above...");
    getTaskDetails();
  }
  
  // ------------------------------------------------------------------
  
  function confirmSorting() {
    // Get the UI service
    var ui = SpreadsheetApp.getUi();
  
    // Show the alert with Yes/No buttons
    var response = ui.alert("Please confirm:", "Did you sort the tasks based on Compound Score in Ascending Order?", ui.ButtonSet.YES_NO);
  
    // Check if user clicked Yes
    if (response === ui.Button.YES) {
      migrateTasks();
    } else {
      ui.alert("Please sort the tasks before continuing.");
    }
  }
  
  
  
  // //----------------------------------------------------------------
  
  
  function migrateTasks() {
  
    // Get all task lists
    const taskLists = Tasks.Tasklists.list().getItems();
  
    // Clear existing tasks
    for (const taskList of taskLists) {
      const tasks = Tasks.Tasks.list(taskList.getId()).getItems();
      for (const task of tasks) {
        Tasks.Tasks.remove(taskList.getId(), task.getId());
      }
    }
  
    // Get the active spreadsheet and sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Try to find the sheet named "main"
    var sheet = spreadsheet.getSheetByName("main");
  
    // Get all tasks from the sheet
    const tasksData = sheet.getDataRange().getValues();
  
    // Skip header row
    tasksData.shift();
  
    // Create tasks in Google Tasks and capture new IDs
    const newTaskIds = [];
    for (let i = 0; i < tasksData.length; i++) {
      const tasksDetails = tasksData[i];
      const listId = tasksDetails[0];
      const taskId = tasksDetails[1];
      const listName = tasksDetails[2];
      const taskTitle = tasksDetails[3];
  
      const newTask = Tasks.Tasks.insert({ title: taskTitle }, listId);
      console.log("Deleted the " + taskId + " and replaced with " + newTask.id);
      newTaskIds.push(newTask.id);
      // Update the corresponding row in the sheet with the new ID
      sheet.getRange(i + 2, 2).setValue(newTask.id); // Assuming old IDs are in the first column and new IDs are placed beside them.
    }
  
      // Creating an altert 
    var ui = SpreadsheetApp.getUi();
    ui.alert('Task Lists Migrated!');
  }
  
  
  
  
  // // ------------------------------------------------------------
  
  function updateTaskDetails(){
  
    // Get list of task list names in the sheet 'listNames'
    getTaskListNames();
    
    results = getTaskData();
    console.log("The Imported data from Google tasks\n");
    console.log(results);
    idValues = getCurrentIdList();
    console.log("The current data in the Google sheet");
    console.log(idValues);
    deleteOldTasks(results, idValues);
    appendNewTasks(results, idValues);
  
    // Creating an altert 
    var ui = SpreadsheetApp.getUi();
    ui.alert('Task Lists Updated!');
  }
  
  // ----------------------------------------------------------------
  
  
  function getTaskData(){
    // Get all task lists
    const taskLists = Tasks.Tasklists.list().getItems();
  
    // Create an empty array to store results
    const results = [];
  
    // Iterate through each task list
    for (const taskList of taskLists) {
      const listName = taskList.getTitle();
  
      // Get tasks within the list
      const tasks = Tasks.Tasks.list(taskList.getId()).getItems();
  
      // Iterate through each task
      for (const task of tasks) {
        results.push({
          listId: taskList.getId(), // Add task list ID
          taskId: task.id,
          listName: listName,
          title: task.getTitle(),
        });
      }
    }
  
    return results;
  }
  
  // -------------------------------------------------------
  
  function getCurrentIdList(){
    // Get the sheet 'main' and lastRow
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("main");
    var lastRow = sheet.getLastRow();
  
  
    // Get the range of the ID values column.
    const idRange = sheet.getRange('B2:B' + lastRow); // Change A:A to your actual column range
  
    // Get the values in the ID values column.
    const idValues = idRange.getValues();
  
    return idValues;
  }
  
  
  // ------------------------------------------------------------
  
  function appendNewTasks(results, idValues){
    console.log("Looking for tasks to append...");
    // Get the sheet 'main' and lastRow
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("main");
    var lastRow = sheet.getLastRow(); 
  
    // Appending New Tasks
    var count = 0;
    for (let i = 0; i < results.length; i++) {
      const row = lastRow + count;
  
      if (idValues.every(row => row[0] != results[i].taskId)){
        sheet.getRange(row + 1, 1, 1, 4).setValues([
          [results[i].listId,results[i].taskId, results[i].listName, results[i].title]
        ]);
        console.log("appended " + results[i].title);
        count++;
      }
  
    }
  }
  
  
  // ------------------------------------------------------------
  
  
  function deleteOldTasks(results, idValues) {
    console.log("Looking for tasks to delete...");
    // Get the sheet 'main' and lastRow
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("main");
  
    // Loop through the values backwards, starting from the last row.
    for (var i = 0; i < idValues.length; i++) {
      var row = idValues[i];
  
      // Loop through each result in results
      let found = false;
  
      for (const result of results) {
        if (result.taskId == row[0]) {
          found = true;
          break;
        }
      }
  
      if (found === false) {
        sheet.deleteRow(i + 2); // Add 1 to compensate for header row (if present).
        console.log("deleted task " + sheet.getRange(i+2, 4));
      }
    }
  
  }
  
  // ---------------------------------------------------------------------------
  
  function getTaskListNames() {
  
    // Get all task lists
    const taskLists = Tasks.Tasklists.list().getItems();
  
    // Get the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    // Try to find the sheet named "listNames"
    var sheet = spreadsheet.getSheetByName("listNames");
  
    // // Get the sheet to write the data to
    // var sheet = SpreadsheetApp.getActiveSheet();
  
    // Clear the sheet
    sheet.clearContents();
  
    // Sort the results array based on title
    taskLists.sort((a, b) => {
      if (a.title < b.title) {
        return -1;
      } else if (a.title > b.title) {
        return 1;
      } else {
        return 0;
      }
    });
  
    // Write the task list information to the sheet
    for (var i = 0; i < taskLists.length; i++) {
      var taskList = taskLists[i];
      sheet.appendRow([taskList.title, taskList.id]);
    }
  }