function main() {

    String.prototype.addQuery = function (obj) {
      return this + "?" + Object.entries(obj).flatMap(([k, v]) => Array.isArray(v) ? v.map(e => `${k}=${encodeURIComponent(e)}`) : `${k}=${encodeURIComponent(v)}`).join("&");
    }
    // https://gist.github.com/tanaikech/70503e0ea6998083fcb05c6d2a857107
    // Snippet used to add query parameters to queries.

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //Grab current sheet 

    var ranges = sheet.getRangeList(["B2:M3", "B3"]).getRanges(); // Grab the following ranges in the sheet. 
    rangeValues = ranges[0].getValues();
    const courseNum= rangeValues[0][1];
    const accesstoken = rangeValues[0][11];
    const completionProgress = ranges[1];

    headers = {
      'Authorization' : 'Bearer '+ accesstoken
    }

    let counter = 0;
    // Can ignore updateCounter as that function is used to update the progress. 
    const updateCounter = (counter) =>  completionProgress.setValue(`${Math.round((counter / 8) * 100 )}%`); // Counter to update progres of script in "B3"

    let studentData = updateCourseNameAndNumberOfStudents(courseNum, sheet, headers); updateCounter(counter++);

    let arrayOfReq = updateModuleAndModuleItems(courseNum, headers, sheet); updateCounter(counter++);
    let arrayOfModuleData = arrayOfReq[0];
    const checkReq = arrayOfReq[1];

    let rowAfterStudents = updateStudentNames(sheet, studentData) + 1; updateCounter(counter++); // Is the row where "Completion Statistics" is written 
    setUpCompletionStats(rowAfterStudents, sheet); updateCounter(counter++);
    sheet.getRange(rowAfterStudents + 1, 3, 1, 1).setValue(studentData.length);
    const completionArray = writeModuleCompletionInfo(arrayOfModuleData, studentData, sheet, courseNum, checkReq); updateCounter(counter++);
    const completionRange = completionArray[1];
    const completionData = completionArray[0];
    const completionCol = colorCompletionBackgrounds(completionData, completionRange, sheet); updateCounter(counter++);
    updateCompletionPercentages(completionData, sheet, completionCol, rowAfterStudents, studentData.length); updateCounter(counter++);
    addPieChartsToSheet(sheet, rowAfterStudents); updateCounter(counter++);
    addProgressUpdatedTime(sheet); updateCounter(counter++);


}


function setUpCompletionStats(rowAfterStudents, sheet){
    // Set up completion statistics with the percentages will eventually be used by the pie charts.

    // Hardcoded values for row
    const completionValues = [
    ["Student count", "", "0", ""],
    ["", "", "", ""],
    ["Completion rate", "" ,"Students", "Percent"],
    ["0%", "", 0, "0%"],
    [">0% and <= 10%", "", 0, "0%"],
    [">10% and <= 20%", "", 0, "0%"],
    [">20% and <= 30%", "", 0, "0%"],
    [">30% and <= 40%", "", 0, "0%"],
    [">40% and <= 50%", "", 0, "0%"],
    [">50% and <= 60%", "", 0, "0%"],
    [">60% and <= 70%", "", 0, "0%"],
    [">70% and <= 80%", "", 0, "0%"],
    [">80% and <= 90%", "", 0, "0%"],
    [">90% and <= 100%", "", 0, "0%"],
    ["no data", "", 0, "0%"],
  ];
    
    try{
      sheet.getRange(rowAfterStudents + 1, 1, completionValues.length, completionValues[0].length).setValues(completionValues).setHorizontalAlignment("left");
      sheet.getRange(`A${rowAfterStudents}`).setValue("Completion statistics").setHorizontalAlignment("left").setFontSize(15);
    }
    catch(e){
      console.error(`An error has occured while writing the harded coded values of completion statistics: ${e}`);
    }
  
    // Calculate A1 notation for bolding student count, number of students, Completion rate, Students, and Percent
    const boldCompletionRangeList = [
      `A${rowAfterStudents}`, `A${rowAfterStudents + 1}`, `C${rowAfterStudents + 1}`, `A${rowAfterStudents + 3}`, `C${rowAfterStudents + 3}`,
      `D${rowAfterStudents + 3}`
    ]

    for (let i = rowAfterStudents + 4 ; i < completionValues.length + rowAfterStudents; i++){
      boldCompletionRangeList.push(`A${i}`)
    }

    try{
      sheet.getRangeList(boldCompletionRangeList).setFontWeight("bold");
      sheet.getRange(rowAfterStudents + 1, 1, completionValues.length, 2).mergeAcross();
    }
    catch(e){
      console.error(`An error ouccred while bolding completion statistics and merging rows: ${e}`);
    }
}

function updateCourseNameAndNumberOfStudents(courseNum, sheet, headers){
    // Grabs course name and number of student and writes it to sheet. Also returns data of all students such as email, name, etc.
    // Note that while it is a bit more efficient to use getRangeList() for this function and combine getting ranges for "A1" and "B5", 
    // it is more clear to the user to seperate the actions in case of changes needed for the future. 

    // Grab name of course
    try{
      let url = 'https://canvas.eee.uci.edu/api/v1/courses/' + courseNum  + "?per_page=100"; 
      let courseData = JSON.parse(UrlFetchApp.fetch(url, {headers:headers}).getContentText());
      let courseName = courseData.name;
      let courseNameCell = sheet.getRange("A1");
      courseString = `Progress report of students in modules in ${courseName}`
      courseNameCell.setValue(courseString).setFontWeight("bold").setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    }
    catch (err){
      console.error(`An error has occurred fetching the name of the course from Canvas: ${err}`);
    }

    // Grabs students
    try{
      let numOfStudentsURL = `https://canvas.eee.uci.edu/api/v1/courses/${courseNum}/users`; 
      const query = {
        per_page: 100,
        enrollment_type: ["student"]
      };
      const endpoint = numOfStudentsURL.addQuery(query);
      let studentData = JSON.parse(UrlFetchApp.fetch(endpoint, {headers:headers}).getContentText());
      let studentSize = studentData.length;

      let countStudentsCell = sheet.getRange("B5");

      countStudentsCell.setValue(`⬇️ ${studentSize} Students`);
      return studentData; // Return response.
    }
    catch (error){
      console.error(`An error has occurred with fetching the student data from Canvas: ${error}`)
    }

}


function updateModuleAndModuleItems(courseNum, headers, sheet){
    // Grabs all the modules and all the data for all the module items and returns them

    // https://canvas.instructure.com/doc/api/modules.html#method.context_modules_api.index
    let moduleData;
    try{
      let listModulesURL = `https://canvas.eee.uci.edu/api/v1/courses/${courseNum}/modules` + "?per_page=100"; 
      moduleData = JSON.parse(UrlFetchApp.fetch(listModulesURL, {headers:headers}).getContentText());
    }
    catch(error){
      console.error(`Error occurred with fetching module data: ${error}`);
    }

    const moduleItemData = [];
    let data = [[], []]
    const checkRequirements = new Map(); // Checks if each module has requirements enabled

    // Grab all the data for each module and module item such as name and id. 
    // Takes a lengthy amount of time but only solution was calling requests to each module and iterating over a list Module items for each module.
    // https://canvas.instructure.com/doc/api/modules.html#method.context_module_items_api.index
    try{
      for (module of moduleData){
        const listModuleItemsURL = `https://canvas.eee.uci.edu/api/v1/courses/${courseNum}/modules/${module.id}/items` + "?per_page=100"; 

        let moduleItemsData = JSON.parse(UrlFetchApp.fetch(listModuleItemsURL, {headers:headers}).getContentText());

        if (moduleItemsData.length == 0){ continue;} /// If module is empty, skip over it and go to next iteration
        else{
          if (moduleItemsData[0].completion_requirement == undefined){ // Check if the first module item has requirements. 
            checkRequirements.set(moduleItemsData[0].module_id, false)
          }
          else{
            checkRequirements.set(moduleItemsData[0].module_id, true)
          }
          for (moduleItem of moduleItemsData){
            data[0].push(module.name);
            data[1].push(moduleItem.title);
            moduleItemData.push([moduleItem.module_id, moduleItem.id]);
          }
        }
      }
    }
    catch (err){
      console.error(`Error has occurred while fetching a module item: ${err}`);
    }

    // Write to the active sheet
    try{
      sheet.getRange(4, 5, 2, data[0].length).setValues(data).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }
    catch(err){
      console.error(`Error has occurred with writing to active sheet: ${err}`)
    }

    // Write to the clearData sheet. Note that this is a hidden sheet. 
    try{
      let moduleRange = `[4, 5, 2, ${data[0].length}]`;
      const clearDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ClearData');
      clearDataSheet.getRange(2, 1).setValue(moduleRange);
    }
    catch(err){
      console.error(`Error has occurred with writing to clearData sheet. Error is: ${err}`)
    }

    return [moduleItemData, checkRequirements];
}

function updateStudentNames(sheet, studentData){
    // Writes student information to the sheet. 

    // Grab value of the dropdown
    let nameSelection;
    try{
      let nameDropdownCell = sheet.getRange("A6:B6");
      nameSelection = nameDropdownCell.getValue();
    }
    catch(err){
      console.error(`An error has occurred while fetching the value of the dropdown list in the sheet on A6:B6: ${err}`);
    }

    const updateStudentData = []
    // Checks what option is selected. Normal name or sortable name.
    // Grabs each student info (name, email) into an array and adds to 2D array. 
    if (nameSelection == "Sortable name"){
      for (student of studentData){
        updateStudentData.push([student.sortable_name, "",  student.email]);
      }
    }
    else{
      for (student of studentData){
        updateStudentData.push([student.name, "", student.email]);
      }
    }

    // https://stackoverflow.com/questions/33935971/how-can-i-clip-text-in-in-cell-in-spreadsheet-with-google-apps-script
    // Updates values and merges rows to fit student name.
    try{
      sheet.getRange(7, 1, updateStudentData.length, 3).setValues(updateStudentData).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      sheet.getRange(7, 1, updateStudentData.length, 2).mergeAcross();
    }
    catch(err){
      console.error(`An error has occurred with writing the student data to the sheet: ${err}`);
    }

    // Writing to the clearData sheet so we know which rows have to be cleared.
    try{
      const clearDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ClearData');
      const val = clearDataSheet.getRange("A2").getValue();
      let array = eval(val);
      let studentRange = `[7, 1, ${updateStudentData.length + 17}, ${array[3] + 4}]`;
      clearDataSheet.getRange(2, 2).setValue(studentRange);
    }
    catch(e){
      console.error(`An error has occurred with writing to the clearData sheet: ${e}`);
    }
    finally{
      return updateStudentData.length + 7;
    }
}

function writeModuleCompletionInfo(arrayOfModuleData, studentData, sheet, courseNum, checkReq){
    // This part takes the longest runtime for the entire script. 
    // For each student, find each module and check if every module item in a module is completed by that student. 
    // Runtime = Number of students * Number of modules * Number of module items in each module
    // At this time, there is no other way other than making API calls for each student, module, and module item to track student progress that I could find. 

    // Map each module to the module id of it's module items
    // https://canvas.instructure.com/doc/api/modules.html#method.context_module_items_api.show
    const moduleMap = new Map();
    try{
      for (let module of arrayOfModuleData){
        let moduleId = module[0];
        let moduleItemId = module[1];
        if (moduleMap.has(moduleId)){
          moduleMap.get(moduleId).push(moduleItemId);
        }
        else{
          const arr = [moduleItemId];
          moduleMap.set(moduleId, arr);
        }
      }
    }
    catch(e){
      console.error(`An error occured while mapping module id to each module item id: ${e}`);
    }

    // Hard coded values of where to start wrtiting "locked" or "completed"
    let startCompletionRow = 7;
    let startCompletionCol = 5;
    const completionData = [];
    // Gran info about each module item and store in 2D array.
    for (const student of studentData){
      const tempStudentStorage = [];
      for (const [moduleId, listOfModuleItems] of moduleMap) {
        let studentData;
        if (checkReq.get(moduleId)){ // If the module has requirements, call the API
              try{
                // API call to check if student completed every module. 
                let moduleCompletionURL = `https://canvas.eee.uci.edu/api/v1/courses/${courseNum}/modules/${moduleId}`; 
                const query = {
                  per_page: 100,
                  student_id: student.id
                };
                const endpoint = moduleCompletionURL.addQuery(query);
                studentData = JSON.parse(UrlFetchApp.fetch(endpoint, {headers:headers}).getContentText());
              }
              catch (e){
                console.error(`Error occurred while fetching data to check if module was completed: ${e}`);
              }
        }
        
        
        // If module is completed, skip going through each module item to save time and just write completed for all it's module items.
        if (checkReq.get(moduleId) == false || studentData.state == "completed" ){ // If module is completed by student or doesn't have requirements, set it as complete for each student. 
          // Multiply "completed" strings by number of module items
          for (let i = 0; i < listOfModuleItems.length; i++){
            tempStudentStorage.push("completed");
          }
        }
        // If module not complete, go through each module item to see the state based on each student.
        else{
          for (const moduleItem of listOfModuleItems){
            let moduleCompleteInfo;
            try{
              // API call for each student, for each module, and for each module item to see if student completed.
              let moduleItemCompletionURL = `https://canvas.eee.uci.edu/api/v1/courses/${courseNum}/modules/${moduleId}/items/${moduleItem}`; 
              const query = {
                per_page: 100,
                student_id: student.id
              };
              const endpoint = moduleItemCompletionURL.addQuery(query);
              moduleCompleteInfo = JSON.parse(UrlFetchApp.fetch(endpoint, {headers:headers}).getContentText());
            }
            catch(e){
              console.error(`Error has occurred while checking a module item given a student id: ${e}`);
            }
            ifCompleted = moduleCompleteInfo.completion_requirement.completed;
            // let ifCompleted;
            // if (moduleCompleteInfo.completion_requirement == undefined){
            //   ifCompleted == "true"
            // }
            // else{
            //   ifCompleted = moduleCompleteInfo.completion_requirement.completed;
            // }
            if (ifCompleted == "true"){ // if module item complete, add "completed" else grab the state ("unlocked", "started", ...).
              tempStudentStorage.push("completed")
            }
            else{
              // Canvas stores states such as "locked", "started", "unlocked" so we can use those for each module item progress
              if (studentData.state == "started"){
                tempStudentStorage.push("to do")
              }
              else{
                tempStudentStorage.push(studentData.state);
              }
            }
          }
        }
      }
      completionData.push(tempStudentStorage);
    }

    // Write the values to the sheet.
    let completionRange;
    try{
      completionRange = sheet.getRange(startCompletionRow, startCompletionCol, completionData.length, completionData[0].length);
      completionRange.setValues(completionData);
    }
    catch(e){
      console.error(`An error occurred while writing to the sheet all the completion values for each module item: ${e}`)
    }
    return [completionData, completionRange]; // Return 2D array that has values of "locked", "started", etc and the range used to write to the sheet such as (7, 5, 10, 10);
}

function colorCompletionBackgrounds(completionData, completionRange, sheet){
    //  Adds the colored backgrounds based on the module item progress and calculates the percentages of each student.
    const colorsArray = [];
    const completionCol = [];
    for (let row of completionData){
      const temp = [];
      let completionCounter = 0;
      for (let col of row){
        // Based on the progress, add a color. 
        if (col == "completed"){
          temp.push("#92d050");
          completionCounter++;
        }
        else if (col == "locked"){
          temp.push("#92cddc");
        }
        else if (col == "to do"){
          temp.push("#ffc000");
        }
        else if (col == "unlocked"){
          temp.push("#ffffff");
        }
      }
      colorsArray.push(temp);
      completionCol.push([`${Math.round((completionCounter / completionData[0].length) * 100 )}%`]);
    }

    try{
      completionRange.setBackgrounds(colorsArray); // Change background of each cell that correlates a module items' progress.
      sheet.getRange(7, 4, completionData.length, 1).setValues(completionCol).setHorizontalAlignment("left");
    }
    catch(e){
      console.error(`Error occurred while setting color backgrounds for module item and calculating student percentages: ${e}`);
    }
    return completionCol;
}

function updateCompletionPercentages(completionData, sheet, completionCol, rowAfterStudents, numOfStudents){
    // Update completion percentages

    // Grab completion for each module
    const completionRow = [];
    for (let i = 0; i < completionData[0].length; i++){
      let completionCounter = 0;
      for (let j = 0; j < completionData.length; j++){
        let status = completionData[j][i];
        if (status == "completed"){
          completionCounter++;
        }
      }
      completionRow.push(`${Math.round((completionCounter / completionData.length) * 100 )}%`);
    }

    try{
      sheet.getRange(6, 5, 1, completionData[0].length).setValues([completionRow]).setHorizontalAlignment("left");
    }
    catch(e){
      console.error(`An error has occurred while writing the completion statisics for the row section: ${e}`);
    }

    try{
      let rowCompletionRange = `[6, 5, 1, ${completionData[0].length}]`;
      const clearDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ClearData');
      clearDataSheet.getRange(2, 3).setValue(rowCompletionRange);
    }
    catch(e){
      console.error(`An error occurred while writing the completion row position to the clearData sheet: ${e}`);
    }

    // Set up completion statisics for completin column. 
    const completionTrack = [[0], [0], [0], [0], [0], [0], [0], [0], [0], [0], [0]];

    for (percentage of completionCol){
      let value = parseInt(percentage);
      if (value == 0){
        completionTrack[0][0] = completionTrack[0][0] + 1;
      }
      else if (value > 0 && value <= 10){
        completionTrack[1][0] = completionTrack[1][0] + 1;
      }
      else if (value > 10 && value <= 20){
        completionTrack[2][0] = completionTrack[2][0] + 1;
      }
      else if (value > 20 && value <= 30){
        completionTrack[3][0] = completionTrack[3][0] + 1;
      }
      else if (value > 30 && value <= 40){
        completionTrack[4][0] = completionTrack[4][0] + 1;
      }
      else if (value > 40 && value <= 50){
        completionTrack[5][0] = completionTrack[5][0] + 1;
      }
      else if (value > 50 && value <= 60){
        completionTrack[6][0] = completionTrack[6][0] + 1;
      }
      else if (value > 60 && value <= 70){
        completionTrack[7][0] = completionTrack[7][0] + 1;
      }
      else if (value > 70 && value <= 80){
        completionTrack[8][0] = completionTrack[8][0] + 1;
      }
      else if (value > 80 && value <= 90){
        completionTrack[9][0] = completionTrack[9][0] + 1;
      }
      else if (value > 90 && value <= 100){
        completionTrack[10][0] = completionTrack[10][0] + 1;
      }

    }

    const completionPercentageTrack = completionTrack.map((element, index) => {
      return [`${Math.round((element[0] / numOfStudents) * 100 )}%`]; // Divides amount completed by number of students.
    });

    const completionMergedTrack = completionTrack.map((element, index) => {
      return [element[0], completionPercentageTrack[index][0]];
    });

    try{
      sheet.getRange(rowAfterStudents + 4, 3, 11, 2).setValues(completionMergedTrack);
    }
    catch(e){
      console.error(`An error has occurred while writing completion values to the sheet: ${e}`);
    }
}

function createPieChart(sheet, rowAfterStudents,  position, width, height, legend){
    // Creates a pie chart given a position, width, height, and boolean if a legend is needed
    // Note that the chart border lines can not transparent or changed.
    // https://issuetracker.google.com/issues/150189790?pli=1
    // seems to be an "intended feature" where u can not change the border of the pie chart background. can change it manually.

    // Referenced:
    // https://stackoverflow.com/questions/64386803/how-to-display-the-values-in-a-pie-chart-google-apps-script

    const chartDataRange = sheet.getRange(rowAfterStudents + 3, 1, 12, 1);   
    const chartDataRange2 = sheet.getRange(rowAfterStudents + 3, 3, 12, 1);   

    let legendOption = "";
    if (legend){
      legendOption = {position: 'right'};
    }
    else{
      legendOption = 'none';
    }

    try{
      const pieChartBuilder = sheet.newChart()
            .addRange(chartDataRange)
            .addRange(chartDataRange2)
            .setOption('legend', legendOption)
            .setOption('colors', ["#FFFFFF", "#ec3b12", "#f84d05", "#fa7e0c", "#fbb716", "#f8ee1b", "#ddf81f", "#aee319", "#7ae51b", "#42e31b", "#0ae21f", "#00ed19"])
            .setChartType(Charts.ChartType.PIE)
            .setPosition(position[0], position[1], position[2], position[3])
            .setOption('title', "")
            .setOption('width',width).setOption('height', height)
            .build();
      sheet.insertChart(pieChartBuilder);
    }
    catch(e){
      console.error(`An error occurred while creating a pie chart: ${e}`);
    }
}

function addPieChartsToSheet(sheet, rowAfterStudents){
    // Adds pie charts to the sheet. 
    // Calculate positions of where to put pie charts.
    const mainPieChartPos = [rowAfterStudents + 2, 5, 0, 0];
    const smallPieChartPos = [4, 1, 0, 25]
    // Passes in sheet, where to put pie chart, width, height, and if to have a legend.
    createPieChart(sheet, rowAfterStudents, mainPieChartPos, 500, 300, true);
    createPieChart(sheet, rowAfterStudents, smallPieChartPos, 125, 75, false);
}

function addProgressUpdatedTime(sheet){
    // Updates the time to when the progress report is ran.
    // https://stackoverflow.com/questions/10182020/how-to-get-the-current-time-in-google-spreadsheet-using-script-editor
    const now = new Date();
    lastUpdated = new Date(now).toLocaleString('en-us');
    sheet.getRange("E3").setValue(`Progress Updated Last on : ${lastUpdated}`); 
}

function clearEverything(){
    // This function is triggered when the clear button is clicked. Clear everything based on values from the clearData sheet.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //Grab current sheet 
    sheet.getRange("B5").setValue(`⬇️ 0 Students`); // Changes number of students in B5 to 0
    sheet.getRange("B3").setValue("0%");
    sheet.getRange("A1").setValue(`Progress report of students in modules in `).setFontWeight("bold");

    // https://webapps.stackexchange.com/questions/151412/how-to-combine-multiple-ranges-with-getrange-in-order-to-clear-the-contents-of-s
    try{
      const clearDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ClearData');
      var ranges = clearDataSheet.getRangeList(["A2", "B2", "C2"]).getRanges();
      for (cell of ranges){
        let range = eval(cell.getValue());
        sheet.getRange(range[0], range[1], range[2], range[3]).clearContent().clearFormat();
      }
    }
    catch(e){
      console.error(`An error occurred while clearing the content: ${e}`)
    }


    // Remove all charts in the current sheet.
    // https://stackoverflow.com/questions/53324528/how-to-delete-a-chart-using-google-script
    try{
      var chts=sheet.getCharts();
      for(var i=0;i<chts.length;i++){
        sheet.removeChart(chts[i]);
      }
    }
    catch(e){
      console.error(`An error occurred while deleting a pie chart: ${e}`)
    }
  }

