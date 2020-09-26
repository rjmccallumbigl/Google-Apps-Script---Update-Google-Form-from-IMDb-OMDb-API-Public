/*************************************************************************************************************************************************************************************
*
* Update Google Form with IMDb URL and allow surveyors to pick new film from list of options.
* https://www.reddit.com/r/googlesheets/comments/iw0n1w/google_form_get_imdb_url_as_input_and_replace/
*
* Directions
* 1. Go to http://www.omdbapi.com/apikey.aspx
* 2. Create your free API key.
* 3. Your API key will be emailed to you. Activate it via the link.
* 4. Link your Google Form to a Google Spreadsheet and add this code to the Google Sheet by clicking Tools -> Script Editor, deleting all the text there, and pasting this script.
* 5. Add your AKI token below to var OMDbApiKey in getMovie().
* 6. (Optional) Add an example IMDb link to var firstMovie in primaryFunction().One has been provided already.
* 7. Create 2 sheets: one called "Movies" and another called "Poll". 
* 8. Run primaryFunction(). If it fails, run it once again.
* 9. Now users can vote (or submit a new movie) using the Form and the vote tally will be seen on the Poll sheet. The Movies sheet will have a database of all IMDb movies in the form so far.
*
* References
* http://googleappscripting.com/json/
* 
*************************************************************************************************************************************************************************************/
function primaryFunction(){
  try{
    //  Declare variables  
    var firstMovie = "https://www.imdb.com/title/tt0264464/";
    getMovie(firstMovie);
    var name = "Select movie or enter IMDB link of your movie suggestion.";
    var movieObject = {};
    movieObject.namedValues = {};
    movieObject.namedValues[name] = firstMovie;
    
//    Initial run, update form and sheets
    onFormSubmit(movieObject);
  } catch(e){
    primaryFunction();
  }
}

/*************************************************************************************************************************************************************************************
*
* Link to OMDb API and return film data from IMDB. Add to Movie sheet.
* 
* @param IMDb_URL {String} The IMDb movie URL passed into the function.
*
*************************************************************************************************************************************************************************************/

function getMovie(IMDb_URL) {
  
  //  Pause script to not trigger API limits
  Utilities.sleep(3000);
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var movieSheetName = "Movies";
  var sheet = spreadsheet.getSheetByName(movieSheetName);
  
  // Add your OMDb token
  var OMDbApiKey = 'xXxXxXxX';
  
  //  Set authentication object parameters
  var headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + OMDbApiKey,
  };
  
  //  Set option parameters
  var options = {
    "method": "POST",
    "headers": headers,
    "muteHttpExceptions": true,
  };
  
  //  Pass IMDb movie ID to omdbapi.com
  var IMDb_ID = return_IMDb_ID(IMDb_URL, "/");
  
  //  Make sure value isn't already on sheet
  var IDHeader = sheet.getDataRange().getDisplayValues()[0].indexOf("imdbID");  
  var flatArray = [];
  
  try{
    flatArray = sheet.getRange(1, IDHeader + 1, sheet.getLastRow(), 1).getDisplayValues().join().split(',');
  } catch (e) {
    console.log(e);
  }
  console.log(flatArray);
  if (flatArray.indexOf(IMDb_ID) < 0){
    var apiUrl = "http://www.omdbapi.com/?i=" + IMDb_ID + "&apikey=" + OMDbApiKey;  
    console.log(apiUrl);
    var response = UrlFetchApp.fetch(apiUrl, options);
    
    //  Return movie data
    var responseText = response.getContentText();
    console.log(responseText);
    var responseTextJSON = JSON.parse(responseText);  
    console.log(responseTextJSON);
    
    // Define an array of all the returned object's keys to act as the Header Row
    var headerRow = Object.keys(responseTextJSON);
    
    // Define an array of all the returned object's values
    var row = headerRow.map(function(key){ return responseTextJSON[key]});
    
    //  Set hyperlink of title to IMDb URL
    row[0] = '=HYPERLINK("' + IMDb_URL + '","'+ row[0] + '")';
    
    //  Define contents while determining if we need to add the header row
    var contents = (sheet.getRange(1, 1).getValue() == headerRow[0]) ? [row] : [headerRow, row];
    
    // Select the spreadsheet range and set values  
    var dataRange = sheet.getRange(sheet.getLastRow() + 1, 1, contents.length, row.length);
    dataRange.setValues(contents);
    
    SpreadsheetApp.flush();
  }
} 

/*************************************************************************************************************************************************************************************
*
* Return IMDb ID from IMDb movie URL.
* 
* @param IMDb_URL {String} The IMDb movie URL passed into the function.
* @param splitter {String} The character that the array will split on.
* @return {String} The IMDb movie ID extracted from the URL.
*
*************************************************************************************************************************************************************************************/

function return_IMDb_ID(IMDb_URL, splitter) {
  
  // Split URL
  var splitArray = IMDb_URL.split(splitter);
  
  //  Return array value with "tt" in the string, which all IMDb movie IDs have
  var getID = splitArray.find(function (value) {
    return /^tt/.test(value);
  });
  
  //  Return the ID
  return getID;
}

/*************************************************************************************************************************************************************************************
*
* Update Google Form
* 
*************************************************************************************************************************************************************************************/

function updateForm() {  
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Movies");
  var formURL = spreadsheet.getFormUrl();
  var form = FormApp.openByUrl(formURL)
  .setTitle("Friday movie suggestion")
  .setDescription('Vote for an existing title OR enter a new suggestion: ' + spreadsheet.getUrl())
  .setConfirmationMessage('Your response has been recorded. Please give the form another minute to regenerate for updates.');
  var range = sheet.getDataRange();
  var rangeValues = range.getDisplayValues();
  var nameHeader = rangeValues[0].indexOf("Title");
  var IDHeader = rangeValues[0].indexOf("imdbID");
  var yearHeader = rangeValues[0].indexOf("Year");
  var ratedHeader = rangeValues[0].indexOf("Rated");
  var releasedHeader = rangeValues[0].indexOf("Released");
  var runtimeHeader = rangeValues[0].indexOf("Runtime");
  var genreHeader = rangeValues[0].indexOf("Genre");
  var actorsHeader = rangeValues[0].indexOf("Actors");
  var plotHeader = rangeValues[0].indexOf("Plot");
  var ratingsHeader = rangeValues[0].indexOf("Ratings");
  var nameItem = "";
  var movieChoice = "";
  var movieChoiceArray = [];
  var movieInfoArray = [];
  
  //  Delete all current form questions
  deleteFormItems(form);
  
  //    Add name prompt
  nameItem = form.addTextItem()
  .setTitle("Name")
  .setRequired(true);
  
  //  Add movie choices
  movieChoice = form.addMultipleChoiceItem()
  .setTitle("Select movie or enter IMDB link of your movie suggestion.")
  .setHelpText("Example: https://www.imdb.com/title/tt8111088");
  
  
  //  Go through each row on sheet to add as a choice, skipping header row
  for (var row = 1; row < rangeValues.length; row++){
    
    movieInfoArray.length = 0;
    movieInfoArray.push(rangeValues[row][nameHeader]);
    movieInfoArray.push("https://www.imdb.com/title/" + rangeValues[row][IDHeader]);
    movieInfoArray.push(rangeValues[row][yearHeader]);
    movieInfoArray.push(rangeValues[row][ratedHeader]);
    movieInfoArray.push(rangeValues[row][releasedHeader]);
    movieInfoArray.push(rangeValues[row][runtimeHeader]);
    movieInfoArray.push(rangeValues[row][genreHeader]);
    movieInfoArray.push(rangeValues[row][actorsHeader]);
    movieInfoArray.push(rangeValues[row][plotHeader]);
    movieInfoArray.push(rangeValues[row][ratingsHeader]);
    
    // Add movie choice
    movieChoiceArray.push(movieChoice.createChoice(movieInfoArray.join(" | ")));
  }
  
  //  Set movie choices
  movieChoice.showOtherOption(true)
  .setRequired(true)
  .setChoices(movieChoiceArray);
  
  // Deletes all triggers in the current project.
  setTriggers(spreadsheet);
}

/*************************************************************************************************************************************************************************************
*
* Delete current form questions.
*
* @param {Object} form This is the current form attached to the spreadsheet.
*
*************************************************************************************************************************************************************************************/

function deleteFormItems(form){
  
  //  Make sure we have the form
  var form = form || FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl()); 
  
  //  Collect form items
  var formItems = form.getItems();
  
  //  Loop through and delete each form item
  for (var count = 0; count < formItems.length; count++){
    form.deleteItem(formItems[count]);
  } 
}

/*************************************************************************************************************************************************************************************
*
* Deletes all triggers in the current project.
*
* @param {Object} spreadsheet This is our primary spreadsheet.
*
*************************************************************************************************************************************************************************************/

function setTriggers(spreadsheet){
  
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  //    Create trigger to capture new form submissions
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(spreadsheet).onFormSubmit()
  .create();
}

// ********************************************************************************************************************************
/**
* A trigger-driven function that updates the sheet and form after a user responds to the form.
*
* @param {Object} e The event parameter for form submission to a spreadsheet;
*     see https://developers.google.com/apps-script/understanding_events
*/

function onFormSubmit(e) {
  
  //  Declare Poll sheet variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Poll");
  var range = sheet.getDataRange();
  var rangeValues = range.getDisplayValues();  
  
  //  Make sure we have a header row in Poll sheet
  if (rangeValues[0][0] != 'Title') {
    var headerRow = ["Title", "URL", "imdbID", "Votes"];
    sheet.appendRow(headerRow);    
  }
  
  //  Return Poll column headers
  var nameHeader = rangeValues[0].indexOf("Title");  
  var URLHeader = rangeValues[0].indexOf("URL");
  var IDHeader = rangeValues[0].indexOf("imdbID");  
  var votesHeader = rangeValues[0].indexOf("Votes");  
  
  //  Flatten array to grab imdbID index easier
  SpreadsheetApp.flush();
  var pollValuesFlattened = sheet.getRange(1, IDHeader + 1, sheet.getLastRow(), 1).getDisplayValues().join().split(',');
  
  //  Declare Movie sheet variables
  var moviesSheet = spreadsheet.getSheetByName("Movies");
  var moviesSheetRange = moviesSheet.getDataRange();
  var movieRangeValues = moviesSheetRange.getDisplayValues();
  
  //  Return Movie column headers
  var movieNameHeader = movieRangeValues[0].indexOf("Title");
  var movieIDHeader = movieRangeValues[0].indexOf("imdbID");
  
  //  Flatten array to grab imdbID index easier
  var movieValuesFlattened = moviesSheet.getRange(1, movieIDHeader + 1, moviesSheet.getLastRow(), 1).getDisplayValues().join().split(',');
  
  var name = "Select movie or enter IMDB link of your movie suggestion.";  
  var returnedString = e.namedValues[name] + '';  
  var IMDb_ID = "";
  var rowContents = [];
  var title = "";
  
  //  Already on Form
  if (returnedString.indexOf("|") > -1){
    IMDb_ID = return_IMDb_ID(returnedString.split(" | ")[1], "/");
    
    //    Confirm it's on Poll sheet, if not add to Poll sheet from Movie sheet
    if (pollValuesFlattened.indexOf(IMDb_ID) < 0){
      
      //    Grab title based on ID
      title = movieRangeValues[movieValuesFlattened.indexOf(IMDb_ID)][movieNameHeader];
      
      //    Add movie data to array
      rowContents.push(title);
      rowContents.push("https://www.imdb.com/title/" + IMDb_ID);
      rowContents.push(IMDb_ID);
      rowContents.push(0);
      
      // Update Poll with movie
      sheet.appendRow(rowContents);
      SpreadsheetApp.flush();
      
      //  Update Poll sheet variables
      range = sheet.getDataRange();
      rangeValues = range.getDisplayValues();  
      nameHeader = rangeValues[0].indexOf("Title");  
      URLHeader = rangeValues[0].indexOf("URL");
      IDHeader = rangeValues[0].indexOf("imdbID");  
      votesHeader = rangeValues[0].indexOf("Votes");  
      pollValuesFlattened = sheet.getRange(1, IDHeader + 1, sheet.getLastRow(), 1).getDisplayValues().join().split(',');
    }
    
    //    Increment poll vote for this movie
    rangeValues[pollValuesFlattened.indexOf(IMDb_ID)][votesHeader] = Number(rangeValues[pollValuesFlattened.indexOf(IMDb_ID)][votesHeader]) + 1;
    
    // Update sheet with new vote  
    range.setValues(rangeValues);
    SpreadsheetApp.flush();
    
  } else {
    //    Not on Poll list, add movie
    IMDb_ID = return_IMDb_ID(returnedString, "/");
    getMovie(returnedString + '');
    SpreadsheetApp.flush();
    
    // Update variables since we updated the Movie sheet
    moviesSheetRange = moviesSheet.getDataRange();
    movieRangeValues = moviesSheetRange.getDisplayValues();
    movieNameHeader = movieRangeValues[0].indexOf("Title");
    movieIDHeader = movieRangeValues[0].indexOf("imdbID");
    movieValuesFlattened = moviesSheet.getRange(1, movieIDHeader + 1, moviesSheet.getLastRow(), 1).getDisplayValues().join().split(',');
    
    //    Grab title based on ID
    title = movieRangeValues[movieValuesFlattened.indexOf(IMDb_ID)][movieNameHeader];
    
    //    Add movie data to array
    rowContents.push(title);
    rowContents.push("https://www.imdb.com/title/" + IMDb_ID);
    rowContents.push(IMDb_ID);
    rowContents.push(1);
    
    // Update Poll with movie
    sheet.appendRow(rowContents);
    SpreadsheetApp.flush();
    
    //  Update form
    try {
      updateForm();
    } catch (e) {
      console.log("Form not up to date");
      console.log(e);
    }
  }
}