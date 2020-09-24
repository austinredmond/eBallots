function updateForm() {
  // This function is an adapted version of the function built and explained by
  // Melinda Waffle at the following URL. Follow the guide to installation.
  // http://wafflebytes.blogspot.com/2016/10/google-script-create-drop-down-list.html
  
  // It creates a drop-down list from a spreadsheet column.
  // Once configured it is very easy to modify the entries of a drop-down menu
  // in a google form.
  
  // This function was designed for eBallots in mind for use by the
  // Alberta Debate and Speech Association.
  // A paper version of the ballot that is being modelled can be found at the 
  // following URL.
  // https://storage.googleapis.com/wzukusers/user-29340614/documents/59e76add0bbd59LpQJYw/CN%20l%20JrH%20Ballot.pdf

  // call your form and connect to the drop-down item
  
  // Edit line 20 by replacing the string with the Form ID that correponds to 
  // your Google Form URL.
  var form = FormApp.openById("Fill in Form ID");
  
  // Each drop-down list corresponds to a question in your Google Form.
  // Each question in the Google Form will have a data-item-ID that can be 
  // found by inspecting the element in the Google form.
  // Edit lines 27, 28, 30, 31, 32, 34, 35, 36  by replacing the string with the
  // found data-item-ID.
  var judgeList = form.getItemById("Fill in the Drop-Down List ID for judges").asListItem();
  var roomList = form.getItemById("Fill in the Drop-Down List ID for rooms").asListItem();
  
  var propositionTeamList = form.getItemById("Fill in the Drop-Down List ID for Prop. Team Codes").asListItem();
  var firstPropositionSpeakerList = form.getItemById("Fill in the Drop-Down List ID for First Prop.").asListItem();
  var secondPropositionSpeakerList = form.getItemById("Fill in the Drop-Down List ID for Second Prop.").asListItem();
  
  var oppositionTeamList = form.getItemById("Fill in the Drop-Down List ID for Propositon Team Codes").asListItem();
  var firstOppositionSpeakerList = form.getItemById("Fill in the Drop-Down List ID for First Opp.").asListItem();
  var secondOppositionSpeakerList = form.getItemById("Fill in the Drop-Down List ID for Second Prop.").asListItem();
  
  // Changes do not need to be made beyond this point.
  
  // identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.getActive();
  var names = ss.getSheetByName("Names List")
  
  var judgeNamesValues = names.getRange(2, 1, names.getMaxRows() - 1).getValues();
  var teamNamesValues = names.getRange(2, 2, names.getMaxRows() - 1).getValues();
  var SpeakerNamesValues = names.getRange(2, 3, names.getMaxRows() - 1).getValues();
  var roomNamesValues = names.getRange(2, 4, names.getMaxRows() - 1).getValues();
  
  var judgeNames = [];
  var teamNames = [];
  var SpeakerNames = [];
  var roomNames = [];
  
  // convert the array ignoring empty cells
  for(var i = 0; i < judgeNamesValues.length; i++)   
    if(judgeNamesValues[i][0] != "")
      judgeNames[i] = judgeNamesValues[i][0];
  
  // convert the array ignoring empty cells
  for(var i = 0; i < teamNamesValues.length; i++)   
    if(teamNamesValues[i][0] != "")
      teamNames[i] = teamNamesValues[i][0];
  
  // convert the array ignoring empty cells
  for(var i = 0; i < SpeakerNamesValues.length; i++)   
    if(SpeakerNamesValues[i][0] != "")
      SpeakerNames[i] = SpeakerNamesValues[i][0];
  
    // convert the array ignoring empty cells
  for(var i = 0; i < roomNamesValues.length; i++)   
    if(roomNamesValues[i][0] != "")
      roomNames[i] = roomNamesValues[i][0];
  
  // populate the drop-down with the array data
  judgeList.setChoiceValues(judgeNames);
  roomList.setChoiceValues(roomNames);

  
  propositionTeamList.setChoiceValues(teamNames);
  firstPropositionSpeakerList.setChoiceValues(SpeakerNames);
  secondPropositionSpeakerList.setChoiceValues(SpeakerNames);

  oppositionTeamList.setChoiceValues(teamNames);
  firstOppositionSpeakerList.setChoiceValues(SpeakerNames);
  secondOppositionSpeakerList.setChoiceValues(SpeakerNames);
  
}