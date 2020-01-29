function myFunction() {
  
}
function onOpen(){
    function updateForm(){ // select list from name
  // call your form and connect to the drop-down item
  var form = FormApp.openById("googleform-id");
  
  var namesList = form.getItemById("data-item-id").asListItem();





// identify the sheet where the data resides needed to populate the drop-down
  var ss =  SpreadsheetApp.openByUrl('link').getSheetByName("redeem_dropdownlist");

  var names =ss;


  // grab the values in the first column of the sheet - use 2 to skip header row

  var namesValues =  names.getRange("B2:B").getValues();
    var studentNames = [];




  // convert the array ignoring empty cells
  for(var i = 0; i < namesValues.length; i++)   
    if(namesValues[i][0] != "")
      studentNames[i] = namesValues[i][0];

  // populate the drop-down with the array data
  namesList.setChoiceValues(studentNames);
 
    }
}// on open
