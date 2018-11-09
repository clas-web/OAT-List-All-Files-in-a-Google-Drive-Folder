//https://ctrlq.org/code/20039-google-picker-with-apps-script
//https://ctrlq.org/code/19854-list-files-in-google-drive-folder
//https://stackoverflow.com/questions/23074352/google-picker-return-the-file-id-to-my-google-script
//https://stackoverflow.com/questions/35641674/google-script-app-list-drive-folder-in-sheets
//http://www.acrosswalls.org/ortext-datalinks/list-google-drive-folder-file-names-urls/
//https://developers.google.com/picker/
//https://developers.google.com/apps-script/reference/drive/drive-app#getfolderbyidid
//https://productforums.google.com/forum/#!topic/docs/FcfEWNZzDeI
//https://ctrlq.org/code/19975-move-file-between-folders
//**********************************************************************************************************************************************************************************************
//Create Menu option when spreadsheet is opened
function onOpen() {
  SpreadsheetApp.getUi().createMenu('List Files in a Google Drive Folder')
  .addItem('Select Folder to List Contents...', 'showPicker')
  .addItem('Organize items based on NinerNET', 'NinerNetSort')
  .addToUi();
}

//**********************************************************************************************************************************************************************************************
//Run code when the form is submitted
function onFormSubmit(e) {
  NinerNetSort();
}
//**********************************************************************************************************************************************************************************************
//Create folder selector
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
  .setWidth(600)
  .setHeight(425)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Folder');
}
//**********************************************************************************************************************************************************************************************
//Authorize script
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
//**********************************************************************************************************************************************************************************************
//Move files into folder by NinerNet
function NinerNetSort(){
  
  //Select folder
  //var folder = DriveApp.getFolderById('0B0hi2JuumvKDfnRmcTB2S1UyU0RHQTRtdXhzSUROYWRTanhIamRlY3RKUGt5Z0tzQXpubkk');
  var folder = DriveApp.getFolderById('1PXTpXg73cIuSnGlgy321L_GnIcMYql32');
  //Select relevant sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  //grab all form submissions
  var entries = sheet.getDataRange().getValues();
  
  //Declare values
  var description; //user's description of the uploaded item(s) from form
  var email; //user's email address from form
  var fileURL; //uploaded file's address
  var fileID; //uploaded file's Google Drive ID
  var fileString; //uploaded file's Google Drive URL (array split to find ID)
  var ninerNETFullAddress; //user's email address from form (array split to find NinerNET)
  var ninerNET; //user's NinerNET
  var ninernet_Subfolder; //user's subfolder based on NinerNET
  var fileURLArray = []; //array of uploaded files to move into subfolders
  
  Logger.log("entries length: " + entries.length); //How many entries do we have
  Logger.log("entries[0] is " + entries[0]); //confirm first "entry" is the header row
  
  //Iterate through each entry on spreadsheet
  for (var i = 1; i <= entries.length; i++) {
    if (i != entries.length){
      
      //Check to make sure entry has not already been sent to subfolder
      if (entries[i][4] != 'Yes') {
        
        fileURL = entries[i][1];
        description = entries[i][2];
        email = entries[i][3];
        
        //Split user's email address so we have their NinerNET
        ninerNETFullAddress = email.toString().split('@');
        ninerNET = ninerNETFullAddress[0];
        Logger.log(ninerNET);
        
        //Push all file ID's to an array so they will be moved into the proper subfolders
        fileURLArray[0] = '';
        
        //No commas detected, thus there's only one uploaded file in this entry. We can add it straight to our array
        if (fileURL.indexOf(',') < 0){
          //Split the file's Google Drive URL to get the file ID
          fileString = fileURL.toString().split('https://drive.google.com/open?id=');
          fileID = fileString[1];
          Logger.log(fileID);
          fileURLArray[1] = fileID;
        } else {
          //Commas detected, meaning a) the user uploaded more than 1 file and b) we'll need to loop through their entry
          // to grab each file ID and push to our array
          var commaArray = fileURL.toString().split(',');
          for (var n = 0; n < commaArray.length; n++) { 
            fileString = commaArray[n].toString().split('https://drive.google.com/open?id=');
            fileID = fileString[1];
            //Add the file IDs to the end of our file ID array
            fileURLArray.push(fileID);
            Logger.log('commaArray['+n+'] '+ commaArray[n]);
          }
        }        
        //*********************
        //Checks if user's NinerNET subfolder exists, if it doesn't, create it
        try {
          //Folder exists
          ninernet_Subfolder = folder.getFoldersByName(ninerNET).next();   
          Logger.log('folder '+ninerNET+ ' exists');
        }
        catch(e) {
          ////Folder doesn't exist, create folder
          ninernet_Subfolder = folder.createFolder(ninerNET);
          Logger.log('folder '+ninerNET+ ' does not exist, creating folder');
        }
        //*********************        
        //"Move" file(s) to subfolder based on file IDs
        Logger.log('fileURLArray.length: ' + fileURLArray.length);
        for (var j = fileURLArray.length-1; j >= 1; j--) {
          if (ninernet_Subfolder.getName() == ninerNET){
            
            //Alter permissions on file
            modifyPermissions(fileURLArray[j],email);
            
            ninernet_Subfolder.addFile(DriveApp.getFileById(fileURLArray[j]));
            
            //Add description to file(s) from user's Google Form entry
            DriveApp.getFileById(fileURLArray[j]).setDescription(description);
            //Remove from parent folder to complete move operation
            folder.removeFile(DriveApp.getFileById(fileURLArray[j]));
            Logger.log(fileURLArray[j]);
            
            //Confirm the file is in the correct subfolder, indicate so on spreadsheet
            Logger.log('Parent folder: ' + DriveApp.getFileById(fileURLArray[j]).getParents().next());
            if (DriveApp.getFileById(fileURLArray[j]).getParents().next() == ninerNET){      
              //"Has this file been moved to the correct user's NinerNET folder?"
              sheet.getRange(i+1, 5).setValue('Yes');              
            }
            
            //Remove entry from array since it has already been assigned a subfolder
            fileURLArray.pop();
          }
        }
        //*********************
      } else if (entries[i][4] == 'Yes') {
        //Skip entry
        Logger.log('Entry '+i+' has already been sent to subfolder');        
      } //*********************
    }
  }    
  Logger.log("Faneto");
}
//**********************************************************************************************************************************************************************************************
//https://developers.google.com/apps-script/reference/drive/file#setOwner(String)
//Add permissions to folder where designated user/committee is the owner, uploader has read access
function modifyPermissions(fileID,email){
  
  //Test folder, comment out when done testing
  //folderID = '1r5ZYvsgeoqK_3zbIMeJoSyyy9kO2FErg';
  
  //var folder = DriveApp.getFolderById(folderID);
  //Logger.log('Owner of folder: ' + folder.getOwner().getEmail());
  //var files = folder.getFiles();    
  //var file;
  var editors = DriveApp.getFileById(fileID).getEditors();  
  
  //Trying Alex out, change to whoever the new owners should be
  //var newOwner = 'dept-clas-oat@uncc.edu';
  //Testing secondary editor
  //editor = 'rmccal14@uncc.edu';
  
  //folder.addEditor(newOwner);
  //folder.setOwner(newOwner);
  
  //while (files.hasNext()){
  //file = files.next();
  
  //editors = file.getEditors();
  //Logger.log('Owner of file '+ file.getName() +': ' + file.getOwner().getEmail());        
  
  //Remove users as editors, add as viewers
  //DriveApp.getFileById(fileID).revokePermissions(email); 
  //DriveApp.getFileById(fileID).addViewer(email);
  
  for each (var f in editors) {         
    Logger.log('Editor: ' + f.getEmail());
    DriveApp.getFileById(fileID).revokePermissions(f.getEmail()); 
    DriveApp.getFileById(fileID).addViewer(f.getEmail());
    DriveApp.getFileById(fileID).setShareableByEditors(FALSE);
  }
  
  //Don't have to set new owner if it goes to team folder
  //file.setOwner(newOwner);
  //Logger.log('New owner of file '+ file.getName() +': ' + file.getOwner().getEmail());
  //file.addEditor(editor); 
  //Logger.log('New editor(s) of file '+ file.getName() +': ' + file.getEditors());
  
  //recheck list of editors
  editors = DriveApp.getFileById(fileID).getEditors();  
  
  for each (var f in editors) {         
    Logger.log('New Editors: ' + f.getEmail());      
  }
  //}
  
  //folder.removeEditor(editor); 
  //file.removeEditor(editor); 
  //folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
}

//**********************************************************************************************************************************************************************************************
// function that Picker passes the folder ID to
function listFolderContents(id) {
  
  //*****************************
  //test case, uncomment to try and comment when done
  //id = '0B0hi2JuumvKDfk5FTmV1c19BbWtwWU9KZDJIQkp2NHBTMDJ4c0RfdjlwQUN4cUhtYzJoVjg';
  //*****************************
  
  //locate the Drive folder with the exact ID  
  var folderID = DriveApp.getFolderById(id); 
  var foldername = folderID.getName();
  var folderURL = folderID.getUrl();
  
  //just in case we need to modify based on Row 1  
  var row1 = ['Name', 'Date Created', 'Folder', 'Type', 'Link', 'Owner', 'Editor(s)', 'Description'];  
  Logger.log(row1);
  var fModifier = 0;
  
  //get the files from the folder
  var contents = folderID.getFiles();
  
  //designate our spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var currentCell = sheet.getCurrentCell();
  
  //designate our Row 1, if not the headers, add the headers
  var currentRow1Range = sheet.getRange(1, 1, 1, 8).getValues();
  Logger.log(currentRow1Range);
  
  if (row1.toString() == currentRow1Range.toString()) {    
    Logger.log("They match");    
  } else {    
    Logger.log("They don't match");    
    sheet.insertRowBefore(1);    
    fModifier = 1;
    for (var j = 1; j <= 8; j++) {
      //var currentValue = sheet.getRange(1,j).getValue();      
      sheet.getRange(1,j).setValue(row1[j-1]);
    }
  }
  //create a filter for our data if one does not exist  
  Logger.log(sheet.getFilter());
  if (sheet.getFilter() == null || fModifier == 1){
    
    Logger.log('Yes, it is null so create a filter');
    if (fModifier == 1){
      sheet.getFilter().remove();    
      Logger.log('The header row changed so alter the filter size automatically');
    }
    sheet.getRange('A1:H').activate();
    sheet.getRange('A1:H').createFilter();
  }else{
    Logger.log('No, it is not null, spreadsheet filter exists already');  
  }  
  //**********************************************************
  //original code, clears the sheet every use and appends original header row
  //clear the sheet  
  //sheet.clear();
  //create header row if it hasn't been created
  //sheet.appendRow(row1);
  //**********************************************************  
  var file; //object
  var name; //name of object
  var link; //link to object
  var date; //date object was created
  var type; //what kind of object is it?
  var info; //metadata on file
  var owner; //owner of file
  var folderLink = '=HYPERLINK("'+folderURL+'","'+foldername+'")'; //link to folder
  var editors; //user(s) with edit access to the file
  var eString; //editors as one string
  while(contents.hasNext()) {
    file = contents.next();
    name = file.getName();
    link = file.getUrl();
    date = file.getDateCreated();
    type = file.getMimeType();
    info = file.getDescription();
    owner = file.getOwner().getEmail();
    editors = file.getEditors();    
    
    for (var i = 0; i < editors.length; i++) {
      editors[i] = editors[i].getEmail();
    }
    eString = editors.toString();
    Logger.log(eString);
    
    sheet.appendRow([name, date, folderLink, type, link, owner, eString, info]);     
  }  
  //********************************************************
  // Get sub-folders in folder
  var sub_folder_names = folderID.getFolders();
  var sub_folder_name; //sub folder object
  var sname; //sub folder name
  var sdate; //sub folder date created
  var slink; //link to sub folder
  var sinfo; //metadata on folder
  var sowner; //owner of folder
  var seditors; //user(s) with edit access to the folder
  var seString; //editors as one string
  while (sub_folder_names.hasNext()) {
    sub_folder_name = sub_folder_names.next();
    sname = sub_folder_name.getName();
    sdate = sub_folder_name.getDateCreated();
    slink = sub_folder_name.getUrl();
    sinfo = sub_folder_name.getDescription();
    sowner = sub_folder_name.getOwner().getEmail();
    seditors = sub_folder_name.getEditors();
    
    for (var i = 0; i < seditors.length; i++) {
      seditors[i] = seditors[i].getEmail();
    }
    seString = seditors.toString();
    Logger.log(seString);
    
    sheet.appendRow ([sname, sdate, folderLink, 'Folder', slink, sowner, seString, sinfo]);    
  }  
  //***********************************************************************************************  
  //Go back to active cell
  currentCell.activate();    
}
//**********************************************************************************************************************************************************************************************