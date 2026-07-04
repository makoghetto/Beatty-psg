// Code.gs
// Replace DRIVE_FOLDER_ID and SHEET_ID.
const DRIVE_FOLDER_ID='PASTE_FOLDER_ID';
const SHEET_ID='PASTE_SHEET_ID';
function doGet(e){
 const sh=SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
 const vals=sh.getDataRange().getValues().slice(1);
 if(e.parameter.action=='list'){
   const out=vals.map(r=>({id:r[0],name:r[1],url:r[2],printed:r[3]}));
   return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
 }
}
function doPost(e){
 const body=JSON.parse(e.postData.contents);
 const sh=SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
 if(body.action=='printed'){
   const vals=sh.getDataRange().getValues();
   for(let i=1;i<vals.length;i++){if(vals[i][0]==body.id){sh.getRange(i+1,4).setValue(true);break;}}
   return ContentService.createTextOutput("OK");
 }
 const bytes=Utilities.base64Decode(body.image.split(',')[1]);
 const blob=Utilities.newBlob(bytes,'image/jpeg','GF_'+Date.now()+'.jpg');
 const file=DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);
 file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
 sh.appendRow([Date.now(),file.getName(),file.getDownloadUrl(),false]);
 return ContentService.createTextOutput("Uploaded");
}
