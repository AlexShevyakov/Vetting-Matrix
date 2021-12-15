// This code generates IDs for any number of rows



function tempID(){
  

  const ss = SpreadsheetApp.openById('1pam32jvPiSoZJj1_S7gsKDa2tmU02Fv7GdKuUxxSJVk');
  const sheet = ss.getSheetByName("Incidents");
    
  for (let i = 2; i <= 30; i ++ ){
    let cell = sheet.getRange(i, 1)
    cell.setValue(randStr(8));
    
  }
}


function randStr(len) {
  let u_str = ''
  while (u_str.length < len) u_str += Math.random().toString(36).substr(2, len - u_str.length);
  return u_str;
}            
       