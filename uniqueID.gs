// Filling IDs for the defined range of the cells in the format of 8 characters
// like: 


  

function unique_ID_screenings(e){
  if (
    e.range.getSheet().getName() === 'Screening' &&
    e.range.columnStart == 2 &&
    e.range.columnEnd == 2 &&
    e.range.rowStart >= 2 &&
    e.range.rowEnd <= 100 &&
    e.range.offset(0, -1).getValue() === ''
  ) {
    e.value !== '' ? e.range.offset(0, -1).setValue(randStr(8)) : null;
  }
}


function unique_ID_inciednts(e){
  
  if (
    e.range.getSheet().getName() === 'Incidents' &&
    e.range.columnStart == 2 &&
    e.range.columnEnd == 2 &&
    e.range.rowStart >= 2 &&
    e.range.rowEnd <= 100 &&
    e.range.offset(0, -1).getValue() === ''
  ) {
    e.value !== '' ? e.range.offset(0, -1).setValue(randStr(8)) : null;
  }
}


function randStr(len) {
  let u_str = ''
  while (u_str.length < len) u_str += Math.random().toString(36).substr(2, len - u_str.length);
  return u_str;
}
