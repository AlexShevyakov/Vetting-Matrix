/**
 * Creates a string of random letters at a set length.
 *
 * @param {number} len The total number of random letters in the string.
 * @param {number} num What type of random number 0. Alphabet with Upper and Lower. 1.Alphanumeric 2. Alphanumeric + characters
 * @return an array of random letters
 * @customfunction
 */

function RANDALPHA(len, num) {
  var text = "";

  //Check if numbers
  if(typeof len !== 'number' ||  typeof num !== 'number'){return text = "NaN"};
  
  var charString = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()<>-=_+:;";
  var charStringRange
  switch (num){
     case 0:
       //Alphabet with upper and lower case
       charStringRange = charString.substr(0,52);
       break;
     case 1:
       //Alphanumeric
       charStringRange = charString.substr(0,62);
       break;
     case 2:
       //Alphanumeric + characters
       charStringRange = charString;
       break;
     default:
       //error reporting
       return text = "Error: Type choice > 2"    
     
  }
  //
  for (var i = 0; i < len; i++)
    text += charStringRange.charAt(Math.floor(Math.random() * charStringRange.length));
    
  return text;
}