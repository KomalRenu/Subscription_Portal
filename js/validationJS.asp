<%'** Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. **'%>
<SCRIPT LANGUAGE=javascript>
<!--
  function isBlank(s) {
    for(var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
        if ((c != ' ') && (c != '\n') && (c != '\t')) return false;
      }
    return true;
  }

  function checkInvalidCharacters(sValue) {
    var aChars;
    var i;

    aChars = invalidChars();

    for (i=0; i < aChars.length; i++) {
      if (sValue.indexOf(aChars.charAt(i)) != -1) return false;
    }

    return true;
  }


  function invalidChars() {
    var aChars

    aChars = "";
<%
  Dim nReservedCount
  For nReservedCount = 0 To UBound(asReservedChars)
      Response.Write "    aChars += ""\" & asReservedChars(nReservedCount) & """;" & vbCr
  Next
%>
    return (aChars);
  }

  function checkIsAlphaNumeric(sValue) {
    var re;

    re = /\W/i;
    return (sValue.search(re) == -1);

  }

  function checkIsNumeric(sNum) {
    var nNum;

    nNum = parseInt(sNum, 10);
    if (isNaN(nNum))
      return(false);
    else
      return(true);
  }


  function checkEmailFormat(sValue){
    var bValid = true;
    var atLocation = 0;

    if(sValue.indexOf('@') == -1){
      bValid = false;
    }
    else{
      atLocation = sValue.indexOf('@');
      if(sValue.indexOf('.', atLocation) == -1) return false;
      if(sValue.indexOf('.', atLocation) == atLocation + 1) return false;
      if(sValue.substring(0,1) == '@') return false;
      if(sValue.charAt(sValue.length - 1) == '.') return false;
    }
    return bValid;
  }


  function checkNumericFormat(sNum) {
    var numbers = "<%Response.Write S_VALID_CHARS_NUM_ADDRESS%>";
    var thisChar;
    var counter = 0;

    for (var i=0; i < sNum.length; i++){
      thisChar = sNum.substring(i, i+1);
      if (numbers.indexOf(thisChar) != -1)
        counter++;
      }

      if (counter == sNum.length)
        return true;
      else
        return false;
  }


  function checkIsDate(sDate){
    var aDate
    var nDay;
    var nMonth;
    var nYear;

    aDate = sDate.split("/");
    if (aDate.length != 3) return(false);

    if (checkIsNumeric(aDate[0]) == false) return (false);
    if (checkIsNumeric(aDate[1]) == false) return (false);
    if (checkIsNumeric(aDate[2]) == false) return (false);

    nMonth = parseInt(aDate[0]);
    nDay   = parseInt(aDate[1]);
    nYear  = parseInt(aDate[2]);


    if (nMonth > 12 || nMonth < 1) return false;
    if ((nMonth == 1 || nMonth == 3 || nMonth == 5 || nMonth == 7 || nMonth == 8 || nMonth == 10 || nMonth == 12) && (nDay > 31 || nDay < 1)) return false;
    if ((nMonth == 4 || nMonth == 6 || nMonth == 9 || nMonth == 11) && (nDay > 30 || nDay < 1)) return false;
    if (nMonth == 2) {
      if (nDay < 1){
        return(false);
      } else {
        if (LeapYear(nYear) == true) {
          if (nDay > 29) return(false);
        }else {
          if (nDay > 28) return(false);
        }
      }
    }

    return true;
  }


  function LeapYear(nYear) {
    if (nYear % 100 == 0) {
      if (nYear % 400 == 0) { return true; }
    } else {
      if ((nYear % 4) == 0) { return true; }
    }
    return false;
  }

//-->
</SCRIPT>