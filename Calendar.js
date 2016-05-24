//This javascript file contains the function required for the calendar
//component for picking dates in prompts and searches
// Month ranges from 0 till 11

var Month;
var	Year;
var textBoxID;
var bIsSeen=false;
var sOperator; //to store the operator name

//Function to determine if a year is a leap year
function isLeapYear (myYear)
{
   if (((myYear % 4)==0) && ((myYear % 100)!=0) || ((myYear % 400)==0))
      return (true);
   else
      return (false);
}

//function to get the number of days in a month
function getDaysOfMonth (myMonth, myYear)
{
   var days;
   if (myMonth==1 || myMonth==3 || myMonth==5 || myMonth==7 || myMonth==8 || myMonth==10|| myMonth==12)
      days=31;
   else if (myMonth==4 || myMonth==6 || myMonth==9 || myMonth==11) days=30;
   else if (myMonth==2)
   {
      if (isLeapYear(myYear)) days=29;
      else days=28;
   }
   return (days);
}

//this function will create the DIV tags where the Calendar table will be placed

function createDivForCalendar()
{
	//first we need to alter the date format
	if (sDateFormat.search("DD") == -1) {
		sDateFormat = sDateFormat.replace("D","DD");
	}
	if (sDateFormat.search("MM") == -1) {
		sDateFormat = sDateFormat.replace("M","MM");
	}
	var oCalendarDiv = getObj('Calendar');
	if (oCalendarDiv == null) {
		document.write("<STYLE TYPE='text/css'>");
		document.write(".arrow {");
			document.write("background-color: #000000;");
			document.write("border:           1px #000000 solid;");
			document.write("color:            #ffffff;");
			document.write("font-family:      Wingdings;");
			document.write("font-size:        8pt;");
			document.write("font-style:       normal;");
			document.write("font-weight:      normal;");
			document.write("text-align:       center;");
			document.write("vertical-align:   middle;");
			document.write("cursor:          hand;");
		document.write("}");
		document.write(".arrownorm {");
	        document.write("color:           #ffffff;");
			document.write("font-style:      normal;");
			document.write("text-decoration: none;");
	        
		document.write("}");
		document.write(".arrowrover {");
	        document.write("color:           #cc3333;");
			document.write("font-style:      normal;");
			document.write("text-decoration: none;");
			document.write("cursor:          hand;");
		document.write("}");
		document.write(".weekday {");
			document.write("background-color: #DDDDBB;");
			document.write("border-bottom:    1px #000000 solid;");
			document.write("color:            #000000;");
			document.write("font-family:      Arial,Helvetica,Sans-Serif;");
			document.write("font-size:        8pt;");
			document.write("font-style:       normal;");
			document.write("font-weight:      normal;");
			document.write("text-align:       center;");
			document.write("vertical-align:   middle;");
		document.write("}");
		document.write(".day {");
	        document.write("background-color: #eeeeee;");
			document.write("color:            #000000;");
			document.write("font-family:      Arial,Helvetica,Sans-Serif;");
			document.write("font-size:        8pt;");
			document.write("font-style:       normal;");
			document.write("font-weight:      normal;");
			document.write("text-align:       center;");
			document.write("vertical-align:   middle;");
		document.write("}");
		document.write(".event {");
	        document.write("background-color: #000000;");
			document.write("border:           1px #000000 solid;");
			document.write("color:            #ffffff;");
			document.write("font-family:      Arial,Helvetica,Sans-Serif;");
			document.write("font-size:        8pt;");
			document.write("font-style:       normal;");
			document.write("font-weight:      normal;");
			document.write("text-align:       center;");
			document.write("vertical-align:   middle;");
		document.write("}");
		document.write(".norm {");
	        document.write("color:           #000000;");
			document.write("font-style:      normal;");
			document.write("text-decoration: none;");
			document.write("cursor:          hand;");
		document.write("}");
		document.write(".arrownorm {");
	        document.write("color:           #ffffff;");
			document.write("font-style:      normal;");
			document.write("text-decoration: none;");
			document.write("cursor:          hand;");
		document.write("}");
		document.write(".inner {");
	        document.write("border-top:       #ffffff 1px solid;");
			document.write("border-bottom:    #777777 1px solid;");
			document.write("border-left:      #ffffff 1px solid;");
			document.write("border-right:     #777777 1px solid;");
		document.write("</STYLE>");
	
		document.write("<DIV ID='Calendar' STYLE='position: absolute;'>");
		document.write("</DIV>");
		removeObj('Calendar');
	}
}

//function to show the Calendar
function showCalendar(myMonth, myYear, myTextBoxID, myLeft,myTop)
{
	if (bIsSeen) {
		hideCalendar();
		bIsSeen=false;
	}
	else {
		makeCalendar(myMonth, myYear, myTextBoxID, myLeft,myTop);
		bIsSeen=true;	
	}
}

//function that makes the Calendar
function makeCalendar(myMonth, myYear, myTextBoxID, myLeft,myTop)
{
//Account for rollover of year and NaN
		

		
		if (isNaN(myMonth) || isNaN(myYear)) {
			myDate = new Date();
			Year = myDate.getFullYear();
			Month= myDate.getMonth()+1;
			
		}
		else if (myMonth==0) {
			Month=12;
			Year=myYear-1;
		}
		else if (myMonth==13) {
			Month=1;
			Year=myYear+1;
		}
		else {
			Month=myMonth;
			Year=myYear;
		}


		
		if (Year < 20) {
			Year = 2000 + Year;

		}
		else if (Year <100) {
			Year = 1900 + Year;
		}

	
		
		//Get the text box object
		textBoxID=getObj(myTextBoxID);
	
		smonth=asDescriptors[Month];
	   
		//start the table here
		var sTableHtml;
		sTableHtml="<table  ID='CalendarTable' class=inner bgcolor='#eeeeee' cellpadding=1 cellspacing=0>";
			sTableHtml=sTableHtml+"<tr><td class=arrow onClick=makeCalendar("+(Month-1)+","+Year+",'"+myTextBoxID+"',"+myLeft+","+myTop+");><IMG SRC='Images/arrow_left_cal.gif'></td>";
			sTableHtml=sTableHtml+"<td colspan=5 align=center class=event>" + smonth + " " + Year + "</td>";
			sTableHtml=sTableHtml+"<td class=arrow onClick=makeCalendar("+(Month+1)+","+Year+",'"+myTextBoxID+"',"+myLeft+","+myTop+");><IMG SRC='Images/arrow_right_cal.gif'></td></tr>";
			sTableHtml=sTableHtml+"<tr>";
			for (i=13; i<20; i++)
				sTableHtml=sTableHtml+"<td class=weekday>" + asDescriptors[i] + "</td>";
			sTableHtml=sTableHtml+"</tr><tr>";
			
		//spaces to leave at the beginning	 
		firstOfMonth = new Date (Year, Month-1, 1);
		leave = firstOfMonth.getDay();
		for (i=0; i<leave; i++)
			sTableHtml=sTableHtml+"<td>&nbsp;</td>";
		
		//Get total number of days in a month
		totalDays = getDaysOfMonth (Month, Year);
	  
		for (day=1; day<=totalDays; day++)	{
			sTableHtml=sTableHtml+"<td class=day onClick=fillTextBox(" + day +",'"+myTextBoxID+"');><span class=norm>&nbsp;" + day + "&nbsp;</span></td>";
			i++;
			if (i % 7 == 0) sTableHtml=sTableHtml+"</tr><tr>";
		}
		
		//we need empty cells in the last row for Netscape
		if (i%7 !=0) {
			var remainder= 7-i%7;
			for (day=1;day<=remainder;day++)
				sTableHtml=sTableHtml+"<td>&nbsp;</td>";
		}
		sTableHtml=sTableHtml+"</tr></table>";
		
		//Display DIV
		displayObj('Calendar');
		
		//write this table to the DIV tag
		writeToDiv(sTableHtml,'Calendar');
		
		//Move table to the required place
		moveObjTo('Calendar', myLeft, myTop);
		
		//hide all select boxes
		showHidePulldowns(false, myLeft, myTop);
		
		//Show Calendar
		showObj('CalendarTable');
		
}

//fills the text box with the data and hides the calendar
function fillTextBox(myDate,myTextBoxID)
{
	//we need to format the date according to the localized date format
	var sDateString = sDateFormat;
	//Year
	if (sDateFormat.search("YYYY") == -1)	{
		Year = Year%100
		if (Year<10) Year = "0" + Year;
		sDateString=sDateString.replace("YY",Year);
	}
	else {
		sDateString=sDateString.replace("YYYY",Year);
	}
	//Date
	if (sDateFormat.search("DD") == -1)	{
		sDateString=sDateString.replace("D",myDate);
	}	
	else {
		if (myDate >= 10) sDateString=sDateString.replace("DD",myDate);
		else			  sDateString=sDateString.replace("DD","0"+myDate);
	}
	//Month
	if (sDateFormat.search("MM") == -1)	{
		sDateString=sDateString.replace("M",Month);	
	}	
	else {
		if (Month >= 10) sDateString=sDateString.replace("MM",Month);
		else			  sDateString=sDateString.replace("MM","0"+Month);
	}
	
	textBoxID=getObj(myTextBoxID);
	
	//Depending on the operator append or not...
	if ((sOperator == "M22")	||  (sOperator == "M57")) {
		if (textBoxID.value.length < 4 ) textBoxID.value= sDateString;
		else textBoxID.value = textBoxID.value + ";" + sDateString;
	}
	else if ((sOperator == "M17") ||  (sOperator == "M44")) {
		if (textBoxID.value.length < 4 ) textBoxID.value= sDateString;
		else {
			var lSeparatorPosition = textBoxID.value.indexOf(";");
			if (lSeparatorPosition > 0 ) {
				textBoxID.value = textBoxID.value.substr(0, lSeparatorPosition+1) + sDateString;
			}
			else {
				textBoxID.value= textBoxID.value + ";" + sDateString;
			}
		}
	}
	else {
		textBoxID.value= sDateString;
	}
	hideCalendar();
	bIsSeen=false;
}


// function to hide the Calendar
function hideCalendar()
{
	hideObj('CalendarTable');
	hideObj('Calendar');
	showHidePulldowns(true, 20, 20);
}

//function to get the month from the text box
function getMonth(myTextBoxID)
{
	var textBox = getObj(myTextBoxID);
	var dateString = textBox.value;
	if (dateString.search(";")!=-1) {
		dateString = dateString.substr(0,dateString.search(";"));
	}
	
	//sDateFormat = getDateFormat(dateString);
	var iPosMonth=sDateFormat.search("M");
	if (dateString.length ==6)	{
		return (parseInt(dateString.substr(iPosMonth,1)));
	}
	else {
		//need to account for bug in parseInt
		//parseInt returns 0 for "08" and "09" :-(
		if (dateString.substr(iPosMonth,1)=="0") {
			return (parseInt(dateString.substr(iPosMonth+1,1)));
		}
		else {
		return (parseInt(dateString.substr(iPosMonth,2)));
		}
	}
}

//function to get the year from the text box
function getYear(myTextBoxID)
{
	var textBox = getObj(myTextBoxID);
	var dateString = textBox.value;
	var iPosYear = sDateFormat.search("YY");


	if ((sDateFormat.length - dateString.length) == 1) {
		iPosYear = iPosYear - 1;
	}
	if ((sDateFormat.length - dateString.length) == 2) {
		iPosYear = iPosYear - 2;
	}
	
	
	if (sDateFormat.length ==10) {
		if (dateString.length ==9) return  (parseInt(dateString.substr(iPosYear-1,4)));
		else return (parseInt(dateString.substr(iPosYear,4), 10));
	}
	else {
		if  ((dateString.length - sDateFormat.length) == 2) {
			return (parseInt(dateString.substr(iPosYear,4), 10));
		}
		else if  ((dateString.length - sDateFormat.length) == 1)  {
			return (parseInt(dateString.substr(iPosYear-1,4), 10));
		}
		else {
			return (parseInt(dateString.substr(iPosYear,2), 10));
		}
	}
}


//function to hide or show the calendar button
//mySelectForm is a string for a select input
//myselectForm is an object for a radio input
function showOrHideCalendarButton(mySelectForm, myButton)
{
	var oRadio = null;
	var data_type = null;
	
	if (typeof(mySelectForm)=="string")	{
		var oSelect = getObj(mySelectForm);
		if (oSelect.selectedIndex >= 0)	{
			data_type = oSelect.options[oSelect.selectedIndex].getAttribute("DATATYPE");
		}
		else {
			data_type=-1
		}
	}
	
	if ((parseInt(data_type) == 16) || (parseInt(data_type) == 15) || (parseInt(data_type) == 14))	{
		showObj(myButton);
	}
	else {	
		hideObj(myButton);	
	}		
}

//function to hide or show the calendar button for a radio input.
function showOrHideCalendarButtonForRadio(sFormID, myButton) {
	var oRadio = null;
	var data_type = -1;
	var myForm = getObj(sFormID);
	for (i=0;i<myForm.elements.length;i++) {
		var e = myForm.elements[i];
		if (e.type == "radio") {
			if (e.checked) {
				oRadio = myForm.elements[i];
				data_type = oRadio.getAttribute("DATATYPE");			
			}
		}
	}
	
	if ((parseInt(data_type) == 16) || (parseInt(data_type) == 15) || (parseInt(data_type) == 14))	{
		showObj(myButton);
	}
	else {	
		hideObj(myButton);	
	}		

}


//function to append to the list box or not....
function updateOperator(mySelectForm)
{
	var oSelect = getObj(mySelectForm);
	sOperator = oSelect.options[oSelect.selectedIndex].value;

	if (sOperator.charAt(0) != 'M')
	{
		sOperator = "M" + sOperator;
	}
}

function showHidePulldowns(bShow, lX, lY) {
//*********************************************************************************************
//Purpose: Hide/Show all the select boxes that come in the way of the calendar
//Inputs:  
//Outputs: 
//*********************************************************************************************
	var oPulldowns = document.getElementsByTagName('SELECT');
	if (oPulldowns) {
		var lPulldowns = 0;
		lPulldowns = oPulldowns.length;
		if (bShow) {
			for (i=0; i < lPulldowns; i++) {
					getObjStyle(oPulldowns.item(i)).visibility = 'visible';
					//oPulldowns.item(i).style.visibility = 'visible';
			}
		} else {
			var oMenu = getObj('Calendar');
			var lMenuWidth = getObjWidth(oMenu);	//.offsetWidth;
			var lMenuHeight = getObjHeight(oMenu);	//.offsetHeight;
			for (i=0; i < lPulldowns; i++) {
					var lLeft = getObjSumLeft(oPulldowns.item(i));
					var lTop = getObjSumTop(oPulldowns.item(i));
					var lWidth = getObjWidth(oPulldowns.item(i));	//.offsetWidth;
					var lHeight = getObjHeight(oPulldowns.item(i));	//.offsetHeight;
					if ((lLeft + lWidth > lX) && (lLeft < lX + lMenuWidth))
						if ((lTop + lHeight > lY) && (lTop < lY + lMenuHeight)) {
							getObjStyle(oPulldowns.item(i)).visibility = 'hidden';
							//oPulldowns.item(i).style.visibility = 'hidden';
						}
			}
		}
	}
	return;
}	