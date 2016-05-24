//-- Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. --//

function getSelectionURL(sURL) {

	sURL = SubmitCheckBoxes(sURL, '', false);
	
	if (lStateID)
		sURL = replaceURLParameter(sURL, 'Index', lStateID);
	if (sView)
		sURL = replaceURLParameter(sURL, 'view', sView);
	if (sMsgID)
		sURL = replaceURLParameter(sURL, 'MsgID', sMsgID);
	sURL = replaceURLValue(sURL, 'Server');
	sURL = replaceURLValue(sURL, 'Project');	

	if (bUseIFrame)
		navIFrame(sURL);
	else
		navWindow(sURL);
}

function replaceURLValue(sURL, sParamater) {
	var lStartPos1, lFinalPos1;
	var sTempStr1, sTempStr2;
	lStartPos1 = sURL.indexOf(sParamater);
	lStartPos1 = sURL.indexOf('=', lStartPos1) + 1;
	lFinalPos1 = sURL.indexOf('=', lStartPos1);
	if(lFinalPos1>lStartPos1)
		sTempStr1 = sURL.substring(lStartPos1, lFinalPos1);
	else
		sTempStr1 = sURL.substring(lStartPos1);
	lFinalPos1 = lStartPos1 + sTempStr1.lastIndexOf('&');
	if(lFinalPos1>lStartPos1)
		sTempStr1 = sURL.substring(lStartPos1, lFinalPos1);
	else
		sTempStr1 = sURL.substring(lStartPos1);
	sTempStr2 = sTempStr1.replace('&', '%26');
	sURL = sURL.replace(sTempStr1, sTempStr2);
	return(sURL);
}

function removeParameterFromURL(sURL, sParameter) {
//******************************************************************************
//Purpose: To remove a parameter from a URL
//Inputs:  sURL, sParameter
//Outputs: A string with the url without the parameter
//******************************************************************************
	var iInitialPos = -1;
	var iFinalPos = -1;
	var iQuestionPos = -1;
	var sTempUpperURL = sURL.toUpperCase();
	var sTempURL = sURL;
	var sSearch = sParameter.toUpperCase() + '=';
	var sLeadChar = '';
	iInitialPos = sTempUpperURL.indexOf(sSearch);
	if (iInitialPos > -1) {
		sLeadChar = sTempUpperURL.substr(iInitialPos - 1, 1);
		if ((sLeadChar != '&') && (sLeadChar != '?')) {
			iInitialPos = sTempUpperURL.indexOf(sSearch, iInitialPos+1);
		}
		if (iInitialPos > -1) {
			iFinalPos = sTempURL.indexOf('&', iInitialPos);
			if (iFinalPos == -1) {
				sLeadChar = sTempURL.substr(iInitialPos - 1, 1);
				if (sLeadChar == '&')
					iInitialPos--;
				sTempURL = sTempURL.substr(0, iInitialPos)
			}
			else {
				iQuestionPos = sTempURL.indexOf('?');
				if (iQuestionPos == (iInitialPos-1))
					sTempURL = sTempURL.substr(0, iInitialPos) + sTempURL.substr(iFinalPos + 1);
				else
					sTempURL = sTempURL.substr(0, iInitialPos - 1) + sTempURL.substr(iFinalPos);
			}
		}
	}
	return sTempURL;
}

function replaceURLParameter(sURL, sFieldToChange, sValueToChange) {
//******************************************************************************
//Purpose: To replace the value of a parameter with the given value
//Inputs:  sURL, sFieldToChange, sValueToChange
//Outputs: A string representing the URL with the new value for the parameter
//******************************************************************************
	sURL = removeParameterFromURL(sURL, sFieldToChange);
	if (sURL.length > 0) {
		if (sURL.substr(-1) == '?')
			sURL += sFieldToChange + '=' + sValueToChange;
		else
			sURL += '&' + sFieldToChange + '=' + sValueToChange;
	}
	else {
		sURL += sFieldToChange + '=' + sValueToChange;
	}
	return sURL;
}

function replaceAllParametersInURL(sURL, sFormName){
//******************************************************************************
//Purpose: get the value of all parameters on the form and replace them on the URL
//Inputs:  sURL, sFormName
//Outputs: A string representing the URL with the new value for the parameters
//******************************************************************************
	var i;	
	var oForm = getObj(sFormName);	
	for(i = 0; i < oForm.elements.length; i++) {		
		var e = oForm.elements[i];
		switch(e.type){
			case "text":
				sURL = replaceURLParameter(sURL, e.name, escape(e.value));
				break;
			case "checkbox":
				if (e.checked)
					sURL = replaceURLParameter(sURL, e.name, 'checked');
				else
					sURL = replaceURLParameter(sURL, e.name, '');
				break;
			case "radio":
				if (e.checked)
					sURL = replaceURLParameter(sURL, e.name, e.value);
				break;
			case "select-one":
				sURL = replaceURLParameter(sURL, e.name, escape(e.value));
				break;
			case "hidden":
				if (!e.getAttribute("NA"))
					sURL = replaceURLParameter(sURL, e.name, escape(e.value));
				break;
		}
	}
	return sURL
}

function moveAnchor(sURL) {
//****************************************************
// Moves an Anchor tag from the middle of the URL to the end
//****************************************************
	var iInitialPos = -1;
	var iFinalPos = -1;
	var sAnchor = '';
	iInitialPos = sURL.indexOf('#');
	if (iInitialPos > -1) {
		iFinalPos = sURL.indexOf('&', iInitialPos);
		if (iFinalPos > -1) {
			sAnchor = sURL.substr(iInitialPos, iFinalPos-iInitialPos);
			sURL = sURL.replace(sAnchor,'') + sAnchor;
		}
	}
	return sURL;
}


function updateComp(sPrefix, sForm, lApplyFlag, lCancelFlag) {
//*************************************************
//Purpose: 
//*************************************************
	var oComp = eval(sForm + '.comp');
	if (oComp) {
		if (aEditorButton[1] == sPrefix) {
			switch (aEditorButton[2]) {
				case 1:
					oComp.value = lApplyFlag;
					break;
				case 3:
					oComp.value = lCancelFlag;
					break;
			}
		}
	}	
}

function navWindow(sNavURL) {
//*************************************************
//Purpose: 
//*************************************************
	window.location = sNavURL;
	return true;
}

function updateHiddenCheckBoxes(form) {
    var i, k;
    var oCheck;
    var oTempCheck;
    
    for(i = 0; i < form.elements.length; i++) {
        oCheck = form.elements(i);
        if (oCheck.name) {
            if((oCheck.name.indexOf('check_') == 0) && (oCheck.type == "hidden")) {
                for(k = 0; k < form.elements(oCheck.name).length; k++) {
                    oTempCheck = form.elements(oCheck.name, k);    
                    if((oTempCheck.type == "checkbox") && (oTempCheck.value == oCheck.value) && !oTempCheck.checked) {
                        oCheck.value = "-2";
                    }
                }
                
            }
        }
    }
}

function SubmitCheckBoxes(sNavURL, sExtraUrl, bHidden)  {

    var oCheck, oTempCheck;
    var oDForm;
    var sTempUrl = '';
    var  i, j, k, urlLength;
    var hiddenElements;
    var bValidHidden;
    
    oDForm = getObj('DrillForm');
    if (!oDForm)
		oDForm = getObj('FilterOnSelections');
    if(oDForm) {
        for(i = 0; i < oDForm.elements.length; i++) {
            bValidHidden = true
            elementsCount = 0;
            oCheck = oDForm.elements[i];
            if (oCheck.name) {
                if((oCheck.name.indexOf('check_') == 0) || (oCheck.name.indexOf('KeepParent') == 0)) {
                    urlLength = (oCheck.name + '=' + oCheck.value).length;
                    j = sNavURL.indexOf(oCheck.name + '=' + oCheck.value);
                    if (j >= 0) {
                        if ((j + urlLength) < sNavURL.length)
                            sNavURL = sNavURL.substring(0, j - 1) + sNavURL.substring(j + urlLength, sNavURL.length);
                        else
                            sNavURL = sNavURL.substring(0, j - 1);
                    }
                    if (oCheck.type == "checkbox" && oCheck.checked) {
                        if(bHidden)
                            sTempUrl = sTempUrl + "<INPUT TYPE='HIDDEN' NAME='" + oCheck.name + "' VALUE='" + oCheck.value + "'/>"
                        else
                            sTempUrl = sTempUrl + '&' + oCheck.name + '=' + oCheck.value;
                    }
                    else if(oCheck.type == "hidden") {
                        hiddenElements = oCheck.value.split(", ");
                        for (l = 0; l < hiddenElements.length; l++) {
                            bValidHidden = true;
                            for(k = 0; k < oDForm.elements(oCheck.name).length; k++) {
                                oTempCheck = oDForm.elements(oCheck.name, k);    
                                if((oTempCheck.type == "checkbox") && (oTempCheck.value == hiddenElements[l]))
                                bValidHidden = false;
                            }
                            if(bValidHidden) {
                                if(bHidden)
                                    sTempUrl = sTempUrl + "<INPUT TYPE='HIDDEN' NAME='" + oCheck.name + "' VALUE='" + hiddenElements[l] + "'/>"
                                else
                                    sTempUrl = sTempUrl + '&' + oCheck.name + '=' + hiddenElements[l];
                            }
                        }
                    }
                }
            }
        }
    }
    return (sNavURL + sTempUrl + sExtraUrl);
}

function formDrillLink(hLink, sDrillParameter) {
//**********************************************************
// Purpose: Form the Drill after the onClick event
//**********************************************************

	hLink.href = 'RebuildReport.asp?' + sDrillParameter + '&' + sBaseURL
}

 function swapPrintMargins(newOrientation) {
  if (currentOrientation == newOrientation) return;
  var frm = document.previewoptionsform;
  if (!frm) return;
  var hg = frm.hg.value;
  var hh = frm.hh.value;
  if (newOrientation == 2) {
   frm.hg.value = frm.hj.value;
   frm.hh.value = frm.hi.value;
   frm.hi.value = hg;
   frm.hj.value = hh;
  } else {
   frm.hg.value = frm.hi.value;
   frm.hh.value = frm.hj.value;
   frm.hi.value = hh;
   frm.hj.value = hg;
  }
  currentOrientation = newOrientation;
  return;
 }