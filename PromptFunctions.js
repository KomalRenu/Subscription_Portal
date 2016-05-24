//-- Copyright 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. --//

	//var bDefault=new Array();

	function gotoAnchor(sCurPin)
	{
		//**********************************************
		//Purpose:  go to a specified anchor
		//Input:     sCurPin
		//Output:
		//**********************************************

		var isNav, isIE

		if ((navigator != null) && (sCurPin > 1)) 
		{
			if ((navigator.appName!=null) && (navigator.appVersion!=null))
			{
				isNav = (navigator.appName == "Netscape");
				isIE  = (navigator.appName.indexOf("Internet Explorer") != -1);
				if (isNav || (isIE && parseInt(navigator.appVersion)>=4 ))
				{
					self.location = "#" + sCurPin;
					//self.navigate "#2"
				}
			}
		}
	}

	function MoveItemsbyListObject(oFromList, oToList)
	{
		//**************************************************************************
		//Purpose:  move selected items from sFromList to sToList referenced by object
		//Input:    sFromList, sToList
		//Output:   sFromList, sToList
		//**************************************************************************
		var oForm = document.PromptForm;
		var i;
		var lLength;
		var lIndexSelected;
		var lNumSelections;
		var aSelArray = new Array();

		// add to right-side
		lLength = oFromList.options.length;
		lNumSelections = 0; 
		
		for (i=0; i<lLength; i++)
		{
			if (oFromList.options[i].selected && oFromList.options[i].value!="-none-")
			{
				if (oToList.options.length==1)	//replace -none- with an item
				{
					if (oToList.options[0].value=="-none-" )
					{
						oToList.options[0] = null;
					}
				}
				var oOption = new Option(oFromList.options[i].text, oFromList.options[i].value, false, false)
				oToList.options[oToList.length] = oOption;
				oToList.options[oToList.length-1].selected = false;
	    		oOption = null;
	    		
	    		lIndexSelected = i;
	    		lNumSelections = lNumSelections + 1;
	    	}
	    }

		//put left-side seleted into a temp array
		for (i=lLength-1;i>=0;i--)
		{
	    	if (oFromList.options[i].selected)
			{
				oFromList.options[i].selected = false;
				aSelArray[aSelArray.length] = i;
			}
		}

		for (i=0; i<aSelArray.length; i++)
		{
			oFromList.options[aSelArray[i]] = null;
		}

		if (oFromList.options.length==0)	 //put -none- when no items
		{
			var oOption = new Option("--" + aDescriptor[18] + "--", "-none-", false, false)
			oFromList.options[oFromList.length] = oOption;
			oFromList.options[oFromList.length-1].selected = false;
			oOption = null;
		}
		
		if ((lIndexSelected > 0) && (oFromList.options.length > 0))
		{			
			if (lIndexSelected >= lNumSelections)
			{
				if (oFromList.options.length > lIndexSelected - lNumSelections + 1)
					oFromList.options[lIndexSelected - lNumSelections + 1].selected = true;
				else
					oFromList.options[oFromList.options.length - 1].selected = true;
			}
			else
			{
				lNumSelections = lNumSelections -1;
				oFromList.options[lIndexSelected - lNumSelections].selected = true;
			}
		}
		else
		{
			oFromList.options[0].selected = true;
		}		
	}

	function AddItemsbyListObjectForObjectPrompt(oFromList, oToList)
	{
		//**************************************************************************
		//Purpose:  move selected items from sFromList to sToList referenced by object
		//Input:    sFromList, sToList
		//Output:   sFromList, sToList
		//**************************************************************************
		var oForm = document.PromptForm;
		var i;
		var j;
		var lLength;
		var aSelArray = new Array();
		var lPosition;
		var lIndexSelected;
		var lNumSelections;
		
		lNumSelections = 0; 

		//get the position of selected item on the ToList
		lPosition = oToList.options.length;
		for (i=0; i<lPosition; i++)
		{
			if (oToList.options[i].selected)
			{
				lPosition = i;
				break;
			}
		}

		// add to right-side
		lLength = oFromList.options.length;
		for (i=0; i<lLength; i++)
		{
			if (oFromList.options[i].selected && oFromList.options[i].value!="-none-")
			{
				if (oToList.options.length==1)	//replace -none- with an item
				{
					if (oToList.options[0].value=="-none-" )
					{
						oToList.options[0] = null;
						if (lPosition > 0)
						{
							lPosition = lPosition - 1;
						}
					}
				}
				//move items after position one down
				for (j=oToList.length-1; j>=lPosition; j--)
				{
					var oOption = new Option(oToList.options[j].text, oToList.options[j].value, false, false);
					oToList.options[j+1] = oOption;
				}

				var oOption = new Option(oFromList.options[i].text, oFromList.options[i].value, false, false);
				oToList.options[lPosition] = oOption;
				oToList.options[lPosition].selected = false;
				lPosition = lPosition + 1;
	    		oOption = null;
	    		
	    		lIndexSelected = i;
	    		lNumSelections = lNumSelections + 1;
	    	}
	    }

		//put left-side seleted into a temp array
		for (i=lLength-1;i>=0;i--)
		{
	    	if (oFromList.options[i].selected)
			{
				oFromList.options[i].selected = false;
				aSelArray[aSelArray.length] = i;
			}
		}

		for (i=0; i<aSelArray.length; i++)
		{
			oFromList.options[aSelArray[i]] = null;
		}

		if (oFromList.options.length==0)	 //put -none- when no items
		{
			var oOption = new Option("--" + aDescriptor[18] + "--", "-none-", false, false)
			oFromList.options[oFromList.length] = oOption;
			oFromList.options[oFromList.length-1].selected = false;
			oOption = null;
		}
		
		if ((lIndexSelected > 0) && (oFromList.options.length > 0))
		{			
			if (lIndexSelected >= lNumSelections)
			{
				if (oFromList.options.length > lIndexSelected - lNumSelections + 1)
					oFromList.options[lIndexSelected - lNumSelections + 1].selected = true;
				else
					oFromList.options[oFromList.options.length - 1].selected = true;
			}
			else
			{
				lNumSelections = lNumSelections -1;
				oFromList.options[lIndexSelected - lNumSelections].selected = true;
			}
		}
		else
		{
			oFromList.options[0].selected = true;
		}
	}

	function BuildUserSelections(lMaxPin) 
	{
		//**************************************************************************
		//Purpose:  collect all the information user selected and put it to a hidden input
		//Input:    lMaxPin
		//Output:   
		//**************************************************************************
		var oForm = document.PromptForm;

		var lLength;
		var i;
		var j;
		var oOptions;
		var oSeletedList;
		var sSelections = '';
		var lSplitCount = 0;
		var bSplit = false;
		var oHiddenInput = new Array();

		var oUserSelections = oForm.UserSelections;

		for (j=1;j<=lMaxPin;j++)
		{
			if (eval('document.PromptForm.Selected_' + parseInt(j)) != null )
			{
				oSeletedList = eval('document.PromptForm.Selected_' + parseInt(j));
				oOptions = oSeletedList.options;
				lLength = oOptions.length;
				if (oOptions[0].value != "-none-") //&& (oOptions[0].value != "-default-"))
				{
				//	oForm.UserSelections.value = oForm.UserSelections.value + parseInt(j) + unescape("%1d") ;
					sSelections = sSelections + parseInt(j) + unescape("%1d") ;
					for(i=0; i<lLength; i++)
					{
						//oForm.UserSelections.value = oForm.UserSelections.value + oOptions[i].value + unescape("%1c") ;
						sSelections = sSelections + oOptions[i].value + unescape("%1c") ;
						if (sSelections.length > 50000) {
							bSplit = true;
							oHiddenInput[lSplitCount] = document.createElement('input');
							oHiddenInput[lSplitCount].type = 'Hidden';
							oHiddenInput[lSplitCount].name = 'split_'+ parseInt(lSplitCount) +'_UserSelections';
							oHiddenInput[lSplitCount].value = sSelections;
							oHiddenInput[lSplitCount].id = 'split_'+ parseInt(lSplitCount) +'_UserSelections';
							oForm.appendChild(oHiddenInput[lSplitCount]);
							sSelections = '';
							lSplitCount ++;
						}
					}
					//oForm.UserSelections.value = oForm.UserSelections.value + unescape("%1d") ;
					sSelections = sSelections + unescape("%1d") ;

				}
			}
		}
		if (bSplit) {
			oHiddenInput[lSplitCount] = document.createElement('input');
			oHiddenInput[lSplitCount].type = 'Hidden';
			oHiddenInput[lSplitCount].name = 'split_'+ parseInt(lSplitCount) +'_UserSelections';
			oHiddenInput[lSplitCount].value = sSelections;
			oHiddenInput[lSplitCount].id = 'split_'+ parseInt(lSplitCount) +'_UserSelections';
			oForm.appendChild(oHiddenInput[lSplitCount]);
			sSelections = '';
			
			oHiddenInput[lSplitCount+1] = document.createElement('input');
			oHiddenInput[lSplitCount+1].type = 'Hidden';
			oHiddenInput[lSplitCount+1].name = 'split';
			oHiddenInput[lSplitCount+1].value = 'UserSelections' + '|' + parseInt(lSplitCount);
			oHiddenInput[lSplitCount+1].id = 'split';
			oForm.appendChild(oHiddenInput[lSplitCount+1]);
			
		}
		else {
			oUserSelections.value = sSelections;
		}

		return(true);
	}

	function GetCurrentAttributeValue(oDropdownList)
	{
		//**************************************************************************
		//Purpose:  get current attribute from dropdown box for HI prompt
		//Input:    oDropdownList
		//Output:   the value of current attribute item
		//**************************************************************************
		var sValue;
		var re = "\x1e";
		var aCurrent= new Array();
		var lLength = oDropdownList.options.length;
		for(var i=0; i<lLength; i++)
		{
			sValue = oDropdownList.options[i].value;
			aCurrent = sValue.split(re);

			if (aCurrent[2] == "1")
			{
				return(aCurrent[0]);
			}
		}
		return(0);
	}

	function GetValue(sValue, re)
	{
		//**************************************************************************
		//Purpose:  get current attribute from dropdown box for HI prompt
		//Input:    oDropdownList
		//Output:   the value of current attribute item
		//**************************************************************************
		var aCurrent= new Array();
		aCurrent = sValue.split(re);
		return(aCurrent[0]);
	}

	function GetCurrentAttributeText(oDropdownList)
	{
		//**************************************************************************
		//Purpose:  get current attribute from dropdown box for HI prompt
		//Input:    oDropdownList
		//Output:   the value of current attribute item
		//**************************************************************************
		var sValue;
		var sText;
		var re = "\x1e";
		var aCurrent= new Array();
		var lLength = oDropdownList.options.length;
		for(var i=0; i<lLength; i++)
		{
			sValue = oDropdownList.options[i].value;
			aCurrent = sValue.split(re);
			sText = aCurrent[1];
			
			if (aCurrent[2] == "1")
			{
				return(sText);
			}
		}
		return(0);
	}

	function CheckAttributeExist(oToList, oDropdownList)
	{
		//**************************************************************************
		//Purpose:  check whether current attribute is in oToList for HI prompt
		//Input:    oToList, oDropdownList
		//Output:   if exist, index of oToList; if not, -1.
		//**************************************************************************
		var reAttribute = "\x1e";
		var aAttribute = new Array();
		var sCurrentAttributeValue = GetCurrentAttributeValue(oDropdownList);
		var sCurrentAttributeText = GetCurrentAttributeText(oDropdownList);
		var lLength = oToList.options.length;
		for(var i=0; i<lLength; i++)
		{
			aAttribute = oToList.options[i].value.split(reAttribute);
			//if (aAttribute[0] == sCurrentAttributeValue && aAttribute.length == 3)
			if (aAttribute[0] == sCurrentAttributeValue && aAttribute[1] == sCurrentAttributeText)
			{
				return(i+1);
			}
		}
		return(-1);
	}

	function CheckAttributeElementExpression(sOptionValue)
	{
		//**************************************************************************
		//Purpose:  check whether an <OPTION> item is attribute or element or expression for HI prompt
		//Input:    sOptionValue, oDropdownList
		//Output:   expression, 2; attribute, 1; element, 0; else -1.
		//**************************************************************************
		var sValue
		var reAttributeElement = ":";
		var reExpression = "\x1b";
		var reAttribute = "\x1e";
		var aCurrent= new Array();
		var aAttribute = new Array();
		var aElement = new Array();

		aCurrent = sOptionValue.split(reExpression);
		if (aCurrent.length > 1)
		{return(2);}

		aAttribute = aCurrent[0].split(reAttribute);
		aElement = aAttribute[0].split(reAttributeElement);
		if (aElement.length > 1)
		{return(0);}
		else if (aCurrent[0] == "-default-" || aCurrent[0] == "-none-")
		{return(-1);}
		else
		{return(1);}
	}

	function AddItemsbyListObjectForHI(oFromList, oToList, oDropdownList, sPin)
	{
		//**************************************************************************
		//Purpose:  add selected items from sFromList to sToList referenced by object
		//Input:    sFromList, sToList, oDropdownList, sPin
		//Output:   sFromList, sToList
		//**************************************************************************
		var oOption;
		var sCurrentAttributeValue;
		var sCurrentAttributeText;
		var aCurrent= new Array();
		var re = "\x1e";
		var aSelArray = new Array();
		var sTemp = "";
		var lPos;
		var lLength
		var lResult;
		var lPin = parseInt(sPin);
		var lIndexSelected;
		var lNumSelections;

		lNumSelections = 0;

		// add to right-side
		if (oFromList)
		{
			lLength = oFromList.options.length;
			lPos= CheckAttributeExist(oToList, oDropdownList);
			sCurrentAttributeValue = GetCurrentAttributeValue(oDropdownList);
			sCurrentAttributeText = GetCurrentAttributeText(oDropdownList);
			for (i=0; i<lLength; i++)
			{
				if (oFromList.options[i].selected && oFromList.options[i].value!="-none-")
				{
					if ((oToList.options.length==1) && ((oToList.options[0].value=="-none-" ) || (oToList.options[0].value=="-default-" )))	//replace -none- with an item
					{
						if (oToList.options[0].value=="-default-" )
						{bDefault[lPin] = true;}
						else
						{bDefault[lPin] = false;}
						oToList.options[0] = null;
					}

					if ( lPos == -1 )
					{
						aCurrent = sCurrentAttributeValue.split(re);
						oOption = new Option(sCurrentAttributeText + ":", sCurrentAttributeValue+unescape("%1e")+sCurrentAttributeText, false, false);
						oToList.options[oToList.length] = oOption;
						oToList.options[oToList.length-1].selected = false;
						oOption = null;
					}
					var oOption = new Option(unescape("%a0")+unescape("%a0")+unescape("%a0")+oFromList.options[i].text, oFromList.options[i].value, false, false)
					lPos = AddItem(sPin, oOption, lPos, oToList);
					lPos += 1; 
					oOption = null;
					
					lIndexSelected = i;
	    			lNumSelections = lNumSelections + 1;
				}
			}

			//put left-side seleted into a temp array
			for (i=lLength-1;i>=0;i--)
			{
				if (oFromList.options[i].selected)
				{
					oFromList.options[i].selected = false;
					aSelArray[aSelArray.length] = i;
				}
			}

			for (i=0; i<aSelArray.length; i++)
			{
				oFromList.options[aSelArray[i]] = null;
			}

			if (oFromList.options.length==0)	 //put -none- when no items
			{
				var oOption = new Option("--" + aDescriptor[18] + "--", "-none-", false, false)
				oFromList.options[oFromList.length] = oOption;
				oFromList.options[oFromList.length-1].selected = false;
				oOption = null;
			}
			aSelArray = null;
			CleanList(oToList);

			if ((lIndexSelected > 0) && (oFromList.options.length > 0))
			{				
				if (lIndexSelected >= lNumSelections)
				{
					if (oFromList.options.length > lIndexSelected - lNumSelections + 1)
						oFromList.options[lIndexSelected - lNumSelections + 1].selected = true;
					else
						oFromList.options[oFromList.options.length - 1].selected = true;
				}
				else
				{
					lNumSelections = lNumSelections -1;
					oFromList.options[lIndexSelected - lNumSelections].selected = true;
				}
			}
			else
			{
				oFromList.options[0].selected = true;
			}

			lLength = oToList.options.length;
			lResult = 0;
			for (i=0;i<lLength;i++)
			{
				sTemp = oToList.options[i].value;
				if (CheckAttributeElementExpression(sTemp)!= 0)
				{lResult += 1;}
			}

			if (lResult > 1)
				displayObj('ANDOR_' + sPin);
		}
	}

	function RemoveItems(oFromList)
	{
		//**************************************************************************
		//Purpose:  remove selected items from sFromList
		//Input:    oFromList
		//Output:   oFromList
		//**************************************************************************
		var aSelArray = new Array();
		var lLength = oFromList.options.length;
		//put oFromList seleted into a temp array
		for (var i=lLength-1;i>=0;i--)
		{
			if (oFromList.options[i].selected)
			{
				oFromList.options[i].selected = false;
				aSelArray[aSelArray.length] = i;
			}
		}

		for (i=0; i<aSelArray.length; i++)
		{
			for(var j=aSelArray[i]; j<oFromList.options.length-1; j++)
			{
				oFromList.options[j].value = oFromList.options[j+1].value;
				oFromList.options[j].text = oFromList.options[j+1].text;
			}
			oFromList.options.length -= 1;
		}

		if (oFromList.options.length==0)	 //put -none- when no items
		{
			var oOption = new Option("--" + aDescriptor[18] + "--", "-none-", false, false)
			oFromList.options[oFromList.length] = oOption;
			oFromList.options[oFromList.length-1].selected = false;
			oOption = null;
		}
		aSelArray = null;
	}

	function RemoveItemsByPin(oFromList, sPin)
	{
		//**************************************************************************
		//Purpose:  remove selected items from sFromList
		//Input:    oFromList
		//Output:   oFromList
		//**************************************************************************
		var aSelArray = new Array();
		var lLength = oFromList.options.length;
		var lPin = parseInt(sPin);
		
		//put oFromList seleted into a temp array
		for (var i=lLength-1;i>=0;i--)
		{
			if (oFromList.options[i].selected)
			{
				oFromList.options[i].selected = false;
				aSelArray[aSelArray.length] = i;
			}
		}

		for (i=0; i<aSelArray.length; i++)
		{
			for(var j=aSelArray[i]; j<oFromList.options.length-1; j++)
			{
				oFromList.options[j].value = oFromList.options[j+1].value;
				oFromList.options[j].text = oFromList.options[j+1].text;
			}
			oFromList.options.length -= 1;
		}

		if (oFromList.options.length==0)	 //put -none- or (default) when no items
		{
			//if (document.all("default_" + sPin).value == '1' || bDefault[lPin])
			if (bDefault[lPin])
			{var oOption = new Option(aDescriptor[19], "-default-", false, false)}
			else
			{var oOption = new Option("--" + aDescriptor[18] + "--", "-none-", false, false)}
			oFromList.options[oFromList.length] = oOption;
			oFromList.options[oFromList.length-1].selected = false;
			oOption = null;
		}
		aSelArray = null;
	}

	function RemoveItemsbyListObjectForHIInList(oFromList, oToList, oDropdownList, sPin)
	{
		//**************************************************************************
		//Purpose:  remove selected items from sFromList to sToList referenced by object
		//Input:    oFromList, oToList, oDropdownList, sPin
		//Output:   oFromList, oToList
		//**************************************************************************
		var sAttributeValue;
		var sElementValue;
		var sCurrentAttributeValue = "";
		var lLength = oFromList.options.length;
		var sText;
		var aCurrent = new Array();
		var sValue;
		var lResult;
		var sTemp;
		var bAllDimension = false;
		var reAttribute = "\x1e";

		if (oDropdownList)
			{sCurrentAttributeValue = GetCurrentAttributeValue(oDropdownList);}
		else
			{bAllDimension = true;}

		for (var lAttributeIndex=0; lAttributeIndex<lLength; lAttributeIndex++)
		{
			sAttributeValue = oFromList.options[lAttributeIndex].value;
			if (CheckAttributeElementExpression(sAttributeValue) == 1)
			{
				sAttributeValue = GetValue(sAttributeValue, reAttribute);
				var lElementIndex=lAttributeIndex+1;
				sElementValue = GetValue(oFromList.options[lElementIndex].value, reAttribute);
				bAttributeSelected = true;
				while ((lElementIndex<lLength) && (CheckAttributeElementExpression(sElementValue) == 0))
				{
					if (oFromList.options[lAttributeIndex].selected)
					{
						oFromList.options[lElementIndex].selected = true;
						if (sAttributeValue==sCurrentAttributeValue || bAllDimension)
						{
							sText = oFromList.options[lElementIndex].text;
							var oElementOption = new Option(sText.substring(3,sText.length), oFromList.options[lElementIndex].value, false, false);
							AddItem(sPin, oElementOption, -1, oToList);
						}
					}
					else
					{
						if (oFromList.options[lElementIndex].selected == true)
						{
							if (sAttributeValue==sCurrentAttributeValue || bAllDimension)
							{
								sText = oFromList.options[lElementIndex].text;
								var oElementOption = new Option(sText.substring(3,sText.length), oFromList.options[lElementIndex].value, false, false);
								AddItem(sPin, oElementOption, -1, oToList);
							}
						}
						else
						{
							bAttributeSelected = false;
						}
					}
					lElementIndex += 1;
					if (lElementIndex < lLength)
					{sElementValue = GetValue(oFromList.options[lElementIndex].value, reAttribute);}
				}
				oFromList.options[lAttributeIndex].selected = bAttributeSelected;
				lAttributeIndex = lElementIndex - 1;
			}
		}
		RemoveItemsByPin(oFromList, sPin);
		CleanList(oFromList);
		lLength = oFromList.options.length;
		lResult = 0;
		for (i=0;i<lLength;i++)
		{
			sTemp = oFromList.options[i].value;
			if (CheckAttributeElementExpression(sTemp)!= 0)
			{lResult += 1;}
		}

		if (lResult < 2)
			removeObj('ANDOR_' + sPin);
	}

	function RemoveItemsbyListObjectForHIInQual(oFromList, oOperatorList, oDropdownList, sPin)
	{
		//**************************************************************************
		//Purpose:  remove selected items from sFromList to sToList referenced by object
		//Input:    oFromList, oToList, oDropdownList, sPin
		//Output:   oFromList, oToList
		//**************************************************************************
		var sAttributeValue;
		var sElementValue;
		var sCurrentAttributeValue;
		var lLength = oFromList.options.length;
		var sText;
		var aCurrent = new Array();
		var sValue;
		var lResult;
		var sTemp;
		var lAttributeIndex;
		var lElementIndex;
		var sFirstValue = "";
		var lType;
		var re = "\x1b";
		var reAttribute = "\x1e";

		for (lAttributeIndex=0; lAttributeIndex<lLength; lAttributeIndex++)
		{
			sAttributeValue = oFromList.options[lAttributeIndex].value;
			lType = CheckAttributeElementExpression(sAttributeValue);
			if ( lType == 1)
			{
				sAttributeValue = GetValue(sAttributeValue, reAttribute);
				lElementIndex=lAttributeIndex+1;
				sElementValue = GetValue(oFromList.options[lElementIndex].value, reAttribute);
				bAttributeSelected = true;
				while ((lElementIndex<lLength) && (CheckAttributeElementExpression(sElementValue) == 0))
				{
					if (oFromList.options[lAttributeIndex].selected)
					{
						oFromList.options[lElementIndex].selected = true;
					}
					else
					{
						if (oFromList.options[lElementIndex].selected == false)
						{
							bAttributeSelected = false;
						}
					}
					lElementIndex += 1;
					if (lElementIndex < lLength)
					{sElementValue = GetValue(oFromList.options[lElementIndex].value, reAttribute);}
				}
				oFromList.options[lAttributeIndex].selected = bAttributeSelected;
				lAttributeIndex = lElementIndex - 1;
			}
			if ((lType == 2) && (sFirstValue == ""))
			{sFirstValue = sAttributeValue;}
		}
		RemoveItemsByPin(oFromList, sPin);
		CleanList(oFromList)

		lLength = oFromList.options.length;
		lResult = 0;
		for (i=0;i<lLength;i++)
		{
			sTemp = oFromList.options[i].value;
			if (CheckAttributeElementExpression(sTemp)!= 0)
			{lResult += 1;}
		}

		if (lResult < 2)
			removeObj('ANDOR_' + sPin);

		aCurrent = sFirstValue.split(re);
		lLength = oDropdownList.options.length;
		for (i=0;i<lLength;i++)
		{
			if (oDropdownList.options[i].value == aCurrent[0])
			{
				oDropdownList.options[i].selected = true;
				break;
			}
		}

		lLength = oOperatorList.options.length;
		for (i=0;i<lLength;i++)
		{
			if (oOperatorList.options[i].value == aCurrent[1])
			{
				oOperatorList.options[i].selected = true;
				break;
			}
		}
	}

	function AddItem(sPin, oOption, lPos, oToList)
	{
		//**************************************************************************
		//Purpose:  add a single <OPTION> item to sToList at index lPOS 
		//Input:    oOption, oToList, lPos
		//Output:   oToList
		//**************************************************************************
		var lLength;
		var lPin = parseInt(sPin);
		
		if (oToList)
		{
			if ((oToList.options.length==1) && ((oToList.options[0].value=="-none-" ) || (oToList.options[0].value=="-default-" )))	//replace -none- with an item
				{
					if (oToList.options[0].value=="-default-" )
						{bDefault[lPin] = true;}
					else
						{bDefault[lPin] = false;}
					oToList.options[0] = null;
				}
			lLength = oToList.options.length;
			if (lPos == -1)
			{
				lPos = lLength;
			}
			for(var i = lLength; i > lPos; i--)
			{
				var oTempOption = new Option(oToList.options[i-1].text, oToList.options[i-1].value, false, false)
				var sDataType1 = oToList.options[i-1].getAttribute("DATATYPE");
				if (!(sDataType1==null)) {
					oTempOption.setAttribute("DATATYPE", sDataType1); 
				}
				oToList.options[i] = oTempOption;
				oToList.options[i].selected = false;
				oTempOption = null;
			}
			//oToList.options.length += 1;
			oToList.options[lPos]= oOption;
			oToList.options[lPos].selected = false;
		}
		return(lPos);
	}

	function AddItemsbyListObjectForExpression(oFromList, oDropdownList, oTextbox, oToList, sPin)
	{
		//**************************************************************************
		//Purpose:  add selected items from sFromList to sToList referenced by object
		//Input:    sFromList, sToList, oDropdownList, sPin, oTextbox
		//Output:   sFromList, sToList, oTextbox
		//**************************************************************************
		var oOption;
		var sCurrentAttributeValueText;
		var sCurrentAttributeValueValue;
		var sCurrentOperator;
		var sInput;
		var sValue;
		var sText;
		var lLength;
		var lInputArrSize;

		sCurrentAttributeValueText = GetSelectedItemText(oFromList);
		sCurrentAttributeValueValue = GetSelectedItemValue(oFromList);
		sCurrentOperator = GetSelectedItemValue(oDropdownList);
		sInput = oTextbox.value;
		sInput = CleanInput(sInput);
		
		lInputArrSize = sInput.split(";").length
		
		if (sCurrentAttributeValueValue == "0")
		{
			DisplayError(aDescriptor[16], aDescriptor[15], sPin);
		}
		else if (sInput == "")
		{
			DisplayError(aDescriptor[16], aDescriptor[12], sPin);
		}
		else if (sCurrentAttributeValueValue != "0"  && sCurrentAttributeValueValue != "1" && sInput != "")
		{
			if (sCurrentOperator == "P1" || sCurrentOperator == "P2")
				{sInput = CleanPercentageSign(sInput) + "%";}
			sValue = BuildSelectValueForExpression(sCurrentAttributeValueValue, sCurrentOperator, sInput);
			sText = BuildSelectTextForExpression(sCurrentAttributeValueText, sCurrentOperator, sInput);

			if (sText == "9")
			{
				DisplayError(aDescriptor[16], aDescriptor[9], sPin);
			}
			else if (sText == "10")
			{
				DisplayError(aDescriptor[16], aDescriptor[10], sPin);
			}
			else if ((sCurrentOperator == "M8" || sCurrentOperator == "M9" || sCurrentOperator == "M10" || sCurrentOperator == "M11" || sCurrentOperator == "M6") 
					&& (lInputArrSize > 1))
			{
				DisplayError(aDescriptor[16], aDescriptor[15], sPin);
				oTextbox.value = "";
			}
			else if (sText != "0")
			{
				oOption = new Option(sText, sValue, false, false);
				AddItem(sPin, oOption, -1, oToList);
				oTextbox.value = "";
				ClearError(sPin);
			}
		}

		lLength = oToList.options.length;
		if (lLength > 1)
		{
			displayObj("ANDOR_" + sPin);
		}
	}

	function RemoveItemsbyListObjectForExpression(oFromList, oDropdownList, oToList, sPin)
	{
		//**************************************************************************
		//Purpose:  add selected items from sFromList to sToList referenced by object
		//Input:    sFromList, sToList, oDropdownList, sPin
		//Output:   sFromList, sToList
		//**************************************************************************
		var oOption;
		var sCurrentAttributeValue;
		var sCurrentOperator;
		var sInput;
		var aCurrent= new Array();
		var aSelArray = new Array();
		var re = "\x1b";
		var reMetric = "\x1e";
		var sFirstValue;
		var sText;
		var	lLength = oFromList.options.length;
		var lPin = parseInt(sPin);
		var bInBlock;
		var aMetric;
		var aTempMetric;
		var sTempMetric;
		var sText;

		//put left-side seleted into a temp array
		for (var i=lLength-1;i>=0;i--)
		{
	    	if (oFromList.options[i].selected)
			{
				oFromList.options[i].selected = false;
				aSelArray[aSelArray.length] = i;
			}
		}

		if (aSelArray.length != 0)
		{
			
			sFirstValue = oFromList.options[aSelArray[aSelArray.length -1]].value;

			if ((aSelArray.length = 1) && (sFirstValue == "-none-"))
			{	
				return;
			}

			for (i=0; i<aSelArray.length; i++)
			{	
				if (oFromList.options[aSelArray[i]].value=="-default-")
					{bDefault[lPin]=true;}
				else
					{oFromList.options[aSelArray[i]] = null;}
			}

			if (oFromList.options.length==0)	 //put -none- when no items
			{
				if (bDefault[lPin])
					{var oOption = new Option(aDescriptor[19], "-default-", false, false)}
				else
					{var oOption = new Option("--" + aDescriptor[18] + "--", "-none-", false, false)}
				oFromList.options[oFromList.length] = oOption;
				oFromList.options[oFromList.length-1].selected = false;
				oOption = null;
			}

			lLength = oFromList.options.length;
			if (lLength < 2)
				removeObj('ANDOR_' + sPin);

			aCurrent = sFirstValue.split(re);
			aMetric = aCurrent[0].split(reMetric);
			lLength = oToList.options.length;
			for (i=0;i<lLength;i++)
			{
				bInBlock = false;
				sTempMetric = oToList.options[i].value;
				aTempMetric = sTempMetric.split(reMetric);

				if ((aTempMetric[0] == aMetric[0]) && (aMetric.length > 1) && (aTempMetric[1] == aMetric[1]))
				{
					oToList.options[i].selected = true;
					bInBlock = true;
					break;
				}
			}

			if ((!bInBlock) && (aCurrent[0] != "-default-"))
			{
				if (aMetric.length > 2)
					{sText = aMetric[1] + "(" + aMetric[3] + ")";}
				else
					{sText = aMetric[1];}
				
				oOption = new Option(sText, aCurrent[0], false, false);
				AddItem(sPin, oOption, 0, oToList);
				oToList.options[0].selected = true;
			}

			lLength = oDropdownList.options.length;
			for (i=0;i<lLength;i++)
			{
				if (oDropdownList.options[i].value == aCurrent[1])
				{
					oDropdownList.options[i].selected = true;
					break;
				}
			}
		}
	}

	function BuildSelectValueForExpression(sObject, sOperator, sInput)
	{
		//**************************************************************************
		//Purpose:  Generate <option> value of a select item for expression prompt
		//Input:    sObject, sOperator, sInput
		//Output:   <option> value of a select item for expression prompt
		//**************************************************************************
		var sValue;
		var re = "\x1e";
		var aCurrent= new Array();
		var reInput

		if ((sOperator == "M22") ||  (sOperator == "M57")) 	//replace , with ;
		{
			var reInput = new RegExp(",", "g");
			sInput = sInput.replace(reInput, ";");
		}
		sValue = sObject + unescape("%1b") + sOperator + unescape("%1b") + sInput;
		return(sValue);
	}

	function FindDelimeter(sValue)
	{
		//*********************************************************************************
		//Purpose:  find delimeter from sValue, delimeter search order is " ", ",", ";" 
		//Input:    sObject, sOperator, sInput
		//Output:   <option> value of a select item for expression prompt
		//*********************************************************************************
		var aResult;
		aResult = sValue.match(";");
		if (aResult != null)
		{return(";")}
		else
		{
			aResult = sValue.match(",");
			if (aResult != null)
				{return(",")}
			else
				{return("")}
		}
	}

	function BuildSelectTextForExpression(sObjectText, sOperator, sInput)
	{
		//**************************************************************************
		//Purpose:  Generate <option> value of a select item for expression prompt
		//Input:    sObject, sOperator, sInput
		//Output:   <option> value of a select item for expression prompt
		//**************************************************************************
		var sValue;
		var re = "\x1e";
		var aCurrent= new Array();
		var sTemp;
		var lLength;
		var sOperatorSymbol = GetOperatorSymbol(sOperator);
		var sDelimeter = FindDelimeter(sInput)

		//aCurrent = sObject.split(re);
		//if (aCurrent.length > 2)
		//{
		//	//attribute qulification
		//	sValue = aCurrent[0] + "(" + aCurrent[2] + ")";
		//}
		//else
		//{
		//	//metric qulification
		//	sValue = aCurrent[0];
		//}

		sValue = sObjectText

		if (sOperator == "M17")
		{
			if (sDelimeter == ";")
			{aCurrent = sInput.split(sDelimeter);}
			else
			{return("9")};

			lLength = aCurrent.length;

			if (lLength == 2)
			{sValue = sValue + " " + aDescriptor[1] + " " + aCurrent[0] + " " + aDescriptor[8] + " " + aCurrent[1];}
			else
			{return("9");}
		}
		else if (sOperator == "M44")
		{
			if (sDelimeter == ";")
			{aCurrent = sInput.split(sDelimeter);}
			else
			{return("10")};

			lLength = aCurrent.length;

			if (lLength == 2)
			{sValue = sValue + " " + aDescriptor[2] + " " + aCurrent[0] + " " + aDescriptor[8] + " " + aCurrent[1];}
			else
			{return("10");}
		}
		else if (sOperator == "M22")
		{
			//if (sDelimeter != "")
			//{
			//	var reInput = new RegExp(sDelimeter, "g");
			//	sTemp = sInput.replace(reInput, ";");
			//}
			//else
			//{sTemp = sInput;}
			var reInput = new RegExp(",", "g");		//handle 1;3,5	
			sTemp = sInput.replace(reInput, ";");
			sValue = sValue + " " + aDescriptor[7] + " (" + sTemp + ")";
		}
		else if (sOperator == "M57")
		{
			var reInput = new RegExp(",", "g");		//handle 1;3,5	
			sTemp = sInput.replace(reInput, ";");
			sValue = sValue + " " + aDescriptor[20] + " (" + sTemp + ")";
		}
		else if (sOperator == "M18")
		{
			//sValue = sValue + " " + aDescriptor[3] + " *" + sInput + "*";
			sValue = sValue + " " + aDescriptor[3] + " " + sInput;
		}
		else if (sOperator == "M43")
		{
			//sValue = sValue + " " + aDescriptor[4] + " *" + sInput + "*";
			sValue = sValue + " " + aDescriptor[4] + " " + sInput;
		}
		else if (sOperator == "P1")
		{
			sValue = sValue + " " + aDescriptor[5] + " " + sInput;
		}
		else if (sOperator == "P2")
		{
			sValue = sValue + " " + aDescriptor[6] + " " + sInput;
		}
		else if (sOperator == "R1")
		{
			sValue = sValue + " " + aDescriptor[5] + " " + sInput;
		}
		else if (sOperator == "R2")
		{
			sValue = sValue + " " + aDescriptor[6] + " " + sInput;
		}
		else
		{
			sValue = sValue + " " + sOperatorSymbol + " " + sInput;
		}
		return(sValue);
	}

	function GetOperatorSymbol(sOperator)
	{
		//**************************************************************************
		//Purpose:  get operator ID from name
		//Input:	sOperator
		//Output:   operator ID
		//**************************************************************************
		if (sOperator == "M17")
			return("between");
		else if (sOperator == "M44")
			return("not between");
		else if (sOperator == "M6")
			return("=");
		else if (sOperator == "M7")
			return("<>");
		else if (sOperator =="M8")
			return(">");
		else if (sOperator =="M10")
			return(">=");
		else if (sOperator == "M9")
			return("<");
		else if (sOperator == "M11")
			return("<=");
		else if (sOperator == "M18")
			return("like");
		else if (sOperator == "M43")
			return("not like");
		else if (sOperator == "M22")
			return("in");
		else if (sOperator == "R1" || sOperator == "P1")
			return("highest");
		else if (sOperator == "R2" || sOperator == "P2")
			return("lowest");
		else
			return("=");
	}

	function GetSelectedItemValue(oList)
	{
		//**************************************************************************
		//Purpose:  get current attribute from list box for Expression prompt
		//Input:    oList
		//Output:   the value of current attribute item
		//**************************************************************************
		var sValue;
		var lLength = oList.options.length;
		var iSelect = 0;
		for(var i=0; i<lLength; i++)
		{
			if (oList.options[i].selected)
			{
				iSelect += 1;
				if (iSelect > 1)
					{return("1");}
				else
					{sValue = oList.options[i].value;}
			}
		}
		if (iSelect == 0)
		{return("0");}
		else
		{return(sValue);}
	}

	function GetSelectedItemText(oList)
	{
		//**************************************************************************
		//Purpose:  get current attribute from list box for Expression prompt
		//Input:    oList
		//Output:   the value of current attribute item
		//**************************************************************************
		var sValue;
		var lLength = oList.options.length;
		var iSelect = 0;
		for(var i=0; i<lLength; i++)
		{
			if (oList.options[i].selected)
			{
				iSelect += 1;
				if (iSelect > 1) 
				{return("1");}
				else
				{sValue = oList.options[i].text;}
			}
		}
		if (iSelect == 0)
		{return("0");}
		else
		{return(sValue);}
	}

	function CleanList(oList)
	{
		//**************************************************************************
		//Purpose:  Clean list box
		//Input:    oList
		//Output:   the value of current attribute item
		//**************************************************************************
		var sValue;
		var lLength = oList.options.length;
		var iSelect = 0;
		for(var i=lLength-1; i>=0; i--)
		{
			sValue = oList.options[i].value;
			if (sValue == "")
			{
				oList.options[i] = null;
			}
		}
	}

	function CleanInput(sInput)
	{
		//**************************************************************************
		//Purpose:  Clean Input value ie. delete all empty value
		//Input:    sInput
		//Output:   the clean value 
		//**************************************************************************
		var sCleanInput;
		var aInput;
		var lInputLength;
		var bFirst;
		var lIndex;

		bFirst = true;
		sCleanInput = "";

		//replace "," with ";"
		//commented out by JILI; Sales > 5,000
		//var reInput = new RegExp(",", "g");
		//sInput = sInput.replace(reInput, ";");

		aInput = sInput.split(";")
		lInputLength = aInput.length;
		for (lIndex = 0; lIndex < lInputLength; lIndex++)
		{
			if (aInput[lIndex] != "")
			{
				if (bFirst)
				{
					sCleanInput = aInput[lIndex];
					bFirst = false;
				}
				else
				{
					sCleanInput = sCleanInput + ";" + aInput[lIndex];
				}
			}
		}
		return(sCleanInput);
	}

	function CleanPercentageSign(sInput)
	{
		//**************************************************************************
		//Purpose:  Clean Input vslue ie. delete all empty value
		//Input:    sInput
		//Output:   the clean value 
		//**************************************************************************
		var sCleanInput;
		
		sCleanInput = sInput
		while (sCleanInput.length > 0 )
		{
			if (sCleanInput.substring(sCleanInput.length-1,sCleanInput.length)== "%")
			{
				sCleanInput = sCleanInput.substring(0,sCleanInput.length-1)
			}
			else
			{
				break;
			}
		}
		return(sCleanInput);
	}

	function DisplayError(sGeneralErrorMssage, sDetailErrorMessage, sPin)
	{
		//**************************************************************************
		//Purpose:  Display error message
		//Input:    sGeneralErrorMssage, sDetailErrorMessage
		//Output:   Error message.
		//**************************************************************************
		var sInnerHTML;

		sInnerHTML = "<TABLE BORDER=\"0\" CELLSPACING=\"0\" CELLPADDING=\"0\"><TR>";
		sInnerHTML = sInnerHTML + "<TD WIDTH=\"11\"><IMG SRC=\"images/1ptrans.gif\" WIDTH=\"11\" HEIGHT=\"1\" BORDER=\"0\" ALT=\"\" /></TD>"
		sInnerHTML = sInnerHTML + "<TD VALIGN=\"TOP\" ALIGN=\"LEFT\" WIDTH=\"23\"><IMG SRC=\"images/promptError_white.gif\" WIDTH=\"23\" HEIGHT=\"23\" ALT=\"Warning!\" BORDER=\"0\" /></TD>"
		sInnerHTML = sInnerHTML + "<TD WIDTH=\"4\"><IMG SRC=\"images/1ptrans.gif\" WIDTH=\"4\" HEIGHT=\"1\" BORDER=\"0\" ALT=\"\" /></TD>"
		sInnerHTML = sInnerHTML + "<TD VALIGN=\"TOP\" ALIGN=\"LEFT\">"
		sInnerHTML = sInnerHTML + "<FONT FACE=\"" + sFontFamily + "\" SIZE=\"" + sSmallFontSize + "\" COLOR=\"#CC0000\"><B>" + sGeneralErrorMssage + "</B><BR /></FONT>"
		sInnerHTML = sInnerHTML + "</TD></TR></TABLE>"

		document.all("GeneralErrorDisplay").innerHTML =  sInnerHTML; //"<IMG SRC=\"images/promptError_white.gif\" WIDTH=\"23\" HEIGHT=\"23\" ALT=\"Warning!\" BORDER=\"0\" />";
		document.all("DetailErrorDisplay_" + sPin).innerHTML = "<IMG SRC=\"Images/promptError_white.gif\" WIDTH=\"23\" HEIGHT=\"23\" BORDER=\"0\"></IMG>" + "<FONT FACE=\"" + sFontFamily + "\" SIZE=\"" + sSmallFontSize + "\" COLOR=\"#CC0000\"><B>" + sDetailErrorMessage + "</B><BR /></FONT>";
		
		//with (document.GeneralErrorDisplay) {
		 // open();
		 // write(sInnerHTML);
		 // close();
		//}

	}
	
	function ClearError(sPin)
	{
		//**************************************************************************
		//Purpose:  Display error message
		//Input:    sGeneralErrorMssage, sDetailErrorMessage
		//Output:   Error message.
		//**************************************************************************
		var oGralErrDisp;
		var oDetailErrDisp;
		
		oGralErrDisp = getObj("GeneralErrorDisplay");
		oDetailErrDisp = getObj("DetailErrorDisplay_" + sPin);
		
		oGralErrDisp.innerHTML =  "";
		oDetailErrDisp.innerHTML = "";
	}

	function selectInOperator(iPromptIndex) {
	//**************************************************************************
	//Purpose:  Select the operator "IN" when the user chooses to load the filter specifications from a file
	//Input:    iPromptIndex
	//**************************************************************************
		var iSelectIndex;
		var oList;
		oList = window.document.PromptForm['Operator_' + iPromptIndex]
		if (oList) {
			if ((oList[oList.selectedIndex].value != 'M22') && (oList[oList.selectedIndex].value != 'M57')) {
				for (iSelectIndex=0; iSelectIndex<oList.length; iSelectIndex++)
					if (oList[iSelectIndex].value == 'M22')
					 oList.selectedIndex = iSelectIndex;
			}
		}
	}

	function checkFileExtension(iPromptIndex) {
	//**************************************************************************
	//Purpose:  Check the extension of the file that the user has choosen to import the filter specifications
	//Input:    iPromptIndex
	//Output:   A boolean indicating if the file extension was ok or not
	//**************************************************************************
		var sFileName;
		var sFileExtension;
		var oFile;
		oFile = window.document.PromptForm['nuXML_TextFile_' + iPromptIndex]
		if (oFile) {
			sFileName = oFile.value;
			sFileExtension = "," + (sFileName.substring(sFileName.length - 3)) + ","
			if (sValidExtensions.indexOf(sFileExtension) == -1) {
				alert(aDescriptor[17]);
				return false;
			}
		}
		return true;
	}
	
	
	function SubmitPromptIndex(sExpression)
	{
		//**************************************************************************
		//Purpose:  
		//Input:    oList
		//Output:   
		//**************************************************************************
		var oModify = getObj(sExpression);
	
		if (oModify!=null){
				if (oModify.length > 1) {
					oModify(0).click();
				}
				else {
					oModify.click();
				}
		}		
	}
	
	
	function SubmitPromptForm(sExpression)
	{
		//**************************************************************************
		//Purpose:  
		//Input:    
		//Output:   
		//**************************************************************************
		var oModify = getObj(sExpression);
	
		if (oModify!=null){			
			return oModify.submit();
		}
	}
	
	function ChangeValue(sObjName, sValue)
	{
		//**************************************************************************
		//Purpose:  
		//Input:    
		//Output:   
		//**************************************************************************
		
		var oElement = getObj(sObjName);
		
		oElement.value = sValue;		
	}