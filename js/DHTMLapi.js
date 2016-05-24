
// Global browser flags

// bIsIE4 (Internet Explorer 4+)
// bIsW3C (Netscape 6/Mozilla)
// bIsNN4 (Netscape 4+)

var bIsIE4 = (document.all) ? true : false;
var bIsW3C = (document.getElementById && !bIsIE4) ? true : false;
var bIsNN4 = (document.layers) ? true : false;

// bIsDHTML capable browser for degradability

var bIsDHTML = (bIsIE4 || bIsW3C || bIsNN4) ? true : false;


function getObj(vObject) {
//*********************************************************************************************
//Purpose: Return an object reference
//Inputs:  vObject - object name or object reference
//Outputs: object reference.
//*********************************************************************************************

	if (typeof vObject == 'string') {
		if (bIsIE4)
	    	return eval("document.all." + vObject);
		else if (bIsW3C)
	   		return document.getElementById(vObject);
	   	else
			if (document.layers[vObject])
				return document.layers[vObject];
			else
		   		return eval("document." + vObject);
	}
	else {
		return vObject;
	}
}

function getObjStyle(vObject) {
//*********************************************************************************************
//Purpose: Return a reference to an objects style property
//Inputs:  vObject - object name or reference
//Outputs: style reference.
//*********************************************************************************************
	if (bIsIE4 || bIsW3C)
		return getObj(vObject).style;
	else
		return getObj(vObject);
}

function getClientWidth() {
//*********************************************************************************************
//Purpose: Return the ClientWidth
//Inputs:  None.
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
        if (bIsIE4) { 
            return document.body.clientWidth;
        } // else is NN4 && W3C
        return innerWidth;
    }
    return 0;
}

function getOffsetWidth() {
//*********************************************************************************************
//Purpose: Return the OffsetWidth
//Inputs:  None.
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
        if (bIsIE4) {
            return document.body.offsetWidth;
        } // else bIsNN4 && W3C
        return outerWidth;
    }
    return 0;
}

function getClientHeight() {
//*********************************************************************************************
//Purpose: Return the ClientHeight
//Inputs:  None.
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
        if (bIsIE4) {
            return document.body.clientHeight;
        } // else bIsNN4 && W3C
        return innerHeight;
    }
    return 0;
}

function getOffsetHeight() {
//*********************************************************************************************
//Purpose: Return the OffsetHeight property
//Inputs:  None.
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
        if (bIsIE4) {
            return document.body.offsetHeight;
        } // else bIsNN4 && W3C
        return outerHeight;
    }
    return 0;
}

function getObjWidth(vObject) {
//*********************************************************************************************
//Purpose: Return the computed width of an object
//Inputs:  Object name or object reference..
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
        if (bIsIE4) {
			return obj.offsetWidth;
        } else if (bIsW3C) {
            return parseInt(document.defaultView.getComputedStyle(obj, "").getPropertyValue("width"));
        } else	{	// else bIsNN4
			return obj.clip.width;
		}
    }
    return 0;
}

function getObjHeight(vObject) {
//*********************************************************************************************
//Purpose: Return the computed height of an object
//Inputs:  Object name or object reference..
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
        if (bIsIE4) {
			return obj.offsetHeight;
		} else if (bIsW3C) {
            return parseInt(document.defaultView.getComputedStyle(obj, "").getPropertyValue("height"));
        } else {	// else is NN4
	        return obj.clip.height;
		}
    }
    return 0;
}

function getObjTop(vObject) {
//*********************************************************************************************
//Purpose: Return the computed top value of an object
//Inputs:  Object name or object reference.
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
        if (bIsIE4) {
			return obj.style.pixelTop;
		} else if (bIsW3C) {
            return parseInt(document.defaultView.getComputedStyle(obj, "").getPropertyValue("top"));
        } else {			// else is NN4
	        return obj.top;
		}
    }
    return 0;
}

function getObjLeft(vObject) {
//*********************************************************************************************
//Purpose: Return the computed left value of an object
//Inputs:  Object name or object reference.
//Outputs: integer.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
        if (bIsIE4) {
			return obj.style.pixelLeft;
		} else if (bIsW3C) {
            return parseInt(document.defaultView.getComputedStyle(obj, "").getPropertyValue("left"));
        } else {	// else is NN4
	        return obj.left;
		}
    }
    return 0;
}

function setObjHeight(vObject, iHeight) {
//*********************************************************************************************
//Purpose: Set the height of an object
//Inputs:  vObject - Object name or object reference, iNewHeight - the new height for this object..
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
		var iNewHeight = parseInt(iHeight);
		if (!isNaN(iNewHeight)) {
			if (bIsIE4) {
			    obj.style.height= iNewHeight;
			} else if (bIsW3C) {
			    obj.style.height = iNewHeight + "px";
			} else { // is NN4
				if (document.layers[obj.id])			
			        document.layers[obj.id].clip.height = iNewHeight;
				else
					obj.clip.height = iNewHeight;
			}
	    }
    }
}

function setObjWidth(vObject, iWidth) {
//*********************************************************************************************
//Purpose: Set the width of an object
//Inputs:  vObject - Object name or object reference, iNewWidth - the new width for this object..
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject)
		var iNewWidth = parseInt(iWidth);
		if (!isNaN(iNewWidth)) {
			if (bIsIE4) {
				obj.style.width = iNewWidth;
			} else if (bIsW3C) {
			    obj.style.width = iNewWidth + "px";
			} else { // is NN4
				if (document.layers[obj.id])
			        document.layers[obj.id].clip.width = iNewWidth;
				else
					obj.clip.width = iNewWidth;
			}
		}
    }
}

function clipObjTo(vObject, iTop, iRight, iBottom, iLeft) {
//*********************************************************************************************
//Purpose: Clip a object to a rectangle.
//Inputs:  vObject - Object name or object reference; iCtop, iCright, iCbottom, iCleft - the rectangles coordinates.
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
		var iCtop = parseInt(iTop);
		var iCright = parseInt(iRight);
		var iCbottom = parseInt(iBottom);
		var iCleft = parseInt(iLeft);
		if (!isNaN(iCtop) && !isNaN(iCright) && !isNaN(iCbottom) && !isNaN(iCleft)) {
			if (bIsIE4 || bIsW3C) {
			    obj.style.clip = "rect(" + iCtop + "px " + iCright + "px " + iCbottom + "px " + iCleft + "px)";
			} else { // bIsNN4
			    obj = document.layers[obj.id];
			    obj.clip.top = iCtop;
			    obj.clip.right = iCright;
			    obj.clip.bottom = iCbottom;
			    obj.clip.left = iCleft;
			}
		}
    }
}

function moveObjTo(vObject, iX, iY) {
//*********************************************************************************************
//Purpose: Move an object to given coordinates
//Inputs:  vObject - Object name or object reference, x - the left coordinate, y - the top coordinate.
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
		var x = parseInt(iX);
		var y = parseInt(iY);
		if (!isNaN(x) && !isNaN(y)) {
			if (bIsIE4 || bIsW3C) {
			    obj.style.left = x + "px";
			    obj.style.top = y + "px";
			} else {
			    document.layers[obj.id].left = x;
			    document.layers[obj.id].top = y;
			}
		}
    }
}

function moveObjBy(vObject, iDx, iDy) {
//*********************************************************************************************
//Purpose: Move an object by the given values (in px).
//Inputs:  vObject - Object name or object reference, dx - the amount to add to the left coordinate, dy - the amount to add to the top coordinate.
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
		var dx = parseInt(iDx);
		var dy = parseInt(iDy);
		if (!isNaN(dx) && !isNaN(dy)) {
			if (bIsIE4 || bIsW3C) {
			    obj.style.left = (getObjLeft(obj.id) + dx) + "px";
			    obj.style.top = (getObjTop(obj.id) + dy) + "px";
			} else {
			    document.layers[obj.id].left += dx;
			    document.layers[obj.id].top += dy;
			}
		}
	}
}

function hideObj(vObject) {
//*********************************************************************************************
//Purpose: Set an objects visibility property to hidden.
//Inputs:  vObject - Object name or object reference
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
        if (bIsIE4 || bIsW3C) {
            obj.style.visibility = "hidden";
        } else {
            document.layers[obj.id].visibility = "hide";
        }
    }
}

function showObj(vObject) {
//*********************************************************************************************
//Purpose: Set an objects visibility property to visible.
//Inputs:  vObject - Object name or object reference
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
		var obj = getObj(vObject);
        if (bIsIE4 || bIsW3C) {
            obj.style.visibility = "visible";
        } else {
            document.layers[obj.id].visibility = "show";
        }
    }
}

function displayObj(vObject) {
//*********************************************************************************************
//Purpose: Force a named object into the document flow (does not work in NN4)
//Inputs:  vObject - Object name or object reference
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
        if (bIsIE4 || bIsW3C) {
			obj = getObj(vObject);
            obj.style.display = "block";
        }
        showObj(vObject);		
    }
    
}

function removeObj(vObject) {
//*********************************************************************************************
//Purpose: Remove a named object from the document flow (does not work in NN4)
//Inputs:  vObject - Object name or object reference
//Outputs: None.
//*********************************************************************************************
    if (bIsDHTML) {
        if (bIsIE4 || bIsW3C) {
			obj = getObj(vObject);
            obj.style.display = "none";
        }
        hideObj(vObject);
    }
}

function getBGColor(vObject) {
//*********************************************************************************************
//Purpose: Return the background color of an object.
//Inputs:  vObject - Object name or object reference.
//Outputs: String - the objects background color
//*********************************************************************************************
	if (bIsDHTML) {
		var obj = getObj(vObject);
		if (bIsIE4)
			return obj.style.backgroundColor;
		else if (bIsW3C)
			return obj.style.bgColor;
		else
			return '';
	}
}

function setBGColor(vObject, sColor) {
//*********************************************************************************************
//Purpose: Set the background color of an object.
//Inputs:  vObject - Object name or object reference, sColor - New color for the object.
//Outputs: None.
//*********************************************************************************************
	if (bIsDHTML) {
		var obj = getObj(vObject);
		if ((bIsIE4) || (bIsW3C))
			obj.style.backgroundColor = sColor;
		else
			obj.style.bgColor = sColor;
	}
}


function getInnerHTML(vObject) {
//*********************************************************************************************
//Purpose: Return the innerHTML of an object (only works in IE).
//Inputs:  vObject - Object name or object reference.
//Outputs: String - the objects innterHTML property
//*********************************************************************************************
	if (bIsDHTML) {
		var obj = getObj(vObject);
		if (bIsIE4 || bIsW3C)
			return obj.innerHTML;
		else
			return '';
	}
}

function setInnerHTML(vObject, sHTMLText) {
//*********************************************************************************************
//Purpose: Set the innerHTML color of an object (only works in IE).
//Inputs:  vObject - Object name or object reference, sHTMLText - New innerHTML for the object.
//Outputs: None.
//*********************************************************************************************
	if (bIsDHTML) {
		var obj = getObj(vObject);
		if ((bIsIE4) || (bIsW3C))
			obj.innerHTML = sHTMLText;
	}
}

function getColor(vObject) {
//*********************************************************************************************
//Purpose: Return the foreground color of an object.
//Inputs:  vObject - Object name or object reference.
//Outputs: String - the objects foreground color
//*********************************************************************************************
	if (bIsDHTML) {
		var obj = getObj(vObject);
		if (bIsIE4)
			return obj.style.color;
		else if (bIsW3C)
			return obj.style.color;
		else
			return '';
	}
}

function setColor(vObject, sColor) {
//*********************************************************************************************
//Purpose: Set the foreground color of an object.
//Inputs:  vObject - Object name or object reference, sColor - New color for the object.
//Outputs: None.
//*********************************************************************************************
	if (bIsDHTML) {
		var obj = getObj(vObject);
		if (bIsIE4)
			obj.style.color = sColor;
		else if (bIsW3C)
			obj.style.color = sColor;
	}
}

function getInverseColor(sColor) {
//*********************************************************************************************
//Purpose: Return the inverse color of the supplied color..
//Inputs:  sColor - The color to be converted.
//Outputs: String - the inverse color..
//*********************************************************************************************
	var iInverseColor;
	var sInverseColor = '';

	// Calculate the inverse value.
	iInverseColor = parseInt(sColorText, 16) - parseInt("FFFFFF", 16);

	// If the inverse is negative, get the absolute value.
	if (iInverseColor < 0)
		iInverseColor = -iInverseColor;
		
	// Parse the inverse value to hex.
	iInverseColor = parseHex(iInverseColor);

	// Save the hex value as a string.
	sInverseColor = iInverseColor.toString();
	
	// If the length of the hex value is less than 6 then append
	// zeros to the beginning.
	while (sInverseColor.length < 6){
		sInverseColor = '0' + sInverseColor;
	}
	
	// Return the hex string.
	return sInverseColor;
}

function parseHex(iDecimalValue) {
//*********************************************************************************************
//Purpose: Convert a decimal value to hex
//Inputs:  iDecimalValue - the decimal value for conversion..
//Outputs: Return the hex value.
//*********************************************************************************************
	var aHexadecimalValues = ['0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F'];
	if (iDecimalValue > 16)
		return parseHex(Math.floor(iDecimalValue / 16)) + '' + aHexadecimalValues[iDecimalValue % 16];
	else
		return aHexadecimalValues[iDecimalValue];
				
}

