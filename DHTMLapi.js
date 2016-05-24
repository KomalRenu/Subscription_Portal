/* Copyright © 1996-2009 MicroStrategy Incorporated, All rights reserved. Confidential. */
var bIsIE4 = (document.all) ? true : false;		// IE 4+
var bIsW3C = (document.getElementById && !bIsIE4) ? true : false;		// N6
var bIsNN4 = (document.layers) ? true : false;	// NC4

var lMouseX = 0;
var lMouseY = 0;

var bIsDHTML = (bIsIE4 || bIsW3C || bIsNN4) ? true : false;

function getObj(vObject) {
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
	if (bIsIE4 || bIsW3C)
		return getObj(vObject).style;
	else
		return getObj(vObject);
}

function getClientWidth() {
    if (bIsDHTML) {
        if (bIsIE4) { 
            return document.body.clientWidth;
        } // else is NN4 && W3C
        return innerWidth;
    }
    return 0;
}

function getClientHeight() {
    if (bIsDHTML) {
        if (bIsIE4) {
            return document.body.clientHeight;
        } // else bIsNN4 && W3C
        return innerHeight;
    }
    return 0;
}

function getObjWidth(vObject) {
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

function getObjSumTop(vObject) {
    if (bIsDHTML) {
		var obj = getObj(vObject);
		var lTop = 0;
        if (bIsIE4) {
			for (var i=0; (obj); i++) {
				lTop += obj.offsetTop;
				obj = obj.offsetParent;
			}
			return lTop;
		} else {
			return getObjTop(vObject);
		}
    }
    return 0;
}

function getObjLeft(vObject) {
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

function getObjSumLeft(vObject) {
    if (bIsDHTML) {
		var obj = getObj(vObject);
		var lLeft = 0;
        if (bIsIE4) {
			for (var i=0; (obj); i++) {
				lLeft += obj.offsetLeft;
				obj = obj.offsetParent;
			}
			return lLeft;
        } else {	// else is NN4
			return getObjLeft(vObject);
		}
    }
    return 0;
}

function setObjHeight(vObject, iHeight) {
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

function clearObjHeight(vObject) {
    if (bIsDHTML) {
		var obj = getObj(vObject);
		if (bIsIE4) {
		    obj.style.height = '';
		} else if (bIsW3C) {
		    obj.style.height = '';
		} else { // is NN4
			if (document.layers[obj.id])			
		        document.layers[obj.id].clip.height = '';
			else
				obj.clip.height = '';
		}
    }
}

function clearObjWidth(vObject) {
    if (bIsDHTML) {
		var obj = getObj(vObject);
		if (bIsIE4) {
		    obj.style.width = '';
		} else if (bIsW3C) {
		    obj.style.width = '';
		} else { // is NN4
			if (document.layers[obj.id])			
		        document.layers[obj.id].clip.width = '';
			else
				obj.clip.width = '';
		}
    }
}

function setObjWidth(vObject, iWidth) {
    if (bIsDHTML) {
		var obj = getObj(vObject);
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

function moveObjTo(vObject, iX, iY) {
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

function hideObj(vObject) {
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
    if (bIsDHTML) {
        if (bIsIE4 || bIsW3C) {
			obj = getObj(vObject);
            obj.style.display = "inline";
        }
        showObj(vObject);		
    }
}

function removeObj(vObject) {
    if (bIsDHTML) {
        if (bIsIE4 || bIsW3C) {
			obj = getObj(vObject);
            obj.style.display = "none";
        }
        hideObj(vObject);
    }
}

function setBGColor(vObject, sColor) {
	if (bIsDHTML) {
		var obj = getObj(vObject);
		if ((bIsIE4) || (bIsW3C))
			obj.style.backgroundColor = sColor;
		else
			obj.style.bgColor = sColor;
	}
}

function getMouse(e) {
	if (bIsNN4) {
		lMouseX = e.pageX - 5;
		lMouseY = e.pageY - 5;
	}
	else if (bIsIE4) {
		lMouseX = event.clientX - 5 + document.body.scrollLeft;		
		lMouseY = event.clientY - 5 + document.body.scrollTop;
	}
	else if (bIsW3C) {
		lMouseX = e.pageX - 5;		
		lMouseY = e.pageY - 5;
	}
}

function writeToDiv(sMessage, sDivID) {
	if (bIsNN4) {
		var oDiv = getObj(sDivID);
		oDiv.document.write(sMessage);
		oDiv.document.close();
	}
	else if (bIsIE4 || bIsW3C) {
		var oDiv = getObj(sDivID);
		if (oDiv)
			oDiv.innerHTML = sMessage;
	}
}

function getEventTarget(e) {
	if (bIsDHTML) {
		if (bIsW3C) {
			return e.target;
		}
		else if (bIsIE4) {
			return event.srcElement;
		}
	}
	return null;
}

function cancelEvent(e) {
	if (bIsDHTML) {
		if (bIsW3C)
			e.preventDefault();
		else if (bIsIE4)
			event.returnValue = false;
	}
}

function findTarget(oTarget, sFind) {
	// Step from the target object up through it's parents until you either find a valid target or there is no more parent.
	if (!bIsW3C) {
		while (oTarget.parentNode) {
			if (oTarget.getAttribute(sFind))
				return oTarget;
			oTarget = oTarget.parentNode;
		}
	}
	else {
		var strTarget = '';
		while(oTarget.parentNode) {
			// For NN6 it has to be a table element to use the get attribute function.
			strTarget = oTarget.toString();
			if ((strTarget == '[object HTMLTableCellElement]') || (strTarget == '[object HTMLTableElement]') || (strTarget == '[object HTMLDivElement]') || (strTarget == '[object HTMLSpanElement]')) {
				if (oTarget.getAttribute(sFind))
					return oTarget;
			}
			oTarget = oTarget.parentNode;
		}
	}
	return null;
}