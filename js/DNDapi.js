/*
 *Constants:
 */

//Drop Constants:
var NO_DROP = 0;
var DROP = 1;

//Drag Constants:
var NO_DRAG = 0;
var DRAG = 1;
var DRAG_REMOVE = 2;

//Orientation:
var VERTICAL = 0;
var HORIZONTAL = 1;

//Elements ID DELIMITER
var DELIMITER = '_DELIM_';

//Elements class:
var ELEMENT_CLASS = 'DNDElement';

//Turn debuging on or off
var DEBUGGING = false;

/*
 * Gloabal vars:
 */
var zones  = new Array();       //All DND Zones of the page

var target = null;          //Target DND element
var targetZone = null;      //Target DND Zone
var source = null;          //Source DND element
var sourceZone = null;      //Source DND Zone

var drag;
var defaultZone = '';


/*
 * IDiv object:
 */

function objIDiv(vObject) {
//*********************************************************************************************
//Purpose: Constructor for an IDiv object, this is an interface that allow
//           us to modify a DIV tag element DHTML properties independently 
//           from the browser platform.
//Inputs:  vObject - DIV name or DIV reference DIV that this IDiv object will modify
//Outputs: a new IDiv object
//*********************************************************************************************/

    //Properties:
    this.obj = getObj(vObject)
    this.id  = vObject.id;
    
    //Methods:
    this.getBGColor = objIDiv_getBGColor;
    this.setBGColor = objIDiv_setBGColor;
    this.getInnerHTML = objIDiv_getInnerHTML;
    this.setInnerHTML = objIDiv_setInnerHTML;
    this.show = objIDiv_show;
    this.hide = objIDiv_hide;
    this.move = objIDiv_move;
    
    return this;
    
}


function objIDiv_getBGColor() {
//*********************************************************************************************
//Purpose: gets the backcolor of a DIV
//Inputs:  none
//Outputs: String - the DIV background color
//*********************************************************************************************/
    return getBGColor(this.obj);
}


function objIDiv_setBGColor(sColor) {
//*********************************************************************************************
//Purpose: Set the background color of a DIV
//Inputs:  sColor - New color for the DIV
//Outputs: None.
//*********************************************************************************************
    return setBGColor(this.obj, sColor);
}

function objIDiv_getInnerHTML() {
//*********************************************************************************************
//Purpose: gets the InnerHTML of a DIV
//Inputs:  none
//Outputs: String - the DIV InnerHTML text
//*********************************************************************************************/
    return getInnerHTML(this.obj);
}

function objIDiv_setInnerHTML(sHTMLText) {
//*********************************************************************************************
//Purpose: Set the innerHTML color of a DIV.
//Inputs:  sHTMLText - New innerHTML for the DIV.
//Outputs: None.
//*********************************************************************************************
    return setInnerHTML(this.obj, sHTMLText);
}

function objIDiv_show() {
//*********************************************************************************************
//Purpose: Shows a DIV on the browser
//Inputs:  none
//Outputs: none
//*********************************************************************************************/
    return displayObj(this.obj);
}

function objIDiv_hide() {
//*********************************************************************************************
//Purpose: Removes a DIV on the browser
//Inputs:  none
//Outputs: none
//*********************************************************************************************/
    return removeObj(this.obj);
}

function objIDiv_move(x, y) {
//*********************************************************************************************
//Purpose: Moves a DIV to the given position in pixels.
//Inputs:  x: the x coordinate
//         y: the y coordinate.
//Outputs: none
//*********************************************************************************************/
    return moveObjTo(this.obj, x, y);
}



/*
 * DNDElement object:
 */

function objDNDElement(value, caption) {
//*********************************************************************************************
//Purpose: Constructor of a DNDElement. This objects are elements on the interface
//           that can be dragged, they belong to 1 or more DNDZones, which associates
//           a DIV tag to each one of them.
//Inputs:  value: The value of the Element
//         caption: The caption of the Element, if empty, the value is used as caption.
//Outputs: a new DNDElement object
//*********************************************************************************************/
    
    //Properties:
    this.value = value;
    this.caption = caption;
    this.index = 0;
    this.id = '';
    
    //Methods:
    this.hightlight  = objDNDElement_highlight;
    this.unhighlight = objDNDElement_unhighlight;
    this.clone = objDNDElement_clone;
    
    return this;
    
}

function objDNDElement_highlight() {
//*********************************************************************************************
//Purpose: Highlights a div tag (put its background in yellow)
//Inputs:  div: an IDiv object to highlight
//Outputs: none
//*********************************************************************************************/
var div;

    div = new objIDiv(getObj(this.id));
    return div.setBGColor('yellow');
}

function objDNDElement_unhighlight() {
//*********************************************************************************************
//Purpose: Unhighlights a div tag (put to its normal background color)
//Inputs:  div: an IDiv object to unhighlight
//Outputs: none
//*********************************************************************************************/
var div;

    div = new objIDiv(getObj(this.id));
    return div.setBGColor('white');
}

function objDNDElement_clone() {
//*********************************************************************************************
//Purpose: Clones this element, returning a new one with the same properties
//Inputs:  none
//Outputs: a new element with the same properties:
//*********************************************************************************************/
var clone;

    clone = new objDNDElement(this.value, this.caption);
    
    return clone;
}



/*
 * DNDZone object:
 */

function objDNDZone(id) {
//*********************************************************************************************
//Purpose: Constructor of a DNDZone. These objects represents zones in the HTML page
//           here the user may drag and drop elements.
//Inputs:  id: this objects id: each of its elements will be associated by this id
//Outputs: a new DNDZone object
//*********************************************************************************************/

    //Properties:
    this.id = id;
    this.elements = Array(0);
    this.maxElements = 0;       //the zone will use this prop iff maxElements > 0
    this.orientation = VERTICAL;
    this.dropMode = DROP;       
    this.dragMode = DRAG_REMOVE;

    //Methods:
    this.add = objDNDZone_add;
    this.remove = objDNDZone_remove;
    this.insert = objDNDZone_insert;
    this.reindex = objDNDZone_reindex;
    this.render = objDNDZone_render;
    
    return this;    
}

function objDNDZone_add(value, caption) {
//*********************************************************************************************
//Purpose: Inserts a new element at the end of the elements array.
//           it the element exceeds the max number of elements, the first element
//           will be removed.
//Inputs:  value: value of the new element
//         caption: caption of the new element
//Outputs: the new element
//*********************************************************************************************/
var elem;
    
    elem = new objDNDElement(value, caption);

    this.insert(elem, this.elements.length); 
            
    return elem;

}

function objDNDZone_insert(oElem, index) {
//*********************************************************************************************
//Purpose: Inserts a DNDElement at position given by index
//           it the element exceeds the max number of elements, the last element
//           will be removed.
//Inputs:  oElem: reference to a DNDElement we want to insert
//         index: position where the element must be inserted
//Outputs: the inserted element
//*********************************************************************************************/
var elem;
var length;

    length = this.elements.length;
    
    elem = oElem.clone();
    
    for(i = length; index < i; i--) {
        this.elements[i] = this.elements[i - 1];
    }
    this.elements[index] = elem;
    this.reindex();
    
    //Remove if we have more than allowed:
    if (this.maxElements > 0) {
        while (this.elements.length > this.maxElements) {
            this.remove(this.elements.length);
        }
    }

    return(elem);

}

function objDNDZone_remove(index) {
//*********************************************************************************************
//Purpose: Removes the element at the given position
//Inputs:  index: index of the element we want to remove
//Outputs: the new element
//*********************************************************************************************/
var elem;
var length;
    
    length = this.elements.length;
    
    for( i = index + 1; i < length; i++) {
        this.elements[i-1] = this.elements[i];
    }
    
    this.elements.length = length - 1;
    this.reindex();
    
    return(elem);

}

function objDNDZone_reindex() {
//*********************************************************************************************
//Purpose: Reindex (puts the correct index) the elements of the elements array
//Inputs:  none
//Outputs: none
//*********************************************************************************************/

    for (i=0; i<this.elements.length; i++) {
         this.elements[i].index = i;
         this.elements[i].id = this.id + DELIMITER + i;
    }

}


function objDNDZone_render() {
//*********************************************************************************************
//Purpose: Render a DND Zone into the HTML page
//Inputs:  none
//Outputs: none
//*********************************************************************************************/
var sHTML;
var div;
var oDiv;
var i;
var re;

	sHTML = "";
	re = /\"/gi;
	
	sHTML = "<TABLE><TR>";
	
	for (i=0; i<this.elements.length;i++) {
		sHTML += "<TD>";
		sHTML += "<DIV CLASS='" + ELEMENT_CLASS + "' ID='" + this.elements[i].id + "'>" + this.elements[i].caption + "</DIV>";
		if (this.dropMode == DROP) sHTML += "<INPUT TYPE=HIDDEN NAME='z" + this.id + "' VALUE=\"" + this.elements[i].value.replace( re, "&quot;") + "\"/>";
		sHTML += "</TD>";
		if (this.orientation == VERTICAL) {
			sHTML += "</TR><TR>";
		}
	}
    //Add a dummy element at the end, to add elements:
	if (this.dropMode == DROP && ((this.maxElements <= 0) || (i < this.maxElements))) {
		sHTML += "<TD>";
		sHTML += "<DIV ID='" + this.id + DELIMITER + this.elements.length + "'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</DIV>";
		//sHTML += "<DIV ID='" + this.id + DELIMITER + this.elements.length + "'>abc</DIV>";
		sHTML += "</TD>";
		if (this.orientation == VERTICAL) {
			sHTML += "</TR><TR>";
		}
	}

	sHTML += "</TR></TABLE>";
	
	oDiv = new objIDiv(this.id);
	oDiv.setInnerHTML(sHTML);
	
	for (i=0; i < this.elements.length; i++) {
		div = getObj(this.elements[i].id);
		
		div.onmousemove = divMouseMove;
		div.onmousedown = divMouseDown;
	}
	if ((this.dropMode == DROP) && ((this.maxElements <= 0) || (this.elements.length < this.maxElements))) {
		div = getObj(this.id + DELIMITER + this.elements.length);
		
		div.onmousemove = divMouseMove;
		div.onmousedown = divMouseDown;
	}

}


function DragStart(div) {
//*********************************************************************************************
//Purpose: Starts the drag operation
//           When dragging starts, the drag elements is unhighlighted and them
//           it starts moving with the mouse until the button is released.
//Inputs:  div: the DIV object name or reference to start dragging from:
//Outputs: none
//*********************************************************************************************/
    source = getDNDElement(div);
    sourceZone = getDNDZone(div);

    selectNewTarget(null, null);
    
    drag.setInnerHTML(source.caption);
    drag.move(event.clientX + document.body.scrollLeft, event.clientY + document.body.scrollTop); 
    drag.show();
}


function DragDrop() {
//*********************************************************************************************
//Purpose: Ends the drag operation
//Inputs:  None
//Outputs:   If the source drag mode is DRAG_REMOVE, it will be removed from the
//               sourceZone
//           If there is a targetZone, the element will be added to that zone.
//           If there is no targetZone, but the element has a default Zone, it will
//               be added to that zone
//*********************************************************************************************/
var index;

    drag.hide();
    
    if (source != null) {
        if (sourceZone.dragMode == DRAG_REMOVE) {
            sourceZone.remove(source.index);
            sourceZone.render();
        }
    
        if (target == null) {
            if (defaultZone != '') {
                targetZone = zones[source.defaultZone];
                index = targetZone.elements.length;
            }
        } else {
            index = target.index;
        }
    
        if (targetZone != null) {
            targetZone.insert(source, index);
            targetZone.render();
        }
                
        source = null;
        sourceZone = null;
        target = null;
        targetZone = null;
    }
}

function divMouseMove() {
//*********************************************************************************************
//Purpose: Handles the mouse move event of a DIV
//Inputs:  None
//Outputs: If necessary highlights the element 
//         If on DND, move the drag element to its new position.
//*********************************************************************************************/
var elem;
var zone;
var div;

var targetDiv;

    elem = getDNDElement(this);
    zone = getDNDZone(this);
    
    if (source != null) { 
    
        drag.move(event.clientX + document.body.scrollLeft, event.clientY + document.body.scrollTop); 
        
        if (zone.dropMode == DROP) {
            selectNewTarget(elem, zone);
        }

    } else {
    
        if ((zone.dragMode != NO_DRAG) && (elem.index < zone.elements.length)) {
            selectNewTarget(elem, zone);
        }
    }
    
    if (bIsIE4) {
        window.event.cancelBubble = true;
        window.event.returnValue = false;
    }
        
    debug("|");
    
}

function divMouseDown() {
//*********************************************************************************************
//Purpose: Handles the mouse down event of a DIV
//Inputs:  None
//Outputs: If necessary starts the DND operation
//*********************************************************************************************/
var elem;
var zone;

    elem = getDNDElement(this);
    zone = getDNDZone(this);
    
    if ((zone.dragMode != NO_DRAG) && (elem.index < zone.elements.length))  {
        DragStart(this);
    } else {
        source = null;
        sourceZone =  null;
    }

}


function MouseMove() {
//*********************************************************************************************
//Purpose: Handles the mouse move event of the window
//Inputs:  None
//Outputs: If on DND operation, moves the drag DIV to its new position, based on the mouse cooridnates.
//*********************************************************************************************/
var div;

    if (source != null) {
        drag.move(event.clientX + document.body.scrollLeft, event.clientY + document.body.scrollTop); 
    }
    
    if (target != null) {
        selectNewTarget(null, null);
    }

    if (bIsIE4) {
        window.event.cancelBubble = true;
        window.event.returnValue = false;
    }
    
    debug(".");
    
}

function MouseUp() {
//*********************************************************************************************
//Purpose: Handles the mouse up event of the window
//Inputs:  None
//Outputs: If on DND operation, calls DragDrop.
//*********************************************************************************************/
    if (source != null) {
        DragDrop();
    }
}

function getDNDElement(div) {
//*********************************************************************************************
//Purpose: Returns the element to which the div object belongs to
//Inputs:  div: the IDiv object to search.
//Outputs: the DNDElement associated with div.
//*********************************************************************************************/
var zone;
var index;
var elem;
    
    zone  = getDNDZone(div);
    index = div.id.substr(div.id.lastIndexOf(DELIMITER) + DELIMITER.length);
    
    if (index < zone.elements.length) {
        elem = zone.elements[index];
    } else {
        elem = new objDNDElement("", "");
        elem.id = zone.id + DELIMITER + zone.elements.length;
        elem.index = zone.elements.length;
    }
    
    return elem;
    
}

function getDNDZone(div) {
//*********************************************************************************************
//Purpose: Returns the zone to which the div object belongs to
//Inputs:  div: the IDiv object to search its zone
//Outputs: the DNDZone to which this elements belong to.
//*********************************************************************************************/
var sZoneId;

    sZoneId = div.id.substr(0, div.id.indexOf(DELIMITER));
    return zones[sZoneId];
    
}


function selectNewTarget(newTarget, newZone) {
//*********************************************************************************************
//Purpose: Sets the new Target and targetZone, if a previous existed, it unhighlights it.
//Inputs:  newTarget, newZone
//Outputs: 
//*********************************************************************************************/

    if (target != null) {
        if ((newTarget == null) || (target.id != newTarget.id)) {
            target.unhighlight();
                    
            target = null;
            targetZone = null;
        }
    }
    
    target = newTarget;
    targetZone = newZone;
          
    if (target != null) {
        target.hightlight();
    }

}

function debug(sText) {
//*********************************************************************************************
//Purpose: Sends the debug text, somewhere we can debug it:
//           The best way to do this, is to have a form with id=form in your document
//           with an textarea with id=text where we can set its value:
//Inputs:  sText: the text we want to display
//Outputs: 
//*********************************************************************************************/
    if (DEBUGGING) {
        document.form.text.value += sText;
    }
}

function initDND() {
//*********************************************************************************************
//Purpose: Initialize DND operations.
//			Assumes a drag DIV exist and creates an IDiv object for it.
//			It also links the window events necessary for DND.
//Inputs:  none
//Outputs: none
//*********************************************************************************************/
    drag = new objIDiv("drag");
    drag.hide();
    
    window.document.onmousemove = MouseMove;
    window.document.onmouseup = MouseUp;
}
