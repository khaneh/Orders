// Selects an item in a selection list
function setSelectValue(pSelectObj, pValue, pDefault) {		
	var myValue, i;
	(pValue == "") ? myValue = pDefault : myValue = pValue;
	for (i = 0; i < pSelectObj.length; i++)
		if (pSelectObj.options[i].value == myValue) { pSelectObj.selectedIndex = i; break; }
	return (1);
}

// Selects an item in a radio group
function setRadioValue(pSelectObj, pValue, pDefault) {
	var myValue, i;
	(pValue == "") ? myValue = pDefault : myValue = pValue;
	for (i = 0; i < pSelectObj.length; i++)
		if (pSelectObj[i].value == myValue) { pSelectObj[i].checked = true; break; }
	return (1);
}

//opens a popup window
function openPopUp(pURL, pName, pW, pH, pX, pY) {
	var myWindow = window.open(pURL, pName, "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=" + pW + ",height=" + pH);
	myWindow.moveTo(pX, pY);
	myWindow.focus();
	return (true);
}

// left trims
function ltrim(pStr) {
	if (pStr == null) return ("");
	var myRegExp = new RegExp("^[ ]*", "gim");
	return (pStr.replace(myRegExp, ""));
}

// checks whether a text input is empty
function isEmpty(pObj) { return (ltrim(pObj.value) == ""); }

// checks whether a text input contains an integer
function isInteger(pObj) { 	return (pObj.value.search(/[^0-9]/) == -1); }

// displays a confirm dialog box then jumps to the specified location
function _confirm(pTxt, pURL) { if (confirm(pTxt)) document.location = pURL; return (1); }
