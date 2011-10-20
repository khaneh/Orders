//
// myLittleTree v1.0
//
// Web site : http://www.mylittletools.net
//	 Email : webmaster@mylittletools.net
//	 (c) 2000-2002, Elian Chrebor, myLittleTools.net,  All rights reserved.
//
// This file is not freeware.
//	 You cannot re-distribute it without any agreement with myLittleTools.net.
// Contact webmaster@mylittletools.net for more info.
//
//	 Logs :
//		20020701 : expandNode function added
//		20020618 : first release
//


// Images
var mlt_blank = new Image(16,22);
var mlt_exp = new Image(16,22);
var mlt_exp_last = new Image(16,22);
var mlt_red = new Image(16,22);
var mlt_red_last = new Image(16,22);
var mlt_node = new Image(16,22);
var mlt_node_last = new Image(16,22);
var mlt_line = new Image(16,22);

// String
var mlt_Error = "Cannot initialize myLittleTree object. Dynamic navigation is disabled.";
var mlt_Alt = "Click to expand or reduce."

// checks browser
var mlt_DOM = (document.getElementById) ? true : false;
var mlt_NS4 = (document.layers) ? true : false;
var mlt_IE = (document.all) ? true : false;
var mlt_IE4 = mlt_IE && !mlt_DOM;
var mlt_Mac = (navigator.appVersion.indexOf("Mac") != -1);
var mlt_IE4M = mlt_IE4 && mlt_Mac;
var mlt_Opera = (navigator.userAgent.indexOf("Opera")!=-1);
var mlt_isOK = !mlt_Opera && !mlt_IE4M && (mlt_DOM);

// Constructor
function myLittleTree_Folder(pId, pFatherId, pDescription, pIcon, pLink, pTarget, pCSSClass)
{
	this.fatherId = pFatherId;
	this.id = pId;
	this.description = pDescription;
	this.link = pLink;
	this.target = pTarget;
	this.icon = pIcon;
	this.CSSClass = pCSSClass;
	this.isOpened = false;
	this.children = new Array();
	this.imageopen = "";
	this.imageclose = "";
}

function myLittleTree(pLitteralName, pImageFolder)
{
	if (arguments.length > 0) 
		this.litteralName = pLitteralName 
	else
	{
		alert(mlt_Error);
		mlt_isOK = false;
	}
	(arguments.length > 1) ? this.imageFolder = pImageFolder : this.imageFolder = "";
	this.folderCount = 0;
	this.root = "";
	this.index = new Array();
	this.tree = new Array();
	this.addFolder = mlt_addFolder;
	this.createHTML = mlt_createHTML;
	this.clickNode = mlt_clickNode;
	this.expandNode = mlt_expandNode;
	mlt_setImages(this.imageFolder);
}

// defines images
function mlt_setImages(pImageFolder)
{
	mlt_blank.src = pImageFolder + "mlt_blank.gif";
	mlt_exp.src = pImageFolder + "mlt_exp.gif";
	mlt_exp_last.src = pImageFolder + "mlt_exp_last.gif";
	mlt_red.src = pImageFolder + "mlt_red.gif";
	mlt_red_last.src = pImageFolder + "mlt_red_last.gif";
	mlt_node.src = pImageFolder + "mlt_node.gif";
	mlt_node_last.src = pImageFolder + "mlt_node_last.gif";
	mlt_line.src = pImageFolder + "mlt_line.gif";
}

// Adds a Node
function mlt_addFolder(pId, pFatherId, pDescription, pIcon, pLink, pTarget, pCSSClass)
{
	var myIcon, myLink, myTarget, myCSSClass;
	(arguments.length > 3) ? myIcon = pIcon : myIcon = "";
	(arguments.length > 4) ? myLink = pLink : myLink = "";
	(arguments.length > 5) ? myTarget = pTarget : myTarget = "";
	(arguments.length > 6) ? myCSSClass = pCSSClass : myCSSClass = "";
	this.tree[this.folderCount] = new myLittleTree_Folder(pId, pFatherId, pDescription, myIcon, myLink, myTarget, myCSSClass);
	this.index[pId] = this.folderCount++;
	(pFatherId != null) ? this.tree[this.index[pFatherId]].children[this.tree[this.index[pFatherId]].children.length] = pId : this.root = pId;
	return (true);
}

// Opens/Closes a Node
function mlt_clickNode(pNodeId)
{
	var myNode = this.tree[this.index[pNodeId]];
	myNode.isOpened = !(myNode.isOpened);
	document.getElementById("IMG_" + myNode.id).src = (myNode.isOpened) ? myNode.imageclose : myNode.imageopen;
	for (var i = 0; i < myNode.children.length; i++)
	{
		if (myNode.isOpened)
		{
			document.getElementById(myNode.children[i]).style.display = "block";
			document.getElementById("IMG_" + myNode.children[i]).src = this.tree[this.index[myNode.children[i]]].imageopen;
		}
		else
		{
			if (this.tree[this.index[myNode.children[i]]].isOpened)
				this.clickNode(myNode.children[i]);
			document.getElementById(myNode.children[i]).style.display = "none";
		}
	}
	return (true);
}

// Expands a full Node
function mlt_expandNode(pNodeId)
{
	if (! mlt_isOK) return (false);
	var myNode = this.tree[this.index[pNodeId]];
	myNode.isOpened = true;
	document.getElementById("IMG_" + myNode.id).src = myNode.imageclose;
	for (var i = 0; i < myNode.children.length; i++)
	{
		//this.expandNode(myNode.children[i]);
		document.getElementById(myNode.children[i]).style.display = "block";
	}
	return (true);
}

// computes HTML code
function mlt_createHTML()
{
	var myIcon = "";
	if (this.tree[this.index[this.root]].icon != "") myIcon = "<IMG SRC=\"" + this.imageFolder + this.tree[this.index[this.root]].icon + "\" BORDER=0 ALIGN=MIDDLE ALT=\"myLittleTree\"> ";
	var myHTML = "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0";
	if (mlt_isOK) myHTML += " STYLE=\"display: block;\"";
	myHTML += " ID=\"" + this.root + "\" SUMMARY=\"" + this.root + "\"><TR><TD NOWRAP>&nbsp;</TD><TD NOWRAP";
	(this.tree[this.index[this.root]].CSSClass == "") ? myHTML += ">" : myHTML += " CLASS=\"" + this.tree[this.index[this.root]].CSSClass + "\">";
	myHTML += myIcon;
	if (this.tree[this.index[this.root]].link != "")
	{
		myHTML += "<A HREF=\"" + this.tree[this.index[this.root]].link + "\"";
		if (this.tree[this.index[this.root]].CSSClass != "") myHTML += " CLASS=\"" + this.tree[this.index[this.root]].CSSClass + "\"";
		(this.tree[this.index[this.root]].target == "") ? myHTML += ">" : myHTML += " TARGET=\"" + this.tree[this.index[this.root]].target + "\">";
	}
	myHTML += this.tree[this.index[this.root]].description;
	if (this.tree[this.index[this.root]].link != "") myHTML += "</A>";
	myHTML += "</TD></TR></TABLE>\n";
	myHTML += mlt_createNodeHTML(this, this.root, "");
	return (myHTML);
}

function mlt_createNodeHTML(pObj, pNodeId, pLineStatus)
{
	var myNode, myIndex, myHTML = "", myCollapse, j, myLineStatus, myIcon, myStyle, myAlt;
	myNode = pObj.tree[pObj.index[pNodeId]];
	for (var i = 0; i < myNode.children.length; i++)
	{
		myIndex = pObj.index[myNode.children[i]];
		(pNodeId == pObj.root) ? myStyle = "block" : myStyle = "none";
		if (pObj.tree[myIndex].children.length > 0)
		{
			if (i == myNode.children.length-1)
			{
				pObj.tree[myIndex].imageopen = mlt_exp_last.src;
				pObj.tree[myIndex].imageclose = mlt_red_last.src;
				myAlt = mlt_Alt;
			}
			else
			{
				pObj.tree[myIndex].imageopen = mlt_exp.src;
				pObj.tree[myIndex].imageclose = mlt_red.src;
				myAlt = mlt_Alt;
			}
		}
		else
		{
			if (i == myNode.children.length-1)
			{
				pObj.tree[myIndex].imageopen = mlt_node_last.src;
				pObj.tree[myIndex].imageclose = mlt_node_last.src;
				myAlt = "";
			}
			else
			{
				pObj.tree[myIndex].imageopen = mlt_node.src;
				pObj.tree[myIndex].imageclose = mlt_node.src;
				myAlt = "";
			}
		}

		myCollapse = "<IMG SRC=\"" + mlt_blank.src + "\" BORDER=0 WIDTH=5 ALT=\"myLittleTree :: spacing\">";
		if (pLineStatus.length >= 1)
			for (j = 0; j < pLineStatus.length ; j++)
				(pLineStatus.charAt(j) == "1") ? myCollapse += "<IMG SRC=\"" + mlt_blank.src + "\" BORDER=0 ALT=\"myLittleTree :: spacing\">" : myCollapse += "<IMG SRC=\"" + mlt_line.src + "\" BORDER=0 ALT=\"myLittleTree :: line\">";
		if (mlt_isOK)
			(pObj.tree[myIndex].children.length > 0) ? myCollapse += "<A HREF=# onclick=\"" + pObj.litteralName + ".clickNode('" + pObj.tree[myIndex].id + "'); return false;\"><IMG SRC=\"" + pObj.tree[myIndex].imageopen + "\" BORDER=0 ID=\"IMG_" + pObj.tree[myIndex].id + "\" ALT=\"" + myAlt + "\"></A>" : myCollapse += "<IMG SRC=\"" + pObj.tree[myIndex].imageopen + "\" BORDER=0 ID=\"IMG_" + pObj.tree[myIndex].id + "\" ALT=\"" + myAlt + "\">";
		else
			myCollapse += "<IMG SRC=\"" + pObj.tree[myIndex].imageclose + "\" BORDER=0 ID=\"IMG_" + pObj.tree[myIndex].id + "\" ALT=\"" + myAlt + "\">";
		myIcon = "";
		if (pObj.tree[myIndex].icon != "") myIcon = "<IMG SRC=\"" + pObj.imageFolder + pObj.tree[myIndex].icon + "\" BORDER=0 ALT=\"myLittleTree\" ALIGN=MIDDLE>";
		myHTML += "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0";
		if (mlt_isOK) myHTML += " STYLE=\"display: " + myStyle + ";\"";
		myHTML += " ID=\"" + pObj.tree[myIndex].id + "\" SUMMARY=\"" + pObj.tree[myIndex].id + "\"><TR><TD NOWRAP>" + myCollapse + "</TD><TD NOWRAP";
		(pObj.tree[myIndex].CSSClass == "") ? myHTML += ">" : myHTML += " CLASS=\"" + pObj.tree[myIndex].CSSClass + "\">";
		myHTML += myIcon + "&nbsp;";
		if (pObj.tree[myIndex].link != "")
		{
			myHTML += "<A HREF=\"" + pObj.tree[myIndex].link + "\"";
			if (pObj.tree[myIndex].CSSClass != "") myHTML += " CLASS=\"" + pObj.tree[myIndex].CSSClass + "\"";
			(pObj.tree[myIndex].target == "") ? myHTML += ">" : myHTML += " TARGET=\"" + pObj.tree[myIndex].target + "\">";
		}
		myHTML += pObj.tree[myIndex].description;
		if (pObj.tree[myIndex].link != "") myHTML += "</A>";
		myHTML += "</TD></TR></TABLE>\n";

		if (pObj.tree[myIndex].children.length > 0)
		{
			myLineStatus = pLineStatus;
			(i == myNode.children.length-1) ? myLineStatus += "1" : myLineStatus += "0";
			myHTML += mlt_createNodeHTML(pObj, pObj.tree[myIndex].id, myLineStatus);
		}
	}
	return (myHTML);
}
