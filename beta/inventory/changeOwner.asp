<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Inventory (5)
PageTitle= "«‰ ﬁ«· „«·ﬂÌ "
SubmenuItem=11
if not Auth(5 , "B") then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------- Submit Trnasfer
'-----------------------------------------------------------------------------------------------------
if request("submit")="«‰ ﬁ«·" then

	itemName	= sqlSafe(request.form("itemName"))
	itemID		= clng(request.form("itemID"))
	fromID		= clng(request.form("fromID"))
	toID		= clng(request.form("toID"))
	qtty		= cdbl(request.form("qtty"))
	accountName = sqlSafe(request.form("accountName"))
	comments	= sqlSafe(left(request.form("comments"),220))	' Max Size of the field is 250
	
	errorFound=false
	mySQL="SELECT SUM(sumQtty) AS sumQtty, AccountID FROM (SELECT SUM((CONVERT(tinyint, InventoryLog.IsInput) - .5) * 2 * InventoryLog.Qtty) AS sumQtty, InventoryLog.owner AS AccountID FROM InventoryLog WHERE (InventoryLog.ItemID = " & itemID & ") AND (InventoryLog.voided = 0) GROUP BY InventoryLog.owner) DERIVEDTBL GROUP BY AccountID HAVING (SUM(sumQtty) <> 0) AND (DERIVEDTBL.AccountID = " & fromID & ")"
	set RSS = Conn.Execute(mySQL)
	if not RSS.eof then
		currentQtty=cdbl(RSS("sumQtty"))
		if currentQtty < qtty then errorFound=true
	else
		errorFound=true
	end if
	RSS.close

	if errorFound then
		Conn.close
		response.redirect "?errMsg="&Server.URLEncode("„ÊÃÊœÌ ﬂ«›Ì ‰Ì” .")
	end if

	mySql="INSERT INTO InventoryLog (ItemID, RelatedID, logDate, Qtty, owner, CreatedBy, IsInput, comments, type) VALUES ("& itemID & ", -5 ,N'"& shamsiToday() & "', "& qtty & ", "& fromID & ", "& session("id") & ", 0, N'«‰ ﬁ«· »Â Õ”«»  " & toID & " (" & comments & ")', 5 )"
	conn.Execute mySql
	
	mySql="INSERT INTO InventoryLog (ItemID, RelatedID, logDate, Qtty, owner, CreatedBy, IsInput, comments, type) VALUES ("& itemID & ", -5 ,N'"& shamsiToday() & "', "& qtty & ", "& toID & ", "& session("id") & ", 1, N'«‰ ﬁ«· «“ Õ”«»  " & fromID & " (" & comments & ")', 5 )"
	conn.Execute mySql
	
	if fromID = "-1" then fromName = "Œ«‰Â ç«Å Ê ÿ—Õ" else fromName = fromID
	if toID = "-1" then toName = "Œ«‰Â ç«Å Ê ÿ—Õ" else toName = toID

	response.write "<br><br>" 
	call showAlert ("«‰ ﬁ«· " & qtty & " ⁄œœ " & itemName & " «“ Õ”«» " & fromName & " »Â Õ”«» " & toName & " «‰Ã«„ ‘œ." , CONST_MSG_INFORM)
	response.end

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Trnasfer Form
'-----------------------------------------------------------------------------------------------------
elseif request("submit")="Ã” ÃÊ" then
	
	if not isNumeric(request("oldItemID")) then
		response.write "<br>Œÿ«!" 
		response.end
	end if 

	set RSS=Conn.Execute ("SELECT * from inventoryItems where oldItemID = " & request("oldItemID") )
	id = RSS("id")
	set RSS=Conn.Execute ("SELECT SUM(DERIVEDTBL.sumQtty) AS sumQttys, DERIVEDTBL.AccountID, dbo.Accounts.AccountTitle FROM (SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty,  dbo.InventoryLog.owner AS AccountID FROM dbo.InventoryLog WHERE (dbo.InventoryLog.ItemID = " & id & " and dbo.InventoryLog.voided=0) GROUP BY dbo.InventoryLog.owner) DERIVEDTBL left outer JOIN dbo.Accounts ON DERIVEDTBL.AccountID = dbo.Accounts.ID GROUP BY DERIVEDTBL.AccountID, dbo.Accounts.AccountTitle having SUM(DERIVEDTBL.sumQtty)<>0")	
	%><BR><BR>
	<FORM METHOD=POST ACTION="?" onsubmit="return checkValidation();">
	<CENTER><H3>«‰ ﬁ«·</H3>	</CENTER>
	<TABLE align=center width=70%>
	<TR height=30>
		<TD align=left valign=top>‰«„ ﬂ«·«:</TD>
		<TD align=right valign=top><b><%=request("accountName")%> <INPUT TYPE="hidden" name="itemName" value="<%=request("accountName")%>"><INPUT TYPE="hidden" name="itemID" value="<%=id%>"></b></TD>
	</TR>
	<TR>
		<TD align=left valign=top>«“ Õ”«»: </TD>
		<TD align=right valign=top>
		
		<TABLE dir=rtl align=center width=100%>
		<TR bgcolor="eeeeee">
			<TD><INPUT TYPE="radio" NAME="" disabled></TD>
			<TD><SMALL>‰«„</SMALL></A></TD>
			<TD align=center><!A HREF="default.asp?s=1"><SMALL>‘„«—Â Õ”«»</SMALL></A></TD>
			<TD align=center><!A HREF="default.asp?s=3"><SMALL> ⁄œ«œ</SMALL></A></TD>
		</TR>
		<%
		tmpCounter=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
				'tmpColor="#DDDDDD"
				'tmpColor2="#EEEEBB"
			End if 

		%>
		<TR bgcolor="<%=tmpColor%>" height=25 style="cursor:hand" onclick="this.getElementsByTagName('input')[0].click();this.getElementsByTagName('input')[0].focus();">
			<TD><INPUT TYPE="radio" NAME="fromID" value="<%=RSS("AccountID")%>"></TD>
			<TD><%
			if RSS("AccountID")="-1" then
				response.write "Œ«‰Â Ã«Å Ê ÿ—Õ" 
			else
				response.write RSS("AccountTitle")
			end if %></TD>
			<TD  align=center dir=ltr><%
			if RSS("AccountID")="-1" then
				response.write "-"
			else
				response.write RSS("AccountID")
			end if %></TD>
			<TD align=center dir=ltr><%=RSS("sumQttys")%></TD>
		</TR>
			  
		<% 
		RSS.moveNext
		Loop
		%>
		</TABLE><br>
		</TD>
	</TR>
	<TR height=30>
		<TD align=left valign=top>»Â Õ”«»: </TD>
		<TD align=right valign=top>
		<SELECT NAME="aaa2"  onchange="hideIT()">
		<option value=1 >Œ«‰Â ç«Å Ê ÿ—Õ</option>
		<option value=2>œÌê—«‰ (‘„«—Â Õ”«»)</option>

		</SELECT>
		<span name="aaa1"  id="aaa1" style="visibility:'hidden'">
		<input type="hidden" Name='tmpDlgArg' value=''>
		<input type="hidden" Name='tmpDlgTxt' value=''>
		<INPUT  dir="LTR"  TYPE="text" NAME="toID" maxlength="10" size="13"  value="-1" onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
		</span><br></TD>
	</TR>
	<TR height=30>
		<TD align=left  valign=top> ⁄œ«œ:</TD>
		<TD align=right valign=top><INPUT dir=ltr TYPE="text" NAME="qtty" size=25 value="<%=preQtty%>" ><!onKeyPress="return maskNumber(this);" ><br><br></TD>
	</TR>
	<TR>
		<TD align=left valign=top> Ê÷ÌÕ« </TD>
		<TD align=right>
			<TEXTAREA NAME="comments" ROWS="4" COLS="" style="width:100%"></TEXTAREA><br><br>
		</TD>
	</TR>
	<TR>
		<TD align=left valign=top></TD>
		<TD align=center valign=top><INPUT TYPE="submit" name="submit" value="«‰ ﬁ«·"></TD>
	</TR>
	</TABLE>
	</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
function checkValidation(){
	found = false;
	len = document.getElementsByName('fromID').length
	for (i = 0; i <len; i++) {
		if (document.getElementsByName('fromID')[i].checked) {
			found = true;
		}
	}
	if ( !found )
	{
		alert('Õ”«» „»œ« —« «‰ Œ«» ﬂ‰Ìœ');
		document.getElementsByName('fromID')[0].focus()
		return false;
	}

	if( !(parseFloat(document.getElementsByName('qtty')[0].value)>0) )
	{
		alert(' ⁄œ«œ —« Ê«—œ ﬂ‰Ìœ');
		document.getElementsByName('qtty')[0].focus();
		return false;
	}

	if(document.getElementsByName('comments')[0].value.length>220)
	{
		alert('ÿÊ· ›Ì·œ  Ê÷ÌÕ«  “Ì«œ «”  \n\nÕœ«ﬂÀ— 220 ( ﬁ—Ì»« 4 Œÿ Å—) „Ì  Ê«‰œ »«‘œ');
		document.getElementsByName('comments')[0].focus();
		return false;
	}

	return true;
}

var dialogActive=false;

function mask(src){ 
	var theKey=event.keyCode;

	if (theKey==13){
		event.keyCode=9
		dialogActive=true
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì  ›’Ì·Ì:"
		var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		if (document.all.tmpDlgTxt.value !="") {
			var myTinyWindow = window.showModalDialog('../ar/dialog_selectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
			dialogActive=false
			if (document.all.tmpDlgArg.value!="#"){
				Arguments=document.all.tmpDlgArg.value.split("#")
				src.value=Arguments[0];
				document.all.accountName.value=Arguments[1];
			}
		}
	}
}

function check(src){ 
	if (!dialogActive){
		badCode = false;
		if (window.XMLHttpRequest) {
			var objHTTP=new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		}
		objHTTP.open('GET','../accounting/xml_CustomerAccount.asp?id='+src.value,false)
		objHTTP.send()
		tmpStr = unescape(objHTTP.responseText)
		document.all.accountName.value=tmpStr;
		}
}

function hideIT()
{
//alert(document.all.aaa2.value)
if(document.all.aaa2.value==2) 
	{
		document.all.aaa1.style.visibility= 'visible'
		document.all.toID.value = ""
		document.all.toID.focus()
	}
	else
	{
		document.all.aaa1.style.visibility= 'hidden'
		document.all.toID.value = "-1"
	}
}

//-->
</SCRIPT>
	<%

	response.end
end if


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Search Form
'-----------------------------------------------------------------------------------------------------


if request("catItem") = "" then
	%><BR>
	<FORM METHOD=POST ACTION=""><BR><center>
		<INPUT TYPE="hidden" name="radif" value="-1">
		<input type="hidden" Name='tmpDlgArg' value=''>
		<input type="hidden" Name='tmpDlgTxt' value=''>
		 ﬂœ ﬂ«·« —« Ê«—œ ﬂ‰Ìœ:   &nbsp;&nbsp;&nbsp;<INPUT  dir="LTR"  TYPE="text" NAME="oldItemID" maxlength="10" size="13"   onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
		&nbsp;&nbsp;<INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ"  style="width:80px;" tabIndex="14">
		</center>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	document.all.oldItemID.focus()
	var dialogActive=false;

	function mask(src){ 
		var theKey=event.keyCode;

		if (theKey==13){
			event.keyCode=9
			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ﬂ«·«Â«Ì «‰»«—"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('dialog_selectInvItem.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				dialogActive=false
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					document.all.accountName.value=Arguments[1];
				}
			}
		}
	}

	function check(src){ 
		if (!dialogActive){
			badCode = false;
			if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
			objHTTP.open('GET','xml_InventoryItem.asp?id='+src.value,false)
			objHTTP.send()
			tmpStr = unescape(objHTTP.responseText)
			document.all.accountName.value=tmpStr;
			}
	}


	//-->
	</SCRIPT>

	<%
end if
%>

	<!--#include file="tah.asp" -->
