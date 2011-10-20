<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Purchase (4)
PageTitle="„—»Êÿ ﬂ—œ‰ Ê—ÊœÌ Â«Ì «‰»«— »Â ”›«—‘ Â«Ì Œ—Ìœ"
SubmenuItem=5
if not Auth(4 , 5) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border: 2px solid white; padding:0; direction: RTL; width:700px;}
	.RcpMainTable Tr {Height:25px; border: 1px solid black;}
	.RcpMainTable Input { font-family:tahoma; font-size: 9pt; border: 1px solid gray; text-align:right; direction: LTR;}
	.RcpMainTable Select { font-family:tahoma; font-size: 9pt;}

	.tblSearch {border: 1 solid #330099; padding:3; direction: RTL;}
	.tblSearch th {font-family:tahoma; font-size: 9pt; border-bottom: solid 1pt black; backGround-Color:#CCCC; text-align:center; font-weight:normal;}
	.tblSearch td {backGround-Color:#EEEEEE; text-align:right;height:25px;}
	.tblSearch input {border: 1px solid gray; }
	.tblSearch select {font-family:tahoma; font-size: 8pt; border: 1 solid gray; width:200px; height:50px;}
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
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
			var myTinyWindow = window.showModalDialog('../inventory/dialog_selectInvItem.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
			dialogActive=false
			if (document.all.tmpDlgArg.value!="#"){
				Arguments=document.all.tmpDlgArg.value.split("#")
				src.value=Arguments[0];
				document.all.itemName.value=Arguments[1];
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
		objHTTP.open('GET','../inventory/xml_InventoryItem.asp?id='+src.value,false)
		objHTTP.send()
		tmpStr = unescape(objHTTP.responseText)
		document.all.itemName.value=tmpStr;
		}
}

function maskVendor(src){ 
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

function checkVendor(src){ 
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
		//document.all['A1'].innerText= objHTTP.status
		//document.all['A2'].innerText= objHTTP.statusText
		//document.all['A3'].innerText= objHTTP.responseText
		document.all.accountName.value=tmpStr;
		}
}


function setColor(obj)
{
	//alert(document.all.purchaseOrderID.length)
	for(i=0; i<document.all.purchaseOrderID.length; i++)
		{
		theTR = document.all.purchaseOrderID[i].parentNode.parentNode
		theTR.setAttribute("bgColor","<%=AppFgColor%>")
		}
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor","#FFFFFF")
}


//-->
</SCRIPT>

<%

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------- Step2 : Find Related Perchuse Order
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="œÊŒ ‰ «Ì‰ ”›«—‘ Œ—Ìœ »Â ¬‰ ÕÊ«·Â Ê—Êœ »Â «‰»«—"then
	purchaseOrderID	= request.form("purchaseOrderID")
	id				= request.form("id")
	goodName		= request.form("goodName")

	if 	id = "" or purchaseOrderID="" or not (isnumeric(purchaseOrderID) and isnumeric(id)) then
		
		response.redirect "?errmsg="& Server.URLEncode("Œÿ«! ÂÌç ”›«—‘ Œ—ÌœÌ «‰ Œ«» ‰‘œÂ «” .&Submit=«‰ Œ«»&id=" & id)
	end if
	
	type1 = 1 
	if purchaseOrderID < 0 then type1 = -1 * purchaseOrderID

	mySql="UPDATE InventoryLog set RelatedID = "& purchaseOrderID & ", type ="& type1 & " where id=" & id
	'response.write "<br>" & mySql
	'response.end
	conn.Execute mySql

	response.redirect "?msg="& Server.URLEncode("Ê—Êœ »Â «‰»«— " & goodName & " »Â ”›«—‘ Œ—Ìœ ‘„«—Â "& purchaseOrderID & " œÊŒ Â ‘œ")
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------- Step2 : Find Related Perchuse Order
'-----------------------------------------------------------------------------------------------------
elseif request("act")="relate" or request("Submit")="Ã” ÃÊ" then

	id				= request("id")
	ChequeDatesFrom	= request.form("ChequeDatesFrom")
	ChequeDatesTo	= request.form("ChequeDatesTo")
	purchOrderID	= request.form("purchOrderID")
	accountID		= request.form("accountID")
	status			= request.form("status")
	item			= request.form("item")

	if id = "" or not isNumeric(id) then
		response.redirect "?msg="& Server.URLEncode("Œÿ«! ÂÌç ﬂ«·«ÌÌ «‰ Œ«» ‰‘œÂ «” .")
	end if 

	set RS3=Conn.Execute ("SELECT InventoryLog.ID, InventoryLog.ItemID, InventoryLog.Qtty, InventoryLog.logDate, InventoryLog.CreatedBy, InventoryLog.comments, InventoryItems.OldItemID, InventoryItems.Name, InventoryItems.Unit, InventoryLog.type FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID WHERE (InventoryLog.id = "& id & ")")	

	response.write "<center><br><br><B>ê«„ œÊ„:</B> Ì«› ‰ ”›«—‘ Œ—Ìœ „—»ÊÿÂ"
	if item="" or not isnumeric(item) then
		item = RS3("OldItemID")
	end if
	%>
	<BR><BR>
	<FORM METHOD=POST ACTION="?">
	<INPUT TYPE="hidden" name="ID" value="<%=id%>">
	<INPUT TYPE="hidden" name="goodName" value="<%=Name%>">
	<TABLE dir=rtl align=center width=700 cellspacing=0 style="background-color:white">
	<TR>
		<TD><B>ﬂ«·«Ì Ê«—œ ‘œÂ »Â «‰Ì«—: </B></TD>
		<!--TD><%=RS3("id")%></TD-->
		<TD dir=ltr align=right>(<%=RS3("logDate")%>)</TD>
		<TD><%=RS3("Name")%></TD>
		<TD><%=RS3("Qtty")%></TD>
		<TD><%=RS3("Unit")%></TD>
		<TD><%=RS3("comments")%> &nbsp;</TD>
	</TR>
	</TABLE>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<TABLE align=center class="RcpMainTable">
	<TR style="height:10pt">
		<td></td>
	</TR>
	<TR>
		<TD align=right> «“  «—ÌŒ : <INPUT dir="LTR" TYPE="text" NAME="ChequeDatesFrom" maxlength="10" size="10"onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=ChequeDatesFrom%>"></TD>
		<TD align=left>
			Õ”«» ›—Ê‘‰œÂ : <INPUT dir="LTR" TYPE="text" NAME="accountID" maxlength="10" size="10" value="<%=accountID%>" onKeyPress='return maskVendor(this);' onBlur='checkVendor(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=25 readonly value="<%=accountName%>" style="background-color:transparent">	</TD>
		<TD align=left>
			Ê÷⁄Ì  :ù 
			<select name="status" style="font-size:8pt">
			<option value="">«‰ Œ«» ﬂ‰Ìœ</option>
			<option value="">---------------</option>
			<option value="OK" <% if status="OK" then %> selected <% end if %>> «ÌÌœ ‘œÂ</option>
			<option value="OUT" <% if status="OUT" then %> selected <% end if %>>Œ«—Ã «“ ‘—ﬂ </option>
			<option value="CANCEL" <% if status="CANCEL" then %> selected <% end if %>>·€Ê ‘œÂ</option>
			</select>
		</TD>
	</TR>
	<TR>
		<TD align=right> «  «—ÌŒ :ù <INPUT dir="LTR" TYPE="text" NAME="ChequeDatesTo" maxlength="10" size="10" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" value="<%=ChequeDatesTo%>"></TD>
		<TD align=left>
			ﬂœ ﬂ«·«Ì «‰»«— : 
			<INPUT dir="LTR" TYPE="text" NAME="item" maxlength="10" size="10" value="<%=item%>" onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="itemName" size=25 readonly value="<%=accountName%>" style="background-color:transparent">
		</TD>
		<TD align=left>
		‘„«—Â : <INPUT TYPE="text" NAME="purchOrderID" size=4 value="<%=purchOrderID%>">
		<INPUT class="GenButton" TYPE="submit" name="submit" value="Ã” ÃÊ" ></TD>
	</TR>
	<TR>
	</TABLE><BR>
	</FORM>
	<FORM METHOD=POST ACTION="?">
	<INPUT TYPE="hidden" name="ID" value="<%=id%>">
	<TABLE dir=rtl align=center width=700 cellspacing=0>
	<tr>
		<TD align=right colspan=3>
			<INPUT TYPE="radio" onclick="setColor(this)" NAME="purchaseOrderID" value="-3"> ﬂ«·«Ì „—ÃÊ⁄Ì 
		</TD>
	</TR>
	<tr>
		<TD align=right colspan=3>
			<INPUT TYPE="radio" onclick="setColor(this)" NAME="purchaseOrderID" value="-6"> Ê—Êœ «“  Ê·Ìœ
		</TD>
	</TR>
	<tr>
		<TD align=right colspan=3>
			<INPUT TYPE="radio" onclick="setColor(this)" NAME="purchaseOrderID" value="-7"> Ê—Êœ «“ «‰»«— ‘Â—Ì«— / œ› — ⁄»«” ¬»«œ<br>
		</TD>
	</TR>
	</TABLE><BR>
	<%
	'-------------------------------------------------------------------------------------------------
	'----------------------------------------------------------------------------- Step3 : Submit form
	'-------------------------------------------------------------------------------------------------
	'SELECT PurchaseOrders.*
	'FROM PurchaseOrders LEFT OUTER JOIN
	'InventoryLog ON PurchaseOrders.ID = InventoryLog.RelatedID
	'WHERE (ISNULL(InventoryLog.RelatedID, 0) = 0) AND (PurchaseOrders.IsService = 0)
	
	conditions = "(PurchaseOrders.IsService = 0)"

	if ChequeDatesFrom <> "" then
		conditions = conditions & " and OrdDate >= N'"& ChequeDatesFrom & "' "
	end if

	if ChequeDatesTo <> "" then
		conditions = conditions & " and OrdDate <= N'"& ChequeDatesTo & "' "
	end if

	if status <> "" then
		conditions = conditions & " and PurchaseOrders.Status = N'"& status & "' "
	end if

	if purchOrderID <> "" then
		if isNumeric(purchOrderID) then
			conditions = conditions & " and PurchaseOrders.ID = "& purchOrderID 
		end if
	end if

	if accountID <> "" then
		if isNumeric(accountID) then
			conditions = conditions & " and PurchaseOrders.Vendor_ID = "& accountID 
		end if
	end if

	if item <> "" then
		if isNumeric(item) then
			conditions = conditions & " and InventoryItems.OldItemID = "& item 
		end if
	end if

	conditions = conditions & " order by PurchaseOrders.ID DESC "

	'if request.form("Submit")="Ã” ÃÊ" then

		'mySQL="SELECT PurchaseOrders.*, Accounts.AccountTitle FROM PurchaseOrders INNER JOIN InventoryItems ON PurchaseOrders.TypeID = InventoryItems.ID INNER JOIN Accounts ON PurchaseOrders.Vendor_ID = Accounts.ID WHERE " & conditions
		'Changed By Kid ! 840219
		mySQL="SELECT DRV.ID AS LogID,DRV.Qtty AS LogQtty, PurchaseOrders.*, Accounts.AccountTitle FROM PurchaseOrders INNER JOIN InventoryItems ON PurchaseOrders.TypeID = InventoryItems.ID INNER JOIN Accounts ON PurchaseOrders.Vendor_ID = Accounts.ID LEFT OUTER JOIN (SELECT ID, RelatedID, Qtty, type FROM InventoryLog WHERE (IsInput = 1) AND (Voided = 0)) DRV ON PurchaseOrders.ID = DRV.RelatedID WHERE " & conditions

		set RS3=Conn.Execute (mySQL)

		%>
		<TABLE dir=rtl align=center cellspacing=1 cellpadding=2 bgcolor=0 style='border:1 solid black;'>
		<!--TR bgcolor="eeeeee" >
			<TD align=center colspan=6><B>Ê—Êœ »Â «‰»«— »« „‰‘« ‰« „‘Œ’</B></TD>
		</TR-->
		<TR bgcolor="CCCCCC" height="25px">
			<TD width="80px"># ”›«—‘ Œ—Ìœ</TD>
			<TD width="65px"> «—ÌŒ</TD>
			<TD width="150px">›—Ê‘‰œÂ</TD>
			<TD width="100px">‰«„ ﬂ«·«</TD>
			<TD width="40px"> ⁄œ«œ</TD>
			<TD width="60px">ﬁÌ„ </TD>
			<TD width="100px"> Ê÷ÌÕ</TD>
			<TD width="50px">Ê—Êœ ﬁ»·Ì</TD>
		</TR>
		<%
		Do while not RS3.eof
			%>
			<TR bgColor="FFFFFF" style="cursor:hand" onclick="this.getElementsByTagName('input')[0].click();this.getElementsByTagName('input')[0].focus();">	
				<TD><INPUT TYPE="radio" NAME="purchaseOrderID" onclick="setColor(this)" VALUE="<%=RS3("ID")%>"> <A target="_blank" HREF="outServiceTrace.asp?od=<%=RS3("ID")%>"><%=RS3("ID")%></A> </TD>		
				<TD><div style="direction:ltr;text-align:center;"><%=RS3("OrdDate")%></div></TD>
				<TD><%=RS3("AccountTitle")%></TD>
				<TD><%=RS3("TypeName")%></TD>
				<TD><%=RS3("Qtty")%></TD>
				<TD><%=Separate(RS3("Price"))%></TD>
				<TD><%=RS3("Comment")%> &nbsp;</TD>
				<TD><A HREF="../inventory/invReport.asp?oldItemID=<%=item%>&logRowID=<%=RS3("LogID")%>" target="_blank"><%=RS3("LogID")%></A> &nbsp; <%=RS3("LogQtty")%></TD>
			</TR>
			<%
		RS3.moveNext
		Loop
		%>
		</TABLE><BR>
		<%
	'end if
	%>
	<BR><CENTER><INPUT TYPE="submit" name="submit" value="œÊŒ ‰ «Ì‰ ”›«—‘ Œ—Ìœ »Â ¬‰ ÕÊ«·Â Ê—Êœ »Â «‰»«—" class="genButton"></CENTER>
	</FORM>
	<%
	response.write "</center><br>"
	response.end
else
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------- Step1 : List Unknown Inventory Inputs
'-----------------------------------------------------------------------------------------------------

	myAnd=" AND "
	myCriteria= ""
	if request("chekDate")="on" then
		FromDate =	sqlSafe(request("FromDate"))
		if FromDate<>"" then
			myCriteria= myCriteria & " AND (InventoryLog.logDate >= '" & FromDate & "')"
		end if
		ToDate =	sqlSafe(request("ToDate"))
		if ToDate <> "" then
			myCriteria= myCriteria & " AND (InventoryLog.logDate <= '" & ToDate & "')"
		end if
	end if

	if request("chekItem")="on" then
		ItemName =	sqlSafe(request("ItemName"))
		if ItemName <> "" then
			myCriteria= myCriteria & " AND (REPLACE(InventoryItems.Name,' ','') LIKE N'%" & replace(ItemName," ","")& "%')"
		end if
		ItemCat =	request("ItemCat")
		if ItemCat <> "-1" then
			myCriteria= myCriteria & " AND (InventoryItemCategoryRelations.Cat_ID='" & ItemCat & "')"
		end if
	end if

	if request("chekDesc")="on" then
		ItemDesc =	request("ItemDesc")
		if ItemDesc <> "" then
			myCriteria= myCriteria & " AND (REPLACE(InventoryLog.comments,' ','') LIKE N'%" & replace(ItemDesc," ","")& "%')"
		end if
	end if

	if request("ResultsInPage")="" then
		ResultsInPage= 20
	else
		ResultsInPage =	cint(request("ResultsInPage"))
	end if

	order = "logDate DESC"


	%>
	<br>
	<hr>
	<TABLE align=center class='tblSearch' cellspacing=1 bgcolor=#0	>
	<FORM METHOD=POST ACTION="?">
	<TR bgcolor=#CCCCCC>
		<TH colspan=2><INPUT TYPE="checkbox" NAME="chekDate" style='border:none;' <%if request("chekDate")="on" then response.write "checked"%>> «—ÌŒ</TD>
		<TH colspan=2><INPUT TYPE="checkbox" NAME="chekItem" style='border:none;' <%if request("chekItem")="on" then response.write "checked"%>>ﬂ«·«</TD>
		<TH colspan=2><INPUT TYPE="checkbox" NAME="chekDesc" style='border:none;' <%if request("chekDesc")="on" then response.write "checked"%>> Ê÷ÌÕ« </TD>
	</TR>
	<TR bgcolor=#F0F0F0>
		<TD>«“:</TD>
		<TD>
			<INPUT dir="LTR" TYPE="text" NAME="FromDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=FromDate%>">
		</TD>
		<TD>‰«„ </TD>
		<TD>
			<INPUT TYPE="text" NAME="ItemName" maxlength="50" style="width:200px;" value="<%=ItemName%>"></TD>
		<TD>‘—Õ</TD>
		<TD>
			<INPUT TYPE="text" NAME="ItemDesc" maxlength="50" style="width:150px;" value="<%=ItemDesc%>"></TD>
	</TR>
	<TR bgcolor=#F0F0F0>
		<TD> «:</TD>
		<TD align=left>
			<INPUT dir="LTR" TYPE="text" NAME="ToDate" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=ToDate%>">
		</TD>
		<TD>ê—ÊÂ</TD>
		<TD>
			<SELECT NAME="ItemCat">
			<option value="-1">«‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">-------------</option>
<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories")
				while not (RS4.eof) %>
					<OPTION value="<%=RS4("ID")%>"<% if trim(ItemCat) = trim(RS4("ID")) then response.write " selected " %>><%=RS4("Name")%> </option>
<%					RS4.MoveNext
				wend
				RS4.close
				Set RS4 = Nothing
				%>
			</SELECT>
		</TD>
		<TD> ⁄œ«œ</TD>
		<TD>
			<INPUT TYPE="text" NAME="ResultsInPage" maxlength="3" style="width:50px;" Value="<%=ResultsInPage%>">
			<INPUT TYPE="submit" Name="Search" Value="Ã” ÃÊ" maxlength="50" style="width:95px;border:1 solid black;">
		</TD>
	</TR>
	</FORM>
	</TABLE>
	<hr>
	<FORM METHOD=POST ACTION="?act=relate">
	<TABLE width=600 align=center class='tblSearch' cellspacing=1 bgcolor=#0>
	<TR bgcolor="eeeeee" >
		<TD align=center colspan=6 height=30><B>Ê—Êœ »Â «‰»«—Â«Ì »« „‰‘« ‰«„⁄·Ê„</B></TD>
	</TR>
	<TR bgcolor="CCCCCC" height=30>
		<TH style="width:70px"> «—ÌŒ</TD>
		<TH style="width:200px">‰«„ ﬂ«·«</TD>
		<TH style="width:50px" > ⁄œ«œ</TD>
		<TH style="width:40px" >Ê«Õœ</TD>
		<TH style="width:150px"> Ê÷ÌÕ</TD>
	</TR>
	<%
	'InventoryLog.type = 1			:	Normal 
	'InventoryLog.RelatedID = -1	:	Not Related
	'InventoryLog.owner = -1		:	PDHCo.
	'InventoryLog.Voided = 0		:	not Voided
	'InventoryLog.IsInput = 1		:	!!
	'
	mySQL="SELECT TOP " & ResultsInPage & " InventoryLog.ID, InventoryLog.ItemID, InventoryLog.Qtty, InventoryLog.logDate, InventoryLog.comments, InventoryItems.OldItemID, InventoryItems.Name, InventoryItems.Unit, InventoryLog.type FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID WHERE (InventoryLog.RelatedID = - 1) AND (InventoryLog.type = 1) AND (InventoryLog.owner = - 1) AND (InventoryLog.Voided = 0) AND (InventoryLog.IsInput = 1) " & myCriteria & " ORDER BY " & order
	set RS3=Conn.Execute (mySQL)	
	if RS3.eof then
%>
		<TR>	
			<TD height=25 align=center bgcolor=#6699CC colspan=6><B>ÂÌç </B></TD>
		</TR>
	
<%
	else
		Do While NOT RS3.eof
%>
			<TR style="cursor:hand" onclick="go2('<%=RS3("id")%>');">	
				<TD dir='LTR'><%=RS3("logDate")%></TD>
				<TD><%=RS3("Name")%></TD>
				<TD><%=RS3("Qtty")%></TD>
				<TD><%=RS3("Unit")%></TD>
				<TD><%=RS3("comments")%> &nbsp;</TD>
			</TR>
<%
		RS3.moveNext
		Loop
	End if
%>
	</TABLE><BR><BR>	
	<CENTER><INPUT TYPE="submit" name="submit" value="«‰ Œ«»" class="genButton" style="width:50pt"></CENTER>
	</FORM>
	<%
	%>

<SCRIPT LANGUAGE="JavaScript">
<!--
function go2(id)
{

	theTR = event.srcElement.parentNode
	theTDs= theTR.getElementsByTagName('TD')
	for(i=0;i<theTDs.length;i++){
		theTDs[i].style.backgroundColor='#FFDDBB';
	}
	window.open('relateInvItemsToPurchOrders.asp?act=relate&id='+id); 
}


//-->
</SCRIPT>
<%
end if
%>
<!--#include file="tah.asp" -->
