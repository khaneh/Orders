<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " ���� ����"
SubmenuItem=3
if not Auth(5 , 3) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->

<BR>
<%
catItem1 = request("catItem")
goodItem1 = request("item")
owner = request("owner")
if not isNumeric(goodItem1) then
	response.write "<br>" 
	call showAlert ("���! �� ���� ���� ��� ����",CONST_MSG_ERROR) 
	response.write "<br>" 
	response.end
end if
purchaseOrderID = request("purchaseOrderID")
if catItem1="" then 
	'catItem1="-1"
	'goodItem1="-1"
	'purchaseOrderID="-1"
elseif goodItem1="" then 
	goodItem1="-1"
	purchaseOrderID="-1"
	owner = "-1"
elseif purchaseOrderID="" then 
	purchaseOrderID="-1"
end if


'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------- Show Inventory Log In Reciept
'-----------------------------------------------------------------------------------------------------
if request("act")="showReceipt" then

	item = request("item")
	qtty = request("qtty")
	response.write "<br><br><br><br><center>"

	set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (OldItemID = "& item & ")" )
	ItemID = RSW("id")
	
	
	oldItemQtty = RSW("qtty")
	newItemQtty = oldItemQtty ' + qtty
	response.write "<li> " & RSW("Name") & " " & qtty & " " & RSW("unit") & " (������ ����: " & newItemQtty & " " & RSW("unit") & ")"

	response.write "<br><br>����� ��� �� ����� ����� ��."
	
	set RSW=Conn.Execute ("SELECT * FROM InventoryLog WHERE (ItemID = "& ItemID & " and Qtty="& Qtty & ") order by id DESC" )
	LogItem_ID = RSW("id")

	catItem1 = "-1"
	goodItem1 = "-1"
	purchaseOrderID = "-1"

	PickupListID=54
	%>
	<BR>	<BR>
	<CENTER>
	<% 	ReportLogRow = PrepareReport ("InventoryLogIn.rpt", "LogItem_ID", LogItem_ID, "/beta/dialog_printManager.asp?act=Fin") %>
	<INPUT TYPE="button" value=" �ǁ ����" Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

	<input type="button" value="���� ����� ����" Class="GenButton" onclick="window.location='itemIn.asp'">

	</CENTER>

	<BR>	
	<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:700; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no></iframe>
	<%

	response.end

end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Voroode Piramon Marjooe submit
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="���� ������� ��� ������" then

	invoice_id = trim(request.form("invoice_id"))
	if not isnumeric(invoice_id) then 
		response.write "<br><br>"
		call showAlert( "����� ������ ���� ��� ������ ���." , CONST_MSG_ERROR)
		response.end
	end if 

	'vvvvvvv ------------------------------------------ start of check for current ItemOUT
	set RSS=Conn.Execute ("SELECT InventoryLog.id, InventoryLog.comments, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.VoidedDate, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.type, InventoryLog.ID, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE (InventoryLog.RelatedInvoiceID = "& invoice_id & ") AND (InventoryLog.IsInput = 1)")
	if not RSS.EOF then
		%>
		<br><br>
		<%
		call showAlert("���� ����� ��� ��� ���� �� �����  ���� ��� ���", CONST_MSG_ALERT)
		response.write "<br><br>" 
		%>
		<TABLE dir=rtl align=center width=600>
			<TR bgcolor="eeeeee" >
				<TD><SMALL>�� ����</SMALL></TD>
				<TD width=200><SMALL>��� ����</SMALL></TD>
				<TD><SMALL>����� </SMALL></TD>
				<TD><SMALL>����</SMALL></TD>
				<TD><SMALL>����� ����</SMALL></TD>
				<TD align=center><SMALL>����� �����</SMALL></TD>
				<TD><SMALL>����</SMALL></TD>
			</TR>
		<%
		tmpCounter=0
		do while not RSS.EOF
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

			%>
			<TR bgcolor="<%=tmpColor%>"  style="height:25pt" <% if RSS("voided") then%> disabled title="��� ��� �� ����� <%=RSS("VoidedDate")%>"<% end if %>>
				<TD  align=right dir=ltr><INPUT TYPE="hidden" name="id" value="<%=RSS("ID")%>"><A HREF="invReport.asp?oldItemID=<%=RSS("oldItemID")%>&logRowID=<%=RSS("ID")%>" target="_blank"><%=RSS("OldItemID")%></A></TD>
				<TD><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><!A HREF="default.asp?itemDetail=<%=RSS("ID")%>"><span style="font-size:10pt"><%=RSS("Name")%></A></TD>
				<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("Qtty")%></span></TD>
				<TD align=right dir=ltr><%=RSS("Unit")%></TD>
				<TD dir=ltr><%=RSS("logDate")%></span></TD>
			<TD align=center><% if RSS("type")= "1" and RSS("RelatedID")<> "-1" then
					%>
						����� ���� 
				     <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A>
				   <%
				   elseif RSS("type")= "2" then
					response.write "<font color=red><b>����� ������</b></font>"
				   elseif RSS("type")= "3" then
					response.write "<font color=green><b>������</b></font>"
					elseif RSS("type")= "4" then
					response.write "<font color=blue><b>����� ����� ���� </b></font>"
					elseif RSS("type")= "5" then
					response.write "<font color=orang><b>������</b></font>"
					elseif RSS("type")= "6" then
					response.write "<font color=#6699CC><b>���� �� �����</b></font>"
					elseif RSS("type")= "7" then
					response.write "<font color=#FF9966><b>���� �� ����� ������</b></font>"
				   else 
					response.write " "
				   end if	%>
				   
				<% if RSS("owner")<> "-1" and RSS("owner")<> "-2" then
					response.write " (������ <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& RSS("owner") &"' target='_blank'> " & RSS("owner") & "</a>)"
				   end if %>
				<% if RSS("comments")<> "-" and RSS("comments")<> "" then
					response.write " <br><br><B>�����:</B>  " & RSS("comments") 
				   end if %>
				   </TD>
				<TD><%=RSS("RealName")%></TD>
			</TR>
				  
			<% 
			RSS.movenext
		loop
		response.write "</table><br><br>" 
		response.end

	end if
	RSS.close
	set RSS = nothing
	'^^^^^^^ ------------------------------------------ end of check for current ItemOUT


	set RSS=Conn.Execute ("SELECT dbo.InventoryItems.Name, dbo.InvoiceLines.AppQtty, dbo.InventoryItems.ID as itemID, dbo.InvoiceItems.RelatedInventoryItemID, dbo.InventoryItems.Unit FROM dbo.InvoiceLines INNER JOIN dbo.InvoiceItems ON dbo.InvoiceLines.Item = dbo.InvoiceItems.ID INNER JOIN dbo.InventoryItems ON dbo.InvoiceItems.RelatedInventoryItemID = dbo.InventoryItems.OldItemID WHERE (dbo.InvoiceLines.Invoice = " & invoice_id & ")")
	st = ""
	response.write "<BR><BR><CENTER>��� ���� ������� ��� ������ ��� �� ����� �� ����� �� ���Ͽ<BR><BR>"

	do while not RSS.EOF
		response.write "<li> " & RSS("AppQtty") & " " & RSS("unit") & " " & RSS("name") 
		RSS.movenext
	loop
	%>
	<FORM METHOD=POST ACTION="">
		<INPUT TYPE="hidden" name="invoice_id" value="<%=invoice_id%>">
		<INPUT TYPE="submit" Name="Submit" Value="����� ���� ������� ��� ������" class=inputBut style="width:170px;" >
		<INPUT TYPE="submit" Name="Submit" Value="������" class=inputBut style="width:70px;" >
	</FORM>
	<%
	response.end
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Voroode Piramon Marjooe submit
'-----------------------------------------------------------------------------------------------------
elseif request.form("Submit")="����� ���� ������� ��� ������" then
	'response.write "<br><br>"
	'call showAlert("��� ���� ���� ����� ����.",CONST_MSG_ERROR)
	'response.end

	invoice_id = trim(request.form("invoice_id"))
	if not isnumeric(invoice_id) then 
		response.write "<br><br>"
		call showAlert( "����� ������ ���� ��� ������ ���." , CONST_MSG_ERROR)
		response.end
	end if 

	set RSS=Conn.Execute ("SELECT dbo.InventoryItems.Name, dbo.InvoiceLines.AppQtty, dbo.InventoryItems.ID as itemID, dbo.InvoiceItems.RelatedInventoryItemID, dbo.InventoryItems.Unit FROM dbo.InvoiceLines INNER JOIN dbo.InvoiceItems ON dbo.InvoiceLines.Item = dbo.InvoiceItems.ID INNER JOIN dbo.InventoryItems ON dbo.InvoiceItems.RelatedInventoryItemID = dbo.InventoryItems.OldItemID WHERE (dbo.InvoiceLines.Invoice = " & invoice_id & ")")
	st = ""
	response.write "<BR><BR><CENTER>������� ��� ������ ��� ���� ����� ��:<BR>"

	do while not RSS.EOF
		mysql = "INSERT INTO dbo.InventoryLog (ItemID, RelatedID, Qtty, logDate, owner, CreatedBy, IsInput, type, RelatedInvoiceID) VALUES (" & RSS("itemID") & ", -3, " & RSS("AppQtty") & ", N'" & shamsiToday() & "', -1, " & session("id") & ", 1, 3, " & invoice_id & ")"
		response.write "<li> " & RSS("AppQtty") & " " & RSS("unit") & " " & RSS("name") 
		Conn.Execute (mysql)
		'response.write "<br>" & mysql
		RSS.movenext
	loop
	response.end

end if
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Submit an Inventory Item Input 3
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="���� ���� �� �����" then
	item = trim(request.form("item"))
	qtty = trim(request.form("qtty"))
	purchaseOrderID = trim(request.form("purchaseOrderID"))
	ownerAcc = request.form("accountID")
	comments = request.form("comments")
	entryDate = request.form("entryDate")

	if Auth(5 , "C") AND "" & entryDate <> "" then ' ��� ����/���� �� ����� ������
		logDate=entryDate
	else
		logDate=shamsiToday()
	end if
	
	if 	not item = "" then

		if purchaseOrderID="" then
			purchaseOrderID = "-1"
		end if

		if item="-1" or qtty="" or qtty="0" then
			response.write "<br><br><br>"
			CALL showAlert ("<B>���! </B><BR>�� ������ ������ ����� ���<br><br><A HREF='itemIn.asp'>�ѐ��</A>",CONST_MSG_ALERT) 
			response.end
		end if

		if qtty<0 then
			response.write "<br><br><br>"
			CALL showAlert ("<B>���! </B><BR>����� ���� ��� ����� ���� ����<br><br><A HREF='itemIn.asp'>�ѐ��</A>",CONST_MSG_ALERT) 
			response.end
		end if

		response.write "<br><br><br><br><center>"

		set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (OldItemID = "& item & ")" )
		ItemID = RSW("id")
		type1 = 1 
		if purchaseOrderID < 0 then type1 = -1 * purchaseOrderID
		mySql="INSERT INTO InventoryLog (ItemID, RelatedID, logDate, Qtty, owner, CreatedBy, IsInput, comments, type) VALUES ("& ItemID & ", "& purchaseOrderID & ",N'"& logDate & "', "& Qtty & ", "& ownerAcc & ", "& session("id") & ", 1, N'" & comments & "', " & type1 & " )"
		conn.Execute mySql
		'RS1.close
		
		response.redirect "ItemIn.asp?act=showReceipt&item=" & item & "&qtty=" & qtty

	end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Submit an Inventory Item Input 2
'-----------------------------------------------------------------------------------------------------
elseif request.form("Submit")="������" then
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function hideIT()
{
//alert(document.all.aaa2.value)
if(document.all.aaa2.value==2) 
	{
		document.all.aaa1.style.visibility= 'visible'
		document.all.accountID.value = ""
		document.all.accountID.focus()
	}
	else
	{
		document.all.aaa1.style.visibility= 'hidden'
		document.all.accountID.value = "-1"
	}
}
//-->
</SCRIPT>
<%
purchaseOrderID = request.form("purchaseOrderID")
if purchaseOrderID = "" then purchaseOrderID = -1
'set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (Status = 'OUT' and ID="& purchaseOrderID & ")" )
set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (ID="& purchaseOrderID & ")" )
if not (RSA.eof) then
	preQtty = RSA("Qtty")
end if
RSA.close

%>
<FORM METHOD=POST ACTION="itemin.asp">
<INPUT TYPE="hidden" name="item" value="<%=request.form("item")%>">
<INPUT TYPE="hidden" name="purchaseOrderID" value="<%=request.form("purchaseOrderID")%>"><BR><BR>
<TABLE border=0 align=center>
<TR>
	<TD align=left valign=top>��� ����</TD>
	<TD align=right><span disabled><%=request("goodName")%></span><br><br></TD>
</TR>

<TR>
	<TD align=left>�����</TD>
	<TD align=right><INPUT dir=ltr TYPE="text" NAME="qtty" size=25 value="<%=preQtty%>" ><!onKeyPress="return maskNumber(this);" ><br></TD>
</TR>
<TR>
	<TD align=left valign=top>������</TD>
	<TD align=right>
	<SELECT NAME="aaa2" onchange="hideIT()" >
		<option value=1 >���� �ǁ � ���</option>
		<option value=2 <%if owner<>"-1" then response.write " selected" %>>����� (����� ����)</option>
	</SELECT>
	<BR>
	<span name="aaa1" id="aaa1" <% if owner="-1" then response.write "style=""visibility:'hidden'"""%>>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<INPUT  dir="LTR"  TYPE="text" NAME="accountID" maxlength="10" size="13"  value="<%=owner%>" onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
	</span></TD>
</TR>
<%
	if Auth(5 , "C") then ' ��� ����/���� �� ����� ������
%>
	<TR>
		<TD align=left>�����</TD>
		<TD align=right><INPUT dir=ltr TYPE="text" NAME="entryDate" size=25 value="<%=shamsiToday()%>" onblur="acceptDate(this)"></TD>
	</TR>
<%
	End if

%>
<TR>
	<TD align=left>�������</TD>
	<TD align=right><TEXTAREA NAME="comments" ROWS="" COLS=""></TEXTAREA>
	<br><br></TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=right></TD>
</TR>
<TR>
	<TD align=center colspan=2>
	<INPUT TYPE="submit" Name="Submit" Value="���� ���� �� �����" class=inputBut style="width:120px;" tabIndex="14"<%
	if not goodItem1<>"-1" then
		response.write " disabled "
	end if
	%>>
	</TD>
</TR>
</TABLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.qtty.focus()

var dialogActive=false;

function mask(src){ 
	var theKey=event.keyCode;

	if (theKey==13){
		event.keyCode=9
		dialogActive=true
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="����� �� ��� ���� ��� ������:"
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


//-->
</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Submit an Inventory Item Input 1
'-----------------------------------------------------------------------------------------------------
elseif request.form("Submit")="�����" then
%>
<FORM METHOD=POST ACTION="itemin.asp">
<INPUT TYPE="hidden" name="item" value="<%=request.form("item")%>">
<TABLE border=0 align=center>
<%
				set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (OldItemID = "& goodItem1 & ")" )
				if RSW.EOF then
					call showAlert ("���! �� ���� ����� ����" , CONST_MSG_ERROR )
					response.end
				end if 
				goodItem1 = RSW("id")
				goodUnit = RSW("unit")
				goodName = RSW("name")
				owner = RSW("owner")
%>
<TR>
	<TD align=right>
	<span disabled><%=goodName%></span><BR>
	<BR>
	����� ���� �� ������ ����:<br>
	<br></TD>
</TR>
<INPUT TYPE="hidden" name="goodName" value="<%=goodName%>">
<TR>
	<TD align=right>
			
<%
				set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (Status = 'OUT' and TypeID="& goodItem1& ") order by OrdDate" )
				'set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (TypeID="& goodItem1& ") order by OrdDate" )
				flg = false
				while not (RSA.eof) %>
					<INPUT TYPE="radio" NAME="purchaseOrderID" value="<%=RSA("ID")%>"<%
					if trim(purchaseOrderID) = trim(RSA("ID")) then
						response.write " chcecked "
						preQtty = RSA("Qtty")
						flg = true
					end if
					%>> ����� <%=RSA("ID")%> (<%=RSA("Qtty")%>&nbsp;<%=goodUnit%>&nbsp;<%=goodName%>) <BR>
<%					RSA.MoveNext
				wend
				RSA.close
				%>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-1" <% if not flg then%>checked<% end if%>> ����� �� ����� ���� ���� ����.<br>

				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-3"> ����� ������ <br>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-6"> ���� �� �����<br>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-7"> ���� ��  ����� ������ / ���� ���� ����<br>
				<% if Auth(5 , 7) then %>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-2"> ����� ������<br>
				<% end if %>
			<br><br></TD>
</TR>
<TR>
	<TD align=center colspan=2>
	<INPUT TYPE="hidden" name="owner" value="<%=owner%>">
	<INPUT TYPE="submit" Name="Submit" Value="������" class=inputBut style="width:120px;" tabIndex="14">
	</TD>
</TR>
</Table>	
</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.purchaseOrderID[0].focus()
//-->
</SCRIPT>
<%
else
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------  Inventory Item Input
'-----------------------------------------------------------------------------------------------------
%>
<TABLE border=0 align=center>
<TR>
	<TD colspan=2 align=center><H3>���� ����</H3></TD>
</TR>
<TR>
	<TD align=left>�� ����<br></TD>
	<TD align=right>
		<FORM METHOD=POST ACTION="itemin.asp">
		<input type="hidden" Name='tmpDlgArg' value=''>
		<input type="hidden" Name='tmpDlgTxt' value=''>
		<INPUT  dir="LTR"  TYPE="text" NAME="item" maxlength="10" size="13"  value="<%=owner%>" onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
	</TD>
</TR>
<TR>
	<TD align=center colspan=2>
	<INPUT TYPE="hidden" name="goodUnit" value="<%=goodUnit%>">
	<INPUT TYPE="hidden" name="goodName" value="<%=goodName%>">
	<INPUT TYPE="submit" Name="Submit" Value="�����" class=inputBut style="width:120px;" tabIndex="14">
	</TD>
</TR>
</TABLE>
</FORM>
<BR><BR>
<TABLE align=center width=50%>
<TR>
	<TD align=center style="border:solid 1pt black">
		<BR>
		<FORM METHOD=POST ACTION="">
		<B> ������� ������ <BR></B><BR>
		����� �ѐ�� �� ����: <INPUT TYPE="text" NAME="invoice_id" dir=ltr ><BR><BR> <INPUT TYPE="submit"  name="submit"  value="���� ������� ��� ������">
		<BR>
		</FORM>
	</TD>
</TR>
</TABLE><BR>

<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.item.focus()

var dialogActive=false;

function mask(src){ 
	var theKey=event.keyCode;

	if (theKey==13){
		event.keyCode=9
		dialogActive=true
		document.all.tmpDlgArg.value="#"
		document.all.tmpDlgTxt.value="����� �� ������� �����"
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
