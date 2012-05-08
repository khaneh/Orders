<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " Ê—Êœ ﬂ«·«"
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
	call showAlert ("Œÿ«! ﬂœ ﬂ«·« »«Ìœ ⁄œœ »«‘œ",CONST_MSG_ERROR) 
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
	response.write "<br><br><br><br><center>"
	if request("id")<>"" then 
		LogItem_ID=request("id")
		set rs=Conn.Execute("select inventoryLog.*,InventoryItems.unit,InventoryItems.name from inventoryLog inner join InventoryItems on inventoryLog.itemID=InventoryItems.id where inventoryLog.id=" & LogItem_ID)
		response.write "<li> " & rs("qtty") & " " & RS("unit") & " "  & RS("Name") & " " 
		if cdbl(rs("owner"))>0 then 
			response.write " Ê”ÿ " & rs("owner") & " "
		end if
		response.write " („ÊÃÊœÌ ÃœÌœ: " & rs("sumQtty") & " " & RS("unit") & "° «—”«·Ì „‘ —Ì: " & rs("sumCusQtty") & " " & rs("unit") &")"
		rs.close
		set rs=nothing
	else
		item = request("item")
		qtty = request("qtty")
		set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (OldItemID = "& item & ")" )
		ItemID = RSW("id")
		
		
		oldItemQtty = RSW("qtty")
		newItemQtty = oldItemQtty ' + qtty
		response.write "<li> " & RSW("Name") & " " & qtty & " " & RSW("unit") & " („ÊÃÊœÌ ÃœÌœ: " & newItemQtty & " " & RSW("unit") & ")"

		set RSW=Conn.Execute ("SELECT * FROM InventoryLog WHERE (ItemID = "& ItemID & " and Qtty="& Qtty & ") order by id DESC" )
		LogItem_ID = RSW("id")
	end if
	response.write "<br><br>ﬂ«·«Ì ›Êﬁ »Â «‰»«— «÷«›Â ‘œ."
	%>
	<BR>	<BR>
	<CENTER>
	<% 	ReportLogRow = PrepareReport ("InventoryLogIn.rpt", "LogItem_ID", LogItem_ID, "/beta/dialog_printManager.asp?act=Fin") %>
	<INPUT TYPE="button" value=" ç«Å —”Ìœ" Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

	<input type="button" value="Ê—Êœ ﬂ«·«Ì »⁄œÌ" Class="GenButton" onclick="window.location='itemIn.asp'">

	</CENTER>

	<BR>	
	<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:700; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no></iframe>
	<%

	response.end

end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Voroode Piramon Marjooe submit
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="Ê—Êœ ÅÌ—«„Ê‰ Â«Ì „—ÃÊ⁄Ì" then

	invoice_id = trim(request.form("invoice_id"))
	if not isnumeric(invoice_id) then 
		response.write "<br><br>"
		call showAlert( "‘„«—Â ›«ﬂ Ê— Ê«—œ ‘œÂ «‘ »«Â «” ." , CONST_MSG_ERROR)
		response.end
	end if 

	'vvvvvvv ------------------------------------------ start of check for current ItemOUT
	set RSS=Conn.Execute ("SELECT InventoryLog.id, InventoryLog.comments, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.VoidedDate, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.type, InventoryLog.ID, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE (InventoryLog.RelatedInvoiceID = "& invoice_id & ") AND (InventoryLog.IsInput = 1)")
	if not RSS.EOF then
		%>
		<br><br>
		<%
		call showAlert("ﬁ»·« «ﬁ·«„ «Ì‰ ÅÌ‘ ‰ÊÌ” »Â «‰»«—  Ê«—œ ‘œÂ «‰œ", CONST_MSG_ALERT)
		response.write "<br><br>" 
		%>
		<TABLE dir=rtl align=center width=600>
			<TR bgcolor="eeeeee" >
				<TD><SMALL>ﬂœ ﬂ«·«</SMALL></TD>
				<TD width=200><SMALL>‰«„ ﬂ«·«</SMALL></TD>
				<TD><SMALL> ⁄œ«œ </SMALL></TD>
				<TD><SMALL>Ê«Õœ</SMALL></TD>
				<TD><SMALL> «—ÌŒ Ê—Êœ</SMALL></TD>
				<TD align=center><SMALL>‘„«—Â ÕÊ«·Â</SMALL></TD>
				<TD><SMALL> Ê”ÿ</SMALL></TD>
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
			<TR bgcolor="<%=tmpColor%>"  style="height:25pt" <% if RSS("voided") then%> disabled title="Õœ› ‘œÂ œ—  «—ÌŒ <%=RSS("VoidedDate")%>"<% end if %>>
				<TD  align=right dir=ltr><INPUT TYPE="hidden" name="id" value="<%=RSS("ID")%>"><A HREF="invReport.asp?oldItemID=<%=RSS("oldItemID")%>&logRowID=<%=RSS("ID")%>" target="_blank"><%=RSS("OldItemID")%></A></TD>
				<TD><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><!A HREF="default.asp?itemDetail=<%=RSS("ID")%>"><span style="font-size:10pt"><%=RSS("Name")%></A></TD>
				<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("Qtty")%></span></TD>
				<TD align=right dir=ltr><%=RSS("Unit")%></TD>
				<TD dir=ltr><%=RSS("logDate")%></span></TD>
			<TD align=center><% if RSS("type")= "1" and RSS("RelatedID")<> "-1" then
					%>
						”›«—‘ Œ—Ìœ 
				     <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A>
				   <%
				   elseif RSS("type")= "2" then
					response.write "<font color=red><b>«’·«Õ „ÊÃÊœÌ</b></font>"
				   elseif RSS("type")= "3" then
					response.write "<font color=green><b>„—ÃÊ⁄Ì</b></font>"
					elseif RSS("type")= "4" then
					response.write "<font color=blue><b> ⁄—Ì› ﬂ«·«Ì ÃœÌœ </b></font>"
					elseif RSS("type")= "5" then
					response.write "<font color=orang><b>«‰ ﬁ«·</b></font>"
					elseif RSS("type")= "6" then
					response.write "<font color=#6699CC><b>Ê—Êœ «“  Ê·Ìœ</b></font>"
					elseif RSS("type")= "7" then
					response.write "<font color=#FF9966><b>Ê—Êœ «“ «‰»«— ‘Â—Ì«—</b></font>"
				   else 
					response.write " "
				   end if	%>
				   
				<% if RSS("owner")<> "-1" and RSS("owner")<> "-2" then
					response.write " («—”«·Ì <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer="& RSS("owner") &"' target='_blank'> " & RSS("owner") & "</a>)"
				   end if %>
				<% if RSS("comments")<> "-" and RSS("comments")<> "" then
					response.write " <br><br><B> Ê÷ÌÕ:</B>  " & RSS("comments") 
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
	response.write "<BR><BR><CENTER>¬Ì« Ê—Êœ ÅÌ—«„Ê‰ Â«Ì „—ÃÊ⁄Ì “Ì— »Â «‰»«— —«  «ÌÌœ „Ì ﬂ‰Ìœø<BR><BR>"

	do while not RSS.EOF
		response.write "<li> " & RSS("AppQtty") & " " & RSS("unit") & " " & RSS("name") 
		RSS.movenext
	loop
	%>
	<FORM METHOD=POST ACTION="">
		<INPUT TYPE="hidden" name="invoice_id" value="<%=invoice_id%>">
		<INPUT TYPE="submit" Name="Submit" Value=" «ÌÌœ Ê—Êœ ÅÌ—«„Ê‰ Â«Ì „—ÃÊ⁄Ì" class=inputBut style="width:170px;" >
		<INPUT TYPE="submit" Name="Submit" Value="«‰’—«›" class=inputBut style="width:70px;" >
	</FORM>
	<%
	response.end
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Voroode Piramon Marjooe submit
'-----------------------------------------------------------------------------------------------------
elseif request.form("Submit")=" «ÌÌœ Ê—Êœ ÅÌ—«„Ê‰ Â«Ì „—ÃÊ⁄Ì" then
	'response.write "<br><br>"
	'call showAlert("«Ì‰ ’›ÕÂ Â‰Ê“ ¬„«œÂ ‰Ì” .",CONST_MSG_ERROR)
	'response.end

	invoice_id = trim(request.form("invoice_id"))
	if not isnumeric(invoice_id) then 
		response.write "<br><br>"
		call showAlert( "‘„«—Â ›«ﬂ Ê— Ê«—œ ‘œÂ «‘ »«Â «” ." , CONST_MSG_ERROR)
		response.end
	end if 

	set RSS=Conn.Execute ("SELECT dbo.InventoryItems.Name, dbo.InvoiceLines.AppQtty, dbo.InventoryItems.ID as itemID, dbo.InvoiceItems.RelatedInventoryItemID, dbo.InventoryItems.Unit FROM dbo.InvoiceLines INNER JOIN dbo.InvoiceItems ON dbo.InvoiceLines.Item = dbo.InvoiceItems.ID INNER JOIN dbo.InventoryItems ON dbo.InvoiceItems.RelatedInventoryItemID = dbo.InventoryItems.OldItemID WHERE (dbo.InvoiceLines.Invoice = " & invoice_id & ")")
	st = ""
	response.write "<BR><BR><CENTER>ÅÌ—«„Ê‰ Â«Ì „—ÃÊ⁄Ì “Ì— Ê«—œ «‰»«— ‘œ:<BR>"

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
if request.form("Submit")="Ê—Êœ ﬂ«·« »Â «‰»«—" then
	item = trim(request.form("item"))
	qtty = CDbl(request.form("qtty"))
	purchaseOrderID = CDbl(request.form("purchaseOrderID"))
	ownerAcc = CDbl(request.form("accountID"))
	comments = request.form("comments")
	entryDate = request.form("entryDate")
	price= request.form("price")
	if price<>"" then 
		price = CDbl(price)
	else
		price = 0
	end if
	gl_update = request.form("gl_update")

	if Auth(5 , "C") AND "" & entryDate <> "" then ' À»  Ê—Êœ/Œ—ÊÃ œ—  «—ÌŒ œ·ŒÊ«Â
		logDate=entryDate
	else
		logDate=shamsiToday()
	end if
	
	if 	not item = "" then

		if purchaseOrderID="" then
			purchaseOrderID = "-1"
		end if

		if item="-1" or qtty=0 then
			response.write "<br><br><br>"
			CALL showAlert ("<B>Œÿ«! </B><BR>ÂÌç ﬂ«·«ÌÌ «‰ Œ«» ‰ﬂ—œÂ «Ìœ<br><br><A HREF='itemIn.asp'>»—ê‘ </A>",CONST_MSG_ALERT) 
			response.end
		end if

		if qtty<0 then
			response.write "<br><br><br>"
			CALL showAlert ("<B>Œÿ«! </B><BR>„ﬁœ«— ﬂ«·« ‰„Ì  Ê«‰œ „‰›Ì »«‘œ<br><br><A HREF='itemIn.asp'>»—ê‘ </A>",CONST_MSG_ALERT) 
			response.end
		end if
		
		if price<=0 and ownerAcc=-1 and purchaseOrderID<>-9 then
			response.write "<br><br><br>"
			CALL showAlert ("<B>Œÿ«! </B><BR>ﬁÌ„  —« »Â ’Ê—  ’ÕÌÕ Ê«—œ ‰„«ÌÌœ<br><br><A HREF='itemIn.asp'>»—ê‘ </A>",CONST_MSG_ALERT) 
			response.end
		end if
		
		set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (OldItemID = "& item & ")" )
		ItemID = RSW("id")
		if ownerAcc=-1 then 
			mySQL = "select * from inventoryLog where owner=-1 and itemID=" & itemID & " and isInput=0 and logDate>='" & logDate & "' and pricedQtty is not null order by logDate desc"
	'		response.write mySQL
			set rs=conn.Execute(mySQL)
			if not rs.eof then 
				response.write "<br><br><br>"
				CALL showAlert ("<B>Œÿ«! </B><BR>»⁄œ «“ «Ì‰  «—ÌŒ Œ—ÊÃ À»  ‘œÂ «” ° ·ÿ›« »⁄œ «“  «—ÌŒ " & rs("logDate") & " Ê—Êœ À»  ﬂ‰Ìœ<br><br><A HREF='itemIn.asp'>»—ê‘ </A>",CONST_MSG_ALERT) 
				response.end
			end if
		end if
		response.write "<br><br><br><br><center>"

		
		type1 = 1 
		if purchaseOrderID < 0 then type1 = -1 * purchaseOrderID
		if purchaseOrderID = -8 then 
			type1 = 1
			purchaseOrderID = request.form("customerItemRequest")
		end if
		if CDbl(price)>0 then 
			pricedQtty = qtty
		else
			pricedQtty = "null"
		end if
		mySql="SET NOCOUNT ON;INSERT INTO InventoryLog (ItemID, RelatedID, logDate, Qtty, owner, CreatedBy, IsInput, comments, type, price, gl_update, pricedQtty) VALUES ("& ItemID & ", "& purchaseOrderID & ",N'"& logDate & "', "& Qtty & ", "& ownerAcc & ", "& session("id") & ", 1, N'" & comments & "', " & type1 & "," & price & " ,0," & pricedQtty & " );select @@identity as newID;"
		'response.write mySQL
		set rs = conn.Execute(mySql)
		newID = rs.Fields("newID").value
		RS.close
		set rs = nothing
		if ownerAcc = -1 then 
			Conn.Execute("insert into InventoryPriceLog (logID,logDate,userID,price) values(" & newID & ",'" & logDate & "'," & session("id") &"," & price & ")")
		end if
		'response.redirect "ItemIn.asp?act=showReceipt&item=" & item & "&qtty=" & qtty
		response.redirect "ItemIn.asp?act=showReceipt&id=" & newID
	end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Submit an Inventory Item Input 2
'-----------------------------------------------------------------------------------------------------
elseif request.form("Submit")="«‰ Œ«»" then
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
customerItemRequest = -1
comment=""
purchaseOrderID = request.form("purchaseOrderID")
if purchaseOrderID = "" then purchaseOrderID = -1
purchaseOrderID = CInt(purchaseOrderID)
'set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (Status = 'OUT' and ID="& purchaseOrderID & ")" )
if purchaseOrderID > 0 then 
	set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (ID="& purchaseOrderID & ")" )
	if not (RSA.eof) then
		set rs = Conn.Execute("select * from VoucherLines where RelatedPurchaseOrderID = " & purchaseOrderID)
		if not rs.eof then 
			qtty = rs("qtty")
			price = rs("price")
			msg = "<a href='../AP/AccountReport.asp?act=showVoucher&voucher=" & rs("voucher_ID") & "'>›Ì„  À»  ‘œÂ œ— ›«ﬂ Ê—</a>"
			gl_update=0
		else
			Qtty = RSA("Qtty")
			price = RSA ("price")
			msg = "<a href='../purchase/outServiceTrace.asp?od=" & purchaseOrderID & "'>ﬁÌ„   Ê«›ﬁÌ À»  ‘œÂ œ— ”›«—‘ Œ—Ìœ</a>"
			gl_update=-1
		end if
		rs.close
	else
		call showAlert ("Œÿ«Ì ⁄ÃÌ»! ”›«—‘ Œ—Ìœ „—»ÊÿÂ ÅÌœ« ‰‘œ" , CONST_MSG_ERROR )
		response.end
	end if
	RSA.close
elseif purchaseOrderID = -1 then 
	price = request.form("price")
	msg="œﬁ  ﬂ‰Ìœ ﬁÌ„  —« ’ÕÌÕ Ê«—œ ‰„«ÌÌœ"
elseif purchaseOrderID = -3 then 
	set rs = Conn.Execute("select * from InventoryLog where id = " & request.form("retID"))
	price = rs("price")
	qtty = rs("qtty")
	gl_update=0
	rs.close
elseif purchaseOrderID = -7 then 
	set rs = Conn.Execute("select * from InventoryLog where id = " & request.form("outID"))
	price = rs("price")
	qtty = rs("qtty")
	gl_update=0
	rs.close
elseif purchaseOrderID = -2 then 
	gl_update=0
	price = request.form("price")
	'------------------------------------------------------------------------------------------------
	' NOT COMPLITED!!!!!!!!!!!!!!!!!!!!!!
	'------------------------------------------------------------------------------------------------
elseif purchaseOrderID = -6 then 
	gl_update=0
	price = request.form("price")
	'------------------------------------------------------------------------------------------------
	' NOT COMPLITED!!!!!!!!!!!!!!!!!!!!!!
	'------------------------------------------------------------------------------------------------
elseif purchaseOrderID = -8 then 
	customerItemRequest = request.form("customerItemRequest")
	mySQL="select InventoryItemRequests.*, Orders.Customer, Accounts.accountTitle from InventoryItemRequests inner join Orders on InventoryItemRequests.order_id=orders.id inner join Accounts on Orders.customer=Accounts.id where InventoryItemRequests.id=" & customerItemRequest
	'response.write mySQL
	set rs = Conn.Execute(mySQL)
	if rs.eof then 
		call showAlert ("Œÿ«Ì ⁄ÃÌ»! ﬂ«·«Ì «—”«·Ì „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ" , CONST_MSG_ERROR )
		response.end
	end if
	if rs("comment")<>"" then comment=" Ê÷ÌÕ: " & rs("comment")
	owner=rs("customer")
	qtty=rs("qtty")
	accountName = rs("accountTitle")
	
	rs.close
	set rs = Conn.Execute("select * from InventoryLog where owner>0 and voided=0 and IsInput=1 and RelatedId=" & customerItemRequest)
	if not rs.eof then 
		call showAlert ("Œÿ«Ì ⁄ÃÌ»! «Ì‰ ﬂ«·« ﬁ»·« Œ«—Ã ‘œÂ" , CONST_MSG_ERROR )
		response.end
	end if
elseif purchaseOrderID = -9 then
	owner=-1
	price=0
end if

'response.write purchaseOrderID 
%>
<FORM METHOD=POST ACTION="itemin.asp">
<INPUT TYPE="hidden" name="item" value="<%=request.form("item")%>">
<INPUT TYPE="hidden" name="customerItemRequest" value="<%=customerItemRequest%>">
<INPUT TYPE="hidden" name="purchaseOrderID" value="<%=request.form("purchaseOrderID")%>"><BR><BR>
<TABLE border=0 align=center>
<TR>
	<TD align=left valign=top>‰«„ ﬂ«·«</TD>
	<TD align=right><span disabled><%=request("goodName")%></span><br><br></TD>
</TR>

<TR>
	<TD align=left> ⁄œ«œ</TD>
	<TD align=right>
		<INPUT dir=ltr TYPE="text" NAME="qtty" size=25 value="<%=Qtty%>" ><onKeyPress="return maskNumber(this);" >
	</TD>
</TR>
<%if customerItemRequest=-1 then %>
<TR>
	<TD align=left>ﬁÌ„ </TD>
	<TD align=right>
		<INPUT dir=ltr TYPE="text" NAME="price" size=25 value="<%=price%>" ><onKeyPress="return maskNumber(this);" >
		<small style="color:red;">*  ÊÃÂ œ«‘ Â »«‘Ìœ ﬂÂ «Ì‰ ﬂ«·« »« «Ì‰ ﬁÌ„  »Â «‰»«— Ê«—œ „Ìù‘Êœ. Å” œ— À»  ¬‰ œ›  »›—„«ÌÌœ</small>
		<br><%=msg%><small style="color:red;">* ﬁÌ„  ﬂ· —« À»  ﬂ‰Ìœ</small>
	</TD>
</TR>
<%end if%>
<TR>
	<TD align=left valign=top>„«·ﬂÌ </TD>
	<TD align=right>
<%if customerItemRequest=-1 then %>
	<SELECT NAME="aaa2" onchange="hideIT()" >
		<option value=1 >Œ«‰Â ç«Å Ê ÿ—Õ</option>
		<option value=2 <%if owner<>"-1" then response.write " selected" %>>œÌê—«‰ (‘„«—Â Õ”«»)</option>
	</SELECT>
<%end if%>
	<BR>
	<span name="aaa1" id="aaa1" <% if owner="-1" then response.write "style=""visibility:'hidden'"""%>>
	<input type="hidden" name="gl_update" value="<%=gl_update%>">
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<INPUT  dir="LTR"  TYPE="text" NAME="accountID" maxlength="10" size="13"  value="<%=owner%>" onKeyPress='return mask(this);' onBlur='check(this);'> &nbsp;&nbsp; <INPUT TYPE="text" NAME="accountName" size=30 readonly  value="<%=accountName%>" style="background-color:transparent">
	</span></TD>
</TR>
<%
	if Auth(5 , "C") then ' À»  Ê—Êœ/Œ—ÊÃ œ—  «—ÌŒ œ·ŒÊ«Â
%>
	<TR>
		<TD align=left> «—ÌŒ</TD>
		<TD align=right><INPUT dir=ltr TYPE="text" NAME="entryDate" size=25 value="<%=shamsiToday()%>" onblur="acceptDate(this)"></TD>
	</TR>
<%
	End if

%>
<TR>
	<TD align=left> Ê÷ÌÕ« </TD>
	<TD align=right><TEXTAREA NAME="comments" ROWS="" COLS=""><%=comment%></TEXTAREA>
	<br><br></TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=right></TD>
</TR>
<TR>
	<TD align=center colspan=2>
	<INPUT TYPE="submit" Name="Submit" Value="Ê—Êœ ﬂ«·« »Â «‰»«—" class=inputBut style="width:120px;" tabIndex="14"<%
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


//-->
</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Submit an Inventory Item Input 1
'-----------------------------------------------------------------------------------------------------
elseif request.form("Submit")="Ã” ÃÊ" then
%>
<FORM METHOD=POST ACTION="itemin.asp">
<INPUT TYPE="hidden" name="item" value="<%=request.form("item")%>">
<TABLE border=0 align=center>
<%
				set RSW=Conn.Execute ("SELECT * FROM InventoryItems WHERE (OldItemID = "& goodItem1 & ")" )
				if RSW.EOF then
					call showAlert ("Œÿ«! ﬂœ ﬂ«·« „⁄ »— ‰Ì” " , CONST_MSG_ERROR )
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
	<br>
	</TD>
</TR>
<INPUT TYPE="hidden" name="goodName" value="<%=goodName%>">
<TR>
	<TD align=right>
			
<%
				set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (Status = 'OUT' and TypeID="& goodItem1& ") order by OrdDate" )
				'set RSA=Conn.Execute ("SELECT * FROM purchaseOrders WHERE (TypeID="& goodItem1& ") order by OrdDate" )
				flg = false
				if not RSA.eof then 
					response.write("”›«—‘ Œ—Ìœ —« «‰ Œ«» ﬂ‰Ìœ:<br>")
					while not (RSA.eof) %>
						<INPUT TYPE="radio" NAME="purchaseOrderID" value="<%=RSA("ID")%>"<%
						if trim(purchaseOrderID) = trim(RSA("ID")) then
							response.write " chcecked "
							preQtty = RSA("Qtty")
							flg = true
						end if
						%>> ‘„«—Â <%=RSA("ID")%> (<%=RSA("Qtty")%>&nbsp;<%=goodUnit%>&nbsp;<%=goodName%>) <BR>
	<%					RSA.MoveNext
					wend
					RSA.close
				end if
				if Auth(5,"E") then 
					response.write ("<input type='hidden' NAME='customerItemRequest' id='customerItemRequest'>")
					set rs=Conn.Execute ("select InventoryItemRequests.*,Accounts.accountTitle from InventoryItemRequests inner join Orders on InventoryItemRequests.order_id=orders.id inner join Accounts on Orders.customer=Accounts.id where InventoryItemRequests.status='new' and CustomerHaveInvItem=1 and InventoryItemRequests.id not in(select RelatedId from InventoryLog where owner>0 and voided=0 and IsInput=1 and ItemID=" & goodItem1 & ") and itemID=" & goodItem1)
					if not rs.eof then response.write ("<br>œ—ŒÊ«” ùÂ«Ì „‘ —Ì:<br>")
					while not rs.eof
						response.write("<INPUT TYPE='radio' name='purchaseOrderID' value='-8' onclick='document.all.customerItemRequest.value=" & rs("id") & "';>")
						response.write ("<a href='../shopfloor/manageOrder.asp?radif=" & rs("order_id") & "'>œ—ŒÊ«”  ÿÌ ”›«—‘ " & rs("order_id") & "</a> Ê”ÿ " & rs("accountTitle") & " <span title='" & rs("comment") & "'>(" & RS("Qtty") &"&nbsp;" & goodUnit &"&nbsp;"& goodName & ")</span> <BR>")
						rs.moveNext
					wend
					rs.close
					set rs = nothing
				end if
				'rs.close
				if Auth(5,"G") then 
				'-------------------------------------------------------------------------------------------------
				'--------------------------------------- NOT IMPLIMENTED!!!!--------------------------------------
				'-------------------------------------------------------------------------------------------------
				set rs=Conn.Execute("select [log].*,InventoryItems.unit from InventoryLog as [log] inner join (select top 1  logDate from InventoryLog where ItemID=" & goodItem1 & " and sumQtty=0 order by logDate desc) as date on [log].logDate>date.logDate inner join InventoryItems on [log].itemID=InventoryItems.ID where isInput = 0 and [log].ItemID=" & goodItem1)
				if not rs.eof then 
				%>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-3"> ﬂ«·«Ì „—ÃÊ⁄Ì <select name="retID">
					<%
					while not rs.eof
					%>
					<option value="<%=rs("id")%>"><%=rs("Qtty") & rs("unit") & " œ—  «—ÌŒ " & rs("logDate")%></option>
					<%
						rs.moveNext
					wend
					%>
				</select><br>
				<%
				end if
				rs.close
				end if
				if Auth(5,"F") then
				'-------------------------------------------------------------------------------------------------
				'--------------------------------------- NOT IMPLIMENTED!!!!--------------------------------------
				'------------------------------------------------------------------------------------------------- 
				%>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-6"> Ê—Êœ «“  Ê·Ìœ<small> ‘„«—Â ”›«—‘:<input name="orderID" type="text" size="6"></small><br>
				<%
				end if
				if Auth(5,"H") then 
				'-------------------------------------------------------------------------------------------------
				'--------------------------------------- NOT IMPLIMENTED!!!!--------------------------------------
				'-------------------------------------------------------------------------------------------------
				set rs=Conn.Execute("select [log].*,InventoryItems.unit from InventoryLog as [log] inner join (select top 1  logDate from InventoryLog where ItemID=" & goodItem1 & " and sumQtty=0 order by logDate desc) as date on [log].logDate>date.logDate inner join InventoryItems on [log].itemID=InventoryItems.ID where isInput = 0 and [log].ItemID=" & goodItem1)
				if not rs.eof then 
				%>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-7"> Ê—Êœ «“  «‰»«— ‘Â—Ì«— / œ› — ⁄»«” ¬»«œ<select name="outID">
				<%
					while not rs.eof
					%>
					<option value="<%=rs("id")%>"><%=rs("Qtty") & rs("unit") & " œ—  «—ÌŒ " & rs("logDate")%></option>
					<%
						rs.moveNext
					wend
					%>
				</select><br>
				<%
				end if
				rs.close
				end if
				if Auth(5 , 7) then %>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-2"> «’·«Õ „ÊÃÊœÌ<br>
				<% end if 
				if Auth(5,"I") then
				%>
				<INPUT TYPE="radio" NAME="purchaseOrderID" value="-9">Ê—Êœ ÷«Ì⁄« <br>
				<%
				end if
				%>
			<br><br></TD>
</TR>
<TR>
	<TD align=center colspan=2>
	<INPUT TYPE="hidden" name="owner" value="<%=owner%>">
	<INPUT TYPE="submit" Name="Submit" Value="«‰ Œ«»" class=inputBut style="width:120px;" tabIndex="14">
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
	<TD colspan=2 align=center><H3>Ê—Êœ ﬂ«·«</H3></TD>
</TR>
<TR>
	<TD align=left>ﬂœ ﬂ«·«<br></TD>
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
	<INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" class=inputBut style="width:120px;" tabIndex="14">
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
		<B> ÅÌ—«„Ê‰ „—ÃÊ⁄Ì <BR></B><BR>
		‘„«—Â »—ê‘  «“ ›—Ê‘: <INPUT TYPE="text" NAME="invoice_id" dir=ltr ><BR><BR> <INPUT TYPE="submit"  name="submit"  value="Ê—Êœ ÅÌ—«„Ê‰ Â«Ì „—ÃÊ⁄Ì">
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
