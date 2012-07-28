<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="»—‰«„Â —Ì“Ì"
SubmenuItem=2
if not Auth(3 , 2) then NotAllowdToViewThisPage()

'OutService Page Request
'By Alix - Last changed: 81/01/13
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<%
catItem1 = request("catItem")
if catItem1="" then catItem1="-1"

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------ Submit an OutService request
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  œ—ŒÊ«”  Œœ„« " then
	order_ID = request.form("radif")
	otypeID = request.form("type")
	comment = request.form("comment")
	Price = request.form("Price")
	qtty = request.form("qtty")
	priceComment = request.form("priceComment")
	if Price = "" then price = "0"
	if comment = "" then comment = "-"
	if qtty = "" then qtty = "0"
	CreatedBy = session("id")
	dueDate = request.form("date1")

	if order_ID="" or otypeID="-1" or comment="" or Price="" then
		response.write "error"
		response.end
	end if

	set RS4 = conn.Execute ("SELECT * FROM OutServices where ID=" & otypeID)
	if (RS4.eof) then
		otype="-unknown-"
	else
		otype=RS4("Name")
	end if
	RS4.close

	mySql="INSERT INTO purchaseRequests (qtty, order_ID, typeName, typeID, comment, ReqDate, price,priceComment, CreatedBy, DueDate, IsService) VALUES ("& qtty & ","& order_ID & ", N'"& otype & "', "& otypeID & ", N'"& comment & "',N'"& shamsiToday() & "', "& price & ",N'"& priceComment & "', "& CreatedBy & ", N'"& DueDate & "' ,1 )"	
	'response.write "<div align=left dir=ltr>"
	'response.write mySql
	'response.end

	conn.Execute mySql
	'RS2.close

	if not request("relatedApprovedInvoiceID")="0" then
		call UnApproveInvoice(request("relatedApprovedInvoiceID"),request("relatedApprovedInvoiceBy"))
	end if

	response.redirect "manageOrder.asp?radif=" & order_ID
end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------- Submit an Inventory Item request
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="À»  œ—ŒÊ«”  ﬂ«·«" then
	order_ID = request.form("radif")
	item = request.form("item")
	comment = request.form("comment")
	qtty = request.form("qtty")
	CreatedBy = session("id")
	if request.form("CustomerHaveInvItem") ="on" then
		CustomerHaveInvItem = 1
	else
		CustomerHaveInvItem = 0
	end if 
	if qtty = "" then qtty = "0"
	if comment = "" then comment = "-"
	if 	not (item = "" or item="-1") then

		if order_ID="" or otypeID="-1" or comment="" or qtty="" then
			response.write "error"
			response.end
		end if

		set RS4 = conn.Execute ("SELECT * FROM InventoryItems where ID=" & item)
		if (RS4.eof) then
			otype="-unknown-"
			unit=RS4("unit")
		else
			otype=RS4("Name")
			unit=RS4("unit")
		end if
		RS4.close
		
		mySql="INSERT INTO InventoryItemRequests (order_ID, ItemName, ItemID, comment, ReqDate, Qtty, unit, CreatedBy, CustomerHaveInvItem) VALUES ("& order_ID & ", N'"& otype & "', "& item & ", N'"& comment & "',N'"& shamsiToday() & "', "& Qtty & ", N'"& unit & "' , "& CreatedBy & " , " & CustomerHaveInvItem & ")"
		conn.Execute mySql
		'RS2.close

		if not request("relatedApprovedInvoiceID")="0" then
			call UnApproveInvoice(request("relatedApprovedInvoiceID"), request("relatedApprovedInvoiceBy"))
		end if

		response.redirect "manageOrder.asp?radif=" & order_ID
	end if
end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Delete an OutReq from an order
'-----------------------------------------------------------------------------------------------------
if request("d")="y" then		
	myRequestID=request("i")
	set RSX=Conn.Execute ("SELECT * FROM purchaseRequests WHERE id = "& myRequestID )	
	if RSX("status")="new" then
	Conn.Execute ("update purchaseRequests SET status = 'del' where id = "& myRequestID )	
	end if

	if not request("relatedApprovedInvoiceID")="0" then
		call UnApproveInvoice(request("relatedApprovedInvoiceID") , request("relatedApprovedInvoiceBy"))
	end if
	response.redirect "manageOrder.asp?radif=" & request("r")
end if

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------- Delete an Inventory Request from an order
'-----------------------------------------------------------------------------------------------------
if request("di")="y" then		
	myRequestID=request("i")
	set RSX=Conn.Execute ("SELECT * FROM InventoryItemRequests WHERE id = "& myRequestID )	
	if RSX("status")="new" then
	Conn.Execute ("update InventoryItemRequests SET status = 'del' where id = "& myRequestID )	
	end if

	if not request("relatedApprovedInvoiceID")="0" then
		call UnApproveInvoice(request("relatedApprovedInvoiceID"), request("relatedApprovedInvoiceBy"))
	end if
	response.redirect "manageOrder.asp?radif=" & request("r")
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------ Main
'-----------------------------------------------------------------------------------------------------
if request("radif")="" then
	%><br><br><br>
	<FORM METHOD=POST ACTION="manageOrder.asp?e=y">
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" style="font-family:Tahoma,arial; font-size: 8pt;">
	<TR>
		<TD>‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ:</TD>
		<TD><INPUT Name="Radif" TYPE="text" maxlength="6" size="6" tabIndex="1" dir="LTR"></TD>
		<TD><INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" style="width:100px;"  class="inputBut"></TD>
	</TR>
	</TABLE>
	</FORM>
	<script language="JavaScript">
	<!--
	document.all.Radif.focus();
	//-->
	</script>
	<%
	response.end
elseif NOT isNumeric(request("radif")) then
	response.write "<BR>"
	call showAlert("«Ì‰ ‘„«—Â ”›«—‘ <B>[ " & request("radif")& " ]</B> «‘ »«Â «” .", CONST_MSG_ERROR )
	response.end 
else
	OrderID=clng(request("radif"))
end if


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------ Details of an Job (Add Request to a JOB)
'------------------------------------------------------------------------- if request("radif") <> "" :
'-----------------------------------------------------------------------------------------------------
set RS1=Conn.Execute ("SELECT orders_trace.*, Invoices.id FROM orders_trace RIGHT OUTER JOIN Orders ON orders_trace.radif_sefareshat = Orders.ID LEFT OUTER JOIN Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (ISNULL(Invoices.issued, 0) = 1) AND (Orders.ID ='"& request("radif") & "') AND (Invoices.Voided = 0)")

set RS2=Conn.Execute ("SELECT * FROM orders_trace WHERE (radif_sefareshat='"& request("radif") & "')")

set RS3=Conn.Execute ("SELECT Invoices.id, Invoices.ApprovedBy FROM Orders LEFT OUTER JOIN Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (ISNULL(Invoices.approved, 0) = 1) AND (Orders.ID ='"& request("radif") & "')")

if RS2.eof then %>
	<FORM METHOD=POST ACTION="manageOrder.asp"><BR><BR>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" style="font-family:Tahoma,arial; font-size: 8pt;">
	<TR>
		<TD>‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ:</TD>
		<TD><INPUT Name="Radif" TYPE="text" maxlength="6" size="6" tabIndex="1" dir="LTR"></TD>
		<TD><INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" style="width:100px;font-family:Tahoma,arial; font-size: 8pt;"></TD>
	</TR>
	</TABLE>
	</FORM>
	<br>
	<%	

	call showAlert("«„ﬂ«‰ ‰œ«—œ ﬂÂ Â„çÌ‰ ”›«—‘Ì ÊÃÊœ œ«‘ Â »«‘œ.", CONST_MSG_ERROR )
	response.end 
end if

if not RS1.eof then
	EDITABLE="NO" 
	response.write "<BR><BR>"
	call showAlert( "Ìﬂ ›«ﬂ Ê—  ’«œ— ‘œÂ „— »ÿ »« »« «Ì‰  ”›«—‘ ‘„«—Â ÅÌœ« ﬂ—œÂ «Ì„.(‘„«—Â ›«ﬂ Ê—: <A HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& RS1("id") & "' target='_blank'>" & RS1("id") & "</a>)<BR> ·–«”  ﬂÂ «„ﬂ«‰Ì »—«Ì  €ÌÌ— œ— «Ì‰ ”›«—‘ ÊÃÊœ ‰œ«—œ.", CONST_MSG_ALERT  ) 

	if Auth(3 , 5) then
		response.write "<br>" 
		call showAlert("Ê·Ì çÊ‰  ‘„« Ã‰«»  "& session("CSRName") &  " Â” Ìœ „Ì  Ê«‰Ìœ Â— «ﬁœ«„Ì „«Ì· »ÊœÌœ «‰Ã«„ œÂÌœ." , CONST_MSG_INFORM)
		EDITABLE="YES" 
	end if
end if 
%>
<!--#include File="../include_JS_InputMasks.asp"-->
<font face="tahoma">

<font face="tahoma">
<% if (not RS3.EOF) and RS1.eof then
	response.write "<BR><BR>"
	call showAlert("Ìﬂ ›«ﬂ Ê—  «ÌÌœ ‘œÂ »« ‘„«—Â <A HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& RS3("id") &"' target='_blank'>" & RS3("id") & "</A> »—«Ì «Ì‰ ”›«—‘ ÊÃÊœ œ«—œ <BR>ﬂÂ »«  €ÌÌ— ”›«—‘  Ê”ÿ ‘„« «“ Õ«·   «ÌÌœ Œ«—Ã ŒÊ«Âœ ‘œ." , CONST_MSG_ALERT ) 
	relatedApprovedInvoiceID = RS3("id")
	relatedApprovedInvoiceBy = RS3("ApprovedBy")
else
	relatedApprovedInvoiceID = 0
end if %>
<font face="tahoma">

<br>
<FORM METHOD=POST disabled>
<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center">
<TR bgcolor="black">
	<TD align="left"><FONT COLOR="YELLOW">‘„«—Â ”›«—‘:</FONT></TD>
	<TD align="right">
		<!-- Radif -->
		<INPUT TYPE="text" disabled maxlength="5" size="5" tabIndex="1" dir="LTR" value="<%=RS2("radif_sefareshat")%>">
		<INPUT TYPE="hidden" NAME="Radif" value="<%=RS2("radif_sefareshat")%>">
	</TD>
	<TD align="left"><FONT COLOR="YELLOW"> «—ÌŒ:</FONT></TD>
	<TD><TABLE border="0">
		<TR>
			<TD dir="LTR">
			<INPUT disabled TYPE="text" maxlength="10" size="8"  value="<%=RS2("order_date")%>">
			<INPUT TYPE="hidden" NAME="OrderDate" value="<%=RS2("order_date")%>">
			</TD>
			<TD dir="RTL"><FONT COLOR="YELLOW"><%=weekdayname(weekday(date))%></FONT></TD>
		</TR>
		</TABLE></TD>
	<TD align="left"><FONT COLOR="YELLOW">”«⁄ :</FONT></TD>
	<TD align="right">
	<INPUT disabled TYPE="text" maxlength="5" size="3" dir="LTR" value="<%=RS2("order_time")%>">
	<INPUT TYPE="hidden" NAME="OrderTime" value="<%=RS2("order_time")%>"></TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left">‰«„ ‘—ﬂ :</TD>
	<TD align="right">
		<!-- CompanyName -->
		<INPUT TYPE="text" NAME="CompanyName" maxlength="50" size="25" tabIndex="2"  value="<%=RS2("company_name")%>"></TD>
	<TD align="left">„Ê⁄œ  ÕÊÌ·:</TD>
	<TD><TABLE border="0">
		<TR>
			<TD dir="LTR"><INPUT TYPE="text" NAME="ReturnDate" maxlength="10" size="8" tabIndex="5" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" value="<%=RS2("return_date")%>"></TD>
			<TD dir="RTL">(?‘‰»Â)</TD>
		</TR>
		</TABLE></TD>
	<TD align="left">”«⁄   ÕÊÌ·:</TD>
	<TD align="right"><INPUT TYPE="text" NAME="ReturnTime" maxlength="6" size="3" tabIndex="6" value="<%=tmpTime%>" dir="LTR" onKeyPress="return maskTime(this);" value="<%=RS2("return_time")%>"></TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left">‰«„ „‘ —Ì:</TD>
	<TD align="right">
		<!-- CustomerName -->
		<INPUT TYPE="text" NAME="CustomerName" maxlength="50" size="25" tabIndex="3" value="<%=RS2("customer_name")%>"></TD>
	<TD align="left">‰Ê⁄ ”›«—‘:</TD>
	<TD> <INPUT TYPE="text" NAME="OrderKind" maxlength="10" size="8" tabIndex="5"  value="<%=RS2("order_kind")%>"> 
</TD>
	<TD align="left">”›«—‘ êÌ—‰œÂ:</TD>
	<TD>
	<INPUT TYPE="text" NAME="SalesPerson" maxlength="10" size="10" tabIndex="5"  value="<%=RS2("salesperson")%>">
	
</TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left"> ·›‰:</TD>
	<TD align="right">
		<!-- Telephone -->
		<INPUT TYPE="text" NAME="Telephone" maxlength="50" size="25" tabIndex="4" value="<%=RS2("telephone")%>"></TD>
	<TD align="left">⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì·:</TD>
	<TD align="right" colspan="4"><INPUT TYPE="text" NAME="OrderTitle" maxlength="255" size="50" tabIndex="9" value="<%=RS2("order_title")%>"></TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left"> ⁄œ«œ:</TD>
	<TD align="right"><INPUT TYPE="text" NAME="Qtty" maxlength="50" size="5" tabIndex="10" dir="LTR" value="<%=RS2("qtty")%>"></TD>
	<TD align="left">”«Ì“:</TD>
	<TD align="right"><INPUT TYPE="text" NAME="Size" maxlength="50" size="8" dir="LTR" tabIndex="11" value="<%=RS2("PaperSize")%>"></TD>
	<TD align="left">Ìﬂ—Ê/œÊ—Ê:</TD>
	<TD><INPUT TYPE="text" NAME="SimplexDuplex" maxlength="10" size="10" tabIndex="5"  value="<%=RS2("SimplexDuplex")%>">
	
</TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left">„—Õ·Â:</TD>
	<TD colspan="2"><INPUT TYPE="text" NAME="Marhale" maxlength="10" size="16" tabIndex="5"  value="<%=RS2("marhale")%>">
	
</TD>
	<TD align="left">Ê÷⁄Ì :</TD>
	<TD colspan="2">
	<INPUT TYPE="text" NAME="Vazyat" maxlength="10" size="16" tabIndex="5"  value="<%=RS2("vazyat")%>">
</TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD colspan="6" height="30px">&nbsp;</TD>
</TR>
</TABLE><br>
</FORM>
	<TABLE cellspacing=0 width=80% align=center>
	<TR bgcolor=white>
		<TD> «—ÌŒ</TD>
		<TD>”«⁄ </TD>
		<TD>„—Õ·Â</TD>
		<TD>Ê÷⁄Ì </TD>
		<TD>‰«„ À»  ﬂ‰‰œÂ</TD>
	</TR>
	<%
	set RS_STEP=Conn.Execute ("SELECT InsertedDate, InsertedTime, StepText, StatusText, Users.RealName FROM OrderTraceLog INNER JOIN Users ON OrderTraceLog.InsertedBy = Users.ID WHERE (OrderTraceLog.[Order] = "& request("radif")  & ") order by OrderTraceLog.ID")
	Do while not RS_STEP.eof	
	%>
	<TR>
		<TD  style="border-bottom: solid 1pt black" dir=ltr align=right><%=RS_STEP("InsertedDate")%> </TD>
		<TD  style="border-bottom: solid 1pt black" dir=ltr align=right>(<%=RS_STEP("InsertedTime")%>)</TD>
		<TD  style="border-bottom: solid 1pt black"><%=RS_STEP("StepText")%></TD>
		<TD  style="border-bottom: solid 1pt black"><%=RS_STEP("StatusText")%></TD>
		<TD  style="border-bottom: solid 1pt black"><%=RS_STEP("RealName")%></TD>
	</TR>
		<%
		RS_STEP.moveNext
	loop
	RS_STEP.close
	Set RS_STEP = nothing

	if NOT EDITABLE="NO" then
	%>
	<TR  bgcolor=white>
		<TD><A HREF="default.asp?orderNum=<%=request("radif")%>&marhale_box=<%=RS2("step")%>"> €ÌÌ— „—Õ·Â</A></TD>
		<td colspan="2" title="»« ﬂ·Ìﬂ —ÊÌ «Ì‰ œﬂ„Â »Â ’Ê—  ŒÊœﬂ«— Ìﬂ «Ì„Ì· »Â „‘ —Ì «—”«· ŒÊ«Âœ ‘œ ﬂÂ Õ«ÊÌ ‘„«—Â „‘ —Ì Ê ‘„«—Â ”›«—‘ ŒÊ«Âœ »Êœ">
			<%
			set rsEmail=Conn.Execute("select accounts.AccountTitle, accounts.Dear1, accounts.FirstName1, accounts.LastName1, orders.ID, orders.Customer,accounts.Email1, orders_trace.order_title from Orders inner join Accounts on orders.Customer=accounts.ID inner join orders_trace on orders_trace.radif_sefareshat=orders.ID where orders.ID=" & request("radif") & " and accounts.EMail1 <> ''")
			if not rsEmail.eof then 
			%>
			<span>
				<form method="post" action="http://my.pdhco.com/sendMail.php">
					<input type="hidden" name="order_id" value="<%=rsEmail("ID")%>">
					<input type="hidden" name="customer_id" value="<%=rsEmail("customer")%>">
					<input type="hidden" name="order_title" value="<%=rsEmail("order_title")%>">
					<input type="hidden" name="Email" value="<%=rsEmail("Email1")%>">
					<input type="hidden" name="AccountTitle" value="<%=rsEmail("AccountTitle")%>">
					<input type="hidden" name="Dear" value="<%=rsEmail("Dear1")%>">
					<input type="hidden" name="FirstName" value="<%=rsEmail("FirstName1")%>">
					<input type="hidden" name="LastName" value="<%=rsEmail("LastName1")%>">
					<input type="submit" name="orderSend" title='<%=rsEmail("email1")%>' value="»Â <%=rsEmail("Dear1") & " " & rsEmail("firstName1") & " " & rsEmail("LastName1")%> «Ì„Ì· ‘Êœ">
				</form>
			</span>
			<%end if%>
		</td>
		<td colspan="2" title="»« ﬂ·Ìﬂ —ÊÌ «Ì‰ œﬂ„Â »Â ’Ê—  ŒÊœﬂ«— Ìﬂ «Ì„Ì· »Â „‘ —Ì «—”«· ŒÊ«Âœ ‘œ ﬂÂ Õ«ÊÌ ‘„«—Â „‘ —Ì Ê ‘„«—Â ”›«—‘ ŒÊ«Âœ »Êœ">
			<%
			set rsEmail=Conn.Execute("select accounts.AccountTitle, accounts.Dear2, accounts.FirstName2, accounts.LastName2, orders.ID, orders.Customer,accounts.Email2, orders_trace.order_title from Orders inner join Accounts on orders.Customer=accounts.ID inner join orders_trace on orders_trace.radif_sefareshat=orders.ID where orders.ID=" & request("radif") & " and accounts.EMail2 <> ''")
			if not rsEmail.eof then 
			%>
			<span>
				<form method="post" action="http://my.pdhco.com/sendMail.php">
					<input type="hidden" name="order_id" value="<%=rsEmail("ID")%>">
					<input type="hidden" name="customer_id" value="<%=rsEmail("customer")%>">
					<input type="hidden" name="order_title" value="<%=rsEmail("order_title")%>">
					<input type="hidden" name="Email" value="<%=rsEmail("Email2")%>">
					<input type="hidden" name="AccountTitle" value="<%=rsEmail("AccountTitle")%>">
					<input type="hidden" name="Dear" value="<%=rsEmail("Dear2")%>">
					<input type="hidden" name="FirstName" value="<%=rsEmail("FirstName2")%>">
					<input type="hidden" name="LastName" value="<%=rsEmail("LastName2")%>">
					<input type="submit" name="orderSend" title='<%=rsEmail("email2")%>' value="»Â <%=rsEmail("Dear2") & " " & rsEmail("firstName2") & " " & rsEmail("LastName2")%> «Ì„Ì· ‘Êœ">
				</form>
			</span>
			<%end if%>
		</td>
	</TR>
	<% end if %>
	</TABLE>

<br>
<table width="700"  align="center">
<tr>
<td  valign=top>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
	<TR bgcolor="black" >
		<TD align="right" colspan=2><FONT COLOR="YELLOW">œ—ŒÊ«” Â«Ì ﬂ«·« «“ «‰»«—:</FONT></TD>
	</TR>
		<TR bgcolor="#CCCCCC" ><td colspan=2>
		<FORM METHOD=POST ACTION="manageOrder.asp" name='findInv'>
			<INPUT TYPE="hidden" name="relatedApprovedInvoiceBy" value="<%=relatedApprovedInvoiceBy%>">
			<INPUT TYPE="hidden" name="relatedApprovedInvoiceID" value="<%=relatedApprovedInvoiceID%>">
			<INPUT TYPE="hidden" name="radif" value="<%=request("radif")%>">
			<SELECT NAME="catItem" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1"  onchange="document.forms['findInv'].submit()"  style="width:300">
			<option value="-1">œ” Â »‰œÌ ﬂ«·« —« «‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">----------------------------------------------</option>
<%
				set RS4 = conn.Execute ("SELECT InventoryItemCategories.Name, InventoryItemCategories.ID FROM InventoryItemCategories INNER JOIN InventoryItemCategoryRelations ON InventoryItemCategories.ID = InventoryItemCategoryRelations.Cat_ID INNER JOIN InventoryItems ON InventoryItemCategoryRelations.Item_ID = InventoryItems.ID WHERE (InventoryItems.outByOrder = 1) GROUP BY InventoryItemCategories.ID, InventoryItemCategories.Name")
				while not (RS4.eof) %>
					<OPTION value="<%=RS4("ID")%>"<%
					if trim(catItem1) = trim(RS4("ID")) then
					response.write " selected "
					end if
					%>>* <%=RS4("Name")%> </option>
<%						RS4.MoveNext
				wend
				RS4.close
				%>
			</SELECT><br><br>
			<SELECT NAME="item" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1" style="width:300">
			<option value="-1">‰Ê⁄ ﬂ«·« —« «‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">----------------------------------------------</option>
			<%
			if catItem1<>"-1" then
				set RS5 = conn.Execute ("SELECT InventoryItems.* FROM InventoryItemCategoryRelations INNER JOIN InventoryItems ON InventoryItemCategoryRelations.Item_ID = InventoryItems.ID WHERE (InventoryItemCategoryRelations.Cat_ID = "& catItem1& ") AND (InventoryItems.outByOrder = 1) ORDER BY Replace([Name],' ','') " )
				while not (RS5.eof) 
			 %>
					<OPTION value="<%=RS5("ID")%>">* <%=RS5("Name")%> (<%=RS5("Unit")%>)</option>
			<%	RS5.MoveNext
				wend
				RS5.close
				%>
			<% end if %>
			</SELECT><br><br>
			<INPUT TYPE="checkbox" NAME="CustomerHaveInvItem" > ﬂ«·«Ì «—”«·Ì ŒÊœ „‘ —Ì <br><br>
			 ⁄œ«œ: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="text" NAME="qtty" size=40 onKeyPress="return maskNumber(this);"><br><br>
			 Ê÷ÌÕ« : <TEXTAREA NAME="comment" ROWS="7" COLS="32"></TEXTAREA>
			<br><center><br>
			<INPUT  class="inputBut" TYPE="submit" Name="Submit" Value="À»  œ—ŒÊ«”  ﬂ«·«"  style="width:120px;" tabIndex="14" <%
			if catItem1="-1" or EDITABLE="NO"  then
				response.write " disabled "
			end if
			%>>
			</center>
		</FORM>
		<hr>

		</FONT></TD>
	</TR>
	<%
	'Gets Request for services list from DB
'set RS3=Conn.Execute ("SELECT dbo.InventoryItemRequests.Comment, dbo.InventoryItemRequests.ID, dbo.InventoryItemRequests.Status, dbo.InventoryItemRequests.ItemName, dbo.InventoryItemRequests.CustomerHaveInvItem, dbo.InventoryItemRequests.Qtty, dbo.InventoryItemRequests.unit, dbo.InventoryItemRequests.ReqDate FROM dbo.InventoryPickuplists FULL OUTER JOIN dbo.InventoryPickuplistItems ON dbo.InventoryPickuplists.id = dbo.InventoryPickuplistItems.pickupListID FULL OUTER JOIN dbo.InventoryItemRequests ON dbo.InventoryPickuplistItems.RequestID = dbo.InventoryItemRequests.ID WHERE (dbo.InventoryItemRequests.Order_ID = "& request("radif") & ") AND (NOT (dbo.InventoryItemRequests.Status = 'del')) GROUP BY dbo.InventoryItemRequests.Comment, dbo.InventoryItemRequests.ID, dbo.InventoryItemRequests.Status, dbo.InventoryItemRequests.ItemName, dbo.InventoryItemRequests.CustomerHaveInvItem, dbo.InventoryItemRequests.Qtty, dbo.InventoryItemRequests.unit, dbo.InventoryItemRequests.ReqDate")
	set RS3=Conn.Execute ("SELECT InventoryItemRequests.*,InventoryPickuplistItems.pickupListID  FROM InventoryItemRequests left outer join InventoryPickuplistItems on InventoryItemRequests.ID=InventoryPickuplistItems.RequestID WHERE InventoryItemRequests.order_ID="& request("radif") )

	'set RS3=Conn.Execute ("SELECT * FROM InventoryItemRequests WHERE (order_ID="& request("radif") & " ) and not status = 'del'")
	%>
		<%
		Do while not RS3.eof
		%>
		<TR bgcolor="#CCCCCC" title="<% 
			Comment = RS3("Comment")
			if Comment<>"-" then
				response.write " Ê÷ÌÕ: " & Comment
			else
				response.write " Ê÷ÌÕ ‰œ«—œ"
			end if
		%>">
			<TD align="right" valign=top><FONT COLOR="black">
			<INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RS3("id")%>" <%
			if RS3("status") = "new" then
				response.write " checked disabled "
			else 
				response.write " disabled "
			end if
			%>><%=RS3("ItemName")%> 
			<%
			if (not isNull(RS3("pickupListID"))) then 
				response.write "<a href='../inventory/default.asp?ed="&RS3("pickupListID")&"'>"
			end if
		%>
			<%
			if RS3("CustomerHaveInvItem")  then
				response.write "<b style='color:red'> «—”«·Ì </b>" 
			end if 
			%>
			
			<small dir=ltr>( ⁄œ«œ: <%=RS3("qtty")%> <%=RS3("unit")%> -  «—ÌŒ: <span dir=ltr><%=RS3("ReqDate")%></span>)</small>
			<%
			if (not isNull(RS3("pickupListID"))) then 
				response.write "</a>"
			end if
			if RS3("status") = "del" then response.write ("<font color=red>Õ–› ‘œÂ</font>")
		%>
			
			</td>
			<td align=left width=5%>
			<%
			
			if RS3("status") = "new" then
			%><a href="manageOrder.asp?di=y&i=<%=RS3("id")%>&r=<%=request("radif")%>&relatedApprovedInvoiceID=<%=relatedApprovedInvoiceID%>&relatedApprovedInvoiceBy=<%=relatedApprovedInvoiceBy%>"><b>Õ–›</b></a><%
			end if %></td>
		</tr>
		<% 
		RS3.moveNext
		Loop
		%>
	</TABLE>
</td>
<td valign=top>
	<FORM METHOD=POST ACTION="manageOrder.asp">
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
	<TR bgcolor="black" >
		<TD align="right" colspan=2><FONT COLOR="YELLOW">œ—ŒÊ«” Â«Ì Œ—Ìœ ”—ÊÌ” Ê ﬂ«·«:</FONT></TD>
	</TR>
		<TR bgcolor="#CCCCCC" ><td colspan=2>
		<center><!INPUT TYPE="submit" Name="Submit" Value="«ÌÃ«œ ”›«—‘ "  style="width:150px;" tabIndex="14"> </center></form>

		<FORM METHOD=POST ACTION="manageOrder.asp">
			<INPUT TYPE="hidden" name="relatedApprovedInvoiceBy" value="<%=relatedApprovedInvoiceBy%>">
			<INPUT TYPE="hidden" name="relatedApprovedInvoiceID" value="<%=relatedApprovedInvoiceID%>">
			<INPUT TYPE="hidden" name="radif" value="<%=request("radif")%>">
			<SELECT NAME="type" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1"  style="width:300">
			<option value="-1">‰Ê⁄ Œœ„  —« «‰ Œ«» ﬂ‰Ìœ: </option>
			<option value="-1">----------------------------------------------</option>
<%
				set RS4 = conn.Execute ("SELECT * FROM OutServices")
				while not (RS4.eof) %>
					<OPTION value="<%=RS4("ID")%>">* <%=RS4("Name")%></option>
<%						RS4.MoveNext
				wend
				RS4.close
				%>
			</SELECT><br>
			<SCRIPT LANGUAGE="JavaScript">
			<!--
			function hideIT()
			{
			if(document.all.tavafogh.checked) 
				{
					document.all.priceTavafoghi.style.visibility= 'visible'
				}
				else
				{
					document.all.price.value= ''
					document.all.priceComment.value= ''
					document.all.priceTavafoghi.style.visibility= 'hidden'
				}
			}
			//-->
			</SCRIPT><br>
			<INPUT TYPE="checkbox" onclick="hideIT()" name="tavafogh"> Ê«›ﬁ ﬁÌ„  Œ«’Ì ’Ê—  ê—› Â «” <BR>
			<div name="priceTavafoghi" id="priceTavafoghi" style="visibility:'hidden'">ﬁÌ„ : &nbsp;<INPUT TYPE="text" NAME="price" ID="price" size=5 onKeyPress="return maskNumber(this);"> ‘—Õ: <INPUT TYPE="text" NAME="priceComment" id="priceComment" size=26 ></div>

			 «—ÌŒÌ ﬂÂ „Ê—œ ‰Ì«“ «” : <INPUT dir=ltr TYPE="text" NAME="date1" size=12 value="<%=shamsiToday()%>" onKeyPress="return maskDate(this);"  onblur="acceptDate(this)" > &nbsp;  ⁄œ«œ: <INPUT dir=ltr TYPE="text" NAME="qtty" size=7  onKeyPress="return maskNumber(this);"><br><br>

			 Ê÷ÌÕ« : <TEXTAREA NAME="comment" ROWS="7" COLS="32"></TEXTAREA>
			<br><br><center>
			<INPUT TYPE="submit" Name="Submit" Value="À»  œ—ŒÊ«”  Œœ„« "  style="width:120px;" tabIndex="14" class="inputBut" <%
			if EDITABLE="NO"  then
				response.write " disabled "
			end if
			%>>
			</center>
		</FORM>
		<hr>

		</FONT></TD>
	</TR>
	<%
	'Gets Request for services list from DB
	'set RS3=Conn.Execute ("SELECT * FROM purchaseRequests WHERE (order_ID="& request("radif") & " ) and not status = 'del'")
	set RS3=Conn.Execute ("SELECT PurchaseRequestOrderRelations.*,purchaseRequests.*,case when isnull(PurchaseOrders.price,-1)=-1 then purchaseRequests.price else purchaseOrders.price end as thisPrice  FROM purchaseRequests LEFT OUTER JOIN PurchaseRequestOrderRelations ON PurchaseRequests.id = PurchaseRequestOrderRelations.Req_ID left outer join PurchaseOrders on PurchaseOrders.ID=PurchaseRequestOrderRelations.ord_id WHERE (order_ID="& request("radif") & " )")
	%>
		
		<%
		Do while not RS3.eof
		%>
		<TR bgcolor="#CCCCCC" title="<% 
			Comment = RS3("Comment")
			if Comment<>"-" then
				response.write " Ê÷ÌÕ: " & Comment
			else
				response.write " Ê÷ÌÕ ‰œ«—œ"
			end if
		%>">
			<TD align="right" valign=top><FONT COLOR="black">
			<INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RS3("id")%>" <%
			if RS3("status") = "new" then
				response.write " checked disabled "
			else 
				response.write " disabled "
			end if
			%>>
			<%
			if (not isNull(RS3("Ord_ID"))) then 
				response.write "<a href='../purchase/outServiceTrace.asp?od="&RS3("Ord_ID")&"'>"
			end if
			%>
			<%=RS3("typeName")%>  <small >( ⁄œ«œ: <%=RS3("qtty")%>° ﬁÌ„ : <%=RS3("thisPrice")%> -  «—ÌŒ: <span dir=ltr><%=RS3("ReqDate")%></span>)</small>
			<%
			if (not isNull(RS3("Ord_ID"))) then 
				response.write "</a>"
			end if
			%>
			</td>
			<td align=left width=5%><%
			if RS3("status") = "new" then
			%><a href="manageOrder.asp?d=y&i=<%=RS3("id")%>&r=<%=request("radif")%>&relatedApprovedInvoiceID=<%=relatedApprovedInvoiceID%>&relatedApprovedInvoiceBy=<%=relatedApprovedInvoiceBy%>"><b>Õ–›</b></a><%
			end if 
			if RS3("status") = "del" then response.write ("<font color=red>Õ–› ‘œÂ</font>")
			%>
			</td>
		</tr>
		<% 
		RS3.moveNext
		Loop
		%>
	</TABLE>
</td>
</tr>
</table>

<%

Conn.Close
%>
</font>
<!--#include file="tah.asp" -->
