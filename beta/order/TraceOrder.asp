<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="����� �����"
SubmenuItem=3
if not Auth(2 , 3) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<%
'Server.ScriptTimeout = 3600
%>
<STYLE>
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	a.aYellow:link {color: yellow;}
	a.aYellow:visited {color: green;}
	a.aYellow:hover {color: gray;}
</STYLE>
<SCRIPT LANGUAGE='JavaScript'>
<!--
function checkValidation(){
	if (document.all.search_box.value != ''){
		return true;
	}
	else{
		document.all.search_box.focus();
		return false;
	}
}
//-->
</SCRIPT>
<font face="tahoma">
<%
if request("act")="" then
%>
	<hr>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="600" align="center" bordercolor="#555599">
	<FORM METHOD=POST ACTION="TraceOrder.asp?act=search" onSubmit="return checkValidation();">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="5" dir="RTL" width="100%">
		<TR bgcolor="#AAAAEE">
			<TD>��� ����� �� ���� �� ����� �����:</TD>
			<TD><INPUT TYPE="text" NAME="search_box" value="<%=request.form("search_box")%>"></TD>
			<TD><INPUT TYPE="submit" NAME="SubmitB" Value="�����" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
			<TD align="left"><% if Auth(2 , 5) then %><A HREF="TraceOrder.asp?act=advancedSearch">������ �������</A><% End If %></TD>
		</TR>
		</TABLE>
	</TD></TR>
	</FORM>
	</TABLE>
	<script language="JavaScript">
	<!--
		document.all.search_box.focus();
	//-->
	</script>
	<hr>
<%
	'
elseif request("act")="search" then
%>
	<hr>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="600" align="center" bordercolor="#555599">
	<FORM METHOD=POST ACTION="TraceOrder.asp?act=search" onSubmit="return checkValidation();">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="5" dir="RTL" width="100%">
		<TR bgcolor="#AAAAEE">
			<TD>��� ����� �� ���� �� ����� �����:</TD>
			<TD><INPUT TYPE="text" NAME="search_box" value="<%=request.form("search_box")%>"></TD>
			<TD><INPUT TYPE="submit" NAME="SubmitB" Value="�����" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
			<TD align="left"><% if Auth(2 , 5) then %><A HREF="TraceOrder.asp?act=advancedSearch">������ �������</A><% End If %></TD>
		</TR>
		</TABLE>
	</TD></TR>
	</FORM>
	</TABLE>
	<script language="JavaScript">
	<!--
		document.all.search_box.focus();
	//-->
	</script>
	<hr>
<%
	
	search=request("search_box")
	if search="" then
		'By Default show Open Orders of Current User
		myCriteria= "Orders.CreatedBy = " & session("ID")
	elseif isNumeric(search) then
		search=clng(search)
		myCriteria= "radif_sefareshat = '"& search & "'"
	else
		search=sqlSafe(search)
		myCriteria= "REPLACE([company_name], ' ', '') LIKE REPLACE(N'%"& search & "%', ' ', '') OR REPLACE([customer_name], ' ', '') LIKE REPLACE(N'%"& search & "%', ' ', '')"
	End If

'	mySQL="SELECT orders_trace.*, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon FROM Orders INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN OrderTraceStatus ON orders_trace.status = OrderTraceStatus.ID WHERE ("& myCriteria & ") AND (Orders.Closed=0) ORDER BY order_date DESC, radif_sefareshat DESC"	
	mySQL="SELECT orders_trace.*, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon , Invoices.ID AS InvoiceID, Invoices.Approved, Invoices.Voided, Invoices.Issued,DRV_invoice.price FROM Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice RIGHT OUTER JOIN Orders INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN OrderTraceStatus ON orders_trace.status = OrderTraceStatus.ID ON InvoiceOrderRelations.[Order] = Orders.ID left outer join (select InvoiceOrderRelations.[Order],SUM(InvoiceLines.Price + InvoiceLines.Vat - InvoiceLines.Discount -InvoiceLines.Reverse) as price from InvoiceOrderRelations inner join Invoices on InvoiceOrderRelations.Invoice=Invoices.ID inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice where Invoices.Voided=0 group by InvoiceOrderRelations.[Order]) DRV_invoice on Orders.id=DRV_invoice.[Order] WHERE ("& myCriteria & ") AND (Orders.Closed=0) ORDER BY orders_trace.order_date DESC, orders_trace.radif_sefareshat DESC"

	set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		tmpCounter=0
%>
	<div align="center" dir="LTR" >
	<table border="1" cellspacing="0" cellpadding="2" dir="RTL"  borderColor="#555588" >
		<TR valign=top bgcolor="#CCCCFF">
			<TD width="40"># �����</TD>
			<TD width="65">����� �����<br>����� �����</TD>
			<TD width="130">��� ���� - �����</TD>
			<TD width="80">����� ���</TD>
			<TD width="36">���</TD>
			<TD width="45">�����</TD>
			<TD width="38">����� ������</TD>
			<TD width="18">���</TD>
			<TD width="30">������</TD>
			<td width="50">����</td>
		</TR>
<%		Do while not RS1.eof

		if isnull(RS1("InvoiceID")) then
			InvoiceStatus="<span style='color:red;'><b>�����</b></span>"
		else
			if RS1("Voided") then
				style="style='color:Red' Title='���� ���'"
			elseif RS1("Issued") then
				style="style='color:Red' Title='���� ���'"
			elseif RS1("Approved") then
				style="style='color:Green' Title='����� ���'"
			else
				style="style='color:#3399FF' Title='����� ����'"
			end if
			InvoiceStatus="<A " & style & " HREF='../AR/AccountReport.asp?act=showInvoice&invoice=" & RS1("InvoiceID")& "' Target='_blank'>" & RS1("InvoiceID") & "</A>"
		end if
			
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
		Else
			tmpColor="#EEEEEE"
		End If 
%>
		<TR valign=top bgcolor="<%=tmpColor%>">
			<TD DIR="LTR"><A HREF="TraceOrder.asp?act=show&order=<%=RS1("radif_sefareshat")%>" target="_blank"><%=RS1("radif_sefareshat")%></A></TD>
			<TD DIR="LTR"><%=RS1("order_date")%><br><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
			<TD><%=RS1("company_name")%><br><span style='color:gray'><%=RS1("customer_name")%></span><br>����:(<%=RS1("telephone")%>)&nbsp;</TD>
			<TD><%=RS1("order_title")%>&nbsp;</TD>
			<TD><%=RS1("order_kind")%></TD>
			<TD><%=RS1("marhale")%></TD>
			<TD><%=RS1("salesperson")%>&nbsp;</TD>
			<TD><IMG SRC="<%=RS1("Icon")%>" WIDTH="20" HEIGHT="20" BORDER=0 ALT="<%=RS1("StatusName")%>"></TD>
			<TD><%=InvoiceStatus%>&nbsp;</TD>
			<td><%if isnull(RS1("price")) then response.write "----" else response.write Separate(RS1("price")) end if %></td>
		</TR>
		<TR bgcolor="#FFFFFF">
			<TD colspan="10" style="height:10px"></TD>
		</TR>
<%			RS1.moveNext
		Loop
%>		<TR bgcolor="#ccccFF">
			<TD colspan="10">����� ����� �����: <%=tmpCounter%></TD>
		</TR>
	</TABLE>
	</div>
	<HR>
<%	elseif request("search_box")<>"" then
%>	<TABLE border="1" cellspacing="0" cellpadding="0" dir="RTL" align="center" width="600">
		<TR bgcolor="#FFFFDD">
			<TD align="center" style="height:40px;font-size:12pt;font-weight:bold;color:red;padding:5px;">��� ����� ���� ����� ����<br>
			(�� ���� ������ ��� �� ���� ��� � �� ���� ���� �����)<br><br>
			���� ������� �� <A HREF="TraceOrder.asp?act=show&order=<%=request("search_box")%>" style="color:blue;">����� � �����</A> ������ ����.</TD>
		</TR>
	</TABLE>
<%		End If
elseif request("act")="show" then
  if isnumeric(request("order")) then
	Order=request("order")
	mySQL="SELECT orders_trace.*, Accounts.ID AS AccID, Accounts.AccountTitle FROM Orders INNER JOIN Accounts ON Orders.Customer = Accounts.ID INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat WHERE (orders_trace.radif_sefareshat='"& Order & "')"
	set RS1=conn.execute (mySQL)
	if RS1.EOF then
		response.write "<BR><BR><BR><BR><CENTER>����� ����� ����� ����</CENTER>"
		response.end
	End If

	if RS1("Status")=2 then
		stamp="<div style='border:2 dashed red;width:150px; text-align:center; padding: 10px;color:red;font-size:15pt;font-weight:bold;'>���� ���</div>"
	End If

%>
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="10">
			<td colspan=2></td>
		</tr>
		<tr height="10">
			<td width="150"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="20">
			<td colspan=2></td>
		</tr>
	</table>

	<CENTER>
		<input type="button" value="�����" Class="GenButton" onclick="window.location='OrderEdit.asp?e=y&radif=<%=Order%>';">&nbsp;
		<% 	ReportLogRow = PrepareReport ("OrderForm.rpt", "Order_ID", Order, "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT TYPE="button" value=" �ǁ " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">
		<input type="button" value="������" Class="GenButton" onclick="window.location='../AR/InvoiceInput.asp?act=submitsearch&query=<%=Order%>';">
		<!-----------------------SAM----------------------------
		<INPUT type='button' value='������' class='GenButton' style='border:1 solid blue;' onclick=''-->
	</CENTER>
	
	<BR>
	<TABLE cellspacing=0 Style="width:80%;border:2 solid #330066" align=center>
	<TR bgcolor=white>
		<TD>�����</TD>
		<TD>����</TD>
		<TD>�����</TD>
		<TD>�����</TD>
		<TD>��� ��� �����</TD>
	</TR>
	<%
	set RS_STEP=Conn.Execute ("SELECT OrderTraceLog.*, Users.RealName FROM OrderTraceLog INNER JOIN Users ON OrderTraceLog.InsertedBy = Users.ID WHERE (OrderTraceLog.[Order] = "& Order  & ") order by OrderTraceLog.ID")
	Do while not RS_STEP.eof	
%>
	<TR style="cursor:pointer" onclick="window.open('viewOrderLog.asp?logid=<%=RS_STEP("id")%>','orderLog','width=700, height=210')">
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
	'set RS_STEP = nothing
%>
	<tr>
		<td  align="center"><A HREF="../shopfloor/default.asp?orderNum=<%=order%>">����� �����</A></td>
		<td colspan="2" title="�� ���� ��� ��� ���� �� ���� ������ �� ����� �� ����� ����� ����� �� �� ���� ����� ����� � ����� ����� ����� ���">
			<%
			set rsEmail=Conn.Execute("select accounts.AccountTitle, accounts.Dear1, accounts.FirstName1, accounts.LastName1, orders.ID, orders.Customer,accounts.Email1, orders_trace.order_title from Orders inner join Accounts on orders.Customer=accounts.ID inner join orders_trace on orders_trace.radif_sefareshat=orders.ID where orders.ID=" & Order & " and accounts.EMail1 <> ''")
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
					<input type="submit" name="orderSend" title='<%=rsEmail("email1")%>' value="�� <%=rsEmail("Dear1") & " " & rsEmail("firstName1") & " " & rsEmail("LastName1")%> ����� ���">
				</form>
			</span>
			<%end if%>
		</td>
		<td colspan="2" title="�� ���� ��� ��� ���� �� ���� ������ �� ����� �� ����� ����� ����� �� �� ���� ����� ����� � ����� ����� ����� ���">
			<%
			set rsEmail=Conn.Execute("select accounts.AccountTitle, accounts.Dear2, accounts.FirstName2, accounts.LastName2, orders.ID, orders.Customer,accounts.Email2, orders_trace.order_title from Orders inner join Accounts on orders.Customer=accounts.ID inner join orders_trace on orders_trace.radif_sefareshat=orders.ID where orders.ID=" & Order & " and accounts.EMail2 <> ''")
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
					<input type="submit" name="orderSend" title='<%=rsEmail("email2")%>' value="�� <%=rsEmail("Dear2") & " " & rsEmail("firstName2") & " " & rsEmail("LastName2")%> ����� ���">
				</form>
			</span>
			<%end if%>
		</td>
	</tr>
	</TABLE><BR>
	<BR>
	<TABLE class="" border="0" cellspacing="0" cellpadding="2" align="center" style="background-color:#CCCCCC; color:black; direction:RTL; width:700; border: 2 solid black;">
	<TR bgcolor="black">
		<TD align="left"><FONT COLOR="YELLOW">����:</FONT></TD>
		<TD align="right" colspan=5 height="25px">
			<span id="customer" style="color:yellow;"><%' after any changes in this span "./Customers.asp" must be revised%>
				<span title="���� ����� ������ ����� ���� ����"><a class="aYellow" href='../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=RS1("AccID")%>'><%=RS1("AccID") & " - "& RS1("AccountTitle")%></a></span>.
			</span>
		</TD>
	</TR>
	
	<TR bgcolor="black" height=30 style="color:yellow;">
		<TD align="left">����� �����:</TD>
		<TD align="right"><%=RS1("radif_sefareshat")%></TD>
		<TD align="left">�����:</TD>
		<TD><span dir="LTR"><%=RS1("order_date")%></span></TD>
		<TD align="left">����:</TD>
		<TD align="right"><%=RS1("order_time")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">��� ����:</TD>
		<TD><%=RS1("company_name")%></TD>
		<TD align="left">���� �����:</TD>
		<TD align="right" dir=LTR><%=RS1("return_date")%></TD>
		<TD align="left">���� �����:</TD>
		<TD align="right" dir=LTR><%=RS1("return_time")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">��� �����:</TD>
		<TD><%=RS1("customer_name")%></TD>
		<TD align="left">��� �����:</TD>
		<TD><%=RS1("order_kind")%></TD>
		<TD align="left">����� ������:</TD>
		<TD><%=RS1("salesperson")%>	</TD>
	</TR>
	<TR height=30>
		<TD align="left">����:</TD>
		<TD><%=RS1("telephone")%></TD>
		<TD align="left">����� ��� ���� ����:</TD>
		<TD colspan="4"><%=RS1("order_title")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">�����:</TD>
		<TD><%=RS1("qtty")%></TD>
		<TD align="left">����:</TD>
		<TD><%=RS1("PaperSize")%></TD>
		<TD align="left">����/����:</TD>
		<TD><%=RS1("SimplexDuplex")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">�����:</TD>
		<TD><%=RS1("marhale")%></TD>
		<TD align="left">�����:</TD>
		<TD colspan="3"><%=RS1("vazyat")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">���� ��:</TD>
		<TD><%=RS1("Price")%></TD>
		<TD colspan="4" height="30px">&nbsp;</TD>
	</TR>
	<!--TR height=30>
		<TD colspan="6" align="center"><input type="button" value="�����" onclick="window.location='OrderEdit.asp?e=y&radif=<%=Order %>';"></TD>
	</TR-->
	</TABLE><BR>

	<table class="CustTable" cellspacing='1' align=center style="width:700; ">
		<tr>
			<td colspan="2" class="CusTableHeader"><span style="width:450;text-align:center;">������� ��</span><span style="width:100;text-align:left;background-color:red;"><input class="GenButton" type="button" value="����� �������" onclick="window.location = '../home/message.asp?RelatedTable=orders&RelatedID=<%=Order%>&retURL=<%=Server.URLEncode("../order/TraceOrder.asp?act=show&order="&Order)%>';"></span></td>
		</tr>
<%
	mySQL="SELECT * FROM Messages INNER JOIN Users ON Messages.MsgFrom = Users.ID WHERE (Messages.RelatedTable = 'orders') AND (Messages.RelatedID = "& Order & ") ORDER BY Messages.ID DESC"
	Set RS = conn.execute(mySQL)
	if NOT RS.eof then

		tmpCounter=0
		Do While NOT RS.eof 
			tmpCounter=tmpCounter+1
%>
			<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
				<td>�� <%=RS("RealName")%><br>
					<%=RS("MsgDate")%> <BR> <%=RS("MsgTime")%>
				</td>
				<td dir='RTL'><%=replace(RS("MsgBody"),chr(13),"<br>")%></td>
			</tr>
<%
			RS.moveNext
		Loop
	else
%>
		<tr class="CusTD3">
			<td colspan="2">��</td>
		</tr>
<%
	end if
	RS.close
%>
	</table><BR>
		<%
		dim fs
		set fs=Server.CreateObject("Scripting.FileSystemObject")
		if not fs.FolderExists(orderFolder & order)=true then
		  fs.CreateFolder (orderFolder & order)  
		end if
		set fs=nothing
		%>
		<a href="<%=orderFolder & order%>">����� ���� �����</a>
		
	<div style='direction:ltr;float:left;'>
	<% ListFolderContents(orderFolder & order) %>
	<div>
	<TABLE border="0" cellspacing="0" cellpadding="2" align="center" style=" color:black; direction:RTL; width:700; ">
	<TR>
		<TD valign=top>
			<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
			<TR bgcolor="black" >
				<TD align="right" colspan=2><FONT COLOR="YELLOW">���������� ���� �� �����:</FONT></TD>
			</TR>
			<%
			'Gets Request for services list from DB
			'-----------------_SAM EDIT THIS 
			set RS3=Conn.Execute ("SELECT InventoryItemRequests.*,InventoryPickuplistItems.pickupListID  FROM InventoryItemRequests left outer join InventoryPickuplistItems on InventoryItemRequests.ID=InventoryPickuplistItems.RequestID WHERE InventoryItemRequests.order_ID="& Order )
			%>
				<%
				Do while not RS3.eof
				%>
				<TR bgcolor="#CCCCCC" title="<% 
					Comment = RS3("Comment")
					if Comment<>"-" then
						response.write "�����: " & Comment
					else
						response.write "����� �����"
					End If
				%>">
					<TD align="right" valign=top><FONT COLOR="black">
					<INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RS3("id")%>" <%
					if RS3("status") = "new" then
						response.write " checked disabled "
					else 
						response.write " disabled "
					End If
					%>>
					<%
					if (not isNull(RS3("pickupListID"))) then 
						response.write "<a href='../inventory/default.asp?ed="&RS3("pickupListID")&"'>"
					end if
				%>
					<%=RS3("ItemName")%>  
					<%
					if RS3("CustomerHaveInvItem")  then
						response.write "<small><b style='color:red'> ������ </b></small>" 
					end if 
					%>
					<small dir=ltr>(�����: <%=RS3("qtty")%> <%=RS3("unit")%> - �����: <span dir=ltr><%=RS3("ReqDate")%></span>)</small>
				<%
					if (not isNull(RS3("pickupListID"))) then 
						response.write "</a>"
					end if
				%>
					</td>
					<td align=left width=5%><%
					if RS3("status") = "del" then
						response.write "<b><small>��� ���</b></small>"
					End If
					%></td>
				</tr>
				<% 
				RS3.moveNext
				Loop
				%>
			</TABLE>
		</TD>

		<TD  valign=top>
			<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" width="350" >
			<TR bgcolor="black" >
				<TD align="right" colspan=2><FONT COLOR="YELLOW">���������� ���� ����� � ����:</FONT></TD>
			</TR>
			<%
			'Gets Request for services list from DB
			set RS3=Conn.Execute ("SELECT PurchaseRequestOrderRelations.*,purchaseRequests.*,case when isnull(PurchaseOrders.price,-1)=-1 then purchaseRequests.price else purchaseOrders.price end as thisPrice FROM purchaseRequests LEFT OUTER JOIN PurchaseRequestOrderRelations ON PurchaseRequests.id = PurchaseRequestOrderRelations.Req_ID left outer join PurchaseOrders on PurchaseOrders.ID=PurchaseRequestOrderRelations.ord_id WHERE (order_ID="& Order & " )")
			%>
				
				<%
				Do while not RS3.eof
				%>
				<TR bgcolor="#CCCCCC" title="<% 
					Comment = RS3("Comment")
					if Comment<>"-" then
						response.write "�����: " & Comment
					else
						response.write "����� �����"
					End If
				%>">
					<TD align="right" valign=top><FONT COLOR="black">
					<INPUT TYPE="checkbox" NAME="outReq" VALUE="<%=RS3("id")%>" <%
					if RS3("status") = "new" then
						response.write " checked disabled "
					else 
						response.write " disabled "
					End If
					%>>
					<%
					if (not isNull(RS3("Ord_ID"))) then 
						response.write "<a href='../purchase/outServiceTrace.asp?od="&RS3("Ord_ID")&"'>"
					end if
					%>
					<%=RS3("typeName")%>  <small >(�����: <%=RS3("qtty")%>� ����: <%=RS3("thisPrice")%> - �����: <span dir=ltr><%=RS3("ReqDate")%></span>)</small>
					<%
					if (not isNull(RS3("Ord_ID"))) then 
						response.write "</a>"
					end if
					%>				
					</td>
					<td align=left width=5%><%
					if RS3("status") = "del" then
						response.write "<b><small>��� ���</b></small>"
					End If
					%>
					</td>
				</tr>
				<% 
				RS3.moveNext
				Loop
				%>
			</TABLE>
		</TD>
	</TR>
	</TABLE><BR>
	<!------------------------------------------------------SAM----------------------------------------------------------->

<%
end if
elseif request("act")="advancedSearch" then
'------  Advanced Search 
%>
<!--#include File="../include_JS_InputMasks.asp"-->
<%
	'Server.ScriptTimeout = 3600
	tmpTime=time
	tmpTime=Hour(tmpTime)&":"&Minute(tmpTime)
	if instr(tmpTime,":")<3 then tmpTime="0" & tmpTime
	if len(tmpTime)<5 then tmpTime=Left(tmpTime,3) & "0" & Right(tmpTime,1)

	if request.form("resultsCount")="" OR not isnumeric(request.form("resultsCount")) then
		resultsCount = 50
	else
		resultsCount = cint(request.form("resultsCount"))
	end if

%>
	<hr>
	<TABLE border="4" cellspacing="0" cellpadding="0" width="700" align="center" bordercolor="#555599">
	<FORM METHOD=POST ACTION="TraceOrder.asp?act=advancedSearch" onSubmit="return checkValidation();">
	<TR><TD>
		<TABLE border="0" cellspacing="0" cellpadding="1" dir="RTL" width="100%" bgcolor="white">
		<TR bgcolor="#EEEEEE">
			<TD><INPUT TYPE="checkbox" NAME="check_sefaresh" onclick="check_sefaresh_Click()" checked></TD>
			<TD>����� �����</TD>
			<TD><INPUT TYPE="text" NAME="az_sefaresh" dir="LTR" value="<%=request.form("az_sefaresh")%>" size="8" maxlength="6" onKeyPress="return maskNumber(this);"></TD>
			<TD>��</TD>
			<TD><INPUT TYPE="text" NAME="ta_sefaresh" dir="LTR" value="<%=request.form("ta_sefaresh")%>" size="8" maxlength="6" onKeyPress="return maskNumber(this);" ></TD>
			<td rowspan="12" style="width:1px" bgcolor="#555599"></td>
			<TD><INPUT TYPE="checkbox" NAME="check_kind" onclick="check_kind_Click()" checked></TD>
			<TD>��� �����</TD>
			<TD colspan="3"><SELECT NAME="order_kind_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
				<OPTION value="����" <%if request.form("order_kind_box")="����" then response.write "selected" %> >����</option>
				<OPTION value="�������" <%if request.form("order_kind_box")="�������" then response.write "selected" %> >�������</option>
				<OPTION value="���� � ����" <%if request.form("order_kind_box")="���� � ����" then response.write "selected" %> >���� � ����</option>
				<OPTION value="�����" <%if request.form("order_kind_box")="�����" then response.write "selected" %> >�����</option>
				<OPTION value="�����" <%if request.form("order_kind_box")="�����" then response.write "selected" %> >�����</option>
				<OPTION value="����" <%if request.form("order_kind_box")="����" then response.write "selected" %> >����</option>
				<OPTION value="����" <%if request.form("order_kind_box")="����" then response.write "selected" %> >����</option>
				<OPTION value="�����" <%if request.form("order_kind_box")="�����" then response.write "selected" %> >�����</option>
				<OPTION value="������" <%if request.form("order_kind_box")="������" then response.write "selected" %> >������</option>
			</SELECT></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD><INPUT TYPE="checkbox" NAME="check_tarikh_sefaresh" onclick="check_tarikh_sefaresh_Click()" checked></TD>
			<TD>����� �����</TD>
			<TD><INPUT TYPE="text" NAME="az_tarikh_sefaresh" dir="LTR" value="<%=request.form("az_tarikh_sefaresh")%>" size="10" onKeyPress="return maskDate(this);" onBlur="if(acceptDate(this))document.all.ta_tarikh_sefaresh.value=this.value;" maxlength="10"></TD>
			<TD>��</TD>
			<TD><INPUT TYPE="text" NAME="ta_tarikh_sefaresh" dir="LTR" value="<%=request.form("ta_tarikh_sefaresh")%>" size="10" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_marhale" onclick="check_marhale_Click()" checked></TD>
			<TD>�����</TD>
			<TD><SELECT NAME="marhale_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
					<%
					set RS_STEP=Conn.Execute ("SELECT * FROM OrderTraceSteps WHERE (IsActive=1)")
					Do while not RS_STEP.eof	
					%>
						<OPTION value="<%=RS_STEP("ID")%>" <%if cint(request.form("marhale_box"))=cint(RS_STEP("ID")) then response.write "selected" %> ><%=RS_STEP("name")%></option>
						<%
						RS_STEP.moveNext
					loop
					RS_STEP.close
					%>
			</SELECT></TD>
			<TD><span id="marhale_not_check_label" style='font-weight:bold;color:red'>�����</span></TD>
			<TD><INPUT TYPE="checkbox" NAME="marhale_not_check" onclick="marhale_not_check_Click();" checked></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD><INPUT TYPE="checkbox" NAME="check_tarikh_tahvil" onclick="check_tarikh_tahvil_Click()" checked></TD>
			<TD>����� �����</TD>
			<TD><INPUT TYPE="text" NAME="az_tarikh_tahvil" dir="LTR" value="<%=request.form("az_tarikh_tahvil")%>" size="10" onblur="acceptDate(this)" maxlength="10" onKeyPress="return maskDate(this);"></TD>
			<TD>��</TD>
			<TD><INPUT TYPE="text" NAME="ta_tarikh_tahvil" dir="LTR" value="<%=request.form("ta_tarikh_tahvil")%>" onblur="acceptDate(this)" maxlength="10" size="10" onKeyPress="return maskDate(this);"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_vazyat" onclick="check_vazyat_Click()" checked></TD>
			<TD>�����</TD>
			<TD><SELECT NAME="vazyat_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
				<%
				set RS_STATUS=Conn.Execute ("SELECT * FROM OrderTraceStatus WHERE (IsActive=1)")
				Do while not RS_STATUS.eof	
				%>
					<OPTION value="<%=RS_STATUS("ID")%>" <%if cint(request.form("vazyat_box"))=cint(RS_STATUS("ID")) then response.write "selected" %> ><%=RS_STATUS("Name")%></option>
					<%
					RS_STATUS.moveNext
				loop
				%>
			</SELECT></TD>
			<TD><span id="vazyat_not_check_label" style='font-weight:bold;color:red'>�����</span></TD>
			<TD><INPUT TYPE="checkbox" NAME="vazyat_not_check" onclick="vazyat_not_check_Click();" checked></TD>
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD colspan="5">&nbsp;</TD>
			<TD><INPUT TYPE="checkbox" NAME="check_closed" onclick="check_closed_Click()" checked></TD>
			<TD colspan="4"><span id="check_closed_label" style='color:black;'>��� ����� ��� ���</span></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD colspan="5">&nbsp;</TD>
			<TD><INPUT TYPE="checkbox" NAME="check_salesperson" onclick="check_salesperson_Click()" checked></TD>
			<TD>����� ������</TD>
			<TD colspan="3">
			<SELECT NAME="salesperson_box" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px'>
<%				set RSV=Conn.Execute ("SELECT RealName FROM Users WHERE Display=1 ORDER BY RealName") 
				Do while not RSV.eof
%>
					<option value="<%=RSV("RealName")%>" <%if RSV("RealName")=request.form("salesperson_box") then response.write " selected "%> ><%=RSV("RealName")%></option>
<%
				RSV.moveNext
				Loop
				RSV.close
%>
			</SELECT></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#EEEEEE">
			<TD><INPUT TYPE="checkbox" NAME="check_company_name" onclick="check_company_name_Click()" checked></TD>
			<TD>��� ����</TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="company_name_box" value="<%=request.form("company_name_box")%>"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_telephone" onclick="check_telephone_Click()" checked></TD>
			<TD>����</TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="telephone_box" value="<%=request.form("telephone_box")%>"></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR>
			<TD><INPUT TYPE="checkbox" NAME="check_customer_name" onclick="check_customer_name_Click()" checked></TD>
			<TD>��� �����</TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="customer_name_box" value="<%=request.form("customer_name_box")%>"></TD>
			<TD><INPUT TYPE="checkbox" NAME="check_order_title" onclick="check_order_title_Click()" checked></TD>
			<TD>����� �����</TD>
			<TD colspan="3"><INPUT TYPE="text" NAME="order_title_box" value="<%=request.form("order_title_box")%>"></TD>
		</TR>
		<TR bgcolor="#555599">
			<td colspan="11" style="height:2px"></td>
		</TR>
		<TR bgcolor="#AAAAEE">
			<td colspan="11" style="height:30px">
			<TABLE align="center" width="50%">
			<TR>
				<TD><INPUT TYPE="submit" Name="Submit" Value="�����" style="font-family:tahoma,arial; font-size:10pt;width:100px;"></TD>
				<TD align="left"><INPUT TYPE="button" Name="Cancel" Value="��� ��" style="font-family:tahoma,arial; font-size:10pt;width:100px;" onclick="window.location='TraceOrder.asp?act=advancedSearch';"></TD>
			</TR>
			<TR>
				<TD align="left">����� �����:</TD>
				<TD>&nbsp;<INPUT TYPE="Text" Name="resultsCount" value="<%=resultsCount%>" maxlength="4" size="4" onKeyPress="return maskNumber(this);"></TD>
			</TR>
			</TABLE>
			</td>
		</TR>
		</TABLE>
	</TD></TR>
	</FORM>
	</TABLE>
	<hr>
	<SCRIPT LANGUAGE='JavaScript'>
	<!--
	function check_sefaresh_Click(){
		if ( document.all.check_sefaresh.checked ) {
			document.all.az_sefaresh.style.visibility = "visible";
			document.all.ta_sefaresh.style.visibility = "visible";
			document.all.az_sefaresh.focus();
		}
		else{
			document.all.az_sefaresh.style.visibility = "hidden";
			document.all.ta_sefaresh.style.visibility = "hidden";
		}
	}

	function check_tarikh_sefaresh_Click(){
		if ( document.all.check_tarikh_sefaresh.checked ) {
			document.all.az_tarikh_sefaresh.style.visibility = "visible";
			document.all.ta_tarikh_sefaresh.style.visibility = "visible";
			document.all.az_tarikh_sefaresh.focus();
		}
		else{
			document.all.az_tarikh_sefaresh.style.visibility = "hidden";
			document.all.ta_tarikh_sefaresh.style.visibility = "hidden";
		}
	}

	function check_tarikh_tahvil_Click(){
		if ( document.all.check_tarikh_tahvil.checked ) {
			document.all.az_tarikh_tahvil.style.visibility = "visible";
			document.all.ta_tarikh_tahvil.style.visibility = "visible";
			document.all.az_tarikh_tahvil.focus();
		}
		else{
			document.all.az_tarikh_tahvil.style.visibility = "hidden";
			document.all.ta_tarikh_tahvil.style.visibility = "hidden";
		}
	}

	function check_company_name_Click(){
		if (document.all.check_company_name.checked) {
			document.all.company_name_box.style.visibility = "visible";
			document.all.company_name_box.focus();
		}
		else{
			document.all.company_name_box.style.visibility = "hidden";
		}
	}

	function check_customer_name_Click(){
		if (document.all.check_customer_name.checked) {
			document.all.customer_name_box.style.visibility = "visible";
			document.all.customer_name_box.focus();
		}
		else{
			document.all.customer_name_box.style.visibility = "hidden";
		}
	}

	function check_kind_Click(){
		if (document.all.check_kind.checked) {
			document.all.order_kind_box.style.visibility = "visible";
			document.all.order_kind_box.focus();
		}
		else{
			document.all.order_kind_box.style.visibility = "hidden";
		}
	}

	function check_marhale_Click(){
		if (document.all.check_marhale.checked) {
			document.all.marhale_box.style.visibility = "visible";
			document.all.marhale_box.focus();
			document.all.marhale_not_check.style.visibility = "visible";
			marhale_not_check_Click();	
		}
		else{
			document.all.marhale_box.style.visibility = "hidden";
			document.all.marhale_not_check.style.visibility = "hidden";
			document.all.marhale_not_check_label.style.color='#BBBBBB'
		}
	}

	function check_vazyat_Click(){
		if (document.all.check_vazyat.checked) {
			document.all.vazyat_box.style.visibility = "visible";
			document.all.vazyat_box.focus();
			document.all.vazyat_not_check.style.visibility = "visible";
			vazyat_not_check_Click();
		}
		else{
			document.all.vazyat_box.style.visibility = "hidden";
			document.all.vazyat_not_check.style.visibility = "hidden";
			document.all.vazyat_not_check_label.style.color='#BBBBBB'
		}
	}

	function check_salesperson_Click(){
		if (document.all.check_salesperson.checked) {
			document.all.salesperson_box.style.visibility = "visible";
			document.all.salesperson_box.focus();
		}
		else{
			document.all.salesperson_box.style.visibility = "hidden";
		}
	}

	function check_telephone_Click(){
		if (document.all.check_telephone.checked) {
			document.all.telephone_box.style.visibility = "visible";
			document.all.telephone_box.focus();
		}
		else{
			document.all.telephone_box.style.visibility = "hidden";
		}
	}

	function check_order_title_Click(){
		if (document.all.check_order_title.checked) {
			document.all.order_title_box.style.visibility = "visible";
			document.all.order_title_box.focus();
		}
		else{
			document.all.order_title_box.style.visibility = "hidden";
		}
	}

	function marhale_not_check_Click(){
		if (document.all.marhale_not_check.checked) {
			document.all.marhale_not_check_label.style.color='red'
		}
		else{
			document.all.marhale_not_check_label.style.color='#BBBBBB'
		}
	}


	function vazyat_not_check_Click(){
		if (document.all.vazyat_not_check.checked) {
			document.all.vazyat_not_check_label.style.color='red'
		}
		else{
			document.all.vazyat_not_check_label.style.color='#BBBBBB'
		}
	}

	function check_closed_Click(){
		if (document.all.check_closed.checked) {
			document.all.check_closed_label.style.color='black'
		}
		else{
			document.all.check_closed_label.style.color='#BBBBBB'
		}
	}


	function Form_Load(){
	<%
	maybeAND = ""
	myCriteria = ""
	If request.form("check_sefaresh") = "on" Then
		if request.form("az_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "radif_sefareshat >= " & request.form("az_sefaresh")
			maybeAND=" AND "
		End If
		if request.form("ta_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "radif_sefareshat <= " & request.form("ta_sefaresh")
			maybeAND=" AND "
		End If
		If (request.form("az_sefaresh") = "") AND (request.form("ta_sefaresh") = "") then
			response.write "document.all.check_sefaresh.checked = false;" & vbCrLf
			response.write "document.all.az_sefaresh.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_sefaresh.style.visibility = 'hidden';" & vbCrLf
		End If
	Else
		response.write "document.all.check_sefaresh.checked = false;" & vbCrLf
		response.write "document.all.az_sefaresh.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_sefaresh.style.visibility = 'hidden';" & vbCrLf

	End If

	If request.form("check_tarikh_sefaresh") = "on" Then
		if request.form("az_tarikh_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "order_date >= '" & request.form("az_tarikh_sefaresh") & "'"
			maybeAND=" AND "
		End If
		if request.form("ta_tarikh_sefaresh") <> "" then
			myCriteria = myCriteria & maybeAND & "order_date <= '" & request.form("ta_tarikh_sefaresh") & "'"
			maybeAND=" AND "
		End If 
		If (request.form("az_tarikh_sefaresh") = "") AND (request.form("ta_tarikh_sefaresh") = "") then
			response.write "document.all.check_tarikh_sefaresh.checked = false;" & vbCrLf
			response.write "document.all.az_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
		End If
	Else
		response.write "document.all.check_tarikh_sefaresh.checked = false;" & vbCrLf
		response.write "document.all.az_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_tarikh_sefaresh.style.visibility = 'hidden';" & vbCrLf
	End If

	If request.form("check_tarikh_tahvil") = "on" Then
		if request.form("az_tarikh_tahvil") <> "" then

			myCriteria = myCriteria & maybeAND & "return_date >= '" & request.form("az_tarikh_tahvil") & "'"
			maybeAND=" AND "
		End If
		if request.form("ta_tarikh_tahvil") <> "" then
			myCriteria = myCriteria & maybeAND & "return_date <= '" & request.form("ta_tarikh_tahvil") & "'"
			maybeAND=" AND "
		End If
		If (request.form("az_tarikh_tahvil") = "") AND (request.form("ta_tarikh_tahvil") = "") then
			response.write "document.all.check_tarikh_tahvil.checked = false;" & vbCrLf
			response.write "document.all.az_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
			response.write "document.all.ta_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
		End If
	Else
		response.write "document.all.check_tarikh_tahvil.checked = false;" & vbCrLf
		response.write "document.all.az_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.ta_tarikh_tahvil.style.visibility = 'hidden';" & vbCrLf
	End If

	If (request.form("check_company_name") = "on" AND request.form("company_name_box") <> "" ) then
		myCriteria = myCriteria & maybeAND & "company_name Like N'%" & request.form("company_name_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_company_name.checked = false;" & vbCrLf
		response.write "document.all.company_name_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If (request.form("check_customer_name") = "on" AND request.form("customer_name_box") <> "")then
		myCriteria = myCriteria & maybeAND & "customer_name Like N'%" & request.form("customer_name_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_customer_name.checked = false;" & vbCrLf
		response.write "document.all.customer_name_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If request.form("check_kind") = "on" then
		myCriteria = myCriteria & maybeAND & "order_kind = N'" & request.form("order_kind_box") & "'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_kind.checked = false;" & vbCrLf
		response.write "document.all.order_kind_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If request.form("check_marhale") = "on" then
		If request.form("marhale_not_check") = "on" then
			myCriteria = myCriteria & maybeAND & "NOT(step = " & request.form("marhale_box") & ")"
		Else
			myCriteria = myCriteria & maybeAND & "step = " & request.form("marhale_box") 
			response.write "document.all.marhale_not_check.checked = false;" & vbCrLf
			response.write "document.all.marhale_not_check_label.style.color='#BBBBBB';"& vbCrLf
		End If
		maybeAND=" AND "
	Else
		response.write "document.all.check_marhale.checked = false;" & vbCrLf
		response.write "document.all.marhale_box.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.marhale_not_check.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.marhale_not_check_label.style.color='#BBBBBB';"& vbCrLf
	End If

	If request.form("check_vazyat") = "on" then
		If request.form("vazyat_not_check") = "on" then
			myCriteria = myCriteria & maybeAND & "NOT(status = " & request.form("vazyat_box") & ")"
		Else
			myCriteria = myCriteria & maybeAND & "status = " & request.form("vazyat_box") 
			response.write "document.all.vazyat_not_check.checked = false;" & vbCrLf
			response.write "document.all.vazyat_not_check_label.style.color='#BBBBBB';"& vbCrLf
		End If
		maybeAND=" AND "
	Else
		response.write "document.all.check_vazyat.checked = false;" & vbCrLf
		response.write "document.all.vazyat_box.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.vazyat_not_check.style.visibility = 'hidden';" & vbCrLf
		response.write "document.all.vazyat_not_check_label.style.color='#BBBBBB';"& vbCrLf
	End If

	If request.form("check_salesperson") = "on" then
		myCriteria = myCriteria & maybeAND & "salesperson = N'" & request.form("salesperson_box") & "'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_salesperson.checked = false;" & vbCrLf
		response.write "document.all.salesperson_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If (request.form("check_order_title") = "on" AND request.form("order_title_box") <> "")then
		myCriteria = myCriteria & maybeAND & "order_title Like N'%" & request.form("order_title_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_order_title.checked = false;" & vbCrLf
		response.write "document.all.order_title_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If (request.form("check_telephone") = "on" AND request.form("telephone_box") <> "")then
		myCriteria = myCriteria & maybeAND & "telephone Like N'%" & request.form("telephone_box") &"%'"
		maybeAND=" AND "
	Else
		response.write "document.all.check_telephone.checked = false;" & vbCrLf
		response.write "document.all.telephone_box.style.visibility = 'hidden';" & vbCrLf
	End If

	If request.form("check_closed") = "on" then
		myCriteria = myCriteria & maybeAND & "Orders.Closed=0"
	Else
		If request("Submit")="�����" then
			response.write "document.all.check_closed.checked = false;" & vbCrLf
			response.write "document.all.check_closed_label.style.color='#BBBBBB';" & vbCrLf
		End If
	End If

	%>
	}
	function checkValidation(){
		return true;
	}

	Form_Load();

	//-->
	</SCRIPT>
<%
	if request("Submit")="�����" then
		IF maybeAND <> " AND " THEN
			response.write "Nothing !!!!!!!!!!"
		ELSE
			mySQL="SELECT orders_trace.*, Orders.closed, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon,DRV_Invoice.price FROM Orders INNER JOIN  orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN  OrderTraceStatus ON orders_trace.status = OrderTraceStatus.ID left outer join (select InvoiceOrderRelations.[Order],SUM(InvoiceLines.Price + InvoiceLines.Vat - InvoiceLines.Discount -InvoiceLines.Reverse) as price from InvoiceOrderRelations inner join Invoices on InvoiceOrderRelations.Invoice=Invoices.ID inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice where Invoices.Voided=0 group by InvoiceOrderRelations.[Order]) DRV_Invoice on Orders.ID=DRV_Invoice.[Order] WHERE ("& myCriteria & ") ORDER BY order_date DESC, radif_sefareshat DESC"	
			set RS1=Conn.Execute (mySQL)
			if not RS1.eof then
				tmpCounter=0
'response.write mySQL
%>
			<div align="center" dir="LTR">
			<TABLE border="1" cellspacing="0" cellpadding="1" dir="RTL" borderColor="#555588">
				<TR bgcolor="#CCCCFF">
					<TD width="44"># �����</TD>
					<TD width="46">����� �����</TD>
					<TD width="64">����� �����</TD>
					<TD width="122">��� ����</TD>
					<TD width="122">��� �����</TD>
					<TD width="84">����� ���</TD>
					<TD width="40">���</TD>
					<TD width="53">�����</TD>
					<TD width="36">����� ������</TD>
					<TD width="18">���</TD>
					<td width="50">���� ��</td>
				</TR>
<%				Do while not RS1.eof and tmpCounter < resultsCount
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
				Else
					tmpColor="#DDDDDD"
				End If

				if RS1("Closed") then
					tmpStyle="background-color:#FFCCCC;"
				else
					tmpStyle=""
				End If
				
%>
				<TR bgcolor="<%=tmpColor%>" title="<%=RS1("StatusName")%>">
					<TD width="40" DIR="LTR"><A HREF="TraceOrder.asp?act=show&order=<%=RS1("radif_sefareshat")%>" target="_blank"><%=RS1("radif_sefareshat")%></A></TD>
					<TD DIR="LTR"><%=RS1("order_date")%></TD>
					<TD DIR="LTR"><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
					<TD><%=RS1("company_name") & "<br>����:("& RS1("telephone")& ")"%>&nbsp;</TD>
					<TD><%=RS1("customer_name")%>&nbsp;</TD>
					<TD><%=RS1("order_title")%>&nbsp;</TD>
					<TD><%=RS1("order_kind")%></TD>
					<TD style="<%=tmpStyle%>"><%=RS1("marhale")%></TD>
					<TD><%=RS1("salesperson")%>&nbsp;</TD>
					<TD><IMG SRC="<%=RS1("Icon")%>" WIDTH="20" HEIGHT="20" BORDER="0"></TD>
					<td><%if isnull(RS1("price")) then response.write "----" else response.write Separate(RS1("price")) end if %></td>
				</TR>
				<TR bgcolor="#FFFFFF">
					<TD colspan="10" style="height:10px"></TD>
				</TR>
<%					RS1.moveNext
				Loop

				if tmpCounter >= resultsCount then
%>					<TR bgcolor="#CCCCCC">
						<TD align="center" colspan="10" style="padding:15px;font-size:12pt;color:red;cursor:hand;" onclick="document.all.resultsCount.focus();"><B>����� ����� ����� ��� �� <%=resultsCount%> ����� ���.</B></TD>
					</TR>
<%				else
%>					<TR bgcolor="#ccccFF">
						<TD colspan="10">����� ����� �����: <%=tmpCounter%></TD>
					</TR>
<%				end if
%>
			</TABLE>
			</div>
			<BR>
<%			else
%>			<TABLE border="1" cellspacing="0" cellpadding="0" dir="RTL" align="center" width="600">
				<TR bgcolor="#FFFFDD">
					<TD align="center" style="height:40px;font-size:12pt;font-weight:bold;color:red">�� ����� ������ �� ��� �����.</TD>
				</TR>
			</TABLE>
			<hr>
<%			End If
		End If
	End If
End If

Conn.Close
%>
</font>
<!--#include file="tah.asp" -->
