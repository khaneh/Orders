<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Home (0)
PageTitle= "��� ���"
SubmenuItem=6
if not Auth("D" , 0) then NotAllowdToViewThisPage()

briefQtty = 10

userID = session("ID")
if Auth("E" , 5) then	'���� ��� ��� ���� �������
	if request("userID") <> "" then 
		userID = cint(request("userID"))
	end if
end if
showAll = request("showAll")

panel = request("panel")
if panel="" then 
	panel=0
else
	panel=cint(panel)
end if

activeTabColor="#336699"
disableTabColor="#CCCCCC"
%>
<!--#include file="top.asp" -->
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: <%=activeTabColor%>; text-align: center; }
	.CusTD1 a {color:#FFFF00; font-size:9pt;}
	.CusTD2 {background-color: <%=disableTabColor%>; text-align: center; }
	.CusTD2 a {color:#888888; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; background-color:#C3DBEB; border:4 solid <%=activeTabColor%>;}
	.RepTextArea {font-family:tahoma;font-size:8pt;width:300px;height:30px;border:none;backGround:Transparent;}
	.RepSelect {font-family:tahoma;font-size:9pt;width:120px;}
	.RepROInputs {font-family:tahoma;font-size:8pt;width:70px;border:none;backGround:Transparent;}

</STYLE>

<table cellspacing=0 cellpadding=0 width="100%" style="border:4 solid <%=AppFgColor%>;" >
<tr><td>
<TABLE cellspacing=0 cellpadding=0 width="100%">
<TR height='15'>
	<TD></TD>
</TR>
<%
	if Auth("E" , 5) then	'���� ��� ��� ���� �������
%> 
	<TR height='25'>
		<TD colspan=10 valign=top>
		<FORM METHOD=POST ACTION="?panel=<%=panel%>&showAll=<%=showAll%>">
		������ ��� ��� �����:

		<select name="userID" class="RepSelect" onchange="submit();">
		<% set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
		Do while not RSV.eof
		%>
			<option value="<%=RSV("ID")%>" <%
				if RSV("ID")=userID then
					response.write " selected "
				end if
				%>><%=RSV("RealName")%></option>
		<%
		RSV.moveNext
		Loop
		RSV.close
		%>
		</select> 
		</FORM>
		</TD>
	</TR>
<%
	end if %>
<TR class='alak' height='25'>

	<TD width=15 >&nbsp;</TD>

<%
	if Auth("D" , 1) then

		mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Orders WHERE (Closed = 0) AND (ID NOT IN (SELECT [ORDER] FROM invoiceorderrelations)) AND CreatedBy = '" & userID & "'"

		Set RS1 = conn.execute(mySQL)
		count = RS1("CNT")
		RS1.close

		if panel=1 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?userID=<%=userID%>&panel=1'>����� ���� ������� (<%=count%>)</A></TD>
<%
	end if %>
	<TD width=5 >&nbsp;</TD>

<%
	if Auth("D" , 2) then 

		mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices WHERE (Voided = 0) AND (Approved = 0) AND (CreatedBy = '" & userID & "')"

		Set RS1 = conn.execute(mySQL)
		count = RS1("CNT")
		RS1.close

		if panel=2 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?userID=<%=userID%>&panel=2'>����� ���� (<%=count%>)</A></TD>
<%
	end if %>
	<TD width=5 >&nbsp;</TD>

<%
	if Auth("D" , 3) then 

		mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices WHERE (Voided = 0) AND (Approved = 1) AND (Issued = 0) AND (ApprovedBy = '" & userID & "')"

		Set RS1 = conn.execute(mySQL)
		count = RS1("CNT")
		RS1.close

		if panel=3 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?userID=<%=userID%>&panel=3'>���� ���� (<%=count%>)</A></TD>
<%
	end if %>
	<TD width=5 >&nbsp;</TD>

<%
	if Auth("D" , 4) then 

		mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (Invoices.IssuedBy = '"& userID & "') AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = 0)"

		Set RS1 = conn.execute(mySQL)
		count = RS1("CNT")
		RS1.close

		if panel=4 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?userID=<%=userID%>&panel=4'>����� ���� (<%=count%>)</A></TD>
<%
	end if %>
	<TD width=5 >&nbsp;</TD>

<%
	if Auth("D" , 5) then 

		mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Accounts WHERE (CSR = '"& userID & "') AND (NOT (ARBalance = 0))"

		Set RS1 = conn.execute(mySQL)
		count = RS1("CNT")
		RS1.close

		if panel=5 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?userID=<%=userID%>&panel=5'>������� (<%=count%>)</A></TD>
<%
	end if %>
	<TD width=5 >&nbsp;</TD>

<%
	if Auth("D" , 6) then 
		mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM EffectiveGLRows INNER JOIN Accounts ON EffectiveGLRows.Tafsil = Accounts.ID WHERE (EffectiveGLRows.ID IN (SELECT MAX(ID) AS MaxID FROM EffectiveGLRows GROUP BY GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (GLAccount = 17011) AND (GL = '" & openGL & "') AND (Ref1 <> N'') AND (COUNT(Ref1) % 2 = 1))) AND (Accounts.CSR = "& userID & ")"

		Set RS1 = conn.execute(mySQL)
		count = RS1("CNT")
		RS1.close

		if panel=6 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?userID=<%=userID%>&panel=6'>�� �ѐ��� (<%=count%>)</A></TD>
<%
	end if %>
	<TD width=5 >&nbsp;</TD>

<%
	if Auth("D" , 7) then 
		if panel=7 then styleClass="CusTD1" else styleClass="CusTD2" 
%> 
		<TD align=center class='<%=styleClass%>'><A HREF='?userID=<%=userID%>&panel=7'>����� ��� ����� ����</A></TD>
<%
	end if %>
	<TD width=5 >&nbsp;</TD>

	<TD width=50 >&nbsp;</TD>
	<TD width=* align=left>&nbsp;</TD>
</TR>
</TABLE>

<TaBlE class="CustTable4" cellspacing="2" cellspacing="2" height=350>
	<Tr>
		<Td height=20>
		</Td>
	</Tr>
	<Tr>
		<Td valign="top" align="center">
		</Td>
	</Tr>
<%
	if Auth("D" , 1) AND panel=1 then '����� ��� ���
%>
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="9" class="CusTableHeader">����� ���� �� ���� �ǘ��� ������</td>
		</tr>
		<%
		if showAll="on" and panel = 1 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Accounts.AccountTitle, orders_trace.radif_sefareshat, orders_trace.order_date, orders_trace.order_kind, orders_trace.order_title, orders_trace.salesperson, orders_trace.vazyat, orders_trace.marhale FROM Orders INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN Accounts ON Orders.Customer = Accounts.ID WHERE (Orders.Closed = 0) AND (Orders.ID NOT IN (SELECT [ORDER] FROM invoiceorderrelations)) AND (Orders.CreatedBy='" & userID & "') ORDER BY Orders.CreatedDate DESC, Orders.ID"
		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">���</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../order/TraceOrder.asp?act=show&order=<%=RS1("radif_sefareshat")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=RS1("radif_sefareshat")%></TD>
					<TD><%=RS1("salesperson")%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("order_date")%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD><%=RS1("order_kind")%>&nbsp;</TD>
					<TD><%=RS1("order_title")%>&nbsp;</TD>
					<TD><%=RS1("marhale")%>&nbsp;</TD>
					<TD><%=RS1("vazyat")%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
%>
			<tr>
				<td colspan="9" class="CusTableHeader" style="text-align:right;"><A HREF="?userID=<%=userID%>&showAll=on&panel=1">����� ���� ...</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="9" class="CusTD3">��</td>
			</tr>
<%
		end if
%>
		</table>
	</Td>
	</Tr>
<%
	end if

	if Auth("D" , 2) AND panel=2 then '������ ��� ����� ����
%>
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="8" class="CusTableHeader" class="CusTableHeader" style="background-color:#CCCC99;">������ ��� �� ����� (����� ����)</td>
		</tr>
		<%
		if showAll="on" and panel = 2 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.TotalReceivable, InvoiceOrderRelations.[Order], orders_trace.vazyat, orders_trace.marhale, Invoices.Customer, Accounts.AccountTitle FROM Accounts INNER JOIN Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID ON Accounts.ID = Invoices.Customer LEFT OUTER JOIN orders_trace RIGHT OUTER JOIN InvoiceOrderRelations ON orders_trace.radif_sefareshat = InvoiceOrderRelations.[Order] ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 0) AND (Invoices.CreatedBy = '" & userID & "') ORDER BY Invoices.CreatedDate DESC, Invoices.ID"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>					<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3"># ������</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3"># �����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand; height: 30px;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
					<TD><%=tmpCounter%></TD>
					<TD><%=RS1("ID")%></TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Creator")%>"><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD align='right'><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD align='center' style="text-decoration:underline;" onclick="window.open('../order/TraceOrder.asp?act=show&order=<%=RS1("Order")%>');event.cancelBubble=true;"><%=RS1("Order")%>&nbsp;</TD>
					<TD align='right'><%=RS1("vazyat")%>&nbsp;</TD>
					<TD align='right'><%=RS1("Marhale")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
%>
			<tr>
				<td colspan="8" class="CusTableHeader" style="text-align:right;background-color:#CCCC99;"><A HREF="?userID=<%=userID%>&showAll=on&panel=2">����� ���� ...</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="8" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
<%
	end if

	if Auth("D" , 3) AND panel=3 then '������ ��� ����� ��� (���� ����)
%>
	<Tr>
	<Td colspan="2" valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="7" class="CusTableHeader" style="background-color:#33BB99;">������ ��� ����� ��� (���� ����)</td>
		</tr>
		<%
		if showAll="on" and panel = 3 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.ApprovedDate, Invoices.TotalReceivable, Users_1.RealName AS Approver, Invoices.Customer, Accounts.AccountTitle FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 1) AND (Invoices.Issued = 0) AND (Invoices.ApprovedBy = '" & userID & "') ORDER BY Invoices.CreatedDate DESC, Invoices.ID"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3"># ������</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=RS1("ID")%></TD>
					<TD><%=RS1("�Creator")%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Approver")%>"><%=RS1("ApprovedDate")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
%>
			<tr>
				<td colspan="7" class="CusTableHeader" style="text-align:right;background-color:#33BB99;"><A HREF="?userID=<%=userID%>&showAll=on&panel=3">����� ���� ...</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="7" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
<%
	end if

	if Auth("D" , 4) AND panel=4 then '������ ��� ���� ��� (����� ����)
%>
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="10" class="CusTableHeader" style="background-color:#CCAA99;">������ ��� ���� ��� (����� ����)</td>
		</tr>
		<%
		if showAll="on" and panel = 4 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Invoices.ApprovedDate, Invoices.TotalReceivable, Invoices.IssuedDate, Invoices.Customer, ARItems.RemainedAmount, Users.RealName AS Creator, Users_1.RealName AS Approver, Users_2.RealName AS Issuer, Accounts.AccountTitle FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Users Users_2 ON Invoices.IssuedBy = Users_2.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.IssuedBy = '"& userID & "') AND (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID DESC"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����� ����</td>
				<td class="CusTD3">����</td>
				<td class="CusTD3">�����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=RS1("ID")%></TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("�Creator")%>"><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Approver")%>"><%=RS1("ApprovedDate")%>&nbsp;</TD>
					<TD align='right'><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Issuer")%>"><%=RS1("IssuedDate")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
					<TD><%=Separate(RS1("RemainedAmount"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
%>
			<tr>
				<td colspan="10" class="CusTableHeader" style="text-align:right;background-color:#CCAA99;"><A HREF="?userID=<%=userID%>&showAll=on&panel=4">����� ���� ...</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="10" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
<%
	end if

	if Auth("D" , 5) AND panel=5 then '���� ��� ����� ���
%>
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="10" class="CusTableHeader" style="background-color:#CC99AA;">���� ��� ����� ���</td>
		</tr>
		<%
		if showAll="on" and panel = 5 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " * FROM Accounts WHERE (CSR = '"& userID & "') AND (NOT ARBalance=0) ORDER BY ARBalance"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3"> # </td>
				<td class="CusTD3"> ����� ���� </td>
				<td class="CusTD3"> ��� ���� ��� </td>
				<td class="CusTD3" width='80'> ����� ����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1

				AccountNo=RS1("ID")
				AccountTitle=RS1("AccountTitle")
				ARBalance=cdbl(RS1("ARBalance"))
				CreditLimit=RS1("CreditLimit")
				contact1 = RS1("Dear1") & " " & RS1("FirstName1") & " " & RS1("LastName1")
				if RS1("Type") = 1 then 
					AccountTitle = AccountTitle & " (�������) "
				end if

				totalBalance=Separate(ARBalance)
				if (ARBalance >= 0 )then
					tempBalanceColor="green"
				else
					tempBalanceColor="red"
				end if

				total = total + ARBalance

				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=AccountNo%>');">
					<TD><%=tmpCounter%>&nbsp;</TD>
					<TD><%=AccountTitle%>&nbsp;</TD>
					<TD><%=contact1%>&nbsp;</TD>
					<TD style="direction:LTR; text-align:right;"><FONT COLOR="<%=tempBalanceColor%>"><%=totalBalance%></FONT>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
%>
			<tr>
				<td colspan="10" class="CusTableHeader" style="text-align:right;background-color:#CC99AA;"><A HREF="?userID=<%=userID%>&showAll=on&panel=5">����� ���� ...</A></td>
			</tr>
<%
			else
%>
			<tr bgcolor='#CC99AA'>
				<td align='right' colspan=3> ���</td>
				<td dir=ltr align='right' > <%=Separate(total)%>&nbsp;</td>

				</td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="10" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
<%
	end if

	if Auth("D" , 6) AND panel=6 then '�� ��� �ѐ���
%>
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="10" class="CusTableHeader" style="background-color:#FF9999;">�� ��� �ѐ���</td>
		</tr>
		<%
		if showAll="on" and panel = 6 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " EffectiveGLRows.GLAccount, EffectiveGLRows.Tafsil, EffectiveGLRows.Amount, EffectiveGLRows.Description, EffectiveGLRows.Ref1, EffectiveGLRows.Ref2, EffectiveGLRows.IsCredit, EffectiveGLRows.GLDocDate, Accounts.CSR, Accounts.AccountTitle FROM EffectiveGLRows INNER JOIN Accounts ON EffectiveGLRows.Tafsil = Accounts.ID WHERE (EffectiveGLRows.ID IN (SELECT MAX(ID) AS MaxID FROM EffectiveGLRows GROUP BY GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (GLAccount = 17011) AND (GL = " & openGL & ") AND (Ref1 <> N'') AND (COUNT(Ref1) % 2 = 1))) AND (Accounts.CSR = "& userID & ")"



		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr class="RepTableHeader">
				<td class="CusTD3">#</td>
				<td class="CusTD3">����� ���</td>
				<td class="CusTD3">���</td>
				<td class="CusTD3">������</td>
				<td class="CusTD3">��������</td>
				<td class="CusTD3">����� ��</td>
			</tr>
<%			Do while not RS1.eof
				Ref1=		RS1("Ref1")
				Ref2=		RS1("Ref2")
				IsCredit=	RS1("IsCredit")
				Description=RS1("Description")
				GLDocDate=	RS1("GLDocDate")
				Amount=		RS1("Amount")
				AccountNo= RS1("Tafsil")

				if IsCredit then
					totalCredit = totalCredit + cdbl(Amount)
				else
					totalDebit = totalDebit + cdbl(Amount)
				end if

				if NOT (Ref1="" AND Ref2="") then tempCounter=tempCounter+1

				if tempCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 

%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=AccountNo%>');">

					<td><%if tempCounter>0 then response.write tempCounter%>
						<INPUT TYPE="hidden" Name="GLAccounts" Value="0">
						<INPUT TYPE="hidden" Name="Accounts" Value="0"></td>
					<td dir=LTR>
						<INPUT readonly Name="GLDocDates" class="RepROInputs" Value="<%=GLDocDate%>"></td>
					<td>
						<TextArea readonly Name="Descriptions" class="RepTextArea" style="cursor: hand;" ><%=replace(Description,"/",".")%></TextArea></td>
<%				if IsCredit then%>
					<td>
						<INPUT TYPE="hidden" Name="IsCredit" Value="<%=IsCredit%>">
					</td>
					<td>
						<INPUT readonly Name="Amounts" class="RepROInputs" Value="<%=Separate(Amount)%>">
					</td>
<%				else%>			
					<td>
						<INPUT readonly Name="Amounts" class="RepROInputs" Value="<%=Separate(Amount)%>">
					</td>
					<td>
						<INPUT TYPE="hidden" Name="IsCredit" Value="<%=IsCredit%>">
					</td>
<%				end if%>
					<td dir=LTR>
						<INPUT readonly Name="CheqDates" class="RepROInputs" Value="<%=Ref2%>">
						<INPUT type="hidden" readonly Name="CheqNos" class="RepROInputs" Value="<%=Ref1%>"></td>
				</tr>
<%				RS1.MoveNext
			Loop
			RS1.Close

			if selectTop<>"" and tempCounter = briefQtty then
%>
			<tr>
				<td colspan="10" class="CusTableHeader" style="text-align:right;background-color:#FF9999;"><A HREF="?userID=<%=userID%>&showAll=on&panel=6">����� ���� ...</A></td>
			</tr>
<%
			else
%>
			<tr bgcolor='#FF9999'>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td><span style="width:300px;text-align:left;overflow:auto;text-overflow:ellipsis;"><B>��� :</B></span></td>
				<td><INPUT readonly class="RepROInputs" Value="<%=Separate(totalDebit)%>"></td>
				<td><INPUT readonly class="RepROInputs" Value="<%=Separate(totalCredit)%>"></td>
				<td>&nbsp;</td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="10" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
<%
	end if

	if Auth("D" , 7) AND panel=7 then '����� ��� ����� ����

	briefQtty = 10
	showAll = request("showAll")
	spanel = request("spanel")
	if spanel="" then 
		spanel=0
	else
		spanel=cint(spanel)
	end if

%>
	<Tr>
	<Td valign="top" align="center">
<!-- -->
	<TaBlE width=100% cellspacing="0" cellspacing="0">
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="9" class="CusTableHeader">����� ���� �� ���� �ǘ��� ������</td>
		</tr>
		<%
		if showAll="on" and spanel = 1 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Accounts.AccountTitle, orders_trace.radif_sefareshat, orders_trace.order_date, orders_trace.order_kind, orders_trace.order_title, orders_trace.salesperson, orders_trace.vazyat, orders_trace.marhale FROM Orders INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN Accounts ON Orders.Customer = Accounts.ID WHERE (Orders.Closed = 0) AND (Orders.ID NOT IN (SELECT [ORDER] FROM invoiceorderrelations)) AND (Accounts.CSR = "& userID & ") ORDER BY Orders.CreatedDate DESC, Orders.ID"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">���</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../order/TraceOrder.asp?act=show&order=<%=RS1("radif_sefareshat")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=RS1("radif_sefareshat")%></TD>
					<TD><%=RS1("salesperson")%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("order_date")%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD><%=RS1("order_kind")%>&nbsp;</TD>
					<TD><%=RS1("order_title")%>&nbsp;</TD>
					<TD><%=RS1("marhale")%>&nbsp;</TD>
					<TD><%=RS1("vazyat")%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop

			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Orders INNER JOIN Accounts ON Orders.Customer = Accounts.ID WHERE (Orders.Closed = 0) AND (Orders.ID NOT IN (SELECT [ORDER] FROM invoiceorderrelations)) AND (Accounts.CSR = "& userID & ")"
				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close
%>
			<tr>
				<td colspan="9" class="CusTableHeader" style="text-align:right;"><A HREF="?userID=<%=userID%>&panel=7&showAll=on&spanel=1">����� ���� ... (����� �� : <%=count%>)</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="9" class="CusTD3">��</td>
			</tr>
<%
		end if
%>
		</table>
	</Td>
	</Tr>
	<Tr>
		<Td colspan="2" valign="top" align="center">
			<hr>
		</Td>
	</Tr>
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="8" class="CusTableHeader" class="CusTableHeader" style="background-color:#CCCC99;">������ ��� �� ����� (����� ����)</td>
		</tr>
		<%
		if showAll="on" and spanel = 2 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & "Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.TotalReceivable, InvoiceOrderRelations.[Order], orders_trace.vazyat, orders_trace.marhale, Invoices.Customer, Accounts.AccountTitle FROM Accounts INNER JOIN Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID ON Accounts.ID = Invoices.Customer LEFT OUTER JOIN orders_trace RIGHT OUTER JOIN InvoiceOrderRelations ON orders_trace.radif_sefareshat = InvoiceOrderRelations.[Order] ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 0) AND (Accounts.CSR = "& userID & ") ORDER BY Invoices.CreatedDate DESC, Invoices.ID"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>					<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3"># ������</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3"># �����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand; height: 30px;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
					<TD><%=tmpCounter%></TD>
					<TD><%=RS1("ID")%></TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Creator")%>"><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD align='right'><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD align='center' style="text-decoration:underline;" onclick="window.open('../order/TraceOrder.asp?act=show&order=<%=RS1("Order")%>');event.cancelBubble=true;"><%=RS1("Order")%>&nbsp;</TD>
					<TD align='right'><%=RS1("vazyat")%>&nbsp;</TD>
					<TD align='right'><%=RS1("Marhale")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
'				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices WHERE (Voided = 0) AND (Approved = 0)"
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Accounts INNER JOIN Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID ON Accounts.ID = Invoices.Customer LEFT OUTER JOIN orders_trace RIGHT OUTER JOIN InvoiceOrderRelations ON orders_trace.radif_sefareshat = InvoiceOrderRelations.[Order] ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 0) AND (Accounts.CSR = "& userID & ")"

				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close
%>
			<tr>
				<td colspan="8" class="CusTableHeader" style="text-align:right;background-color:#CCCC99;"><A HREF="?userID=<%=userID%>&panel=7&showAll=on&spanel=2">����� ���� ... (����� �� : <%=count%> ������ x �����)</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="8" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
	<Tr>
		<Td colspan="2" valign="top" align="center">
			<hr>
		</Td>
	</Tr>
	<Tr>
	<Td colspan="2" valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="7" class="CusTableHeader" style="background-color:#33BB99;">������ ��� ����� ��� (���� ����)</td>
		</tr>
		<%
		if showAll="on" and spanel = 3 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.ApprovedDate, Invoices.TotalReceivable, Users_1.RealName AS Approver, Invoices.Customer, Accounts.AccountTitle FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 1) AND (Invoices.Issued = 0) AND (Accounts.CSR = "& userID & ") ORDER BY Invoices.CreatedDate DESC, Invoices.ID"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3"># ������</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=RS1("ID")%></TD>
					<TD><%=RS1("�Creator")%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Approver")%>"><%=RS1("ApprovedDate")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 1) AND (Invoices.Issued = 0) AND (Accounts.CSR = "& userID & ")"
				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close
%>
			<tr>
				<td colspan="7" class="CusTableHeader" style="text-align:right;background-color:#33BB99;"><A HREF="?userID=<%=userID%>&panel=7&showAll=on&spanel=3">����� ���� ... (����� �� : <%=count%>)</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="7" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
	<Tr>
		<Td colspan="2" valign="top" align="center">
			<hr>
		</Td>
	</Tr>
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="10" class="CusTableHeader" style="background-color:#CCAA99;">������ ��� ���� ��� (����� ����)</td>
		</tr>
		<%
		if showAll="on" and spanel = 4 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

'		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Invoices.ApprovedDate, Invoices.TotalReceivable, Invoices.IssuedDate, Invoices.Customer, ARItems.RemainedAmount, Users.RealName AS Creator, Users_1.RealName AS Approver, Users_2.RealName AS Issuer, Accounts.AccountTitle FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Users Users_2 ON Invoices.IssuedBy = Users_2.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID DESC"
		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Invoices.ApprovedDate, Invoices.TotalReceivable, Invoices.IssuedDate, Invoices.Customer, ARItems.RemainedAmount, Users.RealName AS Creator, Users_1.RealName AS Approver, Users_2.RealName AS Issuer, Accounts.AccountTitle FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Users Users_2 ON Invoices.IssuedBy = Users_2.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = "& userID & ") ORDER BY Invoices.CreatedDate DESC, Invoices.ID DESC"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">����� �����</td>
				<td class="CusTD3">�����</td>
				<td class="CusTD3">����� ����</td>
				<td class="CusTD3">����</td>
				<td class="CusTD3">�����</td>
			</tr>
<%
			tmpCounter=0
			Do while not RS1.eof 
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 
%>
				<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
					<TD style="height:30px;"><%=tmpCounter%></TD>
					<TD style="height:30px;"><%=RS1("ID")%></TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("�Creator")%>"><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Approver")%>"><%=RS1("ApprovedDate")%>&nbsp;</TD>
					<TD align='right'><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title="���� <%=RS1("Issuer")%>"><%=RS1("IssuedDate")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
					<TD><%=Separate(RS1("RemainedAmount"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = "& userID & ")"

				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close

%>
			<tr>
				<td colspan="10" class="CusTableHeader" style="text-align:right;background-color:#CCAA99;"><A HREF="?userID=<%=userID%>&panel=7&showAll=on&spanel=4">����� ���� ... (����� �� : <%=count%>)</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="10" class="CusTD3">��</td>
			</tr>
<%
		end if
		%>
		</table>
	</Td>
	</Tr>
	<Tr>
		<Td colspan="2" valign="top" align="center">
			<hr>
		</Td>
	</Tr>

	</Table>

<!-- -->
	</Td>
	</Tr>
<%
	end if
%>
	<Tr height=15>
		<Td></Td>
	</Tr>
</TaBlE>
</td>
</tr>
</table>

<!--#include file="tah.asp" -->
