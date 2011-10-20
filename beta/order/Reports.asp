<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ ò·Ì"
SubmenuItem=5

briefQtty = 10
showAll = request("showAll")
panel = request("panel")
if panel="" then 
	panel=0
else
	panel=cint(panel)
end if

if not Auth(2 , 6) then NotAllowdToViewThisPage()
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
	.CusTD1 {background-color: <%=activeTabColor%>; text-align: center; font-size:9pt;}
	.CusTD1 a {color:#FFFF00;}
	.CusTD2 {background-color: <%=disableTabColor%>; text-align: center; font-size:9pt;}
	.CusTD2 a {color:#888888;}
	.CusTD3 {background-color: #DDDDDD; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; background-color:#C3DBEB; border:4 solid <%=activeTabColor%>;}
	.RepTextArea {font-family:tahoma;font-size:8pt;width:300px;height:30px;border:none;backGround:Transparent;}
	.RepSelect {font-family:tahoma;font-size:9pt;width:120px;}
	.RepROInputs {font-family:tahoma;font-size:8pt;width:70px;border:none;backGround:Transparent;}

</STYLE>
	<TaBlE class="CustTable4" cellspacing="2" cellspacing="2">
	<Tr>
	<Td valign="top" align="center">
		<table class="CustTable" cellspacing='1' style='width:90%;'>
		<tr>
			<td colspan="9" class="CusTableHeader">”›«—‘ Â«ÌÌ òÂ «’·« ›«ò Ê— ‰œ«—‰œ</td>
		</tr>
		<%
		if showAll="on" and panel = 1 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Accounts.AccountTitle, orders_trace.radif_sefareshat, orders_trace.order_date, orders_trace.order_kind, orders_trace.order_title, orders_trace.salesperson, orders_trace.vazyat, orders_trace.marhale FROM Orders INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN Accounts ON Orders.Customer = Accounts.ID WHERE (Orders.Closed = 0) AND (Orders.ID NOT IN (SELECT [ORDER] FROM invoiceorderrelations)) ORDER BY Orders.CreatedDate DESC, Orders.ID"

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">‘„«—Â</td>
				<td class="CusTD3">„”ÊÊ·</td>
				<td class="CusTD3"> «—ÌŒ</td>
				<td class="CusTD3">„‘ —Ì</td>
				<td class="CusTD3">‰Ê⁄</td>
				<td class="CusTD3">⁄‰Ê«‰</td>
				<td class="CusTD3">„—Õ·Â</td>
				<td class="CusTD3">Ê÷⁄Ì </td>
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
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Orders WHERE (Closed = 0) AND (ID NOT IN (SELECT [ORDER] FROM invoiceorderrelations))"
				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close
%>
			<tr>
				<td colspan="9" class="CusTableHeader" style="text-align:right;"><A HREF="?userID=<%=userID%>&showAll=on&panel=1">«œ«„Â œ«—œ ... ( ⁄œ«œ ﬂ· : <%=count%>)</A></td>
			</tr>
<%
			end if
		else
%>
			<tr>
				<td colspan="9" class="CusTD3">ÂÌç</td>
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
			<td colspan="8" class="CusTableHeader" class="CusTableHeader" style="background-color:#CCCC99;">›«ﬂ Ê— Â«Ì œ— Ã—Ì«‰ ( «ÌÌœ ‰‘œÂ)</td>
		</tr>
		<%
		if showAll="on" and panel = 2 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & "Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.TotalReceivable, InvoiceOrderRelations.[Order], orders_trace.vazyat, orders_trace.marhale, Invoices.Customer, Accounts.AccountTitle FROM Accounts INNER JOIN Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID ON Accounts.ID = Invoices.Customer LEFT OUTER JOIN orders_trace RIGHT OUTER JOIN InvoiceOrderRelations ON orders_trace.radif_sefareshat = InvoiceOrderRelations.[Order] ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID"

		'-------------------------------------------------------------SAM--------------------------------------------------------------------
		mySQLsum="SELECT SUM(sumTotalReceivable) AS sumTotalReceivable from (SELECT MAX(Invoices.TotalReceivable) as sumTotalReceivable FROM Accounts INNER JOIN Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID ON Accounts.ID = Invoices.Customer LEFT OUTER JOIN orders_trace RIGHT OUTER JOIN InvoiceOrderRelations ON orders_trace.radif_sefareshat = InvoiceOrderRelations.[Order] ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 0) GROUP BY Invoices.ID) as drv"
		set RS2 = conn.execute(mySQLsum)
		
		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>					<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3"># ›«ﬂ Ê—</td>
				<td class="CusTD3"> «—ÌŒ</td>
				<td class="CusTD3">„‘ —Ì</td>
				<td class="CusTD3"># ”›«—‘</td>
				<td class="CusTD3">Ê÷⁄Ì </td>
				<td class="CusTD3">„—Õ·Â</td>
				<td class="CusTD3">„»·€</td>
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
					<TD dir="LTR" align='right' title=" Ê”ÿ <%=RS1("Creator")%>"><%=RS1("CreatedDate")%>&nbsp;</TD>
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
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Accounts INNER JOIN Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID ON Accounts.ID = Invoices.Customer LEFT OUTER JOIN orders_trace RIGHT OUTER JOIN InvoiceOrderRelations ON orders_trace.radif_sefareshat = InvoiceOrderRelations.[Order] ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 0)"

				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close
%>
			<tr>
				<td colspan="8" class="CusTableHeader" style="text-align:right;background-color:#CCCC99;"><A HREF="?userID=<%=userID%>&showAll=on&panel=2">«œ«„Â œ«—œ ... ( ⁄œ«œ ﬂ· : <%=count%> ›«ﬂ Ê— x ”›«—‘) òÂ „ÃÊ⁄ ›«ò Ê—Â«Ì ¬‰ <%=Separate(RS2("sumTotalReceivable"))%> —Ì«· „Ì »«‘œ</A></td>
			</tr>
<%					'-------------------------------------------------------------SAM--------------------------------------------------------------------
			RS2.close
			end if
		else
%>
			<tr>
				<td colspan="8" class="CusTD3">ÂÌç</td>
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
			<td colspan="7" class="CusTableHeader" style="background-color:#33BB99;">›«ﬂ Ê— Â«Ì  «ÌÌœ ‘œÂ (’«œ— ‰‘œÂ)</td>
		</tr>
		<%
		if showAll="on" and panel = 3 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.ApprovedDate, Invoices.TotalReceivable, Users_1.RealName AS Approver, Invoices.Customer, Accounts.AccountTitle FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 1) AND (Invoices.Issued = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID"
'-------------------------------------------------------------------------------------SAM----------------------------------------------
		mySQLsum="SELECT SUM(sumTotalReceivable) AS sumTotalReceivable FROM (SELECT MAX(Invoices.TotalReceivable) AS sumTotalReceivable FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 1) AND (Invoices.Issued = 0) GROUP BY Invoices.ID) as drv"
		set RS2 = conn.execute(mySQLsum)

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3"># ›«ﬂ Ê—</td>
				<td class="CusTD3">«ÌÃ«œ ﬂ‰‰œÂ</td>
				<td class="CusTD3"> «—ÌŒ «ÌÃ«œ</td>
				<td class="CusTD3">„‘ —Ì</td>
				<td class="CusTD3"> «—ÌŒ  «ÌÌœ</td>
				<td class="CusTD3">„»·€</td>
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
					<TD><%=RS1("ûCreator")%>&nbsp;</TD>
					<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title=" Ê”ÿ <%=RS1("Approver")%>"><%=RS1("ApprovedDate")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices WHERE (Voided = 0) AND (Approved = 1) AND (Issued = 0)"

				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close
%>
			<tr>
				<td colspan="7" class="CusTableHeader" style="text-align:right;background-color:#33BB99;"><A HREF="?userID=<%=userID%>&showAll=on&panel=3">«œ«„Â œ«—œ ... ( ⁄œ«œ ﬂ· : <%=count%>) òÂ „Ã„Ê⁄ ›«ò Ê—Â«Ì ¬‰ <%=Separate(RS2("sumTotalReceivable"))%> —Ì«· „Ì »«‘œ</A></td>
			</tr>
<%
			RS2.close
			end if
		else
%>
			<tr>
				<td colspan="7" class="CusTD3">ÂÌç</td>
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
			<td colspan="10" class="CusTableHeader" style="background-color:#CCAA99;">›«ﬂ Ê— Â«Ì ’«œ— ‘œÂ ( ”ÊÌÂ ‰‘œÂ)</td>
		</tr>
		<%
		if showAll="on" and panel = 4 then
			selectTop=""
		else
			selectTop="TOP " & briefQtty
		end if

		mySQL="SELECT " & selectTop & " Invoices.ID, Invoices.CreatedDate, Invoices.ApprovedDate, Invoices.TotalReceivable, Invoices.IssuedDate, Invoices.Customer, ARItems.RemainedAmount, Users.RealName AS Creator, Users_1.RealName AS Approver, Users_2.RealName AS Issuer, Accounts.AccountTitle FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Users Users_2 ON Invoices.IssuedBy = Users_2.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID DESC"

		mySQLsum="SELECT SUM(sumTotalReceivable) AS sumTotalReceivable, SUM(sumRemainedAmount) AS sumRemainedAmount FROM (SELECT MAX(Invoices.TotalReceivable) AS sumTotalReceivable, MAX(ARItems.RemainedAmount) AS sumRemainedAmount FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Users Users_2 ON Invoices.IssuedBy = Users_2.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = 0) GROUP BY Invoices.ID) as drv"

		set RS2 = conn.execute(mySQLsum)

		Set RS1 = conn.execute(mySQL)
		if not RS1.eof then
%>
			<tr>
				<td class="CusTD3">#</td>
				<td class="CusTD3">‘„«—Â</td>
				<td class="CusTD3"> «—ÌŒ «ÌÃ«œ</td>
				<td class="CusTD3"> «—ÌŒ  «ÌÌœ</td>
				<td class="CusTD3">„‘ —Ì</td>
				<td class="CusTD3"> «—ÌŒ ’œÊ—</td>
				<td class="CusTD3">„»·€</td>
				<td class="CusTD3">„«‰œÂ</td>
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
					<TD dir="LTR" align='right' title=" Ê”ÿ <%=RS1("ûCreator")%>"><%=RS1("CreatedDate")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title=" Ê”ÿ <%=RS1("Approver")%>"><%=RS1("ApprovedDate")%>&nbsp;</TD>
					<TD align='right'><%=RS1("AccountTitle")%>&nbsp;</TD>
					<TD dir="LTR" align='right' title=" Ê”ÿ <%=RS1("Issuer")%>"><%=RS1("IssuedDate")%>&nbsp;</TD>
					<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
					<TD><%=Separate(RS1("RemainedAmount"))%>&nbsp;</TD>
				</TR>
<%
			RS1.moveNext
			Loop
			RS1.Close

			if selectTop<>"" and tmpCounter = briefQtty then
				mySQL="SELECT ISNULL(COUNT(*), 0) AS CNT FROM Invoices INNER JOIN ARItems ON Invoices.ID = ARItems.Link INNER JOIN Accounts ON Invoices.Customer = Accounts.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.Type = 1) AND (ARItems.RemainedAmount > 0) AND (Accounts.CSR = 0)"

				Set RS1 = conn.execute(mySQL)
				count = RS1("CNT")
				RS1.close

%>
			<tr>
				<td colspan="10" class="CusTableHeader" style="text-align:right;background-color:#CCAA99;"><A HREF="?userID=<%=userID%>&showAll=on&panel=4">«œ«„Â œ«—œ ... ( ⁄œ«œ ﬂ· : <%=count%>) òÂ „Ã„Ê⁄ ¬‰ <%=Separate(RS2("sumTotalReceivable"))%> —Ì«· „Ì »«‘œ. Ê <%=Separate(RS2("sumRemainedAmount"))%> —Ì«· „«‰œÂ ¬‰ „Ì »«‘œ</A></td>
			</tr>
<%
			RS2.close
			end if
		else
%>
			<tr>
				<td colspan="10" class="CusTD3">ÂÌç</td>
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
<!--#include file="tah.asp" -->