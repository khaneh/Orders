<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle=" ‰ŸÌ„«  ﬂ·Ì"
SubmenuItem=8


function Separate(inputTxt)
myMinus=""
input=inputTxt
t=instr(input, ".")
if t>0 then 
	expPart = mid(input, t+1, 2)
	input = left(input, t-1)
end if
if left(input,1)="-" then
	myMinus="-"
	input=right(input,len(input)-1)
end if
if len(input) > 3 then
	tmpr=right(input ,3)
	tmpl=left(input , len(input) - 3 )
	result = tmpr
	while len(tmpl) > 3
		tmpr=right(tmpl,3)
		result = tmpr & "," & result 
		tmpl=left(tmpl , len(tmpl) - 3 )
	wend
	if len(tmpl) > 0 then
		result = tmpl & "," & result
	end if 
else
	result = input
end if 
	if t>0 then 
		result = result & "." & expPart
	end if

	Separate=myMinus & result
end function

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
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; height:100%; background-color:#C3DBEB;}
</STYLE>
	  <TaBlE class="CustTable4" cellspacing="2" cellspacing="2">
	  <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="9" class="CusTableHeader" style="text-align:right;">”›«—‘ Â«ÌÌ òÂ «’·« ›«ò Ê— ‰œ«—‰œ</td>
				</tr>
				<%
				mySQL="SELECT Accounts.AccountTitle, orders_trace.radif_sefareshat, orders_trace.order_date, orders_trace.order_kind, orders_trace.order_title,  orders_trace.salesperson, orders_trace.vazyat, orders_trace.marhale FROM Orders INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN Accounts ON Orders.Customer = Accounts.ID WHERE (Orders.Closed = 0) AND (Orders.ID NOT IN (SELECT [ORDER] FROM InvoiceOrderRelations)) ORDER BY Orders.CreatedDate DESC, Orders.ID"
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
%>
					<tr>
						<td colspan="8" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr>
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
<%					tmpCounter=0
					Do while not RS1.eof 
						tmpCounter = tmpCounter + 1
						if tmpCounter mod 2 = 1 then
							tmpColor="#FFFFFF"
							tmpColor2="#FFFFBB"
						Else
							tmpColor="#DDDDDD"
							tmpColor2="#EEEEBB"
						End if 
						'alert(this.getElementByTagName('td').items(0).data);
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
<%						RS1.moveNext
					Loop
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
					<td colspan="8" class="CusTableHeader" style="text-align:right;">›«ﬂ Ê— Â«Ì œ— Ã—Ì«‰ ( «ÌÌœ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.TotalReceivable, InvoiceOrderRelations.[Order], orders_trace.vazyat,  orders_trace.marhale FROM orders_trace RIGHT OUTER JOIN InvoiceOrderRelations ON orders_trace.radif_sefareshat = InvoiceOrderRelations.[Order] RIGHT OUTER JOIN Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (Invoices.Voided = 0) AND (Invoices.Approved = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID"
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
%>
					<tr>
						<td colspan="8" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr>
						<td class="CusTD3">#</td>
						<td class="CusTD3">‘„«—Â ›«ﬂ Ê—</td>
						<td class="CusTD3">«ÌÃ«œ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ</td>
						<td class="CusTD3">„—»Êÿ »Â ”›«—‘</td>
						<td class="CusTD3">Ê’⁄Ì </td>
						<td class="CusTD3">„—Õ·Â</td>
						<td class="CusTD3">„»·€</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 
						tmpCounter = tmpCounter + 1
						if tmpCounter mod 2 = 1 then
							tmpColor="#FFFFFF"
							tmpColor2="#FFFFBB"
						Else
							tmpColor="#DDDDDD"
							tmpColor2="#EEEEBB"
						End if 
						'alert(this.getElementByTagName('td').items(0).data);
%>
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("ûCreator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("Order")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("vazyat")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("Marhale")%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
						</TR>
<%						RS1.moveNext
					Loop
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
					<td colspan="7" class="CusTableHeader" style="text-align:right;">›«ﬂ Ê— Â«Ì  «ÌÌœ ‘œÂ (’«œ— ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.ApprovedDate, Invoices.TotalReceivable, Users_1.RealName AS Approver FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID WHERE (Invoices.Voided = 0) AND (Invoices.Issued = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID"
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
%>
					<tr>
						<td colspan="7" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr>
						<td class="CusTD3">#</td>
						<td class="CusTD3">‘„«—Â ›«ﬂ Ê—</td>
						<td class="CusTD3">«ÌÃ«œ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ</td>
						<td class="CusTD3"> «ÌÌœ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ  «ÌÌœ</td>
						<td class="CusTD3">„»·€</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 
						tmpCounter = tmpCounter + 1
						if tmpCounter mod 2 = 1 then
							tmpColor="#FFFFFF"
							tmpColor2="#FFFFBB"
						Else
							tmpColor="#DDDDDD"
							tmpColor2="#EEEEBB"
						End if 
						'alert(this.getElementByTagName('td').items(0).data);
%>
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../AR/AccountReport.asp?act=showInvoice&invoice=<%=RS1("ID")%>');">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;"><%=RS1("ID")%></TD>
							<TD><%=RS1("ûCreator")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("CreatedDate")%>&nbsp;</TD>
							<TD><%=RS1("Approver")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("ApprovedDate")%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
						</TR>
<%						RS1.moveNext
					Loop
				end if
				%>
			</table>
		</Td>
	  </Tr>
<!--#include file="tah.asp" -->