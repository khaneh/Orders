<%'
  'This Include file needs the variables below:
  '
  '		Conn	
  '		cusID	(Customer ID)
  '
%>
<script type="text/javascript">
	$(document).ready(function(){
		$(".getName").each(function (index){
			var TD = $(this);
			$.ajax({
				type:"POST",
				url:"/service/json_getName.asp",
				data:{act:TD.attr("act"),id:TD.attr("iID")},
				dataType:"json",
				cache: true
			}).done(function (data){
				TD.html(data.name);
			});
		});

	});
</script>
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;height:30;}
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
</STYLE>
	  <Tr>
		<Td colspan="2" valign="center" align="center">
			<div class="btn-toolbar">
				<% if Auth(2 , 9) then %>
				<div class="btn-group">
				  <button class="btn btn-inverse dropdown-toggle" data-toggle="dropdown">ê—› ‰ «” ⁄·«„
				    <span class="caret"></span>
				  </button>
				  <ul class="dropdown-menu">
				  <%
					  if not AccountIsDisabled then 
						  set rs=Conn.Execute("select * from orderTypes")
						  while not rs.eof
						  	Response.write "<li><a href='../order/order.asp?act=getNew&OrderType=" & rs("id") & "&selectedCustomer=" & cusID & "'>" & rs("name") & "</a></li>"
						  	rs.MoveNext
						  wend
						end if
				  %>
				    
				  </ul>
				</div>
				<% end if %>
				
				<% if Auth(6 , 2) then %>
				<div class="btn-group">
					<input class="btn btn-warning" type="button" value="Ê—Êœ «⁄·«„ÌÂ" onclick="window.open('../AR/MemoInput.asp?act=getMemo&selectedCustomer=<%=cusID%>');" <% if AccountIsDisabled then Response.write " disabled " %> /> 
				</div>
				<% end if %> 
				<% if Auth(6 , 7) then %>
				<div class="btn-group">
					<input class="btn btn-success" type="button" value="œÊŒ ‰" onclick="window.open('../AR/ItemsRelation.asp?sys=AR&act=relate&selectedCustomer=<%=CusID%>');"/> 
				</div>
				<% end if %> 
				<% if Auth(6 , 6) then %>
				<div class="btn-group">
					<input class="btn btn-info" type="button" value="ê“«—‘ Õ”«»" onclick="window.open('../AR/AccountReport.asp?sys=AR&act=show&selectedCustomer=<%=CusID%>');" /> 
				</div>
				<% end if %> 
			</div>
		</Td>
	  </Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="8" class="CusTableHeader">”›«—‘ Â«Ì œ— Ã—Ì«‰</td>
				</tr>
				<%
'				mySQL="SELECT Orders.ID, ISNULL(Invoices.Approved, 0) AS Approved, orders_trace.* FROM orders_trace RIGHT OUTER JOIN Orders ON orders_trace.radif_sefareshat = Orders.ID LEFT OUTER JOIN Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (Orders.Customer = '"& cusID & "') AND (ISNULL(Invoices.Approved, 0) = 0) ORDER BY Orders.CreatedDate DESC, Orders.ID"
'				Changed by kid 820817
				mySQL="SELECT * FROM Orders WHERE (isClosed = 0) and (isOrder=1) AND (Customer = '"& cusID & "') ORDER BY CreatedDate DESC, ID"

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
						<td class="CusTD3"> «—ÌŒ  ÕÊÌ·</td>
						<td class="CusTD3">‰Ê⁄</td>
						<td class="CusTD3">⁄‰Ê«‰</td>
						<td class="CusTD3">„—Õ·Â</td>
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
						<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="window.open('../order/Order.asp?act=show&id=<%=RS1("ID")%>');">
							<TD style="height:30px;"><%=tmpCounter%></TD>
							<TD style="height:30px;" class="orderID"><%=RS1("id")%></TD>
							<TD class="getName" act="user" iID="<%=rs1("createdBy")%>"></TD>
							<TD dir="LTR" align='right'><%=shamsiDate(RS1("createdDate"))%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%if not IsNull(RS1("returnDate")) then shamsiDate(RS1("returnDate"))%>&nbsp;</TD>
							<TD class="getName" act="orderType" iID="<%=RS1("type")%>"></TD>
							<TD><%=RS1("orderTitle")%>&nbsp;</TD>
							<TD class="getName" act="orderStep" iID="<%=RS1("step")%>"></TD>
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
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="7" class="CusTableHeader" style="background-color:#CCCC99;">›«ﬂ Ê— Â«Ì œ— Ã—Ì«‰ ( «ÌÌœ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.TotalReceivable FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID WHERE (Invoices.Customer='"& cusID & "') AND (Invoices.Voided = 0) AND (Invoices.Approved = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID"
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
						<td class="CusTD3">„—»Êÿ »Â ”›«—‘ùÂ«Ì</td>
						<td class="CusTD3">„—»Êÿ »Â «” ⁄·«„ùÂ«Ì</td>
						<td class="CusTD3">„»·€</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 
						RelatedOrders = ""
						RelatedQuotes=""
						tmpsep=""
						mySQL="SELECT [Order] FROM InvoiceOrderRelations WHERE Invoice='"& RS1("ID") & "'"
						Set RS2 = conn.execute(mySQL)
						Do While not RS2.eof 
'							RelatedOrders = RelatedOrders & tmpsep & "<a href='../order/TraceOrder.asp?act=show&order="& RS2("Order") & "' target='_blank'>"& RS2("Order") & "</a>"
							RelatedOrders = RelatedOrders & tmpsep & RS2("Order")
							tmpsep=", "
							RS2.Movenext
						Loop
						RS2.close
						Set Rs2=Nothing
						tmpsep=""
						mySQL="SELECT Quote FROM InvoiceQuoteRelations WHERE Invoice='"& RS1("ID") & "'"
						Set RS2 = conn.execute(mySQL)
						Do While not RS2.eof 
'							RelatedOrders = RelatedOrders & tmpsep & "<a href='../order/TraceOrder.asp?act=show&order="& RS2("Order") & "' target='_blank'>"& RS2("Order") & "</a>"
							RelatedQuotes = RelatedQuotes & tmpsep & RS2("Quote")
							tmpsep=", "
							RS2.Movenext
						Loop
						RS2.close
						Set Rs2=Nothing
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
							<TD dir="LTR" align='right'><%=RelatedOrders%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RelatedQuotes%>&nbsp;</TD>
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
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="7" class="CusTableHeader" style="background-color:#33BB99;">›«ﬂ Ê— Â«Ì  «ÌÌœ ‘œÂ (’«œ— ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Invoices.ID, Invoices.CreatedDate, Invoices.ApprovedDate, Invoices.TotalReceivable, Users.RealName AS Approver FROM Invoices INNER JOIN Users ON Invoices.ApprovedBy = Users.ID WHERE (Invoices.Customer='"& cusID & "') AND (Invoices.Voided = 0) AND (Invoices.Approved = 1) AND (Invoices.Issued = 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID"
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
						<td class="CusTD3"> «ÌÌœ ﬂ‰‰œÂ</td>
						<td class="CusTD3"> «—ÌŒ  «ÌÌœ</td>
						<td class="CusTD3">„—»Êÿ »Â ”›«—‘ Â«Ì</td>
						<td class="CusTD3">„»·€</td>
					</tr>
<%					tmpCounter=0
					Do while not RS1.eof 

						RelatedOrders = ""
						tmpsep=""
						mySQL="SELECT [Order] FROM InvoiceOrderRelations WHERE Invoice='"& RS1("ID") & "'"
						Set RS2 = conn.execute(mySQL)
						Do While not RS2.eof 
'							RelatedOrders = RelatedOrders & tmpsep & "<a href='../order/TraceOrder.asp?act=show&order="& RS2("Order") & "' target='_blank'>"& RS2("Order") & "</a>"
							RelatedOrders = RelatedOrders & tmpsep & RS2("Order")
							tmpsep=", "
							RS2.Movenext
						Loop
						RS2.close
						Set Rs2=Nothing

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
							<TD><%=RS1("Approver")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("ApprovedDate")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RelatedOrders%>&nbsp;</TD>
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
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="10" class="CusTableHeader" style="background-color:#CCAA99;">›«ﬂ Ê— Â«Ì ’«œ— ‘œÂ ( ”ÊÌÂ ‰‘œÂ)</td>
				</tr>
				<%
				mySQL="SELECT Invoices.ID, Invoices.CreatedDate, Users.RealName AS Creator, Invoices.ApprovedDate, Invoices.TotalReceivable, Users_1.RealName AS Approver,  Users_2.RealName AS Issuer, Invoices.IssuedDate, ARItems.RemainedAmount FROM Invoices INNER JOIN Users ON Invoices.CreatedBy = Users.ID INNER JOIN Users Users_1 ON Invoices.ApprovedBy = Users_1.ID INNER JOIN Users Users_2 ON Invoices.IssuedBy = Users_2.ID INNER JOIN ARItems ON Invoices.ID = ARItems.Link and ARItems.Type = 1 and ARItems.voided = 0  WHERE (Invoices.Customer = '"& cusID & "') AND (Invoices.Voided = 0) AND (Invoices.Issued = 1) AND (ARItems.RemainedAmount > 0) ORDER BY Invoices.CreatedDate DESC, Invoices.ID"
				'response.write mysql
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
%>
					<tr>
						<td colspan="10" class="CusTD3">ÂÌç</td>
					</tr>
<%
				else
%>					<tr>
						<td class="CusTD3">#</td>
						<td class="CusTD3">‘„«—Â</td>
						<td class="CusTD3">«ÌÃ«œ</td>
						<td class="CusTD3"> «—ÌŒ «ÌÃ«œ</td>
						<td class="CusTD3"> «ÌÌœ</td>
						<td class="CusTD3"> «—ÌŒ  «ÌÌœ</td>
						<td class="CusTD3">’œÊ—</td>
						<td class="CusTD3"> «—ÌŒ ’œÊ—</td>
						<td class="CusTD3">„»·€</td>
						<td class="CusTD3">„«‰œÂ</td>
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
							<TD><%=RS1("Issuer")%>&nbsp;</TD>
							<TD dir="LTR" align='right'><%=RS1("IssuedDate")%>&nbsp;</TD>
							<TD><%=Separate(RS1("TotalReceivable"))%>&nbsp;</TD>
							<TD><%=Separate(RS1("RemainedAmount"))%>&nbsp;</TD>
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
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="10" class="CusTableHeader" style="background-color:#3AC;">”›«—‘ùÂ«</td>
				</tr>
				<%
				mySQL="select orders.id,orders.orderTitle, orders.qtty, orders.createdDate, invoices.issuedDate, invoices.totalReceivable, orderStatus.name as vazyat,orderSteps.name as marhale,invoices.id as invoice from orders inner join orderStatus on orders.status = orderStatus.id inner join orderSteps on orders.step = orderSteps.id left outer join InvoiceOrderRelations on InvoiceOrderRelations.[Order] = orders.id left outer join Invoices on InvoiceOrderRelations.Invoice=invoices.ID where orders.Customer=" & cusID & " and orders.isOrder=1 order by orders.createdDate desc"
				set rs=Conn.Execute(mySQL)
				if rs.eof then 
				%>
				<tr>
					<td colspan="10" class="CusTD3">ÂÌç</td>
				</tr>
				<%
				else
				%>
				<tr>
					<td class="CusTD3">‘„«—Â ”›«—‘</td>
					<td class="CusTD3">‘—Õ ”›«—‘</td>
					<td class="CusTD3"> Ì—«é</td>
					<td class="CusTD3"> «—ÌŒ ”›«—‘</td>
					<td class="CusTD3"> «—ÌŒ ’œÊ—</td>
					<td class="CusTD3" colspan=4>„»·€</td>
				</tr>
				<%
				tmpColor=""
				while not rs.eof
					if tmpColor<>"#FFFFFF" then
						tmpColor="#FFFFFF"
						tmpColor2="#FFFFBB"
					Else
						tmpColor="#DDDDDD"
						tmpColor2="#EEEEBB"
					End if 
				%>
				<tr bgcolor="<%=tmpColor%>" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'">
					<td title="<%=rs("vazyat") & "° " & rs("marhale")%>"><a href='../order/Order.asp?act=show&id=<%=rs("id")%>'><%=rs("id")%></a></td>
					<td><%=rs("orderTitle")%></td>
					<td><%=rs("qtty")%></td>
					<td><%=shamsiDate(rs("createdDate"))%></td>
					<td title="‰„«Ì‘ ›«ﬂ Ê—"><a href="../AR/AccountReport.asp?act=showInvoice&invoice=<%=rs("invoice")%>"><%=rs("issuedDate")%></a></td>
					<td colspan=4><%=Separate(rs("totalReceivable"))%></td>
				</tr>
				<%	
					rs.moveNext
				wend
				end if
				rs.close
				%>
			</table>
		</td>
	   </tr>
	   <Tr>
		<Td colspan="2" valign="top" align="center">
			<table class="CustTable" cellspacing='1' style='width:90%;'>
				<tr>
					<td colspan="10" class="CusTableHeader" style="background-color:#AAAAEE;">«” ⁄·«„ùÂ«</td>
				</tr>
				<%
				mySQL="select * from Orders where isOrder = 0 and Customer=" & cusID
				set rs=Conn.Execute(mySQL)
				if rs.eof then 
				%>
				<tr>
					<td colspan="10" class="CusTD3">ÂÌç</td>
				</tr>
				<%
				else
				tmpColor=""					
				%>
				<tr>
					<td class="CusTD3">‘„«—Â «” ⁄·«„</td>
					<td class="CusTD3">‘—Õ «” ⁄·«„</td>
					<td class="CusTD3"> «—ÌŒ «” ⁄·«„</td>
					<td class="CusTD3" colspan="6"> Ê÷ÌÕ« </td>
				</tr>
				<%
				
				while not rs.eof
					if tmpColor<>"#FFFFFF" then
						tmpColor="#FFFFFF"
						tmpColor2="#FFFFBB"
					Else
						tmpColor="#DDDDDD"
						tmpColor2="#EEEEBB"
					End if 
				%>
				<tr bgcolor="<%=tmpColor%>" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'">
					<td><a href="../order/order.asp?act=show&id=<%=rs("id")%>"><%=rs("id")%></a></td>
					<td><%=rs("orderTitle")%></td>
					<td><%=shamsiDate(rs("createdDate"))%></td>
					<td colspan="6"><%=rs("Notes")%></td>
				</tr>
				<%
					rs.moveNext
				wend
				end if
				%>
			</table>
		</td>
	   </tr>
