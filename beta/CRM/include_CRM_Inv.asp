	  <Tr><Td colspan="3" height="10px">
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:90%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
</STYLE>


	  &nbsp;</Td></Tr>
	  <Tr>
		<Td colspan="3" valign="top" align="center"  width=500>

				<BR><BR>
		</Td>
	  </Tr>

  	  <Tr>
		<Td valign="top" align="left" height=300>
		</Td>
		<Td valign="top" align="center">
			<table class="CustTable" cellspacing='1'>
				<tr>
					<td colspan="4" class="CusTableHeader">ãæÌæÏí ßÇáÇí ãÔÊÑí ÏÑ ÇäÈÇÑ</td>
				</tr>
				<tr>
					<TD class="CusTableHeader"><SMALL>äÇã ßÇáÇ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ßÏ ßÇáÇ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ÊÚÏÇÏ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>æÇÍÏ</SMALL></A></TD>
				</tr>
				<%
				mySQL="SELECT SUM(DERIVEDTBL.sumQtty) AS sumQttys, DERIVEDTBL.AccountID, Accounts.AccountTitle, InventoryItems.Name, InventoryItems.OldItemID, InventoryItems.Unit FROM (SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty, dbo.InventoryLog.owner AS AccountID, InventoryLog.itemid AS invid FROM dbo.InventoryLog where dbo.InventoryLog.voided=0 GROUP BY dbo.InventoryLog.owner, InventoryLog.itemid) DERIVEDTBL INNER JOIN Accounts ON DERIVEDTBL.AccountID = Accounts.ID INNER JOIN InventoryItems ON DERIVEDTBL.invid = InventoryItems.ID GROUP BY DERIVEDTBL.AccountID, Accounts.AccountTitle, InventoryItems.Name, InventoryItems.OldItemID, InventoryItems.Unit having SUM(DERIVEDTBL.sumQtty) <> 0 and DERIVEDTBL.AccountID="& cusID & ""
				Set RSS = conn.execute(mySQL)
				if RSS.eof then
				%>
					<tr>
						<td class="CusTD3" colspan=4>åí</td>
					</tr>
				<%
				end if
				tmpCounter=0
				while not RSS.eof 
					tmpCounter=tmpCounter+1
				%>
					<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
						<TD><%=RSS("name")%></TD>
						<TD align=center dir=ltr><%=RSS("OldItemID")%></TD>
						<TD align=center dir=ltr><%=RSS("sumQttys")%></TD>
						<TD align=center><%=RSS("Unit")%></TD>
					</tr>
				<%
					RSS.moveNext
				wend
				%>
			</table><BR><BR><BR>

			
			<table class="CustTable" cellspacing='1'>
				<tr>
					<td colspan="8" class="CusTableHeader">ÓÇÈŞå æÑæÏ æ ÎÑæÌ ßÇáÇåÇí ãÔÊÑí Èå ÇäÈÇÑ</td>
				</tr>
				<tr>
					<TD class="CusTableHeader"><SMALL>äÇã ßÇáÇ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ßÏßÇáÇ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>æÑæÏ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ÎÑæÌ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>æÇÍÏ</SMALL></A></TD>
					<TD class="CusTableHeader" align=center><SMALL>ÊÇÑíÎ </SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL> ÍæÇáå </SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ÊæÓØ</SMALL></A></TD>
				</tr>

				<%
				mySQL="SELECT InventoryLog.type, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.comments, InventoryLog.VoidedDate, InventoryLog.IsInput, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.ID, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE (InventoryLog.owner = "&  cusID & ") ORDER BY InventoryLog.ID DESC"
				Set RSS = conn.execute(mySQL)
				if RSS.eof then
				%>
					<tr>
						<td class="CusTD3" colspan="8">åí</td>
					</tr>
				<%
				end if
				tmpCounter=0
				while not RSS.eof 
					tmpCounter=tmpCounter+1
				%>
					<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>" <% if RSS("voided") then%> disabled title="ÍÏİ ÔÏå ÏÑ ÊÇÑíÎ <%=RSS("VoidedDate")%>"<% end if %>>
						<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("name")%></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("OldItemID")%></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"><% if RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"><% if not RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
						<TD align=right dir=ltr><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><%=RSS("Unit")%></TD>
						<TD dir=ltr align=center><%=RSS("logDate")%></span></TD>
						<TD><% if RSS("type")= "2" then
									response.write "<font color=red><b>ÇÕáÇÍ ãæÌæÏí</b></font>"
								elseif RSS("type")= "3" then
									response.write "<font color=green><b>ãÑÌæÚí</b></font>"
								elseif RSS("type")= "4" then
									response.write "<font color=blue><b>ÊÚÑíİ ßÇáÇí ÌÏíÏ </b></font>"
								elseif RSS("type")= "5" then
									response.write "<font color=orang><b>ÇäÊŞÇá</b></font>"
								elseif RSS("type")= "6" then
									response.write "<font color=#6699CC><b>æÑæÏ ÇÒ ÊæáíÏ</b></font>"
								elseif RSS("type")= "7" then
									response.write "<font color=#FF9966><b>æÑæÏ ÇÒ ÇäÈÇÑ ÔåÑíÇÑ</b></font>"
								elseif RSS("RelatedID")= "-1" then %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							   <% else 
									if RSS("IsInput") then 
										%> ÔãÇÑå ÓİÇÑÔ ÎÑíÏ: <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
									else
										%><A HREF="../inventory/default.asp?ed=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
									end if
								end if %>
							<% if trim(RSS("comments"))<> "-" and RSS("comments")<> "" then
								response.write " <br><br><B>ÊæÖíÍ:</B>  " & RSS("comments") 
							   end if %>

						</TD>
						<TD><%=RSS("RealName")%></TD>
					</tr>
				<%
					RSS.moveNext
				wend
				%>
			</table><BR><BR>
			
			<table class="CustTable" cellspacing='1'>
				<tr>
					<td colspan="8" class="CusTableHeader">ÓÇÈŞå æÑæÏ æ ÎÑæÌ ßÇáÇ ÈÑÇí ÓİÇÑÔåÇí ãÔÊÑí Èå ÇäÈÇÑ</td>
				</tr>
				<tr>
					<TD class="CusTableHeader"><SMALL>äÇã ßÇáÇ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ßÏßÇáÇ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ÎÑæÌ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>æÇÍÏ</SMALL></A></TD>
					<TD class="CusTableHeader" align=center><SMALL>ÊÇÑíÎ </SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL> ÍæÇáå </SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ÊæÓØ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ÓİÇÑÔ</SMALL></A></TD>
				</tr>

				<%
				mySQL="select InventoryLog.type, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.comments, InventoryLog.VoidedDate, InventoryLog.IsInput, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.ID, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName,Orders.id as ord FROM InventoryItemRequests inner join Orders on Orders.id = InventoryItemRequests.OrderID inner join InventoryPickuplistItems on InventoryItemRequests.ID=InventoryPickuplistItems.RequestID inner join InventoryLog on InventoryPickuplistItems.pickupListID=InventoryLog.RelatedID INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID where orders.Customer=" & cusID & " and InventoryLog.owner = -1 and InventoryLog.isInput=0 ORDER BY InventoryLog.ID DESC"
				Set RSS = conn.execute(mySQL)
				if RSS.eof then
				%>
					<tr>
						<td class="CusTD3" colspan="8">åí</td>
					</tr>
				<%
				end if
				tmpCounter=0
				while not RSS.eof 
					tmpCounter=tmpCounter+1
				%>
					<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>" <% if RSS("voided") then%> disabled title="ÍÏİ ÔÏå ÏÑ ÊÇÑíÎ <%=RSS("VoidedDate")%>"<% end if %>>
						<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("name")%></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("OldItemID")%></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"></span><%=RSS("Qtty")%></span></TD>
						<TD align=right dir=ltr><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><%=RSS("Unit")%></TD>
						<TD dir=ltr align=center><%=RSS("logDate")%></span></TD>
						<TD><% if RSS("type")= "2" then
									response.write "<font color=red><b>ÇÕáÇÍ ãæÌæÏí</b></font>"
								elseif RSS("type")= "3" then
									response.write "<font color=green><b>ãÑÌæÚí</b></font>"
								elseif RSS("type")= "4" then
									response.write "<font color=blue><b>ÊÚÑíİ ßÇáÇí ÌÏíÏ </b></font>"
								elseif RSS("type")= "5" then
									response.write "<font color=orang><b>ÇäÊŞÇá</b></font>"
								elseif RSS("type")= "6" then
									response.write "<font color=#6699CC><b>æÑæÏ ÇÒ ÊæáíÏ</b></font>"
								elseif RSS("type")= "7" then
									response.write "<font color=#FF9966><b>æÑæÏ ÇÒ ÇäÈÇÑ ÔåÑíÇÑ</b></font>"
								elseif RSS("RelatedID")= "-1" then %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							   <% else 
									if RSS("IsInput") then 
										%> ÔãÇÑå ÓİÇÑÔ ÎÑíÏ: <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
									else
										%><A HREF="../inventory/default.asp?ed=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
									end if
								end if %>
							<% if trim(RSS("comments"))<> "-" and RSS("comments")<> "" then
								response.write " <br><br><B>ÊæÖíÍ:</B>  " & RSS("comments") 
							   end if %>

						</TD>
						<TD><%=RSS("RealName")%></TD>
						<TD><a href="../order/TraceOrder.asp?act=show&order=<%=RSS("ord")%>"><%=RSS("ord")%></a></TD>
					</tr>
				<%
					RSS.moveNext
				wend
				%>
			</table><BR><BR>
			
			
		</Td>
		<Td valign="top" align="center">
			&nbsp;
		</Td>
  	  </Tr>