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
					<td colspan="4" class="CusTableHeader">„ÊÃÊœÌ ﬂ«·« œ— «‰»«—</td>
				</tr>
				<tr>
					<TD class="CusTableHeader"><SMALL>‰«„ ﬂ«·«</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ﬂœ ﬂ«·«</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL> ⁄œ«œ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>Ê«Õœ</SMALL></A></TD>
				</tr>
				<%
				mySQL="SELECT SUM(DERIVEDTBL.sumQtty) AS sumQttys, DERIVEDTBL.AccountID, Accounts.AccountTitle, InventoryItems.Name, InventoryItems.OldItemID, InventoryItems.Unit FROM (SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty, dbo.InventoryLog.owner AS AccountID, InventoryLog.itemid AS invid FROM dbo.InventoryLog where dbo.InventoryLog.voided=0 GROUP BY dbo.InventoryLog.owner, InventoryLog.itemid) DERIVEDTBL INNER JOIN Accounts ON DERIVEDTBL.AccountID = Accounts.ID INNER JOIN InventoryItems ON DERIVEDTBL.invid = InventoryItems.ID GROUP BY DERIVEDTBL.AccountID, Accounts.AccountTitle, InventoryItems.Name, InventoryItems.OldItemID, InventoryItems.Unit having SUM(DERIVEDTBL.sumQtty) <> 0 and DERIVEDTBL.AccountID="& cusID & ""
				Set RSS = conn.execute(mySQL)
				if RSS.eof then
				%>
					<tr>
						<td class="CusTD3" colspan=4>ÂÌç</td>
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
					<td colspan="8" class="CusTableHeader">”«»ﬁÂ Ê—Êœ Ê Œ—ÊÃ ﬂ«·« »Â «‰»«—</td>
				</tr>
				<tr>
					<TD class="CusTableHeader"><SMALL>‰«„ ﬂ«·«</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>ﬂœﬂ«·«</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>Ê—Êœ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>Œ—ÊÃ</SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL>Ê«Õœ</SMALL></A></TD>
					<TD class="CusTableHeader" align=center><SMALL> «—ÌŒ </SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL> ÕÊ«·Â </SMALL></A></TD>
					<TD class="CusTableHeader"><SMALL> Ê”ÿ</SMALL></A></TD>
				</tr>

				<%
				mySQL="SELECT InventoryLog.type, InventoryLog.Voided, InventoryLog.VoidedBy, InventoryLog.comments, InventoryLog.VoidedDate, InventoryLog.IsInput, InventoryItems.Unit, InventoryItems.Name, InventoryItems.OldItemID, InventoryLog.logDate, InventoryLog.Qtty, InventoryLog.RelatedID, InventoryLog.ItemID, InventoryLog.ID, InventoryLog.CreatedBy, InventoryLog.owner, Users.RealName FROM InventoryLog INNER JOIN InventoryItems ON InventoryLog.ItemID = InventoryItems.ID INNER JOIN Users ON InventoryLog.CreatedBy = Users.ID WHERE (InventoryLog.owner = "&  cusID & ") ORDER BY InventoryLog.ID DESC"
				Set RSS = conn.execute(mySQL)
				if RSS.eof then
				%>
					<tr>
						<td class="CusTD3" colspan="8">ÂÌç</td>
					</tr>
				<%
				end if
				tmpCounter=0
				while not RSS.eof 
					tmpCounter=tmpCounter+1
				%>
					<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>" <% if RSS("voided") then%> disabled title="Õœ› ‘œÂ œ—  «—ÌŒ <%=RSS("VoidedDate")%>"<% end if %>>
						<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("name")%></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"><%=RSS("OldItemID")%></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"><% if RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
						<TD align=right dir=ltr><span style="font-size:10pt"><% if not RSS("IsInput") then %><%=RSS("Qtty")%><% end if %></span></TD>
						<TD align=right dir=ltr><% if RSS("voided") then%><div style="position:absolute;width:520;"><hr style="color:red;"></div><% end if %><%=RSS("Unit")%></TD>
						<TD dir=ltr align=center><%=RSS("logDate")%></span></TD>
						<TD><% if RSS("type")= "2" then
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
								elseif RSS("RelatedID")= "-1" then %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							   <% else 
									if RSS("IsInput") then 
										%> ‘„«—Â ”›«—‘ Œ—Ìœ: <A HREF="../purchase/outServiceTrace.asp?od=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
									else
										%><A HREF="../inventory/default.asp?ed=<%=RSS("RelatedID")%>"><%=RSS("RelatedID")%></A><%
									end if
								end if %>
							<% if trim(RSS("comments"))<> "-" and RSS("comments")<> "" then
								response.write " <br><br><B> Ê÷ÌÕ:</B>  " & RSS("comments") 
							   end if %>

						</TD>
						<TD><%=RSS("RealName")%></TD>
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