	  <Tr><Td colspan="3" height="10px">
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
</STYLE>


	  &nbsp;</Td></Tr>
	  <Tr>
		<Td colspan="3" valign="top" align="center">

				<% if Auth(9 , 1) then %><input class="GenButton" type="Button" value="œ—Ì«›  ‰ﬁœ"onclick="window.open('../cashReg/ReceiptInput.asp?act=getReceipt&selectedCustomer=<%=cusID%>');" <% if AccountIsDisabled then %> disabled <% end if %>> <% end if %> 

				<% if Auth("A" , 4) then %><input class="GenButton" type="Button" value="œ—Ì«›  ÕÊ«·Â »«‰ﬂ "onclick="window.open('../bank/submitDraft.asp?selectedCustomer=<%=cusID%>');" <% if AccountIsDisabled then %> disabled <% end if %>> <% end if %> 

				<% if Auth(9 , 2) then %><input class="GenButton" type="button" value="Å—œ«Œ  ‰ﬁœ " onclick="window.open('../cashReg/CashPaymentInput.asp?act=getPayment&selectedCustomer=<%=cusID%>');" <% if AccountIsDisabled then %> disabled <% end if %>> <% end if %> 

				<% if Auth("A" , 1) then %><input class="GenButton" type="button" value="Å—œ«Œ  çﬂ" onclick="window.open('../bank/cheq.asp?act=enterCheque&selectedCustomer=<%=cusID%>');" <% if AccountIsDisabled then %> disabled <% end if %>>		 <% end if %> 
				
				<BR><BR><BR>

		</Td>
	  </Tr>

  	  <Tr>
		<Td valign="top" align="left">
		</Td>
		<Td valign="top" align="center">
			<table class="CustTable" cellspacing='1'>
				<tr>
					<td colspan="2" class="CusTableHeader">Å—œ«Œ ùÂ«</td>
				</tr>
				<%
				mySQL="SELECT Payments.ID, Payments.CreationDate, Users.RealName Creator FROM Payments INNER JOIN Users ON Payments.CreatedBy = Users.ID WHERE (Payments.Account='"& cusID & "') ORDER BY CreationDate DESC, Payments.ID"
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
				%>
					<tr>
						<td class="CusTD3">ÂÌç</td>
					</tr>
				<%
				end if
				tmpCounter=0
				while not RS1.eof 
					tmpCounter=tmpCounter+1
				%>
					<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
						<td><%=tmpCounter%></td>
						<td><a href='../AO/AccountReport.asp?act=showPayment&Payment=<%=RS1("ID")%>' target='_blank'><%=RS1("ID")%><br><%=RS1("Creator")%><br><%=RS1("CreationDate")%></a></td>
					</tr>
				<%
					RS1.moveNext
				wend
				%>
			</table>
		</Td>
		<Td valign="top" align="center">
			<table class="CustTable" cellspacing='1'>
				<tr>
					<td colspan="2" class="CusTableHeader">œ—Ì«› ùÂ«</td>
				</tr>
				<%
				mySQL="SELECT Receipts.ID, Receipts.SYS, Receipts.CreatedDate, Users.RealName AS Creator FROM Receipts INNER JOIN Users ON Receipts.CreatedBy = Users.ID WHERE (Customer='"& cusID & "') ORDER BY Receipts.CreatedDate DESC, Receipts.ID"
				Set RS1 = conn.execute(mySQL)
				if RS1.eof then
				%>
					<tr>
						<td class="CusTD3">ÂÌç</td>
					</tr>
				<%
				end if
				tmpCounter=0
				while not RS1.eof 
					tmpCounter=tmpCounter+1
				%>
					<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
						<td><%=tmpCounter%></td>
						<td><a href='../<%=RS1("SYS")%>/AccountReport.asp?act=showReceipt&receipt=<%=RS1("ID")%>' target='_blank'><%=RS1("ID")%><br><%=RS1("Creator")%><br><%=RS1("CreatedDate")%></a></td>
					</tr>
				<%
					RS1.moveNext
				wend
				%>
			</table>
		</Td>
  	  </Tr>