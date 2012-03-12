<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="ê“«—‘ ”—Å—”  ’‰œÊﬁ"
SubmenuItem=6
if not Auth(9 , 6) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.CClosTable {font-family:tahoma; font-size:9pt; direction: RTL; border:2 solid #330099;}
	.CClosTable td {padding:5;border:1pt solid gray;}
	.CClosTable a {text-decoration:none; color:#222288;}
	.CClosTable a:hover {text-decoration:underline;}
	.CClosTableTitle {background-color: #CCCCFF; text-align: center; font-weight:bold; height:50;}
	.CClosTableHeader {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.CClosTableFooter {background-color: #BBBBBB; direction: LTR; }
	.CClosTR1 {background-color: #DDDDDD;}
	.CClosTR2 {background-color: #FFFFFF;}
	.CClosTR3 {background-color: #EEEEFF;}
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {padding:5;border:1pt solid gray;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTableTitle {background-color: #CCCCFF; text-align: center; font-weight:bold; height:50;}
	.RepTableHeader {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.RepTableFooter {background-color: #BBBBBB; direction: LTR; }
	.RepTR1 {background-color: #DDDDDD;}
	.RepTR2 {background-color: #FFFFFF;}
	.GenInput { font-family:tahoma; font-size: 9pt;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</STYLE>
<%
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

if request("act")="showCashRegReport" AND request("CashRegID")<>"" then 
	CashRegID=request("CashRegID")
	if not isnumeric(CashRegID) then
		response.write "<br>" 
		call showAlert ("ç‰«‰ ’‰œÊﬁÌ ÊÃÊœ ‰œ«—œ",CONST_MSG_ERROR) 
		response.end
	end if
	
	mySQL="SELECT Bankers.Name AS BankerName, CashRegisters.*, Users.RealName FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID INNER JOIN Bankers ON CashRegisters.Banker = Bankers.ID WHERE (CashRegisters.ID='"& CashRegID & "')"
	Set RS1= conn.Execute(mySQL)
	if RS1.eof then 
		response.write "<br>" 
		call showAlert ("ç‰«‰ ’‰œÊﬁÌ ÊÃÊœ ‰œ«—œ",CONST_MSG_ERROR) 
		response.end
	else
		BankerName=RS1("BankerName")
		CashRegName=RS1("NameDate")
		CashierName=RS1("RealName")
		OpenningDate=RS1("OpenDate")
		OpenningAmount=cdbl(RS1("OpeningAmount"))
		totalCashAmountA=Separate(RS1("cashAmountA")) ' SAM
		totalCashAmountB=Separate(RS1("cashAmountB")) ' SAM
		totalChequeAmount=Separate(RS1("chequeAmount"))
		totalChequeQtty=RS1("chequeQtty")
		TotalAmount=Separate(cdbl(RS1("cashAmountA")) + cdbl(RS1("cashAmountB")) + cdbl(RS1("chequeAmount")))

		Set RS1=nothing
	end if

	mySQL="SELECT CashRegisterLines.*, ISNULL(Receipts.CashAmount,0) AS CashAmount, ISNULL(Receipts.ChequeAmount,0) AS ChequeAmount, Receipts.ChequeQtty, Payments.CashAmount AS PayAmount, Accounts_1.AccountTitle AS PayAccTitle, Accounts_2.AccountTitle AS RcpAccTitle, Accounts_1.ID AS PayAccID, Accounts_2.ID AS RcpAccID FROM CashRegisterLines LEFT OUTER JOIN Accounts Accounts_1 INNER JOIN Payments ON Accounts_1.ID = Payments.Account ON CashRegisterLines.Link = Payments.ID LEFT OUTER JOIN Accounts Accounts_2 INNER JOIN Receipts ON Accounts_2.ID = Receipts.Customer ON CashRegisterLines.Link = Receipts.ID WHERE (CashRegisterLines.CashReg = '"& CashRegID & "') ORDER BY CashRegisterLines.ID"
	Set RS1 = conn.execute(mySQL)
	Remained = 0
    Totalcredit = 0
    TotalDebit = 0
	tempCounter = 0

	if Not (RS1.EOF AND OpenningAmount=0) then
%>		
		<br>
		<TABLE class="RepTable" width='100%' align='center'>
		<TR>
			<TD class="RepTableTitle" colspan=9 dir='rtl' align='center'><br><B>ê“«—‘ ’‰œÊﬁ <span dir='LTR'><%=CashRegName%></span> - <%=CashierName%></B><BR>(„Õ·: <%=BankerName%>)<br>
			</TD>
		</TR>
		<TR class="RepTableHeader">
			<TD>#</TD>
			<TD width="50px"> «—ÌŒ</TD>
			<TD width="10px">”«⁄ </TD>
			<TD width="60px"> Ê÷ÌÕ« </TD>
			<td width="30px">«·›/»</td>
			<TD>Õ”«»</TD>
			<TD width="60px">»œÂﬂ«—</TD>
			<TD width="60px">»” «‰ﬂ«—</TD>
			<TD width="60px">„«‰œÂ</TD>
		</TR>
<%		
		if OpenningAmount<>0 then
			Remained = OpenningAmount
			Credit = OpenningAmount
			Totalcredit = OpenningAmount
			tempCounter = 0
%>			<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
				<td><%=tempCounter%></td>
				<td dir='LTR'><%=CashRegName%></td>
				<td colspan='3'>„ﬁœ«— «Ê·ÌÂ ’‰œÊﬁ</td>
				<td dir='LTR' align='right'><%=Separate(Debit)%>&nbsp;</td>
				<td dir='LTR' align='right'><%=Separate(Credit)%>&nbsp;</td>
				<td dir='LTR' align='right'><%=Separate(Remained)%>&nbsp;</td>
			</TR>
<%		end if
		While Not (RS1.EOF)
			tempCounter=tempCounter+1
			if RS1("Type")=1 then
				sourceLink="../AR/AccountReport.asp?act=showReceipt&receipt="& RS1("Link")
'				voidLink="Void.asp?act=voidReceipt&receipt="& RS1("Link")
				Description = "œ—Ì«›  "
				myAND=""

				CashAmount=cdbl(RS1("CashAmount"))
				ChequeAmount=cdbl(RS1("ChequeAmount"))

				if CashAmount<>0 then
					Description=Description & "‰ﬁœ "
					myAND="Ê "
				end if 
				if RS1("ChequeQtty")<>0 then
					Description=Description & myAND & RS1("ChequeQtty") & " ›ﬁ—Â çﬂ "
				end if 
				AccID=RS1("RcpAccID")
				AccTitle=RS1("RcpAccTitle")
				Credit= CashAmount + ChequeAmount
				Debit=0

			elseif RS1("Type")=2 then
				sourceLink="../AR/AccountReport.asp?act=showPayment&payment="& RS1("Link")
'				voidLink="Void.asp?act=voidPayment&payment="& RS1("Link")
				Description = "Å—œ«Œ  ‰ﬁœ"
				AccID=RS1("PayAccID")
				AccTitle=RS1("PayAccTitle")
				Credit=0
				Debit=RS1("PayAmount")
			else
				sourceLink="javascript:void(0);"
'				voidLink="javascript:void(0);"
				Description = ""
				AccountInfo=""
				Credit=0
				Debit=0 
			end if
			TotalDebit = TotalDebit + CLng(Debit)
			Totalcredit = Totalcredit + CLng(Credit)
			Remained = Remained + CLng(Credit) - CLng(Debit)
%>			<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
				<td><%=tempCounter%></td>
				<td dir='LTR'><%=RS1("Date")%></td>
				<td dir='LTR'><%=RS1("Time")%></td>
				<td><A HREF='<%=sourceLink%>' target='_blank'><%=Description%></A></td>
				<td align="center"><%
				if CBool(rs1("isA")) then 
					response.write("<b>«·›</b>")
				else
					response.write("»")
				end if
				%></td>
				<td><A HREF='../AR/AccountReport.asp?act=show&selectedCustomer=<%=AccID%>' target='_blank'><%=AccTitle%></A></td>
				<td dir='LTR' align='right'><%=Separate(Debit)%></td>
				<td dir='LTR' align='right'><%=Separate(Credit) %></td>
				<td dir='LTR' align='right'>
<%					if RS1("Voided")=TRUE then
%>						<div style="position:absolute;width:650;"><hr style="color:red;"></div>
<%						response.write "xx xx xx" '& Separate(remained) & "xx"  
						TotalDebit = TotalDebit - CLng(Debit)
						Totalcredit = Totalcredit - CLng(Credit)
						Remained = Remained - CLng(Credit) + CLng(Debit)
					else
						response.write Separate(remained)
					end if
%>				</td>
			</TR>
<%
			RS1.MoveNext
		Wend
		if Remained>=0 then
			remainedColor="green"
		else
			remainedColor="red"
		end if
%>
			<TR>
				<TD class="RepTableFooter" colspan='6'>&nbsp;&nbsp; : Ã„⁄</td>
				<TD class="RepTableFooter" align='right'><%=Separate(totaldebit)%></td>
				<TD class="RepTableFooter" align='right'><%=Separate(totalcredit)%></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR='<%=remainedColor%>'><%=Separate(Remained)%></FONT></td>
			</TR>
		</TABLE>
		<br>
		<TABLE class="CClosTable" width='90%' align='center' style="page-break-before:always;">
<%	
		mySQL="SELECT ReceivedCheques.ChequeNo, ReceivedCheques.ChequeDate, ReceivedCheques.Description, ReceivedCheques.BankOfOrigin,  receivedCheques.Amount, CashRegisterLines.* FROM ReceivedCheques INNER JOIN Receipts ON ReceivedCheques.Receipt = Receipts.ID INNER JOIN CashRegisterLines ON Receipts.ID = CashRegisterLines.Link WHERE (CashRegisterLines.CashReg = '"& CashRegID & "') AND (CashRegisterLines.Type = 1) ORDER BY CashRegisterLines.[Date], CashRegisterLines.[Time]"
		Set RS1 = conn.execute(mySQL)

		Remained = OpenningAmount
		Credit = OpenningAmount
		Totalcredit = OpenningAmount
		TotalDebit = 0
		tempCounter = 0
%>		<TR class='CClosTR3'>
			<td colspan='7' align='center'>Œ·«’Â ê“«—‘ ’‰œÊﬁ <span dir='LTR'><%=CashRegName%></span> (’‰œÊﬁœ«—: <%=CashierName%>° „Õ·: <%=BankerName%>)</TD>
		</TR>
		<TR>
			<TD colspan='6' class="CClosTableFooter">:„ﬁœ«— ‰ﬁœ</td>
			<TD class="CClosTableFooter" align='right'><span id='cashAmount'><%=Separate(cdbl(totalCashAmountA) + Cdbl(totalCashAmountB))%></span></TD>
		</TR>
		<TR>
			<TD colspan='6' class="CClosTableFooter" style='direction:RTL;text-align=left;'>„ﬁœ«— çﬂ (<span id='chequeQtty'><%=totalChequeQtty%></span> ›ﬁ—Â) :</td>
			<TD class="CClosTableFooter" align='right'><%=totalChequeAmount%></TD>
		</TR>
		<TR>
			<TD colspan='6' class="CClosTableFooter">: Ã„⁄</td>
			<TD class="CClosTableFooter" align='right'><%=TotalAmount%></td>
		</TR>
		<TR class='CClosTR3'>
			<td colspan='7' align='right'>·Ì”  çﬂ Â«:</td>
		</TR>
		<TR class="CClosTableHeader">
			<TD width="10px">#</TD>
			<TD width="100px"> «—ÌŒ œ—Ì«› </TD>
			<TD width="50px">‘„«—Â</TD>
			<TD width="50px">»«‰ﬂ</TD>
			<TD width="70px"> «—ÌŒ çﬂ</TD>
			<TD width="220px"> Ê÷ÌÕ« </TD>
			<TD width="70px">„»·€</TD>
		</TR>
<%
		While Not (RS1.EOF)
			tempCounter=tempCounter+1
			sourceLink="../AR/AccountReport.asp?act=showReceipt&receipt="& RS1("Link")
			Description = "œ—Ì«›  "
			myAND=""

			Amount=Separate(RS1("Amount"))

			if RS1("Voided") then
				tempCounter=tempCounter-1%>
				<TR class='CClosTR4'>
					<td>-</td>
					<td dir='LTR'><A HREF='<%=sourceLink%>' target='_blank'><%=RS1("Date") & " @ " & RS1("Time")%></A></td>
					<td><%=RS1("ChequeNo")%></td>
					<td><%=RS1("BankOfOrigin")%></td>
					<td dir='LTR'><%=RS1("ChequeDate")%></td>
					<td><%=RS1("Description")%></td>
					<td dir='LTR'>
						<div style="position:absolute;width:600;"><hr style="color:red;"></div><div align='right'><%=Amount%></div>
<%			else%>						
				<TR class='<%if tempCounter MOD 2 = 0 then response.write "CClosTR1" else response.write "CClosTR2"%>'>
					<td><%=tempCounter%></td>
					<td dir='LTR'><A HREF='<%=sourceLink%>' target='_blank'><%=RS1("Date") & " @ " & RS1("Time")%></A></td>
					<td><%=RS1("ChequeNo")%></td>
					<td><%=RS1("BankOfOrigin")%></td>
					<td dir='LTR'><%=RS1("ChequeDate")%></td>
					<td><%=RS1("Description")%></td>
					<td dir='LTR' align='right'>
						<%=Amount%>
<%			end if
%>					</td>
				</TR>
<%
			RS1.MoveNext
		Wend
%>
		</TABLE>
		<br>
<%	end if
else
%>
	<BR><BR>
	<TABLE class="RepTable" align='center'>
	<TR class="RepTableTitle">
		<TD colspan=<% if Auth(9,9) then response.write "11" else response.write "6" end if%>> ’‰œÊﬁ „Ê—œ ‰Ÿ— —« «‰ Œ«» ﬂ‰»œ </TD>
	</TR>
<%

	Banker = request("Banker")
	connector = ""
	Criteria = ""

	if isnumeric(Banker) then Banker = cint(Banker) else Banker = 0

	if Banker <> 0 then
		Criteria = Criteria & connector & "(CashRegisters.Banker = " & Banker & ")"
		connector = " AND "
	end if

	Cashier = request("Cashier")
	if isnumeric(Cashier) then Cashier = cint(Cashier) else Cashier = 0

	if Cashier <> 0 then
		Criteria = Criteria & connector & "(CashRegisters.Cashier = " & Cashier & ")"
		connector = " AND "
	end if

	openDateStart = sqlSafe(request("openDateStart"))
	openDateEnd =	sqlSafe(request("openDateEnd"))
	if openDateStart <> "" then
		Criteria = Criteria & connector & "(CashRegisters.openDate >= '" & openDateStart & "')"
		connector = " AND "
	end if
	if openDateEnd <> "" then
		Criteria = Criteria & connector & "(CashRegisters.openDate <= '" & openDateEnd & "')"
		connector = " AND "
	end if

	closeDateStart=	sqlSafe(request("closeDateStart"))
	closeDateEnd =	sqlSafe(request("closeDateEnd"))
	if closeDateStart <> "" then
		Criteria = Criteria & connector & "(CashRegisters.closeDate >= '" & closeDateStart & "')"
		connector = " AND "
	end if
	if closeDateEnd <> "" then
		Criteria = Criteria & connector & "(CashRegisters.closeDate <= '" & closeDateEnd & "')"
		connector = " AND "
	end if

	hasTheDate =	sqlSafe(request("hasTheDate"))
	if hasTheDate <> "" then
		Criteria = Criteria & connector & "(CashRegisters.openDate <= '" & hasTheDate & "') AND (CashRegisters.closeDate >= '" & hasTheDate & "')"
		connector = " AND "
	end if
	
	if Criteria <> "" then Criteria = " WHERE " & Criteria

	Sort = request("Sort")
	if isnumeric(Sort) then Sort = cint(Sort) else Sort = 0

	OrderBy=""
	Select Case Sort
	Case 0:
		OrderBy="ORDER BY CashRegisters.IsOpen DESC, CashRegisters.OpenDate DESC"
	Case 1:
		OrderBy="ORDER BY CashRegisters.NameDate DESC"
	Case 2:
		OrderBy="ORDER BY CashierName"
	Case 3:
		OrderBy="ORDER BY BankerName"
	Case 4:
		OrderBy="ORDER BY CashRegisters.OpenDate DESC"
	Case 5:
		OrderBy="ORDER BY CashRegisters.CloseDate DESC"
	End Select

	PageLink = "Banker=" & Banker & "&Cashier=" & Cashier & "&openDateStart=" & openDateStart & "&openDateEnd=" & openDateEnd & "&closeDateStart=" & closeDateStart & "&closeDateEnd=" & closeDateEnd & "&hasTheDate=" & hasTheDate & "&Sort=" & Sort

%>
	<FORM METHOD=POST ACTION="?">
	<TR class="RepTableTitle">
		<TD colspan=<% if Auth(9,9) then response.write "11" else response.write "6" end if%>><table border=0 width=100%>
		<tr>
			<td align='left'>„Õ·:</td>
			<td colspan=2><SELECT NAME="Banker" class="GenInput">
				<option Value=""> -- «‰ Œ«» ﬂ‰Ìœ -- </option>
<%				mySQL="SELECT * FROM Bankers WHERE (IsBankAccount <> 1)"
				Set RS1 = conn.Execute(mySQL)
				do while not (RS1.eof)
					if cint(RS1("ID")) = Banker then selectedText="selected" else selectedText=""
%>					<option Value="<%=RS1("ID")%>" <%=selectedText%>><%=RS1("Name")%></option>
<%
					RS1.movenext
				loop
				RS1.close
%>				</SELECT>
			</td>
			<td align='left'>’‰œÊﬁœ«—:</td>
			<td colspan=2><SELECT NAME="Cashier" class="GenInput">
				<option Value=""> -- «‰ Œ«» ﬂ‰Ìœ -- </option>
<%				mySQL="SELECT * FROM Users WHERE Display=1 ORDER BY RealName"
				Set RS1 = conn.Execute(mySQL)
				do while not (RS1.eof)
					if cint(RS1("ID")) = Cashier then selectedText="selected" else selectedText=""
%>					<option Value="<%=RS1("ID")%>" <%=selectedText%>><%=RS1("RealName")%></option>
<%
					RS1.movenext
				loop
				RS1.close
%>				</SELECT>

				
			</td>
		</tr>
		<tr>
			<td align='left'>»«“ ‘œ‰:</td>
			<td>«“:<INPUT TYPE="text" NAME="openDateStart" value="<%=openDateStart%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			<td> «:<INPUT TYPE="text" NAME="openDateEnd" value="<%=openDateEnd%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			<td>„— » »—«”«”:</td>
			<td colspan=2><SELECT NAME="Sort" class="GenInput">
				<option Value="0" <%if Sort = 0 then response.write "selected"%>> ÅÌ‘ ›—÷ </option>
				<option Value="1" <%if Sort = 1 then response.write "selected"%>>  «—ÌŒ ’‰œÊﬁ </option>
				<option Value="2" <%if Sort = 2 then response.write "selected"%>> ’‰œÊﬁœ«— </option>
				<option Value="3" <%if Sort = 3 then response.write "selected"%>> „Õ· </option>
				<option Value="4" <%if Sort = 4 then response.write "selected"%>>  «—ÌŒ «ÌÃ«œ </option>
				<option Value="5" <%if Sort = 5 then response.write "selected"%>>  «—ÌŒ »” ‰ </option>
				</SELECT>
			</td>
		</tr>
		<tr>
			<td align='left'>»” Â ‘œ‰:</td>
			<td>«“:<INPUT TYPE="text" NAME="closeDateStart" value="<%=closeDateStart%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			<td> «:<INPUT TYPE="text" NAME="closeDateEnd" value="<%=closeDateEnd%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
			<td align='left'>‘«„·  «—ÌŒ:</td>

			<td>
				<INPUT TYPE="text" NAME="hasTheDate" value="<%=hasTheDate%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);">
			</td>
			<td align=center>
				<INPUT TYPE="submit" class="GenButton" value="‰„«Ì‘">
			</td>
		</tr>
		</table>
		</TD>
	</TR>
	</FORM>
	<FORM METHOD=POST ACTION="?act=showCashRegReport">
<%

	if Auth(9,9) then 
		mySQL="SELECT CashRegisters.ID, MAX(Users.RealName) AS CashierName, MAX(Bankers.Name) AS BankerName, CashRegisters.NameDate, CashRegisters.CloseDate, isnull(SUM(case when CashRegisterLines.isA=1 and CashRegisterLines.Type=1 then ISNULL(Receipts.CashAmount,0) end),0) AS sumCashAmountA, isnull(SUM(case when CashRegisterLines.isA=0 and CashRegisterLines.Type=1 then ISNULL(Receipts.CashAmount,0) end),0) AS sumCashAmountB, SUM(ISNULL(case when CashRegisterLines.Type=1 then Receipts.ChequeAmount end,0)) AS sumChequeAmount, SUM(case when CashRegisterLines.isA=1 and CashRegisterLines.Type=2 then ISNULL(Payments.CashAmount,0) end) AS sumPayAmountA, SUM(case when CashRegisterLines.isA=0 and CashRegisterLines.Type=2 then ISNULL(Payments.CashAmount,0) end) AS sumPayAmountB FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID INNER JOIN Bankers ON CashRegisters.Banker = Bankers.ID LEFT OUTER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg and CashRegisterLines.voided=0 LEFT OUTER JOIN Payments ON CashRegisterLines.Link = Payments.ID LEFT OUTER JOIN Receipts ON CashRegisterLines.Link = Receipts.ID " &Criteria& " GROUP BY CashRegisters.ID,CashRegisters.NameDate, CashRegisters.CloseDate, CashRegisters.IsOpen, CashRegisters.OpenDate, CashRegisters.Banker, CashRegisters.Cashier "& OrderBy
	else
		mySQL="SELECT CashRegisters.ID, Users.RealName AS CashierName, Bankers.Name AS BankerName, CashRegisters.NameDate, CashRegisters.CloseDate FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID INNER JOIN Bankers ON CashRegisters.Banker = Bankers.ID " & Criteria & " " & OrderBy
	end if
	'-------------------Add by Sam
	if Auth(9,9) then 
		mySQLsum = "SELECT SUM(case when CashRegisterLines.isA=1 and CashRegisterLines.Type=1 then ISNULL(Receipts.CashAmount,0) end) AS sumCashAmountA, SUM(case when CashRegisterLines.isA=0 and CashRegisterLines.Type=1 then ISNULL(Receipts.CashAmount,0) end) AS sumCashAmountB, SUM(ISNULL(case when CashRegisterLines.Type=1 then Receipts.ChequeAmount end,0)) AS sumChequeAmount, SUM(case when CashRegisterLines.isA=1 and CashRegisterLines.Type=2 then ISNULL(Payments.CashAmount,0) end) AS sumPayAmountA, SUM(case when CashRegisterLines.isA=0 and CashRegisterLines.Type=2 then ISNULL(Payments.CashAmount,0) end) AS sumPayAmountB FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID INNER JOIN Bankers ON CashRegisters.Banker = Bankers.ID LEFT OUTER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg and CashRegisterLines.voided=0 LEFT OUTER JOIN Payments ON CashRegisterLines.Link = Payments.ID LEFT OUTER JOIN Receipts ON CashRegisterLines.Link = Receipts.ID " &Criteria
		Set rsSUM = Conn.Execute(mySQLsum)
	end if
	Set rs=Server.CreateObject("ADODB.Recordset")'Conn.Execute(mySQL)

	PageSize = 20
	rs.PageSize = PageSize

	rs.CursorLocation=3 'in ADOVBS_INC adUseClient=3
	rs.Open mySQL ,Conn,3
	TotalPages = rs.PageCount

	CurrentPage=1

	if isnumeric(Request.QueryString("p")) then
		pp=clng(Request.QueryString("p"))
		if pp <= TotalPages AND pp > 0 then
			CurrentPage = pp
		end if
	end if

	if not rs.eof then
		rs.AbsolutePage=CurrentPage
	end if

	if rs.eof then
%>		<tr>
			<td colspan=<% if Auth(9,9) then response.write "11" else response.write "7" end if%> class="CusTD3" style='direction:RTL;'>ÂÌç .</td>
		</tr>
<%	else
%>		<tr class="RepTableHeader">
			<td width="20px">#</td>
			<td width="10px"><input disabled type="radio"></td>
			<td>  «—ÌŒ ’‰œÊﬁ </td>
			<td> ’‰œÊﬁœ«— </td>
			<td> „Õ· </td>
			<td>  «—ÌŒ »” ‰ </td>
			<%
			'-------------------Add by Sam
			if Auth(9,9) then 
			%>
			<td>œ—Ì«›  ‰ﬁœ «·›</td>
			<td>œ—Ì«›  ‰ﬁœ »</td>
			<td>œ—Ì«›  çﬂ</td>
			<td>Å—œ«Œ  «·›</td>
			<td>Å—œ«Œ  »</td>
			<%
			end if
			%>
		</tr>
		<%
		'-------------------Add by Sam
		if (Auth(9,9)) then 
			if NOT rsSUM.eof then 
		%>
			<TR class='RepTR1'>
					<TD colspan=6 align="center" style="font-weight: bold;">Ã„⁄ ﬂ·</TD>
					<%
					
					if Auth(9,9) then 
					%>
					<td><%=Separate(rsSUM("sumCashAmountA"))%></td>
					<td><%=Separate(rsSUM("sumCashAmountB"))%></td>
					<td><%=Separate(rsSUM("sumChequeAmount"))%></td>
					<td><%=Separate(rsSUM("sumPayAmountA"))%></td>
					<td><%=Separate(rsSUM("sumPayAmountB"))%></td>
					<%
					end if
					%>
			</TR>
	<%		end if
		end if
		tempCounter=(CurrentPage - 1) * PageSize

		Do While NOT rs.eof AND (rs.AbsolutePage = CurrentPage) 
			tempCounter = tempCounter+1
			CloseDate=rs("CloseDate")
			if isnull(CloseDate) then
				CloseDate="<FONT COLOR='Red'>Â‰Ê“ »«“ «” </FONT>"
			end if
%>
			<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
				<TD><%=tempCounter%></TD>
				<TD><INPUT TYPE="radio" NAME="CashRegID" Value='<%=rs("ID")%>'></TD>
				<TD dir='LTR'><%=rs("NameDate")%></TD>
				<TD><%=rs("CashierName")%></TD>
				<TD><%=rs("BankerName")%></TD>
				<TD dir='LTR'><%=CloseDate%></TD>
				<%
				'-------------------Add by Sam
				if Auth(9,9) then 
				%>
				<td><%=Separate(rs("sumCashAmountA"))%></td>
				<td><%=Separate(rs("sumCashAmountB"))%></td>
				<td><%=Separate(rs("sumChequeAmount"))%></td>
				<td><%=Separate(rs("sumPayAmountA"))%></td>
				<td><%=Separate(rs("sumPayAmountB"))%></td>
				<%
				end if
				%>
			</TR>

<%
			rs.MoveNext
		Loop
%>			<tr class="RepTableTitle">
				<td align='right' colspan=<% if Auth(9,9) then response.write "11" else response.write "6" end if%>>
<%					
		if TotalPages > 1 then
					response.write "’›ÕÂ &nbsp;"
					for i=1 to TotalPages 
						if i = CurrentPage then 
%>							[<%=i%>]&nbsp;
<%						else
%>							<A HREF="?<%=PageLink%>&p=<%=i%>"><%=i%></A>&nbsp;
<%						end if 
						if i mod 16 = 0 then response.write "<br>" 
					next 
			end if
	end if
%>					<BR>
					<P align=center><INPUT TYPE="submit" Value="«‰ Œ«»" class="GenButton"></P>
				</td>
			</tr>
	</TABLE>
	</FORM>
<%
end if
conn.Close
%>
</font>
<!--#include file="tah.asp" -->
