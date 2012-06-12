<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="»” ‰ ’‰œÊﬁ"
SubmenuItem=4
if not Auth(9 , 4) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.CClosTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.CClosTable td {padding:5;border:1pt solid gray;}
	.CClosTable a {text-decoration:none; color:#222288;}
	.CClosTable a:hover {text-decoration:underline;}

	.CClosTableTitle {background-color: #CCCCFF; text-align: center; font-weight:bold; height:50;}
	.CClosTableHeader {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.CClosTableFooter {background-color: #BBBBBB; direction: LTR; }
	.CClosTR1 {background-color: #DDDDDD;}
	.CClosTR2 {background-color: #FFFFFF;}
	.CClosTR3 {background-color: #EEEEFF;}
	.CClosInpTable { font-family:tahoma; font-size: 9pt; padding:0; direction: RTL; width:100%;}
	.CClosInpTable Tr {Height:25px; border: 1px solid black;}
	.CClosInpTable Input { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction: LTR;}
	.CClosInpTable Select { font-family:tahoma; font-size: 9pt; border: 1px solid black; }

	.GenInput { font-family:tahoma; font-size: 9pt;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</STYLE>
<%


function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function


if request("act")="Report" AND request("CashReg")<>"" then 
	CashRegID=request("CashReg")
	if not isnumeric(CashRegID) then
		response.write "<br>" 
		call showAlert ("ç‰«‰ ’‰œÊﬁÌ ÊÃÊœ ‰œ«—œ",CONST_MSG_ERROR) 
		response.end
	end if
'	response.write cashRegID
	mySQL="SELECT CashRegisters.*, Users.RealName FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID WHERE (CashRegisters.IsOpen=1) AND (CashRegisters.ID='"& CashRegID & "')"
	Set RS1= conn.Execute(mySQL)
'	response.write mySQL
	if RS1.eof then 
		response.write "<br>" 
		call showAlert ("ç‰«‰ ’‰œÊﬁÌ ÊÃÊœ ‰œ«—œ",CONST_MSG_ERROR) 
		response.end
	else
		CashRegName=RS1("NameDate")
		theCSRName=RS1("RealName")
		OpenningDate=RS1("OpenDate")
		OpenningAmount=cdbl(RS1("OpeningAmount"))
		cashAmountA=cdbl(RS1("cashAmountA"))
		cashAmountB=cdbl(RS1("cashAmountB"))
		chequeAmount=cdbl(RS1("chequeAmount"))
		chequeQtty=cdbl(RS1("chequeQtty"))

		Set RS1=nothing
	end if

	if OpenningDate > session("OpenGLEndDate") then
'		response.write session("OpenGLEndDate")
		response.write "<br>" 
		call showAlert ("’‰œÊﬁ „—»Êÿ »Â ”«· „«·Ì »⁄œ «“ «Ì‰ «” .",CONST_MSG_ERROR) 
		response.end
	end if
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	mySQL="SELECT ReceivedCheques.ChequeNo, ReceivedCheques.ChequeDate, ReceivedCheques.Description, ReceivedCheques.BankOfOrigin,  receivedCheques.Amount, CashRegisterLines.* FROM ReceivedCheques INNER JOIN Receipts ON ReceivedCheques.Receipt = Receipts.ID INNER JOIN CashRegisterLines ON Receipts.ID = CashRegisterLines.Link WHERE (CashRegisterLines.CashReg = '"& CashRegID & "') AND (CashRegisterLines.Type = 1) ORDER BY CashRegisterLines.[Date], CashRegisterLines.[Time]"
	Set RS1 = conn.execute(mySQL)
	
	Remained = 0
    Totalcredit = 0
    TotalDebit = 0
	tempCounter = 0

	if true then 'Not RS1.EOF OR OpenningAmount<>0 then
%>
		<br>
		<TABLE class="CClosTable" align='center'>
		<TR>
			<TD class="CClosTableTitle" colspan=9 dir='rtl' align='center'><br><B>ê“«—‘ ’‰œÊﬁ <span dir='LTR'><%=CashRegName%></span> - <%=theCSRName%> <INPUT TYPE="hidden" Name="CashRegID" Value="<%=CashRegID%>"></B><br><br>
			</TD>
		</TR>
<%
		Remained = OpenningAmount
		Credit = OpenningAmount
		Totalcredit = OpenningAmount
		tempCounter = 0
%>		<TR class='CClosTR3'>
			<td colspan='6' align='right'>„ﬁœ«— «Ê·ÌÂ ’‰œÊﬁ:</td>
			<td dir='LTR' align='right'><%=Separate(Credit)%>&nbsp;</td>
		</TR>
		<TR>
			<TD colspan='5' class="CClosTableFooter" align="right">:„ﬁœ«— ‰ﬁœ</td>
			<TD colspan="2" class="CClosTableFooter" align='left'>
				<span style="padding:10px;color:red;">«·›: </span>
				<span id='cashAmountA'><%=cashAmountA%></span>
				<span style="padding:10px;color:red;">»: </span> 
				<span id='cashAmountB'><%=cashAmountB%></span>
			</TD>
		</TR>
		<TR>
			<TD colspan='6' class="CClosTableFooter" style='direction:RTL;text-align=left;'>„ﬁœ«— çﬂ (<span id='chequeQtty'><%=Separate(chequeQtty)%></span> ›ﬁ—Â) :</td>
			<TD class="CClosTableFooter" align='right'><%=Separate(chequeAmount)%></TD>
		</TR>
		<TR>

			<TD colspan='6' class="CClosTableFooter">: Ã„⁄</td>
			<TD class="CClosTableFooter" align='right'><%=Separate(cashAmountA + cashAmountB + chequeAmount)%></td>
		</TR>
		<TR class='CClosTR3'>
			<td colspan='7' align='right'>·Ì”  çﬂ Â«:</td>
		</TR>
		<TR class="CClosTableHeader">
			<TD>#</TD>
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
		if Remained>=0 then
			remainedColor="green"
		else
			remainedColor="red"
		end if
%>
		<TR class='CClosTR3'>
			<td colspan='7' align='right'>«ÿ·«⁄«  ·«“„ »—«Ì »” ‰ ’‰œÊﬁ:</td>
		</TR>
			<TR>
				<TD colspan='7' style='border:2 dashed blue;'>
					<table width='100%' class='CClosInpTable'>
					<FORM METHOD=POST ACTION="?act=submitClose" onsubmit="return formValidation();">
					<tr>
						<td>œ— „Ã„Ê⁄ çﬁœ— ﬂ„/“Ì«œ ¬„œø</td>
						<td colspan='3'><INPUT TYPE="text" NAME="ShortOverAmount"  onKeyPress="return maskNumber(this);" onblur="setPrice(this);"> &nbsp; —Ì«· &nbsp;
						<SELECT NAME="ShortOver" onchange="setPrice(this);">
							<OPTION id='opt0' VALUE="">-------</OPTION>
							<OPTION id='opt1' VALUE="1">ﬂ„</OPTION>
							<OPTION id='opt2' VALUE="3">“Ì«œ</OPTION>
						</SELECT></td>
					</tr>
					<tr>
						<td>çﬂ Â« »Â çÂ „Õ·Ì „‰ ﬁ· „Ì ‘Êœø</td>
						<td><SELECT NAME="ChequesNewBanker" onchange="document.all.ChequesNewBankerPass.focus();">
							<OPTION VALUE="">----------------</OPTION>
<%						mySQL="SELECT Bankers.ID, Bankers.Name, Users.RealName FROM Bankers INNER JOIN Users ON Bankers.Responsible = Users.ID WHERE (IsBankAccount = 0) ORDER BY Bankers.Name"
						Set RS1 = conn.execute(mySQL)
						while not RS1.eof
%>							<OPTION VALUE="<%=RS1("ID")%>"><%=RS1("Name")%> (<%=RS1("RealName")%>)</OPTION>
<%							RS1.MoveNext
						wend
%>						</SELECT></td>
						<td>ﬂ·„Â ⁄»Ê—:</td>
						<td><INPUT TYPE="password" NAME="ChequesNewBankerPass" onkeyDown="return myKeyDownHandler();" onKeyPress="return myKeyPressHandler();"></td>
					</tr>
					<tr>
						<td><INPUT checked TYPE="checkbox" style='width:20px;border:0;' NAME="AutoNewCashReg" onclick="checkAutoGen(this);"> « Ê„« Ìﬂ ’‰œÊﬁ ÃœÌœ «ÌÃ«œ ‘Êœ</td>
						<td><span id="Sp1" > «—ÌŒ ’‰œÊﬁ: <INPUT TYPE="text" style="text-align:left;width:80px;" Name="NameDate" value="<%=shamsiToday()%>"  onblur="acceptDate(this)"></span></td>
						<td><span id="Sp2">„ﬁœ«— «Ê·ÌÂ: </span></td>
						<td><span id="Sp3"><INPUT TYPE="text" NAME="OpeningAmount" readonly Value='0'> &nbsp; —Ì«· &nbsp;</span></td>
					</tr>
					<tr>
						<td colspan="4"> ﬂ· „ﬁœ«— ‰ﬁœ „ÊÃÊœ «·›: &nbsp;
						<INPUT readonly tabIndex="98" TYPE="text" style="text-align:left;width:100px; background: transparent;font-weight:bold;" Name="totalRemainedCashA" value="<%=cashAmountA%>"> &nbsp; —Ì«· &nbsp;
						 ﬂ· „ﬁœ«— ‰ﬁœ „ÊÃÊœ »: &nbsp;							
						<INPUT readonly tabIndex="99" TYPE="text" style="text-align:left;width:100px; background: transparent;font-weight:bold;" Name="totalRemainedCashB" value="<%=cashAmountB%>"> &nbsp; —Ì«· &nbsp;
						</td>
					</tr>
					<tr>
						<td>ÅÊ· Â«Ì ‰ﬁœ »Â çÂ ﬂ”Ì  ÕÊÌ· „Ì ‘Êœø</td>
						<td><SELECT NAME="CashAcceptor" onchange="document.all.CashAcceptorPass.focus();">
							<OPTION VALUE="">---------------</OPTION>
<%						mySQL="SELECT * FROM Users WHERE Display=1 ORDER BY RealName"
						Set RS1 = conn.execute(mySQL)
						while not RS1.eof
%>							<OPTION VALUE="<%=RS1("ID")%>"><%=RS1("RealName")%></OPTION>
<%							RS1.MoveNext
						wend

%>						</SELECT></td>
						<td>ﬂ·„Â ⁄»Ê—:</td>
						<td><INPUT TYPE="password" NAME="CashAcceptorPass" onkeyDown="return myKeyDownHandler();" onKeyPress="return myKeyPressHandler();"></td>
					</tr>
					<tr>
						<td colspan="4" align="center">
							<INPUT TYPE="hidden" Name="TempCashAmountA" Value="<%=cashAmountA%>">
							<INPUT TYPE="hidden" Name="TempCashAmountB" Value="<%=cashAmountB%>">
							<INPUT TYPE="hidden" Name="TempChequeAmount" Value="<%=ChequeAmount%>">
							<INPUT TYPE="hidden" Name="TempChequeQtty" Value="<%=ChequeQtty%>">

							<INPUT TYPE="hidden" Name="CashRegID" Value="<%=CashRegID%>">
							<INPUT TYPE="submit" VALUE=" «ÌÌœ" style="width:50px;">
						</td>
					</tr>
					</FORM>
					</table>
				</TD>
			</TR>
		</TABLE>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		document.all.ShortOverAmount.focus();
		//-->
		</SCRIPT>
		<br>
<%	end if
elseif request("act")="submitClose" then

	CashRegID=request.form("CashRegID")
	if not isnumeric(CashRegID) then
		response.redirect "?errmsg=" & Server.URLEncode("’‰œÊﬁ œ—” Ì «‰ Œ«» ‰‘œÂ")
	end if
	mySQL="SELECT * FROM CashRegisters WHERE (IsOpen=1) AND (ID='" & CashRegID & "')"
	Set RS1=conn.execute(mySQL)
	if RS1.eof then
		Conn.close
		response.redirect "?errmsg=" & Server.URLEncode("’‰œÊﬁ œ—” Ì «‰ Œ«» ‰‘œÂ")
	end if

	CashAmountA=cdbl(RS1("cashAmountA"))
	CashAmountB=cdbl(RS1("cashAmountB"))
	ChequeAmount=cdbl(RS1("chequeAmount"))
	ChequeQtty=cdbl(RS1("chequeQtty"))
	nameDate = RS1("NameDate")

	if NOT (CashAmountA = cdbl(request.form("TempCashAmountA")) AND CashAmountB = cdbl(request.form("TempCashAmountB")) AND ChequeAmount = cdbl(request.form("TempChequeAmount")) AND ChequeQtty = cdbl(request.form("TempChequeQtty")) ) then
		Conn.close
		response.redirect "?act=Report&CashReg=" & CashRegID & "&errmsg=" & Server.URLEncode("«ÿ·«⁄«  œ— “„«‰ »” ‰ ’‰œÊﬁ  €ÌÌ— ﬂ—œÂ «” .<br><br> ·ÿ›« œÊ»«—Â çﬂ ﬂ‰Ìœ.")
	end if

	theCashier=RS1("Cashier")
	theBanker=RS1("Banker")
	RS1.close
	Set RS1=nothing

	closeDate = shamsiToday()	
	closeTime = currentTime10()

	if request.form("AutoNewCashReg")="on" then
		AutoNewCashReg = 1
	else
		AutoNewCashReg = 0
	end if

	CashAcceptor = 0
	ChequesNewBanker = -1
	RemainedCashMemo = 0
	ShortOverMemo = 0

	totalRemainedCashA=text2value(request.form("totalRemainedCashA"))
	totalRemainedCashB=text2value(request.form("totalRemainedCashB"))
	if CDbl(totalRemainedCashA) < 0 or CDbl(totalRemainedCashB) < 0 then
		response.redirect "?act=Report&CashReg=" & CashRegID &"&errmsg=" & Server.URLEncode("çÌ øøøøø<br>„‰›Ì ‘œ ’‰œÊﬁ ø!ø!ø!ø!ø!ø Â« Â« !<br>‘ÊŒÌ ‰ﬂ‰ »«»« »—ê—œ œ—” ‘ ﬂ‰.")
	elseif totalRemainedCashA > 0 or totalRemainedCashB then
		CashAcceptor = clng(request.form("CashAcceptor")) 
		mySQL="SELECT Users.Password, Users.Account FROM Users INNER JOIN Accounts ON Users.Account = Accounts.ID WHERE Users.ID='" & CashAcceptor & "' AND Users.Password='" & sqlSafe(request.form("CashAcceptorPass")) & "' "

		Set RS1=conn.execute(mySQL)
		If RS1.EOF Then
			conn.close
			response.redirect "?act=Report&CashReg=" & CashRegID &"&errmsg=" & Server.URLEncode("Õ”«» êÌ—‰œÂ ÅÊ· Â« ‰«œ—”  «” ")
		elseif RS1("Password")<>request.form("CashAcceptorPass") then
			conn.close
			response.redirect "?act=Report&CashReg=" & CashRegID &"&errmsg=" & Server.URLEncode("ﬂ·„Â ⁄»Ê— êÌ—‰œÂ ÅÊ· Â« ‰«œ—”  «” ")
		end if
		cashAcceptorAccount=RS1("Account")
		RS1.close
		Set RS1=Nothing
	end if

	ShortOverAmount=text2value(request.form("ShortOverAmount"))
	if ShortOverAmount < 0 then
		response.redirect "?act=Report&CashReg=" & CashRegID &"&errmsg=" & Server.URLEncode("Œÿ« ... „ﬁœ«— ﬂ„ / “Ì«œ „‰›Ì «” .")
	elseif ShortOverAmount > 0 then
		if request.form("ShortOver")=3 then 
			'3 means Over
			isCredit=1
			ShortOverDescription ="«„«‰  «÷«›Â ‰ﬁœ ’‰œÊﬁ » " & nameDate
		elseif request.form("ShortOver")=1 then
			'1 means Short
			isCredit=0
			ShortOverDescription ="«„«‰  ﬂ”—Ì ‰ﬁœ ’‰œÊﬁ » " & nameDate
		else
			conn.close
			response.redirect "?act=Report&CashReg=" & CashRegID &"&errmsg=" & Server.URLEncode("Œÿ« ... ﬂ„ Ì« “Ì«œø")
		end if
	end if

	if ChequeQtty > 0 then
		ChequesNewBanker = sqlSafe(request.form("ChequesNewBanker")) 
		mySQL="SELECT Users.Password, Users.Account  FROM Bankers INNER JOIN Users ON Bankers.Responsible = Users.ID WHERE (Bankers.ID = '"& ChequesNewBanker & "') AND (Users.Password = N'"& sqlSafe(request.form("ChequesNewBankerPass")) & "')"
		Set RS1=conn.execute(mySQL)
		If RS1.EOF Then
			conn.close
			response.redirect "?act=Report&CashReg=" & CashRegID &"&errmsg=" & Server.URLEncode("ﬂ·„Â ⁄»Ê— êÌ—‰œÂ çﬂ Â« ‰«œ—”  «” ")
		elseif RS1("Password")<>request.form("ChequesNewBankerPass") then
			conn.close
			response.redirect "?act=Report&CashReg=" & CashRegID &"&errmsg=" & Server.URLEncode("ﬂ·„Â ⁄»Ê— êÌ—‰œÂ çﬂ Â« ‰«œ—”  «” ")
		end if
	end if

	'*************************************************************************
	'*****			Handing the Remained Cash to an Account A
	'*************************************************************************
	if totalRemainedCashA > 0 then
			
		' AOMemoType = 1 means (miscellaneous)
		mySQL="INSERT INTO AOMemo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& closeDate & "' , "& session("ID") & ", 1, "& cashAcceptorAccount & ", 0, "& totalRemainedCashA & ", N'«„«‰  „ÊÃÊœÌ ‰ﬁœ ’‰œÊﬁ «·› " & nameDate & "');SELECT @@Identity AS NewMemoID"
		set RS1 = Conn.execute(mySQL).NextRecordSet
		cashMemoA = RS1 ("NewMemoID")
		RemainedCashMemo = cashMemoA
		RS1.close

		'**************************** Creating AOItem for Memo  ****************
		'*** Type = 3 means AOItem is a Memo
		'*** Reason = 5 means (Borrow) and Sys=AO 

		firstGLAccount=	"18001"	'This must be changed... GLAccount For Reason=5	(Misc. Debitors)
		GLAccountA=			"11007"		'This must be changedÖ (Cashier A)
		GLAccountB=			"11005"		'This must be changedÖ (Cashier B)

		mySQL="INSERT INTO AOItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Reason, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount) VALUES ('" &_
		GLAccountA & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& cashAcceptorAccount & "', '"& closeDate & "', 0, 5, 3, '"& cashMemoA & "', '"& totalRemainedCashA & "', N'"& closeDate & "', '"& session("ID") & "', '"& totalRemainedCashA & "')"	
		conn.Execute(mySQL)
		'***------------------------- Creating AOItem for Memo  ----------------

		'**************************** Updating Account AO Balance  ****************
		mySQL="UPDATE Accounts SET AOBalance = AOBalance - '"& totalRemainedCashA & "' WHERE (ID='"& cashAcceptorAccount & "')"
		conn.Execute(mySQL)
		'***------------------------- Updating Account AO Balance  ----------------
	end if
	
	'*************************************************************************
	'*****			Handing the Remained Cash to an Account B
	'*************************************************************************
	if  totalRemainedCashB > 0 then
		' AOMemoType = 1 means (miscellaneous)
		mySQL="INSERT INTO AOMemo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& closeDate & "' , "& session("ID") & ", 1, "& cashAcceptorAccount & ", 0, "& totalRemainedCashB & ", N'«„«‰  „ÊÃÊœÌ ‰ﬁœ ’‰œÊﬁ » " & nameDate & "');SELECT @@Identity AS NewMemoID"
		set RS1 = Conn.execute(mySQL).NextRecordSet
		cashMemoB = RS1 ("NewMemoID")
		RemainedCashMemo = cashMemoB
		RS1.close

		'**************************** Creating AOItem for Memo  ****************
		'*** Type = 3 means AOItem is a Memo
		'*** Reason = 5 means (Borrow) and Sys=AO 

		firstGLAccount=	"18001"	'This must be changed... GLAccount For Reason=5	(Misc. Debitors)
		GLAccountA=			"11007"		'This must be changedÖ (Cashier A)
		GLAccountB=			"11005"		'This must be changedÖ (Cashier B)

		mySQL="INSERT INTO AOItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Reason, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount) VALUES ('" &_
		GLAccountB & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& cashAcceptorAccount & "', '"& closeDate & "', 0, 5, 3, '"& cashMemoB & "', '"& totalRemainedCashB & "', N'"& closeDate & "', '"& session("ID") & "', '"& totalRemainedCashB & "')"	
		conn.Execute(mySQL)
		'***------------------------- Creating AOItem for Memo  ----------------

		'**************************** Updating Account AO Balance  ****************
		mySQL="UPDATE Accounts SET AOBalance = AOBalance - '"& totalRemainedCashB & "' WHERE (ID='"& cashAcceptorAccount & "')"
		conn.Execute(mySQL)
		'***------------------------- Updating Account AO Balance  ----------------
	end if
	
	'*************************************************************************
	'*****			Applying the Cash Register Short / Over Amount to Cashier's Account
	'*************************************************************************
	if ShortOverAmount > 0 then

		mySQL="SELECT Account FROM Users WHERE (ID = '"& theCashier & "')"
		set RS1=conn.execute(mySQL)
		cashierAccount=RS1("Account")
		RS1.close
		set RS1=nothing

		' AOMemoType = 6 means (Short of Cash)
		mySQL="INSERT INTO AOMemo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& closeDate & "' , "& session("ID") & ", 6, "& cashierAccount & ", '"& isCredit & "', "& ShortOverAmount & ", N'"& ShortOverDescription & "');SELECT @@Identity AS NewMemoID"

		set RS1 = Conn.execute(mySQL).NextRecordSet
		MemoID = RS1 ("NewMemoID")
		ShortOverMemo = MemoID
		RS1.close

		'**************************** Creating AOItem for Memo  ****************
		'*** Type = 3 means AOItem is a Memo
		'*** Reason = 5 means (Borrow) and Sys=AO 

		firstGLAccount=	"18001"	'This must be changed... GLAccount For Reason=5	(Misc. Debitors)

		GLAccount=		"11005"	'This must be changed... (Cashier B)

		mySQL="INSERT INTO AOItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Reason, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount) VALUES ('"&_
		GLAccount & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& cashierAccount & "', '"& closeDate & "', '"& isCredit & "', 5, 3, '"& MemoID & "', '"& ShortOverAmount & "', N'"& closeDate & "', '"& session("ID") & "', '"& ShortOverAmount & "')"	
		conn.Execute(mySQL)
		'***------------------------- Creating AOItem for Memo  ----------------

		'**************************** Updating Account AO Balance  ****************
		if isCredit=0 then
			mySQL="UPDATE Accounts SET AOBalance = AOBalance - '"& ShortOverAmount & "' WHERE (ID='"& cashierAccount & "')"
		else
			mySQL="UPDATE Accounts SET AOBalance = AOBalance + '"& ShortOverAmount & "' WHERE (ID='"& cashierAccount & "')"
		end if
			conn.Execute(mySQL)
		'***------------------------- Updating Account AO Balance  ----------------
	end if

	'*************************************************************************
	'*****		Handing Cheques to a Banker  
	'*************************************************************************
	if ChequeQtty > 0 then
		'****************************************************
		'***			UPDATING  ReceivedCheques 
		'***		Set New Banker , and New Status
		'***	Status : 1 (in open cash) --> 6 (received)
		'****************************************************
		mySQL="UPDATE ReceivedCheques SET LastBanker ='"& sqlSafe(request.form("ChequesNewBanker")) & "', LastStatus = 6 WHERE (ID IN (SELECT ReceivedCheques.ID FROM ReceivedCheques INNER JOIN Receipts ON ReceivedCheques.Receipt = Receipts.ID INNER JOIN CashRegisterLines ON Receipts.ID = CashRegisterLines.Link WHERE (CashRegisterLines.CashReg = '"& CashRegID & "') AND (CashRegisterLines.Type = 1) AND (CashRegisterLines.Voided = 0)))"
		conn.execute(mySQL)
	end if

	'*************************************************************************
	'*****		Closing the Cash Register
	'*************************************************************************
	if cashMemoA="" then cashMemoA="null"
	if cashMemoB="" then cashMemoB="null"
	mySQL="UPDATE CashRegisters SET CloseDate=N'"& closeDate & "', CloseTime=N'"& closeTime & "', ClosedBy='"& session("ID") & "', IsOpen = '0', CashAcceptor='" & CashAcceptor & "', RemainedCashMemo='" & RemainedCashMemo & "', ChequesNewBanker = '" & ChequesNewBanker & "', ShortOverAmount='"& ShortOverAmount & "', ShortOverMemo='" & ShortOverMemo & "',cashMemoA=" & cashMemoA & ",cashMemoB=" & cashMemoB & " WHERE (ID='"& CashRegID & "')"
	conn.Execute(mySQL)
	theResultReport="’‰œÊﬁ »Â ”·«„ Ì »” Â ‘œ.<br>"

	if AutoNewCashReg = 1 then
		OpeningAmount=text2value(request("OpeningAmount"))
		'*********************************************************************
		'*****		Creating New Cash Register
		'*********************************************************************
		mySQL="INSERT INTO CashRegisters (Cashier, Banker, OpenDate, NameDate, OpenedBy, IsOpen, IsApproved, OpeningAmount, CashAmountB, CashAmountA, ChequeAmount, ChequeQtty, ShortOverAmount) VALUES ('" &_ 
		theCashier & "', '"& theBanker & "', N'"& closeDate & "', N'"& request("NameDate") & "', '"& session("ID") & "', 1, 0, '"& OpeningAmount & "', '"& OpeningAmount & "', 0, 0, 0, 0)"
		conn.Execute(mySQL)
		theResultReport=theResultReport+"’‰œÊﬁ ÃœÌœ Â„ «ÌÃ«œ ‘œ."
	end if

%>
	<br>
	<TABLE width=70% align='center'>
	<TR>
		<TD align=center bgcolor=#EEFF66 style='border: dashed 1pt Green'><BR><%=theResultReport%><BR><BR></TD>
	</TR>
	</TABLE>
<%
else
	'mySQL="SELECT * FROM CashRegisters WHERE (IsOpen=1) AND (Cashier='"& session("ID") & "')"
	mySQL="SELECT CashRegisters.ID, CashRegisters.NameDate, Users.RealName FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID WHERE (CashRegisters.IsOpen=1)"
	Set RS1= conn.Execute(mySQL)
	if RS1.eof then 
	%><br><br>
		<TABLE width='70%' align='center'>
		<TR>
			<TD align=center bgcolor=#FFBBBB style='border: solid 1pt black'><BR><b>ÂÌç ’‰œÊﬁ »«“Ì ÊÃÊœ ‰œ«—œ ... <br><br>«“ œ”  „‰ ﬂ«—Ì »—«Ì ‘„« »— ‰„Ì ¬Ìœ <br>„ «”›„."</b><BR><BR></TD>
		</TR>
		</TABLE>
	<%	conn.close
		response.end
	else
	%>
	<FORM METHOD=POST ACTION="?act=Report">
		<br><br>
		&nbsp;&nbsp;’‰œÊﬁ „Ê—œ ⁄·«ﬁÂ ŒÊœ —« »—«Ì »” ‰ «‰ Œ«» ﬂ‰Ìœ: <SELECT NAME="CashReg" class="GenButton" onchange="submit();">
			<OPTION VALUE="">--------------------------</OPTION>
<%		Do 
%>			<OPTION VALUE="<%=RS1("ID")%>"><%=RS1("NameDate") & " (" & RS1("RealName")& ")"%></OPTION>
<%			Rs1.moveNext
		Loop while not RS1.eof
%>
		</SELECT> <INPUT TYPE="submit" value="ﬁ»Ê·">
	</FORM>
	<%	Set RS1=nothing
	end if
end if
conn.Close
%>
</font>
<SCRIPT LANGUAGE="JavaScript">
<!--
function setPrice(src){
	if (src.name!='ShortOver'){
		src.value=val2txt(txt2val(src.value));
	}
	if (src.name=='ShortOverAmount'){
		if (txt2val(src.value) != 0)
			document.all.opt1.selected=true;
		else
			document.all.opt0.selected=true;
	}		
	else if (src.name=='ShortOver'){
		if (txt2val(document.all.ShortOver.value)==0){
			document.all.ShortOverAmount.value=0;
		}
		else if (txt2val(document.all.ShortOverAmount.value)==0){
			document.all.opt0.selected=true;
			document.all.ShortOverAmount.focus();
		}
	}
	setTotalRemainedCash();
//	alert(src.name+'='+src.value);
}

function setTotalRemainedCash(){
	with (document.all){
		if (AutoNewCashReg.checked)
			totalRemainedCashB.value=val2txt((ShortOver.value - 2) * txt2val(ShortOverAmount.value) + txt2val(cashAmountB.innerText) - txt2val(OpeningAmount.value))
			//totalRemainedCashA.value=val2txt((ShortOver.value - 2) * txt2val(ShortOverAmount.value) + txt2val(cashAmountB.innerText) - txt2val(OpeningAmount.value))
		else
			totalRemainedCashB.value=val2txt((ShortOver.value - 2) * txt2val(ShortOverAmount.value) + txt2val(cashAmountB.innerText)
	}
}
function checkAutoGen(src){
	if (src.checked){
		document.getElementById("Sp1").style.visibility='visible';
		document.getElementById("Sp2").style.visibility='visible';
		document.getElementById("Sp3").style.visibility='visible';
		document.getElementById("Sp3").getElementsByTagName("Input")[0].focus();
	}
	else{
		document.getElementById("Sp1").style.visibility='hidden';
		document.getElementById("Sp2").style.visibility='hidden';
		document.getElementById("Sp3").style.visibility='hidden';
	}
	setTotalRemainedCash();
}

function formValidation(){
	with (document.all){
		setTotalRemainedCash();
		if (txt2val(chequeQtty.innerText) != 0)
			if (!ChequesNewBanker.value){
				alert(":·ÿ›« „‘Œ’ ﬂ‰Ìœ\n" + "çﬂ Â« »Â ﬂÃ« „‰ ﬁ· „Ì ‘Ê‰œø");
				ChequesNewBanker.focus();
				return false;
			}
			else if (!ChequesNewBankerPass.value){
				alert("ﬂ·„Â ⁄»Ê— —« Ê«—œ ﬂ‰Ìœ");
				ChequesNewBankerPass.focus();
				return false;
			}

		if (totalRemainedCashA.value != 0 || totalRemainedCashB.value != 0)
			if (!CashAcceptor.value){
				alert(":·ÿ›« „‘Œ’ ﬂ‰Ìœ\n" + "ÅÊ· Â«Ì ‰ﬁœ »Â çÂ ﬂ”Ì  ÕÊÌ· „Ì ‘Êœ‰œø");
				CashAcceptor.focus();
				return false;
			}
			else if (!CashAcceptorPass.value){
				alert("ﬂ·„Â ⁄»Ê— —« Ê«—œ ﬂ‰Ìœ");
				CashAcceptorPass.focus();
				return false;
			}
		
	}
	return true;
}
//-->
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
<!--
var tempKeyBuffer;
function myKeyDownHandler(){
	tempKeyBuffer=window.event.keyCode;
}
function myKeyPressHandler(){
//	alert (tempKeyBuffer)
	if (tempKeyBuffer>=65 && tempKeyBuffer<=90){
		window.event.keyCode=tempKeyBuffer+32;
	}
	else if(tempKeyBuffer==186){
		window.event.keyCode=59;
	}
	else if(tempKeyBuffer==188){
		window.event.keyCode=44;
	}
	else if(tempKeyBuffer==190){
		window.event.keyCode=46;
	}
	else if(tempKeyBuffer==191){
		window.event.keyCode=47;
	}
	else if(tempKeyBuffer==192){
		window.event.keyCode=96;
	}
	else if(tempKeyBuffer>=219 && tempKeyBuffer<=221){
		window.event.keyCode=tempKeyBuffer-128;
	}
	else if(tempKeyBuffer==222){
		window.event.keyCode=39;
	}
}
//-->
</SCRIPT>
<!--#include file="tah.asp" -->
