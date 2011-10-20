<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="ê“«—‘ ’‰œÊﬁ"
SubmenuItem=3
if not Auth(9 , 3) then NotAllowdToViewThisPage()

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
</STYLE>
<%
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function


mySQL="SELECT Bankers.Name AS BankerName, CashRegisters.*, Users.RealName FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID INNER JOIN Bankers ON CashRegisters.Banker = Bankers.ID WHERE (CashRegisters.IsOpen=1) AND (Cashier='"& session("ID") & "')"
Set RS1= conn.Execute(mySQL)
if RS1.eof then 
%><br><br>
	<TABLE width='70%' align='center'>
	<TR>
		<TD align=center bgcolor=#FFBBBB style='border: solid 1pt black'><BR><b>Ã‰«» <%=CSRName%> ‘„« ’‰œÊﬁ »«“ ‰œ«—Ìœ ... <br><br>«“ œ”  „‰ ﬂ«—Ì »—«Ì ‘„« »— ‰„Ì ¬Ìœ <br>„ «”›„."</b><BR><BR></TD>
	</TR>
	</TABLE>
<%	conn.close
	response.end
else
	CashRegID=			RS1("ID")
	BankerName=			RS1("BankerName")
	CashRegName=		RS1("NameDate")
	OpenningDate=		RS1("OpenDate")
	OpenningAmount=		cdbl(RS1("OpeningAmount"))
	totalCashAmountA=	Separate(RS1("cashAmountA"))
	totalCashAmountB=	Separate(RS1("cashAmountB"))
	totalChequeAmount=	Separate(RS1("chequeAmount"))
	totalChequeQtty=	RS1("chequeQtty")
	TotalAmount=		Separate(cdbl(RS1("cashAmountA")) + cdbl(RS1("cashAmountB")) + cdbl(RS1("chequeAmount")))

	Set RS1=nothing
end if

if true then ' this 'if' statement is reserved for further changes. It is not neccesary.
'response.write cashRegID
	mySQL="SELECT CashRegisterLines.ID,CashRegisterLines.isA, CashRegisterLines.CashReg, CashRegisterLines.Date, CashRegisterLines.Time, CashRegisterLines.Type, isnull(case when isnull(drv.isA,0)>0 and CashRegisterLines.type=1 then ISNULL(Receipts.CashAmount,0) end,0) AS CashAmountA, isnull(case when isnull(drv.isA,0)<1 and CashRegisterLines.type=1 then ISNULL(Receipts.CashAmount,0) end,0) AS CashAmountB, CashRegisterLines.Link, CashRegisterLines.Voided, ISNULL(Receipts.CashAmount, 0) AS CashAmount, ISNULL(Receipts.ChequeAmount, 0) AS ChequeAmount, Receipts.ChequeQtty, Payments.CashAmount AS PayAmount, Accounts_1.AccountTitle AS PayAccTitle, Accounts_2.AccountTitle AS RcpAccTitle, Accounts_1.ID AS PayAccID, Accounts_2.ID AS RcpAccID FROM CashRegisterLines LEFT OUTER JOIN Accounts AS Accounts_1 INNER JOIN Payments ON Accounts_1.ID = Payments.Account ON CashRegisterLines.Link = Payments.ID LEFT OUTER JOIN Accounts AS Accounts_2 INNER JOIN Receipts ON Accounts_2.ID = Receipts.Customer ON CashRegisterLines.Link = Receipts.ID left outer join (select Receipts.id, isnull(sum(cast(invoices.IsA as int)),-1) as isA from Receipts inner join ARItems on receipts.id=arItems.link inner join ARItemsRelations on arItems.ID=arItemsRelations.CreditARItem inner join ARItems as arItems_D on arItemsRelations.DebitARItem=arItems_D.id inner join Invoices on arItems_D.link=invoices.ID where receipts.SYS='AR' and arItems_D.type=1 group by Receipts.id) as drv on receipts.id=drv.id WHERE (CashRegisterLines.CashReg = '"& CashRegID & "') ORDER BY CashRegisterLines.ID"
'	response.write mySQL
'	response.end
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
			<TD class="RepTableTitle" colspan=10 dir='rtl' align='center'><br><B>ê“«—‘ ’‰œÊﬁ <span dir='LTR'><%=CashRegName%></span> - <%=CSRName%> </B><BR>(„Õ·: <%=BankerName%>)<br>
			</TD>
		</TR>
		<TR class="RepTableHeader">
			<TD width="15px">#</TD>
			<TD width="15px"> x </TD>
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
				<td align='center'><b>x&nbsp;<b></td>
				<td dir='LTR'><%=OpenningDate%></td>
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
				voidLink="Void.asp?act=voidReceipt&receipt="& RS1("Link")
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
				Credit=CashAmount + ChequeAmount
				Debit=0

			elseif RS1("Type")=2 then
				sourceLink="../AR/AccountReport.asp?act=showPayment&payment="& RS1("Link")
				voidLink="Void.asp?act=voidPayment&payment="& RS1("Link")
				Description = "Å—œ«Œ  ‰ﬁœ"
				AccID=RS1("PayAccID")
				AccTitle=RS1("PayAccTitle")
				Credit=0
				Debit=RS1("PayAmount")
			else
				sourceLink="javascript:void(0);"
				voidLink="javascript:void(0);"
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
				<td align='center'><b><A HREF="<%=voidLink%>" onclick="return setVoid(this);">x</A>&nbsp;<b></td>
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
				<td dir='LTR' align='right' title='<%="«·›: "&RS1("CashAmountA") & "° »: " & RS1("CashAmountB")%>'><%=Separate(Credit) %></td>
				<td dir='LTR' align='right'>
<%					if RS1("Voided")=TRUE then
%>						<div style="position:absolute;width:700;"><hr style="color:red;"></div>
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
				<TD class="RepTableFooter" colspan='7'>&nbsp;&nbsp; : Ã„⁄</td>
				<TD class="RepTableFooter" align='right'><%=Separate(totaldebit)%></td>
				<TD class="RepTableFooter" align='right'><%=Separate(totalcredit)%></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR='<%=remainedColor%>'><%=Separate(Remained)%></FONT></td>
			</TR>
		</TABLE>
		<br >
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
			<td colspan='7' align='center'>Œ·«’Â ê“«—‘ ’‰œÊﬁ <span dir='LTR'><%=CashRegName%></span> (’‰œÊﬁœ«—: <%=CSRName%>° „Õ·: <%=BankerName%>)</TD>
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
end if
conn.Close
%>
</font>
<script language="JavaScript">
<!--
function setVoid(obj)
{
	theTR = obj.parentNode.parentNode.parentNode;
	for(i=0; i<theTR.getElementsByTagName("TD").length; i++){
		theTR.getElementsByTagName("TD")[i].setAttribute("bgColor","yellow")
	}
	confirmResult = confirm(' ÊÃÂ! ÅÌ‘ «“ »«ÿ· ﬂ—œ‰ œ—Ì«›  / Å—œ«Œ  »«Ìœ —”Ìœ ﬂ«€–Ì „—»ÊÿÂ «“ „‘ —Ì Å” ê—› Â ‘Êœ\n\n «œ«„Â „Ì œÂÌœø');
	for(i=0; i<theTR.getElementsByTagName("TD").length; i++){
		theTR.getElementsByTagName("TD")[i].setAttribute("bgColor","")
	}
	return confirmResult;
}
//-->
</script>
<!--#include file="tah.asp" -->