<STYLE>
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
</STYLE>
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {padding:5;border:1pt solid gray;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTableTitle {background-color: #CCCCFF; text-align: center; font-weight:bold; height:50;}
	.RepTableHeader {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.RepTableFooter {background-color: #BBBBBB; direction: LTR; }
	.RepTR1 {background-color: #DDDDDD;}
	.RepTR2 {background-color: #FFFFFF;}
	.RepDescSpan {overflow:auto;border:none; width:250px; height:23px; font-size:7pt;}
	.RepGenInput { font-family:tahoma; font-size: 8pt; border: 1 solid black; direction:LTR; width:70px; height:19px;}
</STYLE>
<style>
	.InvTable { font-size: 9pt;}
	.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0FFF0; text-align:right;}
	.InvRowInput4 { font-family:tahoma; font-size: 9pt; border: none; background-color: #FFD3A8; direction:LTR; text-align:right;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; direction: RTL;}
	.InvHeadInput4 { font-family:tahoma; font-size: 9pt; border: none; background-color: #FF9900; text-align:center;}
	.InvGenInput { font-family:tahoma; font-size: 9pt; border: none; }
</style>
<style>
	.RcpTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; }
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: RTL;}
	.RcpMainTableTH { background-color: #C3C300;}
	.RcpMainTableTR { background-color: #CCCC88; border: 0; }
	.RcpRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.RcpRowInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.RcpHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.RcpHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.RcpHeadInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #D0E0FF; text-align:right; direction: RTL;}
	.RcpGenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: LTR;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<style>
	.MmoTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; direction: LTR;}
	.MmoMainTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; background-color: #558855; text-align:right; direction: RTL;}
	.MmoMainTableTH { background-color: #C3C300;}
	.MmoMainTableTR { background-color: #CCCC88; border: 0; }
	.MmoRowInput { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.MmoRowInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #F0F0F0; text-align:right;}
	.MmoHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.MmoHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.MmoHeadInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid black; background-color: #D0E0FF; text-align:right; direction: RTL;}
	.MmoGenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: LTR;}
</style>
<%
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

function Link2TraceQuote(QuoteNo)
	Link2TraceQuote = "<A HREF='../order/Inquiry.asp?act=show&quote="& QuoteNo & "' target='_balnk'>"& QuoteNo & "</A>"
end function

%>
<!--#include File="../include_UtilFunctions.asp"-->
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
if request("act")="search" then
%>
	<br>
	<FORM METHOD=POST ACTION="?act=select" onsubmit="if (document.all.search.value=='') return false;">
	<div dir='rtl'><B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="search">&nbsp;
		<INPUT TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>

	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.search.focus();
	//-->
	</SCRIPT>

	<br><br>
	<FORM METHOD=POST ACTION="?act=select" onsubmit="if (document.all.ItemLink.value=='') return false;">
	‰„«Ì‘ Ìﬂ 
	<SELECT NAME="ItemType" style="font-family:tahoma;font-size:9pt;width:120px;">
<%
	mySQL="SELECT * FROM AXItemTypes ORDER BY ID" 
	defaultSelectedType=2

	Set RS1 = conn.execute(mySQL)
	Do While Not RS1.EOF
		if RS1("ID")= defaultSelectedType then selected="selected" else selected=""
%>		<OPTION Value="<%=RS1("ID")%>" <%=selected%>><%=RS1("Name")%></OPTION>
<%
		RS1.MoveNext
	Loop
	RS1.Close
%>
	</SELECT> Œ«’ »« ‘„«—Â 
		<INPUT TYPE="text" NAME="ItemLink" Class="GenInput">&nbsp;
		<INPUT TYPE="submit" value="Ã” ÃÊ"><br>
	</FORM>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="select" then
	if trim(request("search")) <> "" then
		SA_TitleOrName=request("search")
		SA_Action="return selectOperations();"
		SA_SearchAgainURL="?act=search"
		SA_StepText="" '"ê«„ œÊ„ : «‰ Œ«» Õ”«»"
		SA_ActName		= "select"	
		SA_SearchBox	="search"		
%>
		<FORM METHOD=POST ACTION="?act=show">
		<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	elseif isNumeric(request("ItemType")) AND isNumeric(request("ItemLink")) then
		ItemType= cint(request("ItemType"))
		ItemLink= clng(request("ItemLink"))
		if not fixSys = "-" then sys = fixSys else sys = "AO"

		select case ItemType
		case 1
			Action="showInvoice&invoice="
		case 2
			Action="showReceipt&receipt="
		case 3
			Action="showMemo&sys="& sys & "&memo="
		case 4
			Action="showInvoice&invoice="
		case 5
			Action="showPayment&payment="
		case 6
			Action="showVoucher&voucher="
		case else
			response.redirect "?errmsg=" & Server.URLEncode("„‘Œ’«  Ê«—œ ‘œÂ €·ÿ «” .")
		end select
		response.redirect "?act=" & Action & ItemLink
	else
		response.redirect "?act=search&errmsg=" & Server.URLEncode("„‘Œ’«  Ê«—œ ‘œÂ €·ÿ «” .")
	end if
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function selectOperations(){
		var Arguments = new Array;
		notFound=true;
		for (i=0;i<document.getElementsByName("selectedCustomer").length;i++){
			if(document.getElementsByName("selectedCustomer")[i].checked){
				notFound=false;
			}
		}
		if (notFound)
			return false;
	}
	//-->
	</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="show" AND isnumeric(request("selectedCustomer")) AND request("selectedCustomer")<>"" then
	CustomerID=request("selectedCustomer")
	'sys = request("sys")
	'if sys = "" then sys = "AR"
	reason = request("Reason")
	if reason = "" then reason = 6

	mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
	else
		Sys=			RS1("Acron")
		firstGLAccount=	RS1("GLAccount")
	end if
	RS1.close

	if not fixSys = "-" then sys = fixSys

	mySQL="SELECT AccountTitle FROM Accounts WHERE (ID='"& CustomerID & "')"
	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")
	RS1.close
	
	EndDate=sqlSafe(request("EndDate"))
	StartDate=sqlSafe(request("StartDate"))
	if StartDate="" then
		StartDate=session("OpenGLStartDate") 'yearBeginnig=left(shamsiToday(),4)&"/01/01"
	end if
	if EndDate="" then
		if session("OpenGLEndDate") > shamsiToday() then
			EndDate=shamsiToday()
		else
			EndDate=session("OpenGLEndDate")
		end if
	end if

	nextYear=cint(left(StartDate,4)) + 1 
	prevYear=cint(left(StartDate,4)) - 1 

	nextStartDate=	nextYear & "/01/01" 'right(StartDate,6)
	nextEndDate=	nextYear & "/12/30" 'right(EndDate,6)
	 
	prevStartDate=	prevYear & "/01/01" 'right(StartDate,6)
	prevEndDate=	prevYear & "/12/30" 'right(EndDate,6)

	nextYear=right(nextYear,2)
	prevYear=right(prevYear,2)
%>
	<br>
	<CENTER>
<%	if fixSys = "-" then %>
		<FORM METHOD=POST ACTION="?act=show" ID="sysForm">
		<INPUT TYPE="hidden" name="selectedCustomer" value="<%=CustomerID%>">
		<INPUT TYPE="hidden" name="StartDate" value="<%=StartDate%>">
		<INPUT TYPE="hidden" name="EndDate" value="<%=EndDate%>">

		<TABLE align='center' cellpadding='5'>
		<TR>
			<TD>”Ì” „: </TD>
<%
			mySQL="SELECT * FROM AXItemReasons WHERE Display=1 ORDER BY ID"
			set RS1=conn.execute(mySQL)
			while not RS1.eof
%>				<TD><INPUT TYPE="radio" NAME="Reason" value="<%=RS1("ID")%>" <% if cint(reason)=cint(RS1("ID")) then response.write "checked "%> onclick="document.getElementById('sysForm').submit()"><%=RS1("Name")%><br></TD>
<%				RS1.movenext
			wend
%>		
		</TR>
		</TABLE>
		</FORM>
<% 
	else 
		response.write "<br><br>" 
		sys = fixSys 
	end if
%>
	[ ê“«—‘ Õ”«» ] &nbsp; 
	<%
	if SYS="AR" then
%>
		<A HREF="?act=showDebit&selectedCustomer=<%=CustomerID%>&sys=AR">[ê“«—‘ „«‰œÂ »œÂÌù] </A>
<%
	elseif SYS="AP" then
%>
		<A HREF="?act=showCredit&selectedCustomer=<%=CustomerID%>&sys=AP">[ê“«—‘ „«‰œÂ »” «‰ﬂ«—Ì] </A>
<%
	end if
%>
	<BR><BR>
	</CENTER>
<%
	' Changed By Kid 82/11/05 for using stored procedure 
	mySQL="EXEC proc_"& sys & "_StatementReport_Partial '"& CustomerID & "', '"& StartDate & "', '"& EndDate & "'" ', " & reason & "" 

	Set RS1 = conn.execute(mySQL)
	
	Remained = 0
	Totalcredit = 0
	TotalDebit = 0
	tempCounter = -1
	if Not (RS1.EOF) then
%>
		<TABLE class="RepTable" width='90%' align='center'>
		<TR>
			<TD colspan=8 dir='rtl' align='center'>
			</TD>
		</TR>
		<TR>
			<TD class="RepTableTitle" colspan=8 dir='rtl' align='center'>
				<br>
				<FORM METHOD=POST ACTION="?act=show&sys=<%=sys%>&selectedCustomer=<%=CustomerID%>" ID="dateForm">
				ê“«—‘ Õ”«» <A target="_blank" HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=CustomerID%>"><%=customerName%> [<%=CustomerID%>]</A><br><br>
				«“  «—ÌŒ <INPUT class="RepGenInput" TYPE="text" NAME="StartDate" Value="<%=StartDate%>" OnBlur="return acceptDate(this);">  « <INPUT class="RepGenInput" TYPE="text" NAME="EndDate" Value="<%=EndDate%>"OnBlur="return acceptDate(this);"> <INPUT Class="GenButton" TYPE="button" Value=" ‰„«Ì‘ "onclick="if(acceptDate(document.getElementsByName('StartDate')[0]) && acceptDate(document.getElementsByName('EndDate')[0])) document.getElementById('dateForm').submit()"> &nbsp; 
				<%if (sys="AR" AND Auth(6 , "J")) OR (sys="AP" AND Auth(7 , "A")) OR (sys="AO" AND Auth("B" , 7 )) then%>
					<% 	ReportLogRow = PrepareReport (sys&"Statement.rpt", "Account"&chr(1)&"StartDate"&chr(1)&"EndDate", CustomerID&chr(1)&StartDate&chr(1)&EndDate, "/beta/dialog_printManager.asp?act=Fin") %>
					<INPUT Class="GenButton" style="border:1 solid green;" TYPE="button" value=" ç«Å " onclick="printThisReport(this,<%=ReportLogRow%>);">
				<%end if%>
				<BR><BR>
				<INPUT Class="GenButton" TYPE="button" Value=" <- <%=prevYear%> " onclick="window.location='?act=show&reason=<%=reason%>&sys=<%=sys%>&selectedCustomer=<%=CustomerID%>&startDate=<%=prevStartDate%>&endDate=<%=prevEndDate%>';">
				<INPUT Class="GenButton" TYPE="button" Value=" <%=nextYear%> -> " onclick="window.location='?act=show&reason=<%=reason%>&sys=<%=sys%>&selectedCustomer=<%=CustomerID%>&startDate=<%=nextStartDate%>&endDate=<%=nextEndDate%>';">
				</FORM>

			</TD>
		</TR>
		<TR class="RepTableHeader">
			<TD>#</TD>
			<TD> «—ÌŒ</TD>
			<TD width=70>⁄ÿ›</TD>
			<TD> Ê÷ÌÕ« </TD>
			<TD>»œÂﬂ«—</TD>
			<TD>»” «‰ﬂ«—</TD>
			<TD width=70>„«‰œÂ</TD>
		</TR>
<%
		While Not (RS1.EOF)
			tempCounter=tempCounter+1
			
			Select Case RS1("Type")
			Case 1 :
				'================================== Type = 1
				'================================== 
				sourceLink="<a href='?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— " & RS1("Link") & "</a>"
				Description = RS1("Description")

				if RS1("Orders")<>"" then
					tempWriteAnd=""
					Description = Description & "(”›«—‘ "
					orders = split (RS1("orders"),", ")
					for i=0 to ubound(orders)
						Description = Description & tempWriteAnd & Link2Trace(orders(i))
						tempWriteAnd=" Ê "
					next
					Description = Description & ")"
				end if
			
			Case 2 :
				'================================== Type = 2
				'================================== 
				sourceLink="<a href='?act=showReceipt&receipt="& RS1("Link") & "' target='_blank'>" & "œ—Ì«›  " & "</a>"

			Case 3 :
				'================================== Type = 3
				'================================== 
				sourceLink="<a href='?act=showMemo&sys="& sys & "&memo="& RS1("Link") & "' target='_blank'>" & "«⁄·«„ÌÂ " & "</a>"

			Case 4 :
				'================================== Type = 4
				'================================== 
'				sourceLink="<a href='?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "»—ê‘  «“ ›—Ê‘" & RS1("Link") & "</a>"
'				Description = "»«»  »—ê‘  «“ ›—Ê‘ ‘„«—Â "& RS1("InvNo")
				sourceLink="<a href='?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "»—ê‘  «“ ›—Ê‘ " & RS1("Link") & "</a>"
				Description = RS1("Description")

				if RS1("Orders")<>"" then
					tempWriteAnd=""
					Description = Description & "(”›«—‘ "
					orders = split (RS1("orders"),", ")
					for i=0 to ubound(orders)
						Description = Description & tempWriteAnd & Link2Trace(orders(i))
						tempWriteAnd=" Ê "
					next
					Description = Description & ")"
				end if
			
			Case 5 :
				'================================== Type = 5
				'================================== 
				sourceLink="<a href='?act=showPayment&payment="& RS1("Link") & "' target='_blank'>" & "Å—œ«Œ  </a>"
				Description = RS1("Description")

			Case 6 :
				'================================== Type = 6
				'================================== 
				sourceLink="<a href='?act=showVoucher&voucher="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— Œ—Ìœ </a>"

			Case 0 :
				'================================== Type = 6
				'================================== 
				sourceLink="„«‰œÂ ﬁ»·"

			Case Else:
				'================================== Unknown
				'================================== 
				sourceLink="<a href='javascript:void(0);'>" & "‰«‘‰«”" & "</a>"
			End Select
			if not isnull(RS1("Description")) then 
				Description = replace(RS1("Description"),chr(13),"<br>")
				Description = replace(Description,"/",".")
			else
				Description = ""
			end if

			if RS1("IsCredit")=True then
				Credit=cdbl(RS1("AmountOriginal"))
				Debit=0
			else
				Credit=0
				Debit=cdbl(RS1("AmountOriginal"))
			end if

			if RS1("Voided") then 

%>			<TR bgcolor=#FFEEEE style='color:#999999'>
				<td> # <%tempCounter = tempCounter -1 %></td>
				<td dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td><%=sourceLink%></td>
				<td><span class='RepDescSpan'><%=Description%></span></td>
				<td dir='LTR' align='right'><%=Separate(Debit)%></td>
				<td dir='LTR' align='right'><%=Separate(Credit) %></td>
				<td dir='LTR' align='right'>&nbsp;</td>
			</TR>
<%
			else
				TotalDebit = TotalDebit + cdbl(Debit)
				Totalcredit = Totalcredit + cdbl(Credit)
				Remained = Remained + Cdbl(Credit) - Cdbl(Debit)
%>			<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
				<td><%=tempCounter %></td>
				<td dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td><%=sourceLink%></td>
				<td><span class='RepDescSpan'><%=Description%></span></td>
				<td dir='LTR' align='right'><%=Separate(Debit)%></td>
				<td dir='LTR' align='right'><%=Separate(Credit) %></td>
				<td dir='LTR' align='right'><%=Separate(remained)%></td>
			</TR>
<%
			end if
			RS1.MoveNext
		Wend
		if Remained>=0 then
			remainedColor="green"
		else
			remainedColor="red"
		end if
%>
			<TR>
				<TD class="RepTableFooter" colspan='4'>&nbsp;&nbsp; : Ã„⁄</span></td>
				<TD class="RepTableFooter" align='right'><%=Separate(totaldebit)%></td>
				<TD class="RepTableFooter" align='right'><%=Separate(totalcredit)%></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR='<%=remainedColor%>'><%=Separate(Remained)%></FONT></td>
			</TR>
		</TABLE>
		<br>
<%	end if

'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showCredit" AND isnumeric(request("selectedCustomer")) then

	mySQL="SELECT AccountTitle FROM Accounts WHERE (ID='"& request("selectedCustomer") & "')"
	CustomerID=request("selectedCustomer")
	'sys = request("sys")
	'if sys = "" then sys = "AP"

	reason = request("Reason")
	if reason = "" then reason = 2

	mySQL2="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
	Set RS2=Conn.execute(mySQL2)
	if RS2.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
	else
		Sys=			RS2("Acron")
		firstGLAccount=	RS2("GLAccount")
	end if
	RS2.close

	if not fixSys = "-" then sys = fixSys

	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")
	RS1.close

	mySQL="SELECT "& sys & "Items.EffectiveDate, "& sys & "Items.Voided, "& sys & "Items.CreatedBy, "& sys & "Items.IsCredit, "& sys & "Items.Type, "& sys & "Items.Link, "& sys & "Items.AmountOriginal, "& sys & "Items.RemainedAmount, "& sys & "Memo.Description, Receipts.Number AS RcpNo, ISNULL(Receipts.CashAmount, 0) AS RcvCash, ISNULL(Receipts.ChequeAmount, 0) AS RcvCheq, Vouchers.Title AS VouchTitle, ISNULL(Payments.CashAmount, 0) AS PaidCash, ISNULL(Payments.ChequeAmount, 0) AS PaidCheq FROM Payments RIGHT OUTER JOIN "& sys & "Items ON Payments.ID = "& sys & "Items.Link LEFT OUTER JOIN "& sys & "Memo ON "& sys & "Items.Link = "& sys & "Memo.ID LEFT OUTER JOIN Receipts ON "& sys & "Items.Link = Receipts.ID LEFT OUTER JOIN Vouchers ON "& sys & "Items.Link = Vouchers.ID WHERE ("& sys & "Items.reason = '"& reason & "') AND ("& sys & "Items.Account = '"& CustomerID & "') AND ("& sys & "Items.IsCredit=1) ORDER BY "& sys & "Items.EffectiveDate, "& sys & "Items.ID" 

	Set RS1 = conn.execute(mySQL)

	%>

		<br>
		<CENTER>
			<% if fixSys = "-" then %>
			<FORM METHOD=POST ACTION="?act=showDebit" Name="sysForm">
			<INPUT TYPE="hidden" name="selectedCustomer" value="<%=CustomerID%>">

			<TABLE align='center' cellpadding='5'>
			<TR>
				<TD>”Ì” „: </TD>
			<%
				mySQL2="SELECT * FROM AXItemReasons WHERE Display=1 ORDER BY ID"
				set RS2=conn.execute(mySQL2)
				while not RS2.eof
			%>				<TD><INPUT TYPE="radio" NAME="Reason" value="<%=RS2("ID")%>" <% if cint(reason)=cint(RS2("ID")) then response.write "checked "%> onclick="document.getElementById('sysForm').submit()"><%=RS2("Name")%><br></TD>
			<%				RS2.movenext
				wend
			%>		
			</TR>
			</TABLE>


			</FORM>
			<% else 
				response.write "<br><br>" 
				sys = fixSys 
			end if
			%>
			<A HREF="?act=show&selectedCustomer=<%=CustomerID%>&sys=<%=SYS%>">[ ê“«—‘ Õ”«» ]</A> &nbsp; [ê“«—‘ „«‰œÂ »” «‰ﬂ«—Ì]
			<BR><BR>
		</CENTER>

	<%
	tempCounter = 0
	if Not (RS1.EOF) then
%>
		<TABLE class="RepTable" width='90%' align='center'>
		<TR>
			<TD colspan=8 dir='rtl' align='center'>
			</TD>
		</TR>
		<TR>
			<TD class="RepTableTitle" colspan=8 dir='rtl' align='center'><br>
				<B><A target="_blank" HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=CustomerID%>">„—»Êÿ »Â <%=customerName%> (<%=CustomerID%>)</A></B><br><br>
			</TD>
		</TR>
		<TR class="RepTableHeader">
			<TD>#</TD>
			<TD> «—ÌŒ</TD>
			<TD width=70>⁄ÿ›</TD>
			<TD> Ê÷ÌÕ« </TD>
			<TD width=70>„»·€</TD>
			<TD width=70>„«‰œÂ</TD>
		</TR>
<%
		TotalCredit=0
		While Not (RS1.EOF)
			tempCounter=tempCounter+1
			
			'================================== Type = 2
			'================================== 
			if RS1("Type")=2 then
				sourceLink="<a href='?act=showReceipt&receipt="& RS1("Link") & "' target='_blank'>" & "œ—Ì«›  " & "</a>"
				Description = "œ—Ì«›  "
				myAND=""
				CashAmount=		cdbl(RS1("RcvCash"))
				ChequeAmount=	cdbl(RS1("RcvCheq"))
				if CashAmount<>0 then
					Description=Description & "‰ﬁœ "
					myAND="Ê "
				end if 
				if ChequeAmount<>0 then
					Description=Description & myAND & "çﬂ "
				end if 
				if clng(RS1("RcpNo"))<>0 then
					Description=Description & "ÿÌ —”Ìœ ‘„«—Â "& RS1("RcpNo")
				end if
			
			'================================== Type = 3
			'================================== 
			elseif RS1("Type")=3 then
				sourceLink="<a href='?act=showMemo&sys="& sys & "&memo="& RS1("Link") & "' target='_blank'>" & "«⁄·«„ÌÂ " & "</a>"
				Description = "<span style='overflow:auto;border:none; width:250px; height:23px; font-size:7pt;'>" & RS1("Description") & "</span>" 'RS1("Description")

			'================================== Type = 5
			'================================== 
			elseif RS1("Type")=5 then
				sourceLink="<a href='?act=showPayment&payment="& RS1("Link") & "' target='_blank'>" & "Å—œ«Œ  </a>"
				Description = "Å—œ«Œ  "
				myAND=""
				CashAmount=		cdbl(RS1("PaidCash"))
				ChequeAmount=	cdbl(RS1("PaidCheq"))
				if CashAmount<>0 then
					Description=Description & "‰ﬁœ "
					myAND="Ê "
				end if 
				if ChequeAmount<>0 then
					Description=Description & myAND & "çﬂ "
				end if 

			'================================== Type = 6
			'================================== 
			elseif RS1("Type")=6 then
				sourceLink="<a href='?act=showVoucher&voucher="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— Œ—Ìœ </a>"
				Description = replace(RS1("VouchTitle") , "/", ".")

			'================================== Unknown
			'================================== 
			else
				sourceLink="<a href='javascript:void(0);'>" & "‰«‘‰«”" & "</a>"
				Description = " «—ÌŒ : <span dir='LTR'>(" & RS1("EffectiveDate") & ")</span>"
			end if

'			Remained = -1 * cdbl(RS1("RemainedAmount"))
'			Total = -1 * cdbl(RS1("AmountOriginal"))
'			Changed by kid 820824

			CreditRemained = cdbl(RS1("RemainedAmount"))
			Credit = cdbl(RS1("AmountOriginal"))


			if RS1("Voided") then 

%>			<TR bgcolor=#FFEEEE style='color:#999999'>
				<td> # <%tempCounter = tempCounter -1 %></td>
				<td dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td><%=sourceLink%></td>
				<td><%=Description%></td>
				<td dir='LTR' align='right'><%=Separate(Credit) %></td>
				<td dir='LTR' align='right'>&nbsp;</td>
			</TR>
<%
			else
				TotalCredit = TotalCredit + Credit
%>			<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
				<td><%=tempCounter %></td>
				<td dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td><%=sourceLink%></td>
				<td><%=Description%></td>
				<td dir='LTR' align='right'><%=Separate(Credit) %></td>
				<td dir='LTR' align='right'><%=Separate(CreditRemained)%></td>
			</TR>
<%
			end if
			RS1.MoveNext
		Wend
%>
			<TR>
				<TD class="RepTableFooter" colspan='4'>&nbsp;&nbsp; : Ã„⁄ »” «‰ﬂ«—</span></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="Green"><%=Separate(TotalCredit)%></FONT></td>
				<TD class="RepTableFooter" >&nbsp;</td>
			</TR>
<%		mySQL="SELECT ISNULL(SUM(AmountOriginal),0) AS totalDebit, ISNULL(SUM(RemainedAmount),0) AS totalCredRemained FROM "& sys & "Items WHERE (Account = '"& CustomerID & "') AND (IsCredit = 0) AND (Voided = 0) GROUP BY Account"
		Set RS1=Conn.execute(mySQL)
		if RS1.eof then
			TotalDebit = 0
'			TotalDebitRemained = 0
		else
			TotalDebit = cdbl(RS1("totalDebit"))
'			TotalDebitRemained = cdbl(RS1("totalDebitRemained"))
		end if
		
		Summ= TotalCredit - TotalDebit
		SummText="&nbsp;"
		if Summ < 0 then
			Summ= -1 * Summ
			SummColor="Red"
			SummText="(»œÂﬂ«—)"
		elseif Summ > 0 then
			SummColor="Green"
			SummText="(»” «‰ﬂ«—)"
		end if
%>
			<TR>
				<TD class="RepTableFooter" colspan='4'>&nbsp;&nbsp; : Ã„⁄ »œÂﬂ«— </span></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="Red"><%=Separate(TotalDebit)%></FONT></td>
				<TD class="RepTableFooter" >&nbsp;</td>
			</TR>
			<TR>
				<TD class="RepTableFooter" colspan='4'>&nbsp;&nbsp; : Ã„⁄ ﬂ·</span></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="<%=SummColor%>"><%=Separate(Summ)%></FONT></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="<%=SummColor%>"><%=SummText%></FONT></td>
			</TR>
		</TABLE>
		<br>
<%	end if

'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showDebit" AND isnumeric(request("selectedCustomer")) then
	mySQL="SELECT AccountTitle FROM Accounts WHERE (ID='"& request("selectedCustomer") & "')"
	CustomerID=request("selectedCustomer")
	'sys = request("sys")
	'if sys = "" then sys = "AR"

	reason = request("Reason")
	if reason = "" then reason = 1

	mySQL2="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
	Set RS2=Conn.execute(mySQL2)
	if RS2.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
	else
		Sys=			RS2("Acron")
		firstGLAccount=	RS2("GLAccount")
	end if
	RS2.close

	if not fixSys = "-" then sys = fixSys

	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")
	RS1.close

	mySQL="SELECT "& sys & "Items.EffectiveDate, "& sys & "Items.Voided, "& sys & "Items.CreatedBy, "& sys & "Items.IsCredit, "& sys & "Items.Type, "& sys & "Items.Link, "& sys & "Items.AmountOriginal, "& sys & "Items.RemainedAmount, "& sys & "Memo.Description, Receipts.Number AS RcpNo, ISNULL(Receipts.CashAmount, 0) AS RcvCash, ISNULL(Receipts.ChequeAmount, 0) AS RcvCheq, Invoices.IsA, Invoices.Number AS InvNo, ISNULL(Payments.CashAmount, 0) AS PaidCash, ISNULL(Payments.ChequeAmount, 0) AS PaidCheq FROM Payments RIGHT OUTER JOIN "& sys & "Items ON Payments.ID = "& sys & "Items.Link LEFT OUTER JOIN "& sys & "Memo ON "& sys & "Items.Link = "& sys & "Memo.ID LEFT OUTER JOIN Receipts ON "& sys & "Items.Link = Receipts.ID LEFT OUTER JOIN Invoices ON "& sys & "Items.Link = Invoices.ID WHERE ("& sys & "Items.reason = '"& reason & "') AND ("& sys & "Items.Account = '"& CustomerID & "') AND ("& sys & "Items.IsCredit=0) ORDER BY "& sys & "Items.EffectiveDate, "& sys & "Items.ID" 

'response.write "<br>" & mySQL
'response.end
	Set RS1 = conn.execute(mySQL)

	%>

		<br>
		<CENTER>
			<% if fixSys = "-" then %>
			<FORM METHOD=POST ACTION="?act=showDebit" Name="sysForm">
			<INPUT TYPE="hidden" name="selectedCustomer" value="<%=CustomerID%>">

			<TABLE align='center' cellpadding='5'>
			<TR>
				<TD>”Ì” „: </TD>
			<%
				mySQL2="SELECT * FROM AXItemReasons WHERE Display=1 ORDER BY ID"
				set RS2=conn.execute(mySQL2)
				while not RS2.eof
			%>				<TD><INPUT TYPE="radio" NAME="Reason" value="<%=RS2("ID")%>" <% if cint(reason)=cint(RS2("ID")) then response.write "checked "%> onclick="document.getElementById('sysForm').submit()"><%=RS2("Name")%><br></TD>
			<%				RS2.movenext
				wend
			%>		
			</TR>
			</TABLE>


			</FORM>
			<% else 
				response.write "<br><br>" 
				sys = fixSys 
			end if
			%>
			<A HREF="?act=show&selectedCustomer=<%=CustomerID%>&sys=<%=SYS%>">[ ê“«—‘ Õ”«» ]</A> &nbsp; [ê“«—‘ „«‰œÂ »œÂÌ]
			<BR><BR>
		</CENTER>

	<%
	tempCounter = 0
	if Not (RS1.EOF) then
%>
		<TABLE class="RepTable" width='90%' align='center'>
		<TR>
			<TD colspan=8 dir='rtl' align='center'>
			</TD>
		</TR>
		<TR>
			<TD class="RepTableTitle" colspan=8 dir='rtl' align='center'><br>
				<B><A target="_blank" HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=CustomerID%>">„—»Êÿ »Â <%=customerName%> (<%=CustomerID%>)</A></B><br><br>
			</TD>
		</TR>
		<TR class="RepTableHeader">
			<TD>#</TD>
			<TD> «—ÌŒ</TD>
			<TD width=70>⁄ÿ›</TD>
			<TD> Ê÷ÌÕ« </TD>
			<TD width=70>„»·€</TD>
			<TD width=70>„«‰œÂ</TD>
		</TR>
<%
		TotalDebit=0
		While Not (RS1.EOF)
			tempCounter=tempCounter+1
			
			'================================== Type = 1
			'================================== 
			if RS1("Type")=1 then
				sourceLink="<a href='?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— " & RS1("Link") & "</a>"
				if RS1("IsA") then
					Description = "»«»  ›«ﬂ Ê— ‘„«—Â "& RS1("InvNo")
				else
					Description = ""
				end if
				Set RS2 = conn.Execute("SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice = "& RS1("Link") & ")")
				if not RS2.eof then
					tempWriteAnd=""
					Description = Description & "(”›«—‘ "
					Do while not RS2.eof 
						Description = Description & tempWriteAnd & Link2Trace(RS2("Order"))
						tempWriteAnd=" Ê "
						RS2.moveNext
					Loop
					Description = Description & ")"
				end if
				RS2.close
				Set RS2=nothing

			
			'================================== Type = 2
			'================================== 
			elseif RS1("Type")=2 then
				sourceLink="<a href='?act=showReceipt&receipt="& RS1("Link") & "' target='_blank'>" & "œ—Ì«›  " & "</a>"
				Description = "œ—Ì«›  "
				myAND=""
				CashAmount=		cdbl(RS1("RcvCash"))
				ChequeAmount=	cdbl(RS1("RcvCheq"))
				if CashAmount<>0 then
					Description=Description & "‰ﬁœ "
					myAND="Ê "
				end if 
				if ChequeAmount<>0 then
					Description=Description & myAND & "çﬂ "
				end if 
				if clng(RS1("RcpNo"))<>0 then
					Description=Description & "ÿÌ —”Ìœ ‘„«—Â "& RS1("RcpNo")
				end if
			
			'================================== Type = 3
			'================================== 
			elseif RS1("Type")=3 then
				sourceLink="<a href='?act=showMemo&sys="& sys & "&memo="& RS1("Link") & "' target='_blank'>" & "«⁄·«„ÌÂ " & "</a>"
				Description = "<span style='overflow:auto;border:none; width:250px; height:23px; font-size:7pt;'>" & RS1("Description") & "</span>" 'RS1("Description")

			'================================== Type = 4
			'================================== 
			elseif RS1("Type")=4 then
				sourceLink="<a href='?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "»—ê‘  «“ ›—Ê‘" & RS1("Link") & "</a>"
				Description = "&nbsp;»«»  »—ê‘  «“ ›—Ê‘ ‘„«—Â "& RS1("InvNo")

			'================================== Type = 5
			'================================== 
			elseif RS1("Type")=5 then
				sourceLink="<a href='?act=showPayment&payment="& RS1("Link") & "' target='_blank'>" & "Å—œ«Œ  </a>"
				Description = "Å—œ«Œ  "
				myAND=""
				CashAmount=		cdbl(RS1("PaidCash"))
				ChequeAmount=	cdbl(RS1("PaidCheq"))
				if CashAmount<>0 then
					Description=Description & "‰ﬁœ "
					myAND="Ê "
				end if 
				if ChequeAmount<>0 then
					Description=Description & myAND & "çﬂ "
				end if 

			'================================== Type = 6
			'================================== 
			elseif RS1("Type")=6 then
				sourceLink="<a href='?act=showVoucher&voucher="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— Œ—Ìœ </a>"
				Description = " «—ÌŒ : <span dir='LTR'>(" & RS1("EffectiveDate") & ")</span>"

			'================================== Unknown
			'================================== 
			else
				sourceLink="<a href='javascript:void(0);'>" & "‰«‘‰«”" & "</a>"
				Description = " «—ÌŒ : <span dir='LTR'>(" & RS1("EffectiveDate") & ")</span>"
			end if

'			Remained = -1 * cdbl(RS1("RemainedAmount"))
'			Total = -1 * cdbl(RS1("AmountOriginal"))
'			Changed by kid 820824

			DebitRemained = cdbl(RS1("RemainedAmount"))
			Debit = cdbl(RS1("AmountOriginal"))


			if RS1("Voided") then 

%>			<TR bgcolor=#FFEEEE style='color:#999999'>
				<td> # <%tempCounter = tempCounter -1 %></td>
				<td dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td><%=sourceLink%></td>
				<td><%=Description%></td>
				<td dir='LTR' align='right'><%=Separate(Debit) %></td>
				<td dir='LTR' align='right'>&nbsp;</td>
			</TR>
<%
			else
				TotalDebit = TotalDebit + Debit
%>			<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
				<td><%=tempCounter %></td>
				<td dir='LTR' align='right'><%=RS1("EffectiveDate")%></td>
				<td><%=sourceLink%></td>
				<td><%=Description%></td>
				<td dir='LTR' align='right'><%=Separate(Debit) %></td>
				<td dir='LTR' align='right'><%=Separate(DebitRemained)%></td>
			</TR>
<%
			end if
			RS1.MoveNext
		Wend
%>
			<TR>
				<TD class="RepTableFooter" colspan='4'>&nbsp;&nbsp; : Ã„⁄ »œÂﬂ«—</span></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="Red"><%=Separate(TotalDebit)%></FONT></td>
				<TD class="RepTableFooter" >&nbsp;</td>
			</TR>
<%		mySQL="SELECT ISNULL(SUM(AmountOriginal),0) AS totalCred, ISNULL(SUM(RemainedAmount),0) AS totalCredRemained FROM "& sys & "Items WHERE (Account = '"& CustomerID & "') AND (IsCredit = 1) AND (Voided = 0) GROUP BY Account"
		Set RS1=Conn.execute(mySQL)
		if RS1.eof then
			TotalCredit = 0
'			TotalCreditRemained = 0
		else
			TotalCredit = cdbl(RS1("totalCred"))
'			TotalCreditRemained = cdbl(RS1("totalCredRemained"))
		end if
		
		Summ= TotalCredit - TotalDebit 
		SummText="&nbsp;"
		if Summ < 0 then
			Summ= -1 * Summ
			SummColor="Red"
			SummText="(»œÂﬂ«—)"
		elseif Summ > 0 then
			SummColor="Green"
			SummText="(»” «‰ﬂ«—)"
		end if
%>
			<TR>
				<TD class="RepTableFooter" colspan='4'>&nbsp;&nbsp; : Ã„⁄ »” «‰ﬂ«— </span></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="Green"><%=Separate(TotalCredit)%></FONT></td>
				<TD class="RepTableFooter" >&nbsp;</td>
			</TR>
			<TR>
				<TD class="RepTableFooter" colspan='4'>&nbsp;&nbsp; : Ã„⁄ ﬂ·</span></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="<%=SummColor%>"><%=Separate(Summ)%></FONT></td>
				<TD class="RepTableFooter" align='right'><FONT COLOR="<%=SummColor%>"><%=SummText%></FONT></td>
			</TR>
		</TABLE>
		<br>
<%	end if

'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showPayment" AND request("payment") <> "" then
	payment=request("payment")
	if not(isnumeric(payment)) then
		response.write "<br>"
		CALL showAlert ("Œÿ« ! <br><br><A HREF='InvoiceDetails.asp' style='font-size:7pt;'>Ã” ÃÊÌ „Ãœœ</A>",CONST_MSG_ERROR) 
		response.end
	end if

	'mySQL="SELECT Payments.*, PaidCash.Description FROM Payments LEFT OUTER JOIN PaidCash ON Payments.ID = PaidCash.Payment WHERE (Payments.ID = '"& Payment & "')"
	'Changed By Kid 840503, Adding the CashRegister in wich the payment is registered
	'mySQL="SELECT Payments.*, PaidCash.Description, CashRegisters.ID AS CashReg, CashRegisters.Cashier, CashRegisters.IsOpen, CashRegisters.NameDate FROM CashRegisters INNER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg RIGHT OUTER JOIN Payments ON CashRegisterLines.Link = Payments.ID LEFT OUTER JOIN PaidCash ON Payments.ID = PaidCash.Payment WHERE (Payments.ID = '"& Payment & "') "
	'Changed By Kid 860828, Correcting the join criteria , enforcing payment type (CashRegisterLines.Type = 2)
	mySQL="SELECT Payments.*, PaidCash.Description, CashRegisters.ID AS CashReg, CashRegisters.Cashier, CashRegisters.IsOpen, CashRegisters.NameDate FROM CashRegisters INNER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg RIGHT OUTER JOIN Payments ON CashRegisterLines.Type = 2 AND CashRegisterLines.Link = Payments.ID LEFT OUTER JOIN PaidCash ON Payments.ID = PaidCash.Payment WHERE (Payments.ID = '"& Payment & "') "

	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		response.write "<br>"
		CALL showAlert ("ÅÌœ« ‰‘œ.",CONST_MSG_ERROR) 
		response.end
	end if

	sys=RS1("sys")
	mySQL2 = "SELECT "& sys & "Items.ID, CreatedDate, CreationTime, EffectiveDate, Voided, Users.RealName FROM Users INNER JOIN "& sys & "Items ON Users.ID = "& sys & "Items.CreatedBy WHERE (Type = 5) AND (Link = "& Payment & ")" 
	Set RS2 = conn.Execute(mySQL2)
	AnyItemID =		RS2("ID")
	voided =		RS2("voided")
	EffectiveDate=	RS2("EffectiveDate")
	Creator=		RS2("RealName")
	CreatedDate=	RS2("CreatedDate")
	CreationTime=	RS2("CreationTime")
	RS2.close

	Description=	RS1("Description")

	CashReg= 		RS1("CashReg") & ""
	Cashier=		RS1("Cashier")
	IsOpen =		RS1("IsOpen")
	NameDate =		RS1("NameDate") & ""

	Account=		RS1("Account")
	CashAmount=		cdbl(RS1("CashAmount"))
	ChequeAmount=	cdbl(RS1("ChequeAmount"))
	TotalAmount=	CashAmount + ChequeAmount


	mySQL="SELECT ID,AccountTitle FROM Accounts WHERE (ID='"& Account & "')"

	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")
	RS1.close

	if Voided then
		stamp="<img src='/images/voided.gif'>"
	else
		stamp=""
	end if 

	if IsOpen="True" AND Cashier=session("ID") then
		CashRegLink="<A Target='_blank' HREF='../cashReg/CashRegReport.asp'>Â‰Ê“ œ— ’‰œÊﬁ </A>"
	elseif CashReg <> "" AND Auth(9 , 6) then		'ê“«—‘ ”—Å—”  ’‰œÊﬁ
		CashRegLink="<A Target='_blank' HREF='../cashReg/CashRegAdminReport.asp?act=showCashRegReport&CashRegID=" & CashReg & "'>’‰œÊﬁ " & replace(NameDate,"/",".") & "</A>"
	else
		CashRegLink=""
	end if
%>
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="10">
			<td width="150"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="50">
			<td colspan=2></td>
		</tr>
	</table>
	<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" align="center" style="border:1 solid navy;">
	<tr bgcolor='#CCCCEE' height="40px">
		<td colspan="6" align='center' style="border-bottom:1 solid navy;">—”Ìœ Å—œ«Œ  ‘„«—Â <%=payment%><br>„—»Êÿ »Â <a href="?act=show&selectedCustomer=<%=Account%>"><%=CustomerName%></a>.</td>
	</tr>
	<tr bgcolor='#CCCCEE' height="20px">
		<td align='left'> À» : </td>
		<td align='right' dir="LTR">&nbsp;<%=CreatedDate & " " & CreationTime%>&nbsp;</td>
		<td colspan='2' align='center'><%=CashRegLink%>&nbsp;</td>
		<td align='left'> À»  ﬂ‰‰œÂ: </td>
		<td align='right'>&nbsp;<%=Creator%>&nbsp;</td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan='5' align='left'>  «—ÌŒ („ÊÀ—): </td>
		<td align='right' dir="LTR">&nbsp;<%=EffectiveDate%>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="5"> „»·€ ‰ﬁœ: </td>
		<td align='right' dir="LTR" class="RcpRowInput"><%=Separate(CashAmount)%></td>
	</tr>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="6"> çﬂ: <br></td>
	</tr>
<%
	i=0
	mySQL="SELECT PaidCheques.*, Bankers.Name as BankOfOrigin FROM PaidCheques INNER JOIN Bankers ON PaidCheques.Banker = Bankers.ID WHERE (PaidCheques.Payment = "& Payment & ")"
	Set RS1 = conn.Execute(mySQL)
	if not(RS1.eof) then
%>
	<tr>
		<td class="RcpHeadInput" align='center' width="25px"> # </td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value="‘„«—Â çﬂ" size="10" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value=" «—ÌŒ" size="10" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="»«‰ﬂ" size="10" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value=" Ê÷ÌÕ" size="15" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="„»·€" size="15" tabindex="9999"></td>
	</tr>
<%
	else
%>
	<tr class="RcpHeadInput" >
		<td align='center' width="25px"> - - </td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="10" value="‰œ«—œ"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="10"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="10"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="15"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="15"></td>
	</tr>
<%
	end if
	while not(RS1.eof) 
		i=i+1
%>
			<tr bgcolor='#F0F0F0' height="20">
				<td align='center' width="25px"><%=i%></td>
				<td dir="LTR" class="RcpRowInput" title='ÃÂ  ⁄„·Ì«  ç«Å çﬂ ﬂ·Ìﬂ ﬂ‰Ìœ'><A HREF='/beta/bank/chequePrint.asp?PaidChequesID=<%=RS1("ID")%>'><%=RS1("ChequeNo")%></A></td>
				<td dir="LTR" class="RcpRowInput"><%=RS1("ChequeDate")%></td>
				<td dir="RTL" class="RcpRowInput"><%=RS1("BankOfOrigin")%></td>
				<td dir="LTR" class="RcpRowInput"><%=RS1("Description")%></td>
				<td dir="LTR" class="RcpRowInput"><%=Separate(RS1("Amount"))%></td>
			</tr>
<%
		RS1.moveNext
	wend %>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="6">  Ê÷ÌÕ« : <%=Description%> <br></td>
	</tr>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="5" align="left"> Ã„⁄:&nbsp;&nbsp;</td>
		<td align='right' style="font-size: 9pt; direction:LTR; background-color=#C0D0F0;"><%=Separate(TotalAmount)%></td>
	</tr>
		<tr bgcolor='#CCCC88' height='30px'>
			<td colspan="6" align='center'>
<%	if Auth(9 , 7) AND NOT voided then			' Has the Priviledge to VOID the RECEIPT/PAYMENT %>
				<INPUT class="GenButton" TYPE="button" Value=" «»ÿ«· " onclick="VoidPayment();">
<%	end if '--------------------------------------------------EDIT BY SAM---------------------------------------------------%>
			<% '	ReportLogRow = PrepareReport ("InvoicePrintForm.rpt", "Payment", Payment, "/beta/dialog_printManager.asp?act=Fin") %>
			<input class='GenButton' type='button' value='ç«Å çﬂ' onclick="changeURL('checkPrint','/beta/bank/chequePrint.asp?payment=<%=payment%>');">
			<!--input class='GenButton' type='button' value='ç«Å »—êÂ çﬂ' onclick='printThisReport(this,<%=ReportLogRow%>);'-->
		</tr>
	</TABLE>
	<Br>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function VoidPayment(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ Å—œ«Œ  —« '»«ÿ·' ﬂ‰Ìœø\n\n"))
			window.location="../cashReg/Void.asp?act=voidPayment&Payment=<%=Payment%>";
	}
	//-->
	function changeURL(winName, newURL) {
		win = window.open("", winName);
		win.location.href = newURL;
	}


	</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showVoucher" then
	ON ERROR RESUME NEXT
		vouch = clng(request("voucher"))
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— Œ—Ìœ „⁄ »— ‰Ì” .")
		end if
	ON ERROR GOTO 0
	sys = "AP"
	set RSV=Conn.Execute ("SELECT * FROM Vouchers WHERE (ID="& vouch & ")" )
	if RSV.eof then
		conn.close
		response.redirect "top.asp?errMsg=" & Server.URLEncode("ç‰Ì‰ ›«ﬂ Ê— Œ—ÌœÌ ÊÃÊœ ‰œ«—œ.")
	end if
	VendorID =		RSV("VendorID") 
	VouchTitle=		RSV("title")
	CreationDate=	RSV("CreationDate")
	CreationTime=	RSV("CreationTime")
	SavedFileName=	RSV("SavedFileName")
	Comments=		RSV("Comment")
	TotalPrice=		RSV("TotalPrice")
	Verified=		RSV("Verified")
	Voided =		RSV("voided")
	EffectiveDate =	RSV("EffectiveDate")
	Paid =			RSV("Paid")
	totalVat=		RSV("totalVat")
	number =		RSV("Number")

	set RSL=Conn.Execute ("SELECT VoucherLines.*, PurchaseOrders.Status FROM VoucherLines INNER JOIN PurchaseOrders ON VoucherLines.RelatedPurchaseOrderID = PurchaseOrders.ID WHERE (VoucherLines.Voucher_ID = "& vouch & ")" )
	set RSF=Conn.Execute ("SELECT AccountTitle,APBalance FROM Accounts WHERE (ID = "& VendorID & ")" )

	APBalance=		cdbl(RSF("APBalance"))

	Set RS2 = conn.Execute("SELECT ID FROM "& sys & "Items WHERE (Type = 6) AND (Link = "& vouch & ")")
	if not RS2.EOF then	AnyItemID = RS2("ID")
	RS2.close
	Set RS2 = Nothing

	if Voided then
		stamp="<img src='/images/voided.gif'>"
	elseif Verified then
		stamp="<img src='/images/Approved.gif'>"
	else
		stamp=""
	end if 

%>	
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="10">
			<td width="50"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="50">
			<td colspan=2></td>
		</tr>
	</table>
	<TABLE align=center Style="border:1 solid #330099;">
	<TR>
		<TD align=center colspan=2 height='25'><B>Ã“ÌÌ«  ›«ﬂ Ê—</B></td>
	</TR>
	<TR>
		<TD colspan=2 style='height:1px; background-color:#330099;'></td>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP>›—Ê‘‰œÂ: </td>
		<TD align=right VALIGN=TOP><B><%=RSF("AccountTitle")%> </B>
		</TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP> —«“ Õ”«» ›—Ê‘‰œÂ: </td>
		<TD align=right VALIGN=TOP><B>
		<%if APBalance<0 then %>
			<font color=red>»œÂﬂ«—: <span dir=ltr><%=Separate(APBalance)%></span> —Ì«· </font>
		<%elseif APBalance>0 then %>
			<font color=green>»” «‰ﬂ«—: <span dir=ltr><%=Separate(APBalance)%></span> —Ì«· </font>
		<%else %>
			0
		<%end if %></B>
		</TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP>⁄‰Ê«‰ ›«ﬂ Ê— : </TD>
		<TD align=right VALIGN=TOP><B><%=VouchTitle%></B></TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP> «—ÌŒ «ÌÃ«œ: </TD>
		<TD align=right VALIGN=TOP><B><span dir=ltr><%=CreationDate%></span> (”«⁄  <%=CreationTime%>)</B></TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP> «—ÌŒ „ÊÀ—: </TD>
		<TD align=right VALIGN=TOP><B><span dir=ltr><%=EffectiveDate%></span></B></TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP>›«Ì·  ’ÊÌ— ›«ﬂ Ê— : </TD>
		<TD align=right VALIGN=TOP><B>
		<% 	if SavedFileName<>"-" and SavedFileName<>"" then %>
		<A HREF="vouchers/<%=SavedFileName%>" target=_blank>œ—Ì«›  ›«Ì·  ’ÊÌ— ›«ﬂ Ê—</A>
		<%else%>-
		<% end if %>
		
		</B></TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP> Ê÷ÌÕ« : </TD>
		<TD align=right VALIGN=TOP><B><%=Comments%></B></TD>
	</TR>
	<TR>
		<TD align=left VALIGN=TOP>‘„«—Â ›«ﬂ Ê— «’·Ì:</TD>
		<TD align=right VALIGN=TOP><B><%=number%></B></TD>
	</TR>
	<TR>
		<TD align=right VALIGN=TOP colspan=2>
			<TABLE dir=rtl align=center width=400>
			<TR bgcolor="eeeeee" >
				<TD><SMALL>—œÌ› ›«ﬂ Ê—</SMALL></TD>
				<TD align='center'><SMALL>›Ì</SMALL></TD>
				<TD><SMALL> ⁄œ«œ</SMALL></TD>
				<TD><SMALL>ﬁÌ„  (—Ì«·)</SMALL></TD>
				<TD><SMALL>Ê÷⁄Ì  ”›«—‘ „—»ÊÿÂ</SMALL></TD>
			</TR>
			<%
			tmpCounter=0
			flag = ""
			Do while not RSL.eof
				tmpCounter = tmpCounter + 1
				if tmpCounter mod 2 = 1 then
					tmpColor="#FFFFFF"
					tmpColor2="#FFFFBB"
				Else
					tmpColor="#DDDDDD"
					tmpColor2="#EEEEBB"
				End if 

			%>
			<TR bgcolor="<%=tmpColor%>" >
				<TD><%=RSL("LineTitle")%></TD>
				<TD>
<%				if cdbl(RSL("qtty")) > 0 then 
					response.write(Separate(fix(cdbl(RSL("price"))/cdbl(RSL("qtty")))))
				else
					response.write("!")
				end if
%>
				</TD>
				<TD><%=RSL("qtty")%></TD>
				<TD><%=Separate(cdbl(RSL("price")))%></TD>
				<TD><%=RSL("status")%> &nbsp;&nbsp;&nbsp;( <A HREF="../purchase/outServiceTrace.asp?od=<%=RSL("RelatedPurchaseOrderID")%>">‘„«—Â :<%=RSL("RelatedPurchaseOrderID")%></A>)</TD>
				<% if RSL("status")<>"OK" then flag=" disabled" %>
			</TR>
				 
			<% 
			RSL.moveNext
			Loop
			%>
			<TR>
				<TD colspan=1 align=left>„«·Ì«  »— «—“‘ «›“ÊœÂ: </TD>
				<TD colspan=5><%=Separate(TotalVat)%> —Ì«·</td>
			</TR>			
			<TR>
				<TD colspan=1 align=left>Ã„⁄ ﬂ·</TD>
				<TD colspan=5><%=Separate(TotalPrice)%> —Ì«·</td>
			</TR>
			</TABLE><br>
		</TD>
	</TR>
<%	if NOT Voided then 
%>
	<TR>
		<TD colspan=2 style='height:1px; background-color:#330099;'></td>
	</TR>
	<TR>
		<TD colspan=2 align=center>
<%	 
		if NOT Verified then 
			if flag=" disabled" then
				response.write "<br>"
				CALL showAlert ("»—ŒÌ «“ «ﬁ·«„ «Ì‰ ›«ﬂ Ê—  «ÌÌœ ‰‘œÂ «‰œ.<br>·–« «„ﬂ«‰  «ÌÌœ ﬂ· ¬‰ ÊÃÊœ ‰œ«—œ.",CONST_MSG_ALERT) 
			end if
%>			<BR>
			<FORM METHOD=POST ACTION="../AP/verify.asp">
				<INPUT TYPE="hidden" name="VouchID" value=<%=vouch%>>
<%				if Auth(7 , 2) then ' Has the privilege to APPROVE Voucher
%>
					<INPUT TYPE="submit" Name="Submit" Value=" «ÌÌœ" class="GenButton" style="width:150px;" tabIndex="14" <%=flag%>>
					<br><br>
<%				end if
%>
				<INPUT TYPE="button" Value=" ÊÌ—«Ì‘ " class="GenButton" style="width:150px;" onclick="window.location='../AP/voucherEdit.asp?act=editVoucher&VoucherID=<%=vouch%>';">
				&nbsp;&nbsp;
				<INPUT TYPE="button" Value=" Õ–› " class="GenButton" style="width:150px;border:1 solid red;" onclick="if (confirm('¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— Œ—Ìœ —« Õ–› ﬂ‰Ìœø\n\n'))			window.location='../AP/voucherEdit.asp?act=delVoucher&VoucherID=<%=vouch%>';">
			</FORM>
<%		else
			response.write "<br>«Ì‰ ›«ﬂ Ê—  «ÌÌœ ‘œÂ «” . <br><br>"
			if Auth(7 , 9) then ' Has the privilege to VOID Voucher
%>				<INPUT class="GenButton" style="width:150px;border:1 solid red;" type="button" value=" «»ÿ«· " onclick="VoidVoucher();">

				<br><br>
<%			end if
		end if 
%>
		</TD>
	</TR>
	<TR>
		<TD colspan=2 style='height:1px; background-color:#330099;'></td>
	</TR>
	<TR>
		<TD colspan=2 align=center>
<%
		if Paid then
			response.write "<br>„»·€ «Ì‰ ›«ﬂ Ê— Å—œ«Œ  ‘œÂ «” . <br><br>"
		else
			if Verified then 
				if flag=" disabled" then
					response.write "<br>"
					CALL showAlert ("»—ŒÌ «“ «ﬁ·«„ «Ì‰ ›«ﬂ Ê—  «ÌÌœ ‰‘œÂ «‰œ.<br>·–« «„ﬂ«‰ Å—œ«Œ  ¬‰ ÊÃÊœ ‰œ«—œ.",CONST_MSG_ALERT) 
				end if 
			else
				response.write "<br>"
				CALL showAlert ("«Ì‰ ›«ﬂ Ê—  «ÌÌœ ‰‘œÂ «” .<br>·–« «„ﬂ«‰ Å—œ«Œ  ¬‰ ÊÃÊœ ‰œ«—œ.",CONST_MSG_ALERT) 
				flag=" disabled" 
			end if
%>
			<BR>
			<FORM METHOD=POST ACTION="../bank/cheq.asp?act=enterCheque">
				<INPUT TYPE="hidden" name="VouchID" value=<%=vouch%>>
				<INPUT TYPE="submit" Name="Submit" Value=" Å—œ«Œ  " class="GenButton" style="width:150px;" tabIndex="14" <%=flag%>>
			</FORM>
<%
		end if
	end if 
%>
		</TD>
	</TR>
	</TABLE>
	<script language="JavaScript">
	<!--
	function VoidVoucher(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '»«ÿ·' ﬂ‰Ìœø\n"))
			window.location='../AP/voucherEdit.asp?act=voidVoucher&VoucherID=<%=vouch%>';
	}
	//-->
	</script>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showInvoice" AND request("invoice") <> "" then
	InvoiceID=request("invoice")
	if not(isnumeric(InvoiceID)) then
		response.write "<br>" 
		call showAlert ("Œÿ« œ— ‘„«—Â ›«ﬂ Ê—",CONST_MSG_ERROR) 
		response.end
	end if
	InvoiceID=clng(InvoiceID)
	'mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
	'Changed By kid 840502 extracting Creator, Approver, Issuer & Voider
	mySQL="SELECT Invoices.*, Users_1.RealName AS Creator, Users_2.RealName AS Approver, Users_3.RealName AS Voider, Users_4.RealName AS Issuer FROM Invoices LEFT OUTER JOIN Users Users_4 ON Invoices.IssuedBy = Users_4.ID LEFT OUTER JOIN Users Users_3 ON Invoices.VoidedBy = Users_3.ID LEFT OUTER JOIN Users Users_2 ON Invoices.ApprovedBy = Users_2.ID LEFT OUTER JOIN Users Users_1 ON Invoices.CreatedBy = Users_1.ID WHERE (Invoices.ID ='"& InvoiceID & "')"

	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		response.write "<br>" 
		call showAlert ("ÅÌœ« ‰‘œ",CONST_MSG_ERROR) 
		response.end
	end if
	sys = "AR"

	customerID=		RS1("Customer")
	totalPrice=		cdbl(RS1("totalPrice"))
	totalDiscount=	cdbl(RS1("totalDiscount"))
	totalReverse=	cdbl(RS1("totalReverse"))
	totalVat =		cdbl(RS1("totalVat"))
	creationDate=	RS1("CreatedDate")
	ApproveDate=	RS1("ApprovedDate")
	issueDate=		RS1("IssuedDate")
	VoidDate=		RS1("VoidedDate")
	InvoiceNo=		RS1("Number") 
	Voided=			RS1("Voided")
	Issued=			RS1("Issued")
	Approved=		RS1("Approved")
	IsReverse=		RS1("IsReverse")
	if RS1("IsA") = TRUE then IsA=1 else IsA=0
	Creator =		RS1("Creator")
	Approver =		RS1("Approver")
	Issuer =		RS1("Issuer")
	Voider =		RS1("Voider")

	if IsReverse then itemType=4 else itemType=1

	Set RS2 = conn.Execute("SELECT ID FROM "& sys & "Items WHERE (Type = "& itemType & ") AND (Link = "& InvoiceID & ")")
	if not RS2.EOF then	AnyItemID = RS2("ID")

	TotalReceivable= totalPrice - totalDiscount - totalReverse + totalVat

	mySQL="SELECT ID,AccountTitle FROM Accounts WHERE (ID='"& customerID & "')"

	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")

	RS1.close

	if Voided then
		stamp="<img src='/images/voided.gif'>"
	elseif Issued then
		stamp="<img src='/images/Issued.gif'>"
	elseif Approved then
		stamp="<img src='/images/Approved.gif'>"
	else
		stamp=""
	end if 

	if IsReverse then
		HeaderColor="#FF9900"
	else
		HeaderColor="#C3C300"
	end if

%>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>

	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr height="10">
			<td width="280"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="20">
			<td colspan=2></td>
		</tr>
	</table>

	<table align="center" Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="InvTable">
		<tr>
		<TD colspan="1"><TABLE Border="0" Width="100%" Cellspacing="0" Cellpadding="1" Dir="RTL" bgColor="<%=HeaderColor%>">
			<TR>
				<TD colspan=2>&nbsp;&nbsp;<%if IsReverse then response.write "<B>›«ﬂ Ê— »—ê‘  «“ ›—Ê‘</B>"%></TD>
				<TD align="left">‘„«—Â ›«ﬂ Ê—:</TD>
				<TD width=20%>&nbsp;<INPUT readonly class="InvGenInput" NAME="InvoiceID" value="<%=InvoiceID%>" style="direction:ltr" TYPE="text" maxlength="10" size="10"></TD>
			</TR>
			<TR>
				<TD align="left">Õ”«»:</TD>
				<TD align="right">
					<span id="customer">
						<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><span><A HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=customerID%>" target="_blank"><%=CustomerName%></A></span>
					</span>	<input type="hidden" Name='InvoiceID' value='<%=InvoiceID%>'></TD>
				<TD align="left"> «—ÌŒ ’œÊ—:</TD>
				<TD width="20%"><table border="0">
					<tr>
						<td dir="ltr">
							<input readonly class="InvGenInput" NAME="InvoiceDate" TYPE="text" maxlength="10" size="10" value="<%=issueDate%>">
						</td>
						<td dir="rtl"><%="ø‘‰»Â"%></td>
					</tr>
					</table></TD>
			</TR>
			<TR>
				<TD align="left" width="100px">„—»Êÿ »Â ”›«—‘:</TD>
				<TD align="right">
					<span id="orders">
<%
					tempWriteAnd=""

					mySQL="SELECT InvoiceOrderRelations.[Order], orders_trace.order_kind, orders_trace.order_title FROM InvoiceOrderRelations LEFT OUTER JOIN orders_trace ON InvoiceOrderRelations.[Order] = orders_trace.radif_sefareshat WHERE (InvoiceOrderRelations.Invoice = '"& InvoiceID & "')"

					Set RS1 = conn.Execute(mySQL)
					while not(RS1.eof) 
						response.write "<input type='hidden' name='selectedOrders' value='"& RS1("Order") & "'>"
						response.write tempWriteAnd & Link2Trace(RS1("Order"))
						response.write " [ " & RS1("Order_Title") & " (" & RS1("Order_Kind") & ") ]"
						tempWriteAnd=" Ê "
						RS1.moveNext
					Wend
					RS1.close
					Set RS1 = Nothing 
%>					</span>&nbsp;
				</TD>
<%				if IsA then %>	
					<TD align="left">‘„«—Â:</td>
					<TD><TABLE border="0">
						<TR>
							<TD dir="LTR">
								<INPUT readonly class="InvGenInput" NAME="InvoiceNo" value="<%=InvoiceNo%>" style="border:1px solid black;" TYPE="text" maxlength="10" size="10"></td>
							</TD>
							<td dir="RTL">
									<B>«·›</B> &nbsp;
							</td>
						</TR>
						</TABLE></TD>
<%				else%>
					<TD Colspan="2">&nbsp;</TD>
<%				end if%>
			</TR>
<%
			tempWriteAnd=""
			relatedQuotesText=""
			hasRelatedQuotes = false
			mySQL="SELECT InvoiceQuoteRelations.[Quote], Quotes.[order_kind], Quotes.[order_title] FROM InvoiceQuoteRelations INNER JOIN Quotes ON InvoiceQuoteRelations.[Quote] = Quotes.[ID] WHERE (InvoiceQuoteRelations.[Invoice] = '"& InvoiceID & "')"

			Set RS1 = conn.Execute(mySQL)
			While Not(RS1.eof) 
				'response.write "<input type='hidden' name='selectedQuotes' value='"& RS1("Quote") & "'>"
				relatedQuotesText = relatedQuotesText & tempWriteAnd & Link2TraceQuote(RS1("Quote"))
				relatedQuotesText = relatedQuotesText & " [ " & RS1("Order_Title") & " (" & RS1("Order_Kind") & ") ]"
				tempWriteAnd=" Ê "
				hasRelatedQuotes = true
				RS1.moveNext
			Wend
			RS1.close
			Set RS1 = Nothing 

			If hasRelatedQuotes Then 
%>
				<TR bgcolor='#AAAAEE' height="30">
					<TD align="left" width="100px">„—»Êÿ »Â «” ⁄·«„ :</TD>
					<TD align="right" colspan="3">
						&nbsp; 
						<span id="quotes"><%=relatedQuotesText%></span>&nbsp;
					</TD>
				</TR>
<%			End If %>
			</TABLE></TD>
		</tr>
		<tr bgcolor="#AAAA55" >
			<td>
			<TABLE Border="0" Cellspacing="0" Cellpadding="5" Dir="RTL" >
			<TR>
				<td width='70' bgcolor="#BBBBBB" align='left'>«ÌÃ«œ ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#BBBBBB" align='right' title="<%=CreationDate%>"><%=Creator%></td>
<%			if Approved then%>
				<td width='70' bgcolor="#77BB99" align='left'> «ÌÌœ ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#77BB99" align='right' title="<%=ApproveDate%>"><%=Approver%></td>
<%			end if
			if Issued then%>
				<td width='70' bgcolor="#7799BB" align='left'>’«œ— ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#7799BB" align='right' title="<%=IssueDate%>"><%=Issuer%></td>
<%			end if
			if Voided then%>
				<td width='70' bgcolor="#BB7799" align='left'>«»ÿ«· ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#BB7799" align='right' title="<%=VoidDate%>"><%=Voider%></td>
<%			end if
%>
			</TR>
			</TABLE>
			</td>
		</tr>
		<tr bgcolor='#F0F0F0'>
		<TD colspan="1"><div>
		<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" align="center" class="InvTable">
		<TR bgcolor='#CCCC88'>
			<td align='center' width="25px"> # </td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" value="¬Ì „" size="4" ></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" value=" Ê÷ÌÕ« " size="30"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="ÿÊ·" size="2"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="⁄—÷" size="2"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value=" ⁄œ«œ" size="3"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="›—„" size="2"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value=" ⁄œ«œ „ÊÀ—" size="6"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="›Ì" size="7"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="ﬁÌ„ " size="9"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value=" Œ›Ì›"size="7"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="»—ê‘ " size="5"></td>
			<!-------------------------------SAM----------------------------------------->
			<td><INPUT readonly class="InvHeadInput4" TYPE="text" Value="„«·Ì« " size="6"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="ﬁ«»· Å—œ«Œ " size="9"></td>
		</TR>
<%		
		i=0
		mySQL="SELECT * FROM InvoiceLines WHERE (Invoice='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		while not(RS1.eof) 
			i=i+1
			Price =		cdbl(RS1("Price"))
			AppQtty=	cdbl(RS1("AppQtty"))
			Discount =	cdbl(RS1("Discount"))
			Reverse =	cdbl(RS1("Reverse"))
'	SAM
			Vat =		cdbl(RS1("Vat"))

%>
			<TR bgcolor='#F0F0F0' height="20px">
				<td align='center' width="25px"><%=i%></td>
				<td class="InvRowInput" dir="LTR"><%=RS1("Item")%></td>
				<td class="InvRowInput2" dir="RTL" width="170px"><%=RS1("Description")%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Length")))%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Width")))%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Qtty")))%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Sets")))%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(AppQtty)%></td>
				<td class="InvRowInput" dir="LTR"><%if AppQtty = 0 then response.write 0 else response.write Separate(Price/AppQtty)%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(Price)%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(Discount)%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(Reverse)%></td>
				<td class="InvRowInput4" dir="LTR"><%=Separate(Vat)%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(Price - Discount - Reverse + Vat)%></td>
			</TR>
<%
			RS1.moveNext
		wend
		RS1.close
%>
		<TR>
			<td colspan="13" height="2px" bgcolor="#CCCC88"></td>
		</TR>
		<TR bgcolor='#CCCC88' height="20px">
			<td colspan="9"> &nbsp; </td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalPrice)%></td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalDiscount)%></td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalReverse)%></td>
			<td align='right' dir="LTR" bgcolor="#FF9900"><%=Separate(TotalVat)%></td>
			<td align='right' dir="LTR" bgcolor="#AAF0FF"><%=Separate(TotalReceivable)%></td>
		</TR>
		<TR bgcolor='#CCCC88' height="20px">
			<td colspan="10"> &nbsp; </td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Pourcent(TotalDiscount,TotalPrice) & "% Œ›Ì›"%></td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Pourcent(TotalReverse,TotalPrice) & "%»—ê‘ "%></td>
			<td>&nbsp;</td>
		</TR>
		<TR bgcolor='<%=HeaderColor%>' height="20px">
			<td align=center colspan='13' style="padding:2;">
<%			if voided then
				'This is checked in order to prevent additional checks later.
				'Because when an invoice is voided, definitely it is has been issued before.
			elseif Issued then 
				if Auth(6 , "A") then
					' Has the Priviledge to CHANGE the Invoice / Rev. Invoice after it has been issued
%>						<input class="GenButton" type="button" value=" «’·«Õ«  Ã“ÌÌ " onclick="window.location='../AR/InvoiceEdit.asp?act=editInvoice&invoice=<%=InvoiceID%>';"> 
<%
				end if
				if Auth(6 , "F") then
					' Has the Priviledge to VOID the Invoice / Rev. Invoice
%>						<input class="GenButton" style="border:1 solid red;" title="”‰œ Õ”«»œ«—Ì œ«—œ" type="button" value=" «»ÿ«· " onclick="VoidInvoice();"> 
<%
				else
					if Auth(6 , "M") then
						set myRS=conn.Execute("SELECT * FROM ARItems WHERE (Type = '"& itemType & "') AND (Link='"& InvoiceID & "') AND GL_Update=1")
						if not myRS.eof then 					
								' Has the Priviledge to VOID the Invoice not has gl_update / Rev. Invoice
			%>						<input class="GenButton" style="border:1 solid red;" title="”‰œ Õ”«»œ«—Ì ‰œ«—œ" type="button" value=" «»ÿ«· " onclick="VoidInvoiceOnly();"> 
			<%
						end if
					end if
				end if
			else
				' Is not Issued
				if Approved then
					if Auth(6 , "D") then					' Has the Priviledge to ISSUE the Invoice / Rev. Invoice
						if Auth(6 , "I") then				
							' can ISSUE the Invoice / Rev. Invoice on another Date
%>							<INPUT class="InvGenInput" style="text-align:left;direction:LTR;" TYPE="text" maxlength="10" size="10" TYPE="Text" value="<%=shamsiToday()%>" NAME="IssueDate" OnBlur="return acceptDate(this);">
							<INPUT class="GenButton" TYPE="button" value=" ’œÊ— ›«ﬂ Ê— " onclick="IssueInvoice();">
<%						else
%>							<INPUT class="GenButton" TYPE="button" value=" ’œÊ— ›«ﬂ Ê— " onclick="IssueInvoice();">
<%						end if
					end if
				elseif Not(hasRelatedQuotes) then
					'Is not approved
					if Auth(6 , "C") then					' Has the Priviledge to APPROVE the Invoice / Rev. Invoice
%>						<INPUT class="GenButton" TYPE="button" value="  «ÌÌœ ›«ﬂ Ê— " onclick="ApproveInvoice();">
<%					end if
					
				end if

				if IsReverse AND Auth(6 , 5) then	' Has the Priviledge to EDIT the Rev. Invoice
%>					<input class="GenButton" type="button" value=" ÊÌ—«Ì‘ " onclick="window.location='../AR/InvoiceEdit.asp?act=editInvoice&invoice=<%=InvoiceID%>';"> 
<%				elseif Auth(6 , 3)then				' Has the Priviledge to EDIT the Invoice
%>					<input class="GenButton" type="button" value=" ÊÌ—«Ì‘ " onclick="window.location='../AR/InvoiceEdit.asp?act=editInvoice&invoice=<%=InvoiceID%>';"> 
<%				end if

				if Auth(6 , "G") then
					' Has the Priviledge to REMOVE the Pre-Invoice / Pre-Rev. Invoice
%>						<input class="GenButton" style="border:1 solid red;" type="button" value=" Õ–› ÅÌ‘ ‰ÊÌ” " onclick="RemovePreInvoice();"> 
					<%
				end if
			end if

'---------- Remarked by kid 840924 , mr. vazehi wanted so
'
'			if IsReverse AND Auth(6 , 4) then	' Has the Priviledge to INPUT a Rev. Invoice
'% >				<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InvoiceInput.asp?act=copyInvoice&invoice=< %=InvoiceID% >';"> 
'< %			elseif Auth(6 , 1)then				' Has the Priviledge to INPUT an Invoice
'% >				<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InvoiceInput.asp?act=copyInvoice&invoice=< %=InvoiceID% >';"> 
'< %			end if
'
'--------- End of Remark

'---------- Added by kid 850816, they requested to have the copy option but with priviledges
'
			if Auth(6 , "L") then	' Has the Priviledge to COPY an Invoice / Rev. Invoice
				if IsReverse AND Auth(6 , 4) then	' Has the Priviledge to INPUT a Rev. Invoice
%>					<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InvoiceInput.asp?act=copyInvoice&invoice=<%=InvoiceID%>';"> 
<%				elseif Auth(6 , 1)then				' Has the Priviledge to INPUT an Invoice
%>					<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InvoiceInput.asp?act=copyInvoice&invoice=<%=InvoiceID%>';"> 
<%				end if
			end if
'---------- End of 850816 additions


			if Auth(6 , "E") then					' Has the Priviledge to PRINT the Invoice / Rev. Invoice
%>				
				<% 	ReportLogRow = PrepareReport ("InvoicePre.rpt", "Invoice_ID", InvoiceID, "/beta/dialog_printManager.asp?act=Fin") %>
				<INPUT Class="GenButton" style="border:1 solid blue;" TYPE="button" value=" ç«Å ÅÌ‘ù‰ÊÌ” " 
				onclick="window.location='../AR/InvoicePrint.asp?r=<%=ReportLogRow%>';">
				<% 	ReportLogRow = PrepareReport ("InvoiceNew.rpt", "Invoice_ID", InvoiceID, "/beta/dialog_printManager.asp?act=Fin") %>
				<INPUT Class="GenButton" style="border:1 solid blue;" TYPE="button" value=" ç«Å " 
				onclick="window.location='../AR/InvoicePrint.asp?r=<%=ReportLogRow%>';">
<!--				onclick="printThisReport(this,<%=ReportLogRow%>);"-->
<%			'----------------------------------------------SAM-----------------------------------------------------
			if Issued then %>
				<INPUT class='GenButton' style='border:1 solid blue;' type='button' value=' Ê·Ìœ ﬂœ Å—œ«Œ  «Ì‰ —‰ Ì' onclick="window.location='../AR/ePayment.asp?Invoice=<%=InvoiceID%>';">
<%			end if
			'------------------------------------------------------------------------------------------------------
			end if
%>
			</td>
		</TR>
		</TABLE>
		</div></TD>
		</tr>
	</table>
	<br>
	<table align=center><tr>
<%
	mySQL="SELECT InvoiceOrderRelations.*, PurchaseOrders.ID AS PurchaseOrdersID, PurchaseOrders.Status AS PurchaseOrdersStatus, PurchaseOrders.TypeName AS TypeName FROM PurchaseOrders FULL OUTER JOIN PurchaseRequestOrderRelations RIGHT OUTER JOIN PurchaseRequests INNER JOIN InvoiceOrderRelations ON PurchaseRequests.Order_ID = InvoiceOrderRelations.[Order] ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID WHERE (InvoiceOrderRelations.Invoice ='"& InvoiceID & "') and PurchaseRequests.Status<> 'del'"
	Set RS1 = conn.Execute(mySQL)
	tempWriteAnd = "Œ—Ìœ Œœ„«  „—»Êÿ »Â «Ì‰ ›«ﬂ Ê—: <hr>"
	if not(RS1.eof) then
%>
		<td valign=top style="border:solid 1pt black">
			<div align=right>
<%
		while not(RS1.eof) 
			PurchaseOrdersStatus=RS1("PurchaseOrdersStatus")
			if PurchaseOrdersStatus="NEW" then
				status = "ÃœÌœ"
			elseif PurchaseOrdersStatus="OUT" then
				status = "Œ«—Ã «“ ‘—ﬂ "
			elseif PurchaseOrdersStatus="RETURN" then
				status = "»—ê‘ Â »Â ‘—ﬂ "
			elseif PurchaseOrdersStatus="CANCEL" then
				status = "·€Ê ‘œÂ"
			elseif PurchaseOrdersStatus="OK" then
				status = " «ÌÌœ ‘œÂ"
			end if
			response.write tempWriteAnd & "<li> <a target='_blank' href='../purchase/outServiceTrace.asp?od="&RS1("PurchaseOrdersID")&"'>" & RS1("TypeName") & "</a> - " & status
			tempWriteAnd=""
			RS1.moveNext
		wend
%>	
			<br>	
			</div>
		</td>
<%
	end if
%>
		<td width=30>
		</td>
<%
	mySQL="SELECT InvoiceOrderRelations.*, InventoryPickuplistItems.ItemName AS ItemName, InventoryPickuplistItems.pickupListID AS pickupListID, InventoryPickuplistItems.Qtty AS Qtty, InventoryPickuplistItems.unit AS unit, InventoryPickuplistItems.CustomerHaveInvItem FROM InventoryPickuplistItems INNER JOIN InventoryPickuplists ON InventoryPickuplistItems.pickupListID = InventoryPickuplists.id INNER JOIN InvoiceOrderRelations ON InventoryPickuplistItems.Order_ID = InvoiceOrderRelations.[Order] WHERE (InvoiceOrderRelations.Invoice = '"& InvoiceID & "' AND NOT InventoryPickuplists.Status = N'del')"


	Set RS1 = conn.Execute(mySQL)
	tempWriteAnd = "ÕÊ«·Â Â«Ì «‰»«— „—»Êÿ »Â «Ì‰ ›«ﬂ Ê—: <hr>"
	if not(RS1.eof) then
%>
		<td valign=top style="border:solid 1pt black">
			<div align=right>
<%
		while not(RS1.eof) 
			response.write tempWriteAnd & "<li><a target='_blank' href='../inventory/default.asp?show="&RS1("pickupListID")&"'>" & RS1("ItemName") & "</a> - " & RS1("Qtty") & RS1("unit")
			if RS1("CustomerHaveInvItem") then
			response.write " <B>(«—”«·Ì</B>)" 
			end if
			tempWriteAnd=""
			RS1.moveNext
		wend
%>
			<br>	
			</div>
		</td>
<%	end if
%>	
	</tr></table>
	<br>
	<table class="CustTable" cellspacing='1' align=center>
		<tr>
			<td colspan="2" class="CusTableHeader"><span style="width:450;text-align:center;">Ì«œœ«‘  Â«</span><span style="width:100;text-align:left;background-color:red;"><input class="GenButton" type="button" value="‰Ê‘ ‰ Ì«œœ«‘ " onclick="window.location = '../home/message.asp?RelatedTable=invoices&RelatedID=<%=InvoiceID%>&retURL=<%=Server.URLEncode("../AR/AccountReport.asp?act=showInvoice&invoice="&InvoiceID)%>';"></span></td>
		</tr>
<%
	mySQL="SELECT * FROM Messages INNER JOIN Users ON Messages.MsgFrom = Users.ID WHERE (Messages.RelatedTable = 'invoices') AND (Messages.RelatedID = "& InvoiceID & ") ORDER BY Messages.ID DESC"
	Set RS1 = conn.execute(mySQL)
	if NOT RS1.eof then
%>
<%
		tmpCounter=0
		Do While NOT RS1.eof 
			tmpCounter=tmpCounter+1
%>
			<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
				<td>«“ <%=RS1("RealName")%><br>
					<%=RS1("MsgDate")%> <BR> <%=RS1("MsgTime")%>
				</td>
				<td dir='RTL'><%=replace(RS1("MsgBody"),chr(13),"<br>")%></td>
			</tr>
<%
			RS1.moveNext
		Loop
	else
%>
		<tr class="CusTD3">
			<td colspan="2">ÂÌç</td>
		</tr>
<%
	end if
	RS1.close


%>
	</table><BR><BR>
	<br><!--CENTER>ÅÌ‘ ‰„«Ì‘ ç«Å:</CENTER-->
	<BR>
	<script language="JavaScript">
	<!--
	function ApproveInvoice(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« ' «ÌÌœ' ﬂ‰Ìœø\n\n"))
			window.location="../AR/invoiceEdit.asp?act=approveInvoice&invoice=<%=InvoiceID%>";
	}
	function IssueInvoice(){
	<% if Auth(6 , "I") then%>
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '’«œ—' ﬂ‰Ìœø\n\n"))
			window.location="../AR/invoiceEdit.asp?act=IssueInvoice&invoice=<%=InvoiceID%>&issueDate="+document.all.IssueDate.value;
	<% else %>
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '’«œ—' ﬂ‰Ìœø\n\n"))
			window.location="../AR/invoiceEdit.asp?act=IssueInvoice&invoice=<%=InvoiceID%>";
	<% end if%>
	}
	function VoidInvoice(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '»«ÿ·' ﬂ‰Ìœø\n")){

			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value=" Ê÷ÌÕÌ œ— »«—Â œ·Ì· «»ÿ«· ›«ﬂ Ê—:"
			window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false
			window.location="../AR/invoiceEdit.asp?act=voidInvoice&invoice=<%=InvoiceID%>&comment="+escape(document.all.tmpDlgTxt.value);
		}
	}
	function VoidInvoiceOnly(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '»«ÿ·' ﬂ‰Ìœø\n")){

			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value=" Ê÷ÌÕÌ œ— »«—Â œ·Ì· «»ÿ«· ›«ﬂ Ê—:"
			window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false
			window.location="../AR/invoiceEdit.asp?act=voidInvoiceOnly&invoice=<%=InvoiceID%>&comment="+escape(document.all.tmpDlgTxt.value);
		}
	}
	function RemovePreInvoice(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ÅÌ‘ ‰ÊÌ” —« Õ–› ﬂ‰Ìœø\n"))
			window.location='../AR/InvoiceEdit.asp?act=removePreInvoice&invoice=<%=InvoiceID%>';
	}
	//-->
	</script>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------- I N Q U I R Y ------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showInquiry" AND request("inquiry") <> "" then
	InquiryID=request("inquiry")
	if not(isnumeric(InquiryID)) then
		response.write "<br>" 
		call showAlert ("Œÿ« œ— ‘„«—Â «” ⁄·«„",CONST_MSG_ERROR) 
		response.end
	end if
	InquiryID=clng(InquiryID)
	'mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
	'Changed By kid 840502 extracting Creator, Approver, Issuer & Voider
	mySQL="SELECT Inquiries.*, Users_1.RealName AS Creator, Users_2.RealName AS Approver, Users_3.RealName AS Voider, Users_4.RealName AS Issuer FROM Inquiries LEFT OUTER JOIN Users Users_4 ON Inquiries.IssuedBy = Users_4.ID LEFT OUTER JOIN Users Users_3 ON Inquiries.VoidedBy = Users_3.ID LEFT OUTER JOIN Users Users_2 ON Inquiries.ApprovedBy = Users_2.ID LEFT OUTER JOIN Users Users_1 ON Inquiries.CreatedBy = Users_1.ID WHERE (Inquiries.ID ='"& InquiryID & "')"

	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		response.write "<br>" 
		call showAlert ("ÅÌœ« ‰‘œ",CONST_MSG_ERROR) 
		response.end
	end if
	sys = "AR"

	customerID=		RS1("Customer")
	totalPrice=		cdbl(RS1("totalPrice"))
	totalDiscount=	cdbl(RS1("totalDiscount"))
	totalReverse=	cdbl(RS1("totalReverse"))
	totalVat =		cdbl(RS1("totalVat"))
	creationDate=	RS1("CreatedDate")
	ApproveDate=	RS1("ApprovedDate")
	issueDate=		RS1("IssuedDate")
	VoidDate=		RS1("VoidedDate")
	InquiryNo=		RS1("Number") 
	Voided=			RS1("Voided")
	Issued=			RS1("Issued")
	Approved=		RS1("Approved")
	IsReverse=		RS1("IsReverse")
	if RS1("IsA") = TRUE then IsA=1 else IsA=0
	Creator =		RS1("Creator")
	Approver =		RS1("Approver")
	Issuer =		RS1("Issuer")
	Voider =		RS1("Voider")

	if IsReverse then itemType=4 else itemType=1

	Set RS2 = conn.Execute("SELECT ID FROM "& sys & "Items WHERE (Type = "& itemType & ") AND (Link = "& InquiryID & ")")
	if not RS2.EOF then	AnyItemID = RS2("ID")

	TotalReceivable= totalPrice - totalDiscount - totalReverse + totalVat

	mySQL="SELECT ID,AccountTitle FROM Accounts WHERE (ID='"& customerID & "')"

	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")

	RS1.close

	if Voided then
		stamp="<img src='/images/voided.gif'>"
	elseif Issued then
		stamp="<img src='/images/Issued.gif'>"
	elseif Approved then
		stamp="<img src='/images/Approved.gif'>"
	else
		stamp=""
	end if 

	if IsReverse then
		HeaderColor="#FF9900"
	else
		HeaderColor="#C3C300"
	end if

%>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>

	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr height="10">
			<td width="280"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="20">
			<td colspan=2></td>
		</tr>
	</table>

	<table align="center" Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="InvTable">
		<tr>
		<TD colspan="1"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgColor="<%=HeaderColor%>">
			<TR>
				<TD colspan=2>&nbsp;&nbsp;<%if IsReverse then response.write "<B>›«ﬂ Ê— »—ê‘  «“ ›—Ê‘</B>"%></TD>
				<TD align="left">‘„«—Â «” ⁄·«„:</TD>
				<TD width=20%>&nbsp;<INPUT readonly class="InvGenInput" NAME="InquiryID" value="<%=InquiryID%>" style="direction:ltr" TYPE="text" maxlength="10" size="10"></TD>
			</TR>
			<TR>
				<TD align="left">Õ”«»:</TD>
				<TD align="right">
					<span id="customer">
						<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><span><A HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=customerID%>" target="_blank"><%=CustomerName%></A></span>
					</span>	<input type="hidden" Name='InquiryID' value='<%=InquiryID%>'></TD>
				<TD align="left"> «—ÌŒ ’œÊ—:</TD>
				<TD width="20%"><TABLE border="0">
					<TR>
						<TD dir="LTR">
							<INPUT readonly class="InvGenInput" NAME="InquiryDate" TYPE="text" maxlength="10" size="10" value="<%=issueDate%>">
						</TD>
						<TD dir="RTL"><%="ø‘‰»Â"%></TD>
					</TR>
					</TABLE></TD>
			</TR>
			<TR>
				<TD align="left" width="100px">„—»Êÿ »Â ”›«—‘:</TD>
				<TD align="right">
					<span id="orders">
<%
					tempWriteAnd=""

					mySQL="SELECT InquiryOrderRelations.[Order], orders_trace.order_kind, orders_trace.order_title FROM InquiryOrderRelations LEFT OUTER JOIN orders_trace ON InquiryOrderRelations.[Order] = orders_trace.radif_sefareshat WHERE (InquiryOrderRelations.Inquiry = '"& InquiryID & "')"

					Set RS1 = conn.Execute(mySQL)
					while not(RS1.eof) 
						response.write "<input type='hidden' name='selectedOrders' value='"& RS1("Order") & "'>"
						response.write tempWriteAnd & Link2Trace(RS1("Order"))
						response.write " [ " & RS1("Order_Title") & " (" & RS1("Order_Kind") & ") ]"
						tempWriteAnd=" Ê "
						RS1.moveNext
					wend
%>					</span>&nbsp;
				</TD>
<%				if IsA then %>	
					<TD align="left">‘„«—Â:</td>
					<TD><TABLE border="0">
						<TR>
							<TD dir="LTR">
								<INPUT readonly class="InvGenInput" NAME="InquiryNo" value="<%=InquiryNo%>" style="border:1px solid black;" TYPE="text" maxlength="10" size="10"></td>
							</TD>
							<td dir="RTL">
									<B>«·›</B> &nbsp;
							</td>
						</TR>
						</TABLE></TD>
<%				else%>
					<TD Colspan="2">&nbsp;</TD>
<%				end if%>
			</TR></TABLE></TD>
		</tr>
		<tr bgcolor="#AAAA55" >
			<td>
			<TABLE Border="0" Cellspacing="0" Cellpadding="5" Dir="RTL" >
			<TR>
				<td width='70' bgcolor="#BBBBBB" align='left'>«ÌÃ«œ ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#BBBBBB" align='right' title="<%=CreationDate%>"><%=Creator%></td>
<%			if Approved then%>
				<td width='70' bgcolor="#77BB99" align='left'> «ÌÌœ ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#77BB99" align='right' title="<%=ApproveDate%>"><%=Approver%></td>
<%			end if
			if Issued then%>
				<td width='70' bgcolor="#7799BB" align='left'>’«œ— ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#7799BB" align='right' title="<%=IssueDate%>"><%=Issuer%></td>
<%			end if
			if Voided then%>
				<td width='70' bgcolor="#BB7799" align='left'>«»ÿ«· ﬂ‰‰œÂ:</td>
				<td width='90' bgcolor="#BB7799" align='right' title="<%=VoidDate%>"><%=Voider%></td>
<%			end if
%>
			</TR>
			</TABLE>
			</td>
		</tr>
		<tr bgcolor='#F0F0F0'>
		<TD colspan="1"><div>
		<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" align="center" class="InvTable">
		<TR bgcolor='#CCCC88'>
			<td align='center' width="25px"> # </td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" value="¬Ì „" size="4" ></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" value=" Ê÷ÌÕ« " size="30"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="ÿÊ·" size="2"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="⁄—÷" size="2"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value=" ⁄œ«œ" size="3"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="›—„" size="2"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value=" ⁄œ«œ „ÊÀ—" size="6"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="›Ì" size="7"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="ﬁÌ„ " size="9"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value=" Œ›Ì›"size="7"></td>
			<td><INPUT readonly class="InvHeadInput" TYPE="text" Value="»—ê‘ " size="5"></td>
			<!-------------------------------SAM----------------------------------------->
			<td><INPUT readonly class="InvHeadInput4" TYPE="text" Value="„«·Ì« " size="6"></td>
			<td><INPUT readonly class="InvHeadInput2" TYPE="text" Value="ﬁ«»· Å—œ«Œ " size="9"></td>
		</TR>
<%		
		i=0
		mySQL="SELECT * FROM InquiryLines WHERE (Inquiry='"& InquiryID & "')"
		Set RS1 = conn.Execute(mySQL)
		while not(RS1.eof) 
			i=i+1
			Price =		cdbl(RS1("Price"))
			AppQtty=	cdbl(RS1("AppQtty"))
			Discount =	cdbl(RS1("Discount"))
			Reverse =	cdbl(RS1("Reverse"))
'	SAM
			Vat =		cdbl(RS1("Vat"))

%>
			<TR bgcolor='#F0F0F0' height="20px">
				<td align='center' width="25px"><%=i%></td>
				<td class="InvRowInput" dir="LTR"><%=RS1("Item")%></td>
				<td class="InvRowInput2" dir="RTL" width="170px"><%=RS1("Description")%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Length")))%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Width")))%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Qtty")))%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(cdbl(RS1("Sets")))%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(AppQtty)%></td>
				<td class="InvRowInput" dir="LTR"><%if AppQtty = 0 then response.write 0 else response.write Separate(Price/AppQtty)%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(Price)%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(Discount)%></td>
				<td class="InvRowInput" dir="LTR"><%=Separate(Reverse)%></td>
				<td class="InvRowInput4" dir="LTR"><%=Separate(Vat)%></td>
				<td class="InvRowInput2" dir="LTR"><%=Separate(Price - Discount - Reverse + Vat)%></td>
			</TR>
<%
			RS1.moveNext
		wend
		RS1.close
%>
		<TR>
			<td colspan="13" height="2px" bgcolor="#CCCC88"></td>
		</TR>
		<TR bgcolor='#CCCC88' height="20px">
			<td colspan="9"> &nbsp; </td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalPrice)%></td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalDiscount)%></td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Separate(TotalReverse)%></td>
			<td align='right' dir="LTR" bgcolor="#FF9900"><%=Separate(TotalVat)%></td>
			<td align='right' dir="LTR" bgcolor="#AAF0FF"><%=Separate(TotalReceivable)%></td>
		</TR>
		<TR bgcolor='#CCCC88' height="20px">
			<td colspan="10"> &nbsp; </td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Pourcent(TotalDiscount,TotalPrice) & "% Œ›Ì›"%></td>
			<td align='right' dir="LTR" bgcolor="#F0F0F0"><%=Pourcent(TotalReverse,TotalPrice) & "%»—ê‘ "%></td>
			<td>&nbsp;</td>
		</TR>
		<TR bgcolor='<%=HeaderColor%>' height="20px">
			<td align=center colspan='13' style="padding:2;">
<%			if voided then
				'This is checked in order to prevent additional checks later.
				'Because when an Inquiry is voided, definitely it is has been issued before.
			elseif Issued then 
				if Auth(6 , "A") then
					' Has the Priviledge to CHANGE the Inquiry / Rev. Inquiry after it has been issued
%>						<input class="GenButton" type="button" value=" «’·«Õ«  Ã“ÌÌ " onclick="window.location='../AR/InquiryEdit.asp?act=editInquiry&inquiry=<%=InquiryID%>';"> 
<%
				end if
				if Auth(6 , "F") then
					' Has the Priviledge to VOID the Inquiry / Rev. Inquiry
%>						<input class="GenButton" style="border:1 solid red;" type="button" value=" «»ÿ«· " onclick="VoidInquiry();"> 
<%
				end if
			else
				' Is not Issued
				if Approved then
					if Auth(6 , "D") then					' Has the Priviledge to ISSUE the Inquiry / Rev. Inquiry
						if Auth(6 , "I") then				
							' can ISSUE the Inquiry / Rev. Inquiry on another Date
%>							<INPUT class="InvGenInput" style="text-align:left;direction:LTR;" TYPE="text" maxlength="10" size="10" TYPE="Text" value="<%=shamsiToday()%>" NAME="IssueDate" OnBlur="return acceptDate(this);">
							<INPUT class="GenButton" TYPE="button" value=" ’œÊ— «” ⁄·«„ " onclick="IssueInquiry();">
<%						else
%>							<INPUT class="GenButton" TYPE="button" value=" ’œÊ— «” ⁄·«„ " onclick="IssueInquiry();">
<%						end if
					end if
				else
					'Is not approved
					if Auth(6 , "C") then					' Has the Priviledge to APPROVE the Inquiry / Rev. Inquiry
%>						<INPUT class="GenButton" TYPE="button" value="  «ÌÌœ «” ⁄·«„ " onclick="ApproveInquiry();">
<%					end if
					
				end if

				if IsReverse AND Auth(6 , 5) then	' Has the Priviledge to EDIT the Rev. Inquiry
%>					<input class="GenButton" type="button" value=" ÊÌ—«Ì‘ " onclick="window.location='../AR/InquiryEdit.asp?act=editInquiry&inquiry=<%=InquiryID%>';"> 
<%				elseif Auth(6 , 3)then				' Has the Priviledge to EDIT the Inquiry
%>					<input class="GenButton" type="button" value=" ÊÌ—«Ì‘ " onclick="window.location='../AR/InquiryEdit.asp?act=editInquiry&inquiry=<%=InquiryID%>';"> 
<%				end if

				if Auth(6 , "G") then
					' Has the Priviledge to REMOVE the Pre-Inquiry / Pre-Rev. Inquiry
%>						<input class="GenButton" style="border:1 solid red;" type="button" value=" Õ–› ÅÌ‘ ‰ÊÌ” " onclick="RemovePreInquiry();"> 
					<%
				end if
			end if

'---------- Remarked by kid 840924 , mr. vazehi wanted so
'
'			if IsReverse AND Auth(6 , 4) then	' Has the Priviledge to INPUT a Rev. Inquiry
'% >				<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InvoiceInput.asp?act=copyInvoice&invoice=< %=InvoiceID% >';"> 
'< %			elseif Auth(6 , 1)then				' Has the Priviledge to INPUT an Inquiry
'% >				<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InvoiceInput.asp?act=copyInvoice&invoice=< %=InvoiceID% >';"> 
'< %			end if
'
'--------- End of Remark

'---------- Added by kid 850816, they requested to have the copy option but with priviledges
'
			if Auth(6 , "L") then	' Has the Priviledge to COPY an Inquiry / Rev. Inquiry
				if IsReverse AND Auth(6 , 4) then	' Has the Priviledge to INPUT a Rev. Inquiry
%>					<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InquiryInput.asp?act=copyInquiry&inquiry=<%=InquiryID%>';"> 
<%				elseif Auth(6 , 1)then				' Has the Priviledge to INPUT an Inquiry
%>					<input class="GenButton" type="button" value="ﬂÅÌ ÅÌ‘ ‰ÊÌ”" onclick="window.location='../AR/InquiryInput.asp?act=copyInquiry&inquiry=<%=InquiryID%>';"> 
<%				end if
			end if
'---------- End of 850816 additions


			if Auth(6 , "E") then					' Has the Priviledge to PRINT the Inquiry / Rev. Inquiry
%>				
				<% 	ReportLogRow = PrepareReport ("InquiryNew.rpt", "Inquiry_ID", InquiryID, "/beta/dialog_printManager.asp?act=Fin") %>
				<INPUT Class="GenButton" style="border:1 solid blue;" TYPE="button" value=" ç«Å " 
				onclick="window.location='../AR/InquiryPrint.asp?r=<%=ReportLogRow%>';">
<!--				onclick="printThisReport(this,<%=ReportLogRow%>);"-->
<%			'----------------------------------------------SAM-----------------------------------------------------
			if Issued then %>
				<INPUT class='GenButton' style='border:1 solid blue;' type='button' value=' Ê·Ìœ ﬂœ Å—œ«Œ  «Ì‰ —‰ Ì' onclick="window.location='../AR/ePayment.asp?Inquiry=<%=InquiryID%>';">
<%			end if
			'------------------------------------------------------------------------------------------------------
			end if
%>
			</td>
		</TR>
		</TABLE>
		</div></TD>
		</tr>
	</table>
	<br>
	<table align=center><tr>
<%
	mySQL="SELECT InquiryOrderRelations.*, PurchaseOrders.ID AS PurchaseOrdersID, PurchaseOrders.Status AS PurchaseOrdersStatus, PurchaseOrders.TypeName AS TypeName FROM PurchaseOrders FULL OUTER JOIN PurchaseRequestOrderRelations RIGHT OUTER JOIN PurchaseRequests INNER JOIN InquiryOrderRelations ON PurchaseRequests.Order_ID = InquiryOrderRelations.[Order] ON PurchaseRequestOrderRelations.Req_ID = PurchaseRequests.ID ON PurchaseOrders.ID = PurchaseRequestOrderRelations.Ord_ID WHERE (InquiryOrderRelations.Inquiry ='"& InquiryID & "') and PurchaseRequests.Status<> 'del'"
	Set RS1 = conn.Execute(mySQL)
	tempWriteAnd = "Œ—Ìœ Œœ„«  „—»Êÿ »Â «Ì‰ «” ⁄·«„: <hr>"
	if not(RS1.eof) then
%>
		<td valign=top style="border:solid 1pt black">
			<div align=right>
<%
		while not(RS1.eof) 
			PurchaseOrdersStatus=RS1("PurchaseOrdersStatus")
			if PurchaseOrdersStatus="NEW" then
				status = "ÃœÌœ"
			elseif PurchaseOrdersStatus="OUT" then
				status = "Œ«—Ã «“ ‘—ﬂ "
			elseif PurchaseOrdersStatus="RETURN" then
				status = "»—ê‘ Â »Â ‘—ﬂ "
			elseif PurchaseOrdersStatus="CANCEL" then
				status = "·€Ê ‘œÂ"
			elseif PurchaseOrdersStatus="OK" then
				status = " «ÌÌœ ‘œÂ"
			end if
			response.write tempWriteAnd & "<li> <a target='_blank' href='../purchase/outServiceTrace.asp?od="&RS1("PurchaseOrdersID")&"'>" & RS1("TypeName") & "</a> - " & status
			tempWriteAnd=""
			RS1.moveNext
		wend
%>	
			<br>	
			</div>
		</td>
<%
	end if
%>
		<td width=30>
		</td>
<%
	mySQL="SELECT InquiryOrderRelations.*, InventoryPickuplistItems.ItemName AS ItemName, InventoryPickuplistItems.pickupListID AS pickupListID, InventoryPickuplistItems.Qtty AS Qtty, InventoryPickuplistItems.unit AS unit, InventoryPickuplistItems.CustomerHaveInvItem FROM InventoryPickuplistItems INNER JOIN InventoryPickuplists ON InventoryPickuplistItems.pickupListID = InventoryPickuplists.id INNER JOIN InquiryOrderRelations ON InventoryPickuplistItems.Order_ID = InquiryOrderRelations.[Order] WHERE (InquiryOrderRelations.Inquiry = '"& InquiryID & "' AND NOT InventoryPickuplists.Status = N'del')"


	Set RS1 = conn.Execute(mySQL)
	tempWriteAnd = "ÕÊ«·Â Â«Ì «‰»«— „—»Êÿ »Â «Ì‰ ›«ﬂ Ê—: <hr>"
	if not(RS1.eof) then
%>
		<td valign=top style="border:solid 1pt black">
			<div align=right>
<%
		while not(RS1.eof) 
			response.write tempWriteAnd & "<li><a target='_blank' href='../inventory/default.asp?show="&RS1("pickupListID")&"'>" & RS1("ItemName") & "</a> - " & RS1("Qtty") & RS1("unit")
			if RS1("CustomerHaveInvItem") then
			response.write " <B>(«—”«·Ì</B>)" 
			end if
			tempWriteAnd=""
			RS1.moveNext
		wend
%>
			<br>	
			</div>
		</td>
<%	end if
%>	
	</tr></table>
	<br>
	<table class="CustTable" cellspacing='1' align=center>
		<tr>
			<td colspan="2" class="CusTableHeader"><span style="width:450;text-align:center;">Ì«œœ«‘  Â«</span><span style="width:100;text-align:left;background-color:red;"><input class="GenButton" type="button" value="‰Ê‘ ‰ Ì«œœ«‘ " onclick="window.location = '../home/message.asp?RelatedTable=Inquiries&RelatedID=<%=InquiryID%>&retURL=<%=Server.URLEncode("../Inquiry/AccountReport.asp?act=showInquiry&inquiry="&InquiryID)%>';"></span></td>
		</tr>
<%
	mySQL="SELECT * FROM Messages INNER JOIN Users ON Messages.MsgFrom = Users.ID WHERE (Messages.RelatedTable = 'Inquiries') AND (Messages.RelatedID = "& InquiryID & ") ORDER BY Messages.ID DESC"
	Set RS1 = conn.execute(mySQL)
	if NOT RS1.eof then
%>
<%
		tmpCounter=0
		Do While NOT RS1.eof 
			tmpCounter=tmpCounter+1
%>
			<tr class="<%if (tmpCounter MOD 2) = 1 then response.write "CusTD3" else response.write "CusTD4" %>">
				<td>«“ <%=RS1("RealName")%><br>
					<%=RS1("MsgDate")%> <BR> <%=RS1("MsgTime")%>
				</td>
				<td dir='RTL'><%=replace(RS1("MsgBody"),chr(13),"<br>")%></td>
			</tr>
<%
			RS1.moveNext
		Loop
	else
%>
		<tr class="CusTD3">
			<td colspan="2">ÂÌç</td>
		</tr>
<%
	end if
	RS1.close
%>
	</table><BR><BR>
	<br><!--CENTER>ÅÌ‘ ‰„«Ì‘ ç«Å:</CENTER-->
	<BR>
	<script language="JavaScript">
	<!--
	function ApproveInquiry(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« ' «ÌÌœ' ﬂ‰Ìœø\n\n"))
			window.location="../AR/InquiryEdit.asp?act=approveInquiry&inquiry=<%=InquiryID%>";
	}
	function IssueInquiry(){
	<% if Auth(6 , "I") then%>
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '’«œ—' ﬂ‰Ìœø\n\n"))
			window.location="../AR/InquiryEdit.asp?act=IssueInquiry&inquiry=<%=InquiryID%>&issueDate="+document.all.IssueDate.value;
	<% else %>
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '’«œ—' ﬂ‰Ìœø\n\n"))
			window.location="../AR/InquiryEdit.asp?act=IssueInquiry&inquiry=<%=InquiryID%>";
	<% end if%>
	}
	function VoidInquiry(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '»«ÿ·' ﬂ‰Ìœø\n")){

			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value=" Ê÷ÌÕÌ œ— »«—Â œ·Ì· «»ÿ«· ›«ﬂ Ê—:"
			window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false
			window.location="../AR/InquiryEdit.asp?act=voidInquiry&inquiry=<%=InquiryID%>&comment="+escape(document.all.tmpDlgTxt.value);
		}
	}
	function RemovePreInquiry(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ÅÌ‘ ‰ÊÌ” —« Õ–› ﬂ‰Ìœø\n"))
			window.location='../AR/InquiryEdit.asp?act=removePreInquiry&inquiry=<%=InquiryID%>';
	}
	//-->
	</script>
<%
'----------------------------------------- E N D   I N Q U I R Y -------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showReceipt" then
	ReceiptID=request("receipt")
	if Trim(ReceiptID)="" OR NOT isnumeric(ReceiptID) then
		conn.close
		response.redirect "top.asp?errmsg=" & Server.URLEncode("Œÿ«! œ— ‘„«—Â œ—Ì«› .")
	end if
	ReceiptID=clng(ReceiptID)

	'mySQL="SELECT Receipts.*, ReceivedCash.Description FROM Receipts LEFT OUTER JOIN ReceivedCash ON Receipts.ID = ReceivedCash.Receipt WHERE (Receipts.ID = '"& ReceiptID & "')"
	'Changed By Kid 840503, Adding the CashRegister in wich the receipt is registered
	mySQL="SELECT Receipts.*, ReceivedCash.Description, CashRegisters.ID AS CashReg, CashRegisters.Cashier, CashRegisters.NameDate, CashRegisters.IsOpen FROM CashRegisters INNER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg RIGHT OUTER JOIN Receipts ON CashRegisterLines.Link = Receipts.ID LEFT OUTER JOIN ReceivedCash ON Receipts.ID = ReceivedCash.Receipt WHERE (Receipts.ID = '"& ReceiptID & "')"
	
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "top.asp?errmsg=" & Server.URLEncode("Œÿ«! <br> ç‰Ì‰ œ—Ì«› Ì ÊÃÊœ ‰œ«—œ.")
	end if

	sys=RS1("sys")
	mySQL2 = "SELECT "& sys & "Items.ID, CreatedDate, CreationTime, EffectiveDate, Voided, Users.RealName FROM Users INNER JOIN "& sys & "Items ON Users.ID = "& sys & "Items.CreatedBy WHERE (Type = 2) AND (Link = "& ReceiptID & ")"
	Set RS2 = conn.Execute(mySQL2)
	if not rs2.eof then
		AnyItemID =		RS2("ID")
		voided =		RS2("voided")
		EffectiveDate=	RS2("EffectiveDate")
		Creator=		RS2("RealName")
		CreatedDate=	RS2("CreatedDate")
		CreationTime=	RS2("CreationTime")
	end if
	RS2.close

	CashReg= 		RS1("CashReg") & ""
	Cashier=		RS1("Cashier")
	IsOpen =		RS1("IsOpen")
	NameDate =		RS1("NameDate") & ""

	customerID=		RS1("Customer")
	TotalAmount=	cdbl(RS1("TotalAmount"))
	CashAmount=		cdbl(RS1("CashAmount"))
	DepositAmount=	clng(RS1("DepositAmount"))

	'-----------------
	' --- This Must be removed !!! 
	' --- Mistake in data schema design, Receipt description must be stored in [Receipts] not in [ReceivedCash]
	Description = RS1("Description")
	'------------------


	mySQL="SELECT ID,AccountTitle FROM Accounts WHERE (ID='"& customerID & "')"

	Set RS1 = conn.Execute(mySQL)
	AccountNo=RS1("ID")
	customerName=RS1("AccountTitle")
	RS1.close

	if Voided then
		stamp="<img src='/images/voided.gif'>"
	else
		stamp=""
	end if 

	if IsOpen="True" AND Cashier=session("ID") then
		CashRegLink="<A Target='_blank' HREF='../cashReg/CashRegReport.asp'>Â‰Ê“ œ— ’‰œÊﬁ </A>"
	elseif CashReg <> "" AND Auth(9 , 6) then		'ê“«—‘ ”—Å—”  ’‰œÊﬁ
		CashRegLink="<A Target='_blank' HREF='../cashReg/CashRegAdminReport.asp?act=showCashRegReport&CashRegID=" & CashReg & "'>’‰œÊﬁ " & replace(NameDate,"/",".") & "</A>"
	else
		CashRegLink=""
	end if
%>
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="10">
			<td width="150"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="50">
			<td colspan=2></td>
		</tr>
	</table>
	<TABLE class="RcpTable" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" align="center" style="border:1 solid navy;">
	<tr bgcolor='#CCCCEE' height="40px">
		<td colspan="6" align='center' style="border-bottom:1 solid navy;"><A HREF="../cashReg/ReceiptInput.asp?act=ShowReceipt&id=<%=ReceiptID%>">—”Ìœ œ—Ì«›  ‘„«—Â <%=ReceiptID%></A><br>„—»Êÿ »Â <a href="?act=show&selectedCustomer=<%=customerID%>"><%=CustomerName%></a>.</td>
	</tr>
	<tr bgcolor='#CCCCEE' height="20px">
		<td align='left'> À» : </td>
		<td align='right' dir="LTR">&nbsp;<%=CreatedDate & " " & CreationTime%>&nbsp;</td>
		<td colspan='2' align='center'><%=CashRegLink%>&nbsp;</td>
		<td align='left'> À»  ﬂ‰‰œÂ: </td>
		<td align='right'>&nbsp;<%=Creator%>&nbsp;</td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="5" align='left'>  «—ÌŒ („ÊÀ—): </td>
		<td align='right' dir="LTR">&nbsp;<%=EffectiveDate%>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="5"> „»·€ ‰ﬁœ: </td>
		<td align='right' dir="LTR" class="RcpRowInput"><%=Separate(CashAmount)%></td>
	</tr>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="6"> çﬂ: <br></td>
	</tr>
<%
	i=0
	mySQL="SELECT * FROM ReceivedCheques WHERE (Receipt='"& ReceiptID & "')"
	Set RS1 = conn.Execute(mySQL)
	if not(RS1.eof) then
%>
	<tr>
		<td class="RcpHeadInput" align='center' width="25px"> # </td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value="‘„«—Â çﬂ" size="10" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" value=" «—ÌŒ" size="10" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="»«‰ﬂ" size="10" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value=" Ê÷ÌÕ" size="15" tabindex="9999"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" Value="„»·€" size="15" tabindex="9999"></td>
	</tr>
<%
	else
%>
	<tr class="RcpHeadInput" >
		<td align='center' width="25px"> - - </td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="10" value="‰œ«—œ"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="10"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="10"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="15"></td>
		<td class="RcpHeadInput"><INPUT class="RcpHeadInput" readonly TYPE="text" size="15"></td>
	</tr>
<%
	end if
	while not(RS1.eof) 
		i=i+1
%>
			<tr bgcolor='#F0F0F0' height="20">
				<td align='center' width="25px"><%=i%></td>
				<td dir="LTR" class="RcpRowInput"><%=RS1("ChequeNo")%></td>
				<td dir="LTR" class="RcpRowInput"><%=RS1("ChequeDate")%></td>
				<td dir="RTL" class="RcpRowInput"><%=RS1("BankOfOrigin")%></td>
				<td dir="RTL" class="RcpRowInput"><%=RS1("Description")%></td>
				<td dir="LTR" class="RcpRowInput"><%=Separate(cdbl(RS1("Amount")))%></td>
			</tr>
<%
		RS1.moveNext
	wend %>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height="20px">
		<td colspan="6">  Ê÷ÌÕ« : <%=Description%> <br></td>
	</tr>
	<tr>
		<td colspan="6"></td>
	</tr>
	<tr bgcolor='#CCCC88' height='20px'>
		<td colspan="5" align="left"> Ã„⁄:&nbsp;&nbsp;</td>
		<td align='right' style="font-size: 9pt; direction:LTR; background-color=#C0D0F0;"><%=Separate(TotalAmount)%></td>
	</tr>
<%	if Auth(9 , 7) AND NOT voided then			' Has the Priviledge to VOID the RECEIPT/PAYMENT %>
		<tr bgcolor='#CCCC88' height='30px'>
			<td colspan="6" align='center'><INPUT class="GenButton" TYPE="button" Value=" «»ÿ«· " onclick="VoidReceipt();"></td>
		</tr>
<%	end if%>
	</TABLE>
	<Br>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function VoidReceipt(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ œ—Ì«›  —« '»«ÿ·' ﬂ‰Ìœø\n\n"))
			window.location="../cashReg/Void.asp?act=voidReceipt&Receipt=<%=ReceiptID%>";
	}
	//-->
	</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showMemo" AND request("memo") <> "" then
	sys = request("sys")
	if sys = "" then sys = "AR"
	MemoID=request("memo")
	if not(isnumeric(MemoID)) then
			response.end
	end if
	mySQL="SELECT AXMemoTypes.Name AS TypeName, Users.RealName AS Creator, "& sys & "Memo.* FROM "& sys & "Memo INNER JOIN AXMemoTypes ON "& sys & "Memo.Type = AXMemoTypes.ID INNER JOIN Users ON "& sys & "Memo.CreatedBy = Users.ID WHERE ("& sys & "Memo.ID = '"& MemoID& "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
%>
		<table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> ÅÌœ« ‰‘œ <br></td></tr></table>
<%			response.end
	end if

	Set RS2 = conn.Execute("SELECT ID,EffectiveDate,Voided,Link FROM "& sys & "Items WHERE (Type = 3) AND (Link = "& MemoID & ")")
	AnyItemID = RS2("ID")
	Voided = RS2("Voided")
	EffectiveDate=RS2("EffectiveDate")
	Item=RS2("ID")
	RS2.close
	'---------------SAM-----------------------
	'response.write("SELECT glDoc FROM EffectiveGLRows WHERE sys = '"&sys&"' AND Link ="& Item)
	set RS2 = conn.Execute("SELECT glDoc FROM EffectiveGLRows WHERE sys = '"&sys&"' AND Link ="& Item)
	if NOT RS2.EOF then 
		glDoc = RS2("glDoc")
	else 
		glDoc = 0
	end if
	RS2.close
	
	customerID=RS1("Account")
	creationDate=RS1("CreatedDate")
	Creator=RS1("Creator")
	MemoTypeName=RS1("TypeName")
	Description=RS1("Description")
	Amount=RS1("Amount")

	if RS1("IsCredit")=true then
		MemoCreDeb="»” «‰ﬂ«—"
	else
		MemoCreDeb="»œÂﬂ«—"
	end if
	'--------------------SAM-------------------
	if RS1("Type")=7 then 
		if RS1("IsCredit") then
			mySQL = "select fromItemLink as linkedMemoID, fromItemType as linkedMemoSYS from InterMemoRelation where toItemLink=" & memoID & " and toItemType='" & sys & "'"
		else 
			mySQL = "select toItemLink as linkedMemoID, toItemType as linkedMemoSYS from InterMemoRelation where fromItemLink=" & memoID & " and fromItemType='" & sys & "'"
		end if
		set RS21 = conn.Execute(mySQL)
		linkedMemoID=0
		linkedMemoSYS=""
		if NOT RS21.eof then 
			linkedMemoID = CDbl(RS21("linkedMemoID"))
			linkedMemoSYS = RS21("linkedMemoSYS")
		end if
		RS21.close
	end if
	mySQL="SELECT * FROM Accounts WHERE (ID='"& customerID & "')"

	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")

	RS1.close

	if Voided then
		stamp="<img src='/images/voided.gif'>"
	else
		stamp=""
	end if 

%>
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="10">
			<td width="1"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="50">
			<td colspan=2></td>
		</tr>
	</table>
	<table class="MmoMainTable" Cellspacing="1" Cellpadding="5" Width="500" align="center">
		<tr class="MmoMainTableTH">
		<TD colspan="10"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="0"><TR>
			<TD>Õ”«»: <a href="?act=show&selectedCustomer=<%=customerID%>"><%=CustomerName%></a>.</TD>
			<TD>«ÌÃ«œ ﬂ‰‰œÂ: <%=Creator%>.</TD>
			<TD align="left"> «—ÌŒ(„ÊÀ—):</TD>
			<TD width="50"><TABLE class="MmoTable">
				<TR>
					<TD dir="LTR">
						<INPUT readonly class="MmoGenInput" TYPE="text" value="<%=EffectiveDate%>" size="10">
					</TD>
					<TD dir="RTL"> </TD>
				</TR>
				</TABLE></TD>
			</TR></TABLE></TD>
		</tr>
		<tr class="MmoMainTableTH" height="25">
			<TD> ‰Ê⁄ «⁄·«„ÌÂ </TD>
			<TD> »œÂﬂ«—/»” «‰ﬂ«— </TD>
			<TD> ‘—Õ </TD>
			<TD> „»·€ </TD>
		</tr>
		<tr class="MmoMainTableTR">
			<TD valign="top">
				<INPUT readonly class="MmoRowInput" size="10" value="<%=MemoTypeName%>">
				<% '---------------------------------------------SAM--------------------------------------------
				if glDoc>0 then
					response.write("<div><a href='../accounting/GLMemoDocShow.asp?id=" & glDoc & "'>‰„«Ì‘ ”‰œ Õ”«»œ«—Ì</a></div>")
				end if
				if linkedMemoID > 0 then 
					response.write("<div><a href='AccountReport.asp?act=showMemo&sys=" & linkedMemoSYS & "&memo=" & linkedMemoID & "'>‰„«Ì‘ «⁄·«„ÌÂ „ ‰«Ÿ—</a></div>")
				else 
					response.write("<div>Œÿ«Ì ⁄ÃÌ»! «⁄·«„ÌÂ „ ‰«Ÿ— ÅÌœ« ‰‘œ!</div>")
				end if
				%>
			</TD>
			<TD valign="top"><INPUT readonly class="MmoRowInput" size="10" value="<%=MemoCreDeb%>"></TD>
			<TD valign="top"><TEXTAREA readonly class="MmoRowInput" dir="RTL" ROWS="3" COLS="30"><%=Description%></TEXTAREA></TD>
			<TD valign="top"><INPUT readonly class="MmoRowInput" Dir="LTR" TYPE="text" size="15" value="<%=Separate(Amount)%>"></TD>
		</tr>
<%		
		if NOT voided then
			if	(sys="AR" AND Auth(6 , "H")) OR (sys="AP" AND Auth(7 , 8 )) OR (sys="AO" AND Auth("B" , 6 )) then
				' Has the Priviledge to VOID the AR/AP/AO MEMO %>
				<tr bgcolor='#CCCC88' height='30px'>
					<td colspan="5" align='center'><INPUT class="GenButton" TYPE="button" Value=" «»ÿ«· " onclick="VoidMemo();"></td>
				</tr>
<%			end if
		end if%>
	</table>
	<br>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function VoidMemo(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ «⁄·«„ÌÂ —« '»«ÿ·' ﬂ‰Ìœø\n\n"))
			window.location="../AO/MemoVoid.asp?act=VoidMemo&Sys=<%=sys%>&Memo=<%=MemoID%>";
	}
	//-->
	</SCRIPT>
<%
end if

'---- Show Relations
if not AnyItemID ="" then

	dim typeNamesArray(6) 
	typeNamesArray(1) = "›«ﬂ Ê—"
	typeNamesArray(2) = "œ—Ì«› "
	typeNamesArray(3) = "«⁄·«„ÌÂ"
	typeNamesArray(4) = "»—ê‘ "
	typeNamesArray(5) = "Å—œ«Œ "
	typeNamesArray(6) = "›«ﬂ Ê— Œ—Ìœ"

	response.write "<center>"
	mySQL2="SELECT "& sys & "ItemsRelations.ID, "& sys & "ItemsRelations.Amount, "&sys & "Items.ID AS ItemID, "&sys & "Items.Type, "&sys & "Items.Link FROM "& sys & "ItemsRelations INNER JOIN "&sys & "Items ON "& sys & "ItemsRelations.Debit"& sys & "Item = "& sys & "Items.ID WHERE ("& sys & "ItemsRelations.Credit"& sys & "Item = "& AnyItemID & ") "
	Set RS2 = conn.Execute(mySQL2)

	if NOT RS2.EOF then
%>		<TABLE border=1 cellpadding=2 cellspacing=1 bgColor=#000099>
		<TR bgColor=#CCCCEE>
			<TD colspan=3>—«»ÿÂ Â« (œÊŒ Â ‘œÂ)</TD>
		</TR>
<%
		while Not RS2.EOF
'			retURL=Server.URLEncode("AccountReport.asp?msg=—«»ÿÂ Õ–› ‘œ.")
%>
		<TR>
			<TD bgColor=#DDDDDD><A style='color:red;' HREF="ItemsRelation.asp?act=unrelate&sys=<%=sys%>&relation=<%=RS2("ID")%>"><B>x</B></A></TD>
			<TD bgColor=#F0F0F0 align=right><%=separate(cdbl(RS2("Amount")))%> —Ì«·</TD>
			<TD bgColor=#DDDDDD style='cursor:hand;' onclick="window.open('../accounting/ShowItem.asp?sys=<%=sys%>&Item=<%=RS2("ItemID")%>');"><%=typeNamesArray(cint(RS2("Type"))) & " ‘„«—Â "& RS2("Link")%></TD>
		</TR>
<%

			RS2.movenext
		wend
%>
		</TABLE>
		<br><br>
<%
	end if
	RS2.close

	mySQL2="SELECT "& sys & "ItemsRelations.ID, "& sys & "ItemsRelations.Amount, "&sys & "Items.ID AS ItemID, "&sys & "Items.Type, "&sys & "Items.Link FROM "& sys & "ItemsRelations INNER JOIN "&sys & "Items ON "& sys & "ItemsRelations.Credit"& sys & "Item = "& sys & "Items.ID WHERE ("& sys & "ItemsRelations.Debit"& sys & "Item = "& AnyItemID & ") "
	Set RS2 = conn.Execute(mySQL2)

	if NOT RS2.EOF then
%>		<TABLE border=1 cellpadding=2 cellspacing=1 bgColor=#000099>
		<TR bgColor=#CCCCEE>
			<TD colspan=3>—«»ÿÂ Â« (œÊŒ Â ‘œÂ)</TD>
		</TR>
<%
		while Not RS2.EOF
'			retURL=Server.URLEncode("AccountReport.asp?msg=—«»ÿÂ Õ–› ‘œ.")
%>
		<TR>
			<TD bgColor=#DDDDDD><A style='color:red;' HREF="ItemsRelation.asp?act=unrelate&sys=<%=sys%>&relation=<%=RS2("ID")%>"><B>x</B></A></TD>
			<TD bgColor=#F0F0F0 align=right><%=separate(cdbl(RS2("Amount")))%> —Ì«·</TD>
			<TD bgColor=#DDDDDD style='cursor:hand;' onclick="window.open('../accounting/ShowItem.asp?sys=<%=sys%>&Item=<%=RS2("ItemID")%>');"><%=typeNamesArray(cint(RS2("Type"))) & " ‘„«—Â "& RS2("Link")%></TD>
		</TR>
<%

			RS2.movenext
		wend
%>
		</TABLE>
		<br><br>
<%
	end if

	response.write "</center>"

end if
conn.Close

'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
%>
</font>
