<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , 8) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
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
<%
if request("act")="show" then

	'ON ERROR RESUME Next
	FromDate =		sqlSafe(request("FromDate"))
	ToDate = 		sqlSafe(request("ToDate"))
	Ord =			request("Ord")
	If Ord="" Then Ord = 0

	'mySQL="SELECT account, accountTitle, sum(case arItems.[type] when 1 then amountOriginal else 0 end) as totalDebit, sum(case arItems.[type] when 2 then amountOriginal else 0 end) as totalCredit, count(account) as Items FROM arItems inner join accounts on arItems.account = accounts.id where voided = 0 and effectiveDate between N'" & FromDate & "' and N'" & ToDate & "' group by account,accountTitle" & order


	select case Ord
	case "1":
		order="account"
	case "-1":
		order="account DESC"
	case "2":
		order="AccountTitle"
	case "-2":
		order="AccountTitle DESC"
	case "3":
		order="totalDebit"
	case "-3":
		order="totalDebit DESC"
	case "4":
		order="totalCredit"
	case "-4":
		order="totalCredit DESC"
	Case "5":
		order = "remainDebit"
	Case "-5":
		order = "remainDebit DESC"
	case "6":
		order="remainCredit"
	case "-6":
		order="remainCredit DESC"
	case else:
		order="account"
		Ord=1
	end select

	
%>
	<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
	
<%
	mySQL = "select sum(remDebit) as sumRemDebit, sum(remCredit) as sumRemCredit, sum(totalDebit + remDebit) as sumTotalDebit, sum(totalCredit + remCredit) as sumTotalCredit, sum(case when (totalDebit+remDebit)>(totalCredit+remCredit) then (totalDebit+remDebit)-(totalCredit+remCredit) else 0 end) as remainDebit, sum(case when (totalDebit+remDebit)<(totalCredit+remCredit) then (totalCredit+remCredit)-(totalDebit+remDebit) else 0 end) as remainCredit, sum(Items) as totalItems from (SELECT CASE WHEN remTbl.amount > 0 THEN 0 ELSE - remTbl.amount END AS remDebit, CASE WHEN remTbl.amount > 0 THEN remTbl.amount ELSE 0 END AS remCredit, account, accountTitle, sum(case arItems.IsCredit when 0 then amountOriginal else 0 end) as totalDebit, sum(case arItems.IsCredit when 1 then amountOriginal else 0 end) as totalCredit, count(account) as Items FROM arItems inner join accounts on arItems.account = accounts.id LEFT OUTER JOIN (SELECT account as remAccount, ISNULL(SUM(AmountOriginal *(CONVERT(tinyint, IsCredit) - .5) * 2 ), 0) as amount from arItems WHERE voided = 0 AND EffectiveDate < N'"&FromDate&"' GROUP BY account) remTbl on remTbl.remAccount = arItems.account WHERE voided = 0 and effectiveDate between N'" & FromDate & "' and N'" & ToDate & "' GROUP BY account,accountTitle,remTbl.amount) drvTBL"
'response.write mySQL
	set rs=Conn.Execute (mySQL)
	if rs.eof then
%>	<tr>
		<td bgcolor="#BBBBBB" height="30" colspan="7" align=center><b>ÂÌç .</b></td>
	</tr>
<%	Else %>
	<TR bgcolor="#CCCCEE" >
		<TD colspan=2 rowspan=2 title=" <%=rs("totalItems")%> ¬Ì „ ">
			«“  «—ÌŒ <B><%=replace(FromDate,"/",".")%></B>  «  «—ÌŒ <B><%=replace(ToDate,"/",".")%></B><br>
			«“  ›’Ì·Ì <B><%=FromTafsil%></B>  «  ›’Ì·Ì <B><%=ToTafsil%></B>
		</TD>
		<TD width=70 >ê—œ‘ »œÂﬂ«—		</TD>
		<TD width=70 >ê—œ‘ »” «‰ﬂ«—		</TD>
		<TD width=70 >„«‰œÂ »œÂﬂ«—		</TD>
		<TD width=70 >„«‰œÂ »” «‰ﬂ«—	</TD>
	</TR>
	<TR bgcolor="#CCCCEE" >
		<TD title='<%=Separate(rs("sumRemDebit"))%> »œÂò«—Ì ÅÌ‘Ì‰ ' width=70 ><%=Separate(rs("sumTotalDebit"))%></TD>
		<TD title='<%=Separate(rs("sumRemCredit"))%> »” «‰ò«—Ì ÅÌ‘Ì‰ ' width=70 ><%=Separate(rs("sumTotalCredit"))%></TD>
		<TD width=70 ><%=Separate(rs("remainDebit"))%></TD>
		<TD width=70 ><%=Separate(rs("remainCredit"))%></TD>
	</TR>
	<TR bgcolor="black" height="2">
		<TD colspan="6" style="padding:0;"></TD>
	</TR>
<%
	rs.close

	if ord<0 then
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>6 6 6</span>"
	else
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>5 5 5</span>"
	end if
%>
	<TR bgcolor="eeeeee" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
		<TD width=50  onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>">‘„«—Â Õ”«»			<%if abs(ord)=1 then response.write arrow%></TD>
		<TD width='*' onclick='go2Page(1,2);' style="<%if abs(ord)=2 then response.write style%>">⁄‰Ê«‰ Õ”«»		<%if abs(ord)=2 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-3);' style="<%if abs(ord)=3 then response.write style%>">ê—œ‘ »œÂﬂ«—		<%if abs(ord)=3 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-4);' style="<%if abs(ord)=4 then response.write style%>">ê—œ‘ »” «‰ﬂ«—		<%if abs(ord)=4 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-5);' style="<%if abs(ord)=5 then response.write style%>">„«‰œÂ »œÂﬂ«—		<%if abs(ord)=5 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-6);' style="<%if abs(ord)=6 then response.write style%>">„«‰œÂ »” «‰ﬂ«—	<%if abs(ord)=6 then response.write arrow%></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
	</TR>
<%		
	SumCredit=0
	SumDebit=0
	SumCreditRemained=0
	SumDebitRemained=0
	tmpCounter=0

	mySQL="SELECT CASE WHEN remTbl.amount > 0 THEN 0 ELSE - remTbl.amount END AS remDebit, CASE WHEN remTbl.amount > 0 THEN remTbl.amount ELSE 0 END AS remCredit, account, accountTitle, sum(case arItems.IsCredit when 0 then amountOriginal else 0 end) + CASE WHEN remTbl.amount > 0 THEN 0 ELSE - remTbl.amount END as totalDebit, sum(case arItems.IsCredit when 1 then amountOriginal else 0 end) + CASE WHEN  remTbl.amount > 0 THEN remTbl.amount ELSE 0 END as totalCredit, case when (sum(case arItems.IsCredit when 0 then amountOriginal else 0 end) + CASE WHEN remTbl.amount > 0 THEN 0 ELSE - remTbl.amount END) > (sum(case arItems.IsCredit when 1 then amountOriginal else 0 end) + CASE WHEN  remTbl.amount > 0 THEN remTbl.amount ELSE 0 END) then (sum(case arItems.IsCredit when 0 then amountOriginal else 0 end) + CASE WHEN remTbl.amount > 0 THEN 0 ELSE - remTbl.amount END)-(sum(case arItems.IsCredit when 1 then amountOriginal else 0 end) + CASE WHEN  remTbl.amount > 0 THEN remTbl.amount ELSE 0 END) else 0 end as remainDebit, case when (sum(case arItems.IsCredit when 0 then amountOriginal else 0 end) + CASE WHEN remTbl.amount > 0 THEN 0 ELSE - remTbl.amount END) < (sum(case arItems.IsCredit when 1 then amountOriginal else 0 end) + CASE WHEN  remTbl.amount > 0 THEN remTbl.amount ELSE 0 END) then -(sum(case arItems.IsCredit when 0 then amountOriginal else 0 end) + CASE WHEN remTbl.amount > 0 THEN 0 ELSE - remTbl.amount END)+(sum(case arItems.IsCredit when 1 then amountOriginal else 0 end) + CASE WHEN  remTbl.amount > 0 THEN remTbl.amount ELSE 0 END) else 0 end as remainCredit, count(account) as Items FROM arItems inner join accounts on arItems.account = accounts.id LEFT OUTER JOIN (SELECT account as remAccount, ISNULL(SUM(AmountOriginal *(CONVERT(tinyint, IsCredit) - .5) * 2 ), 0) as amount from arItems WHERE voided = 0 AND EffectiveDate < N'"&FromDate&"' GROUP BY account) remTbl on remTbl.remAccount = arItems.account WHERE voided = 0 and effectiveDate between N'" & FromDate & "' and N'" & ToDate & "' GROUP BY account,accountTitle,remTbl.amount ORDER BY " & order
	
	Set rs=Server.CreateObject("ADODB.Recordset")'Conn.Execute(mySQL)

	PageSize = 50
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
			<td bgcolor="#BBBBBB" height="30" colspan="7" align=center><b>ÂÌç .</b></td>
		</tr>
<%	else
		Do While NOT rs.eof AND (rs.AbsolutePage = CurrentPage)
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
				<TD dir=ltr align=right><A href="AccountReport.asp?act=show&reason=6&sys=AR&selectedCustomer=<%=rs("account")%>&startDate=<%=FromDate%>&endDate=<%=ToDate%>"><%=rs("account")%></A></TD>
				<TD title=" <%=rs("Items")%> ¬Ì „ "><%=rs("accountTitle")%></TD>
				<TD title='<%=Separate(rs("remDebit"))%> »œÂò«—Ì ÅÌ‘Ì‰ ' dir=ltr align=right><span dir=ltr><%=Separate(rs("totalDebit"))%></span></TD>
				<TD title='<%=Separate(rs("remCredit"))%> »” «‰ò«—Ì ÅÌ‘Ì‰ ' dir=ltr align=right><span dir=ltr><%=Separate(rs("totalCredit"))%></span></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(rs("remainDebit"))%></span></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(rs("remainCredit"))%></span></TD>
			</TR>
			  
	<% 
		rs.moveNext
		Loop

		if TotalPages > 1 then
			pageCols=20
%>			
			<TR bgcolor="eeeeee" >
				<TD colspan=6 height=2 bgcolor=0></TD>
			</TR>
			<TR class="RepTableTitle">
				<TD bgcolor="#CCCCEE" height="30" colspan="6">
				<table width=100% cellspacing=0 style="cursor:hand;color:gray;">
				<tr>
					<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
						<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
						&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>,0);">’›ÕÂ »⁄œ &gt;</a>
					</td>
				</tr>
				<tr>
<%				for i=1 to TotalPages 
					if i = CurrentPage then 
%>						<td style="color:black;"><b>[<%=i%>]</b></td>
<%					else
%>						<td onclick="go2Page(<%=i%>,0);"><%=i%></td>
<%					end if 
					if i mod pageCols = 0 then response.write "</tr><tr>" 
				next 

%>				</tr>
				</table>
				</TD>
			</TR>
<%		end if
%>
		</TABLE>
		<br>

		<SCRIPT LANGUAGE="JavaScript">
		function go2Page(p,ord) {
			if(ord==0){
				ord=<%=Ord%>;
			}
			else if(ord==<%=Ord%>){
				ord= 0-ord;
			}
			str = '?act=show&FromDate=' + escape('<%=FromDate%>') + '&ToDate=' + escape('<%=ToDate%>') + '&Ord=' + escape(ord) + '&p=' + escape(p);  
			window.location = str;
		}
		</SCRIPT>
<%	
	End if
End If
End If 
%>
<!--#include file="tah.asp" -->
