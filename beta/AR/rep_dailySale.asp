<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , 1) then NotAllowdToViewThisPage()
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
	.Inp1 {font-size: 9pt;background-color: #CCE;border: none;direction: LTR;}
</STYLE>
<% 
if Auth("C" , 2) then 
	input_date_start=	sqlSafe(request("input_date_start"))
	input_date_end=		sqlSafe(request("input_date_end"))
else
	input_date_start=	sqlSafe(request("input_date"))
	input_date_end=		sqlSafe(request("input_date"))
end if

	if request("fullyApplied")="on" then 
		fullyApplied=1
	else 
		fullyApplied=0
	end if
'ON ERROR RESUME Next
Ord =	request("Ord")	
page=	request("page")	
select case Ord
	case "1":
		order="EffectiveDate"
	case "-1":
		order="EffectiveDate DESC"
	case "2":
		order="AccountTitle"
	case "-2":
		order="AccountTitle DESC"
	case "3":
		order="AmountOriginal"
	case "-3":
		order="AmountOriginal DESC"
	case "4":
		order="RemainedAmount"
	case "-4":
		order="RemainedAmount DESC"
	Case "5":
		order = "arBalance"
	Case "-5":
		order = "arBalance DESC"
	case "6":
		order="FullyApplied"
	case "-6":
		order="FullyApplied DESC"
	case else:
		order="EffectiveDate"
		Ord=1
end select	
mySQL="select SUM(AmountOriginal) as AmountOriginal,SUM(RemainedAmount) as RemainedAmount, MAX(arBalance) as arBalance, count(ARItems.ID) AS totalItems from ARItems inner join Accounts on ARItems.Account=Accounts.ID where (ARItems.EffectiveDate between '" & input_date_start & "' and '" & input_date_end & "') and ARItems.Type=1"
if fullyApplied=1 then mySQL=mySQL&" AND FullyApplied=0"

%>
<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">	
<%
set rs=Conn.Execute (mySQL)
if rs.eof then
%>	<tr>
		<td bgcolor="#BBBBBB" height="30" colspan="7" align=center><b>ÂÌç .</b></td>
	</tr>
<%	Else %>
	<TR bgcolor="#CCCCEE" >
		<TD colspan=2 rowspan=2 title=" <%=rs("totalItems")%> ¬Ì „ ">
			<form method=post>
				&nbsp;«“  «—ÌŒ <input class=Inp1 type=text name='input_date_start' size=10 value='<%=input_date_start%>' style="width:55px;">
				&nbsp; «  «—ÌŒ <input class=Inp1 type=text name='input_date_end' size=10 value='<%=input_date_end%>' style="width:55px;">
				&nbsp;›ﬁÿ  ”ÊÌÂ ‰‘œÂùÂ«<input type=CHECKBOX name='fullyApplied' <% if fullyApplied=1 then response.write("checked") %>>
				<input type=submit value='»ê—œ'>
				<input type=hidden name="Ord" value='<%=Ord%>'>
				<input type=hidden name="Page" value='<%=page%>'>
			</form>
		</TD>

		<TD width=70 >Ã„⁄ „»·€</TD>
		<TD width=70 >Ã„⁄ „«‰œÂ</TD>
		<TD width=70 >Ã„⁄ „«‰œÂ Õ”«»ùÂ«</TD>
		<TD width=70 ></TD>
	</TR>
	<TR bgcolor="#CCCCEE" >
		<TD width=70 dir=ltr align=right><b><%=Separate(rs("AmountOriginal"))%></b></TD>
		<TD width=70 dir=ltr align=right><b><%=Separate(rs("RemainedAmount"))%></b></TD>
		<TD width=70 dir=ltr align=right><b><%=Separate(rs("arBalance"))%></b></TD>
		<TD></TD>
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
		<TD width=50 onclick='go2Page(1,-1);' style="<%if abs(ord)=1 then response.write style%>"> «—ÌŒ<%if abs(ord)=1 then response.write arrow%></TD>
		<TD width='*'  onclick='go2Page(1,-2);' style="<%if abs(ord)=2 then response.write style%>">‰«„ Õ”«»<%if abs(ord)=2 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-3);' style="<%if abs(ord)=3 then response.write style%>">„»·€<%if abs(ord)=3 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-4);' style="<%if abs(ord)=4 then response.write style%>">„«‰œÂ<%if abs(ord)=4 then response.write arrow%></TD>
		
		<TD width=70 onclick='go2Page(1,-5);' style="<%if abs(ord)=5 then response.write style%>">„«‰œÂ Õ”«»<%if abs(ord)=5 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-6);' style="<%if abs(ord)=6 then response.write style%>">Ê÷⁄Ì <%if abs(ord)=6 then response.write arrow%></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
	</TR>
<%		
	SumAmount=0
	SumRemain=0
	tmpCounter=0

	mySQL="select Accounts.arBalance,Accounts.AccountTitle, ARItems.* from ARItems inner join Accounts on ARItems.Account=Accounts.ID where (ARItems.EffectiveDate between '" & input_date_start & "' and '" & input_date_end & "')  and ARItems.Type=1" 
	if fullyApplied=1 then mySQL=mySQL&" AND ARItems.FullyApplied=0"
	mySQL= mySQL & " ORDER BY " & order
	Set rs=Server.CreateObject("ADODB.Recordset")'Conn.Execute(mySQL)

	PageSize = 50
	rs.PageSize = PageSize
'response.write mySql
	rs.CursorLocation=3 'in ADOVBS_INC adUseClient=3
	rs.Open mySQL ,Conn,3
	TotalPages = rs.PageCount

	CurrentPage=1

	if isnumeric(Page) then
		pp=clng(Page)
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
				<TD><%=Separate(rs("EffectiveDate"))%></TD>
				<TD title="„‘«ÂœÂ Õ”«»"><A href='../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rs("Account")%>'><%=rs("AccountTitle")%></A></TD>
				<TD dir=ltr align=right title="„‘«ÂœÂ ›«ﬂ Ê— „—»ÊÿÂ"><a href='AccountReport.asp?act=showInvoice&invoice=<%=rs("Link")%>'><%=Separate(rs("AmountOriginal"))%></a></TD>
				<TD dir=ltr align=right><%=Separate(rs("RemainedAmount"))%></TD>
				<TD dir=ltr align=right title="ê“«—‘ Õ”«»" <% if cdbl(rs("arBalance"))<0 then response.write("style='background-color:#F55;'") %>><a href='AccountReport.asp?sys=AR&act=show&selectedCustomer=<%=rs("Account")%>'><%=Separate(rs("arBalance"))%></a></TD>
				<TD ><% 
					if rs("FullyApplied")="True" then 
						response.write(" ”ÊÌÂ ‘œÂ")
					else
						response.write("<b style='color:red;'> ”ÊÌÂ ‰‘œÂ</b>")
					end if
					%>
				</TD>
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
	end if
end if
%>
		</TABLE>
		<SCRIPT LANGUAGE="JavaScript">
		function go2Page(p,ord) {
			if(ord==0){
				ord=<%=Ord%>;
			}
			else if(ord==<%=Ord%>){
				ord= 0-ord;
			}
			document.getElementsByName('Page')[0].value=p;
			document.getElementsByName('Ord')[0].value=ord;
			document.forms[0].submit();
			//str = '?act=show&FromDate=' + escape('<%=FromDate%>') + '&ToDate=' + escape('<%=ToDate%>') + '&Ord=' + escape(ord) + '&p=' + escape(p);  
			//window.location = str;
		}
		</SCRIPT>

<!--#include file="tah.asp" -->
