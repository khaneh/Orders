<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ê“«—‘ »—ê‘  «“ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , "A") then NotAllowdToViewThisPage()
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
	fromDate =	request("fromDate")
	toDate = 	request("toDate")
	Ord =		request("Ord")
	item =		request("item")
	cat =		request("cat")
	
	If Ord="" Then Ord = 0
	select case Ord
	case "1":
		order="invoices.id"
	case "-1":
		order="invoices.id DESC"
	case "2":
		order="InvoiceOrderRelations.[Order]"
	case "-2":
		order="InvoiceOrderRelations.[Order] DESC"
	case "3":
		order="orders_trace.order_title"
	case "-3":
		order="orders_trace.order_title DESC"
	case "4":
		order="orders_trace.customer_name,orders_trace.company_name"
	case "-4":
		order="orders_trace.customer_name DESC,orders_trace.company_name DESC"
	case "5":
		order="invoices.totalReceivable"
	case "-5":
		order="invoices.totalReceivable DESC"
	case "6":
		order="invoices.totalReverse"
	case "-6":
		order="invoices.totalReverse DESC"
	case "7":
		order="users.RealName"
	case "-7":
		order="users.RealName DESC"
	
	case else:
		order="invoices.id"
		Ord=1
	end select
%>
	<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
	<TR bgcolor="eeeeee" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
		<TD width=50 onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>">‘„«—Â ›«ﬂ Ê— <%if abs(ord)=1 then response.write arrow%></TD>
		<td onclick='go2Page(1,2);' style="<%if abs(ord)=2 then response.write style%>">‘„«—Â ”›«—‘<%if abs(ord)=2 then response.write arrow%></td>
		<td onclick='go2Page(1,3);' style="<%if abs(ord)=3 then response.write style%>">⁄‰Ê«‰ ”›«—‘<%if abs(ord)=3 then response.write arrow%></td>
		<td onclick='go2Page(1,4);' style="<%if abs(ord)=4 then response.write style%>">‰«„ „‘ —Ì<%if abs(ord)=4 then response.write arrow%></td>
		<td onclick='go2Page(1,5);' style="<%if abs(ord)=5 then response.write style%>">Ã„⁄ ›«ﬂ Ê—<%if abs(ord)=5 then response.write arrow%></td>
		<td onclick='go2Page(1,6);' style="<%if abs(ord)=6 then response.write style%>">Ã„⁄ »—ê‘ <%if abs(ord)=6 then response.write arrow%></td>
		<td onclick='go2Page(1,7);' style="<%if abs(ord)=7 then response.write style%>">’«œ— ﬂ‰‰œÂ<%if abs(ord)=7 then response.write arrow%></td>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=7 height=2 bgcolor=0></TD>
	</TR>
	
<%
	myCond=""
	if item <>"" then
		myCond = " and invoices.id in (select invoice from invoiceLines where reverse>0 and item=" & item & ") "
	elseif cat<>"" then 
		myCond = " and invoices.id in (select invoice from invoiceLines where reverse>0 and item in (select invoiceItem from InvoiceItemCategoryRelations where InvoiceItemCategory = " & cat & ")) "
	end if 
	mySQL="select InvoiceOrderRelations.Invoice, InvoiceOrderRelations.[Order], orders_trace.order_title, orders_trace.company_name, orders_trace.customer_name, isnull(users.RealName,'-') as issuedBy, invoices.totalReceivable, invoices.totalReverse, orders_trace.salesperson, invoices.customer from Invoices left outer join InvoiceOrderRelations on invoices.id = InvoiceOrderRelations.Invoice left outer join Orders_trace on  InvoiceOrderRelations.[Order] = orders_trace.radif_sefareshat left outer join Users on invoices.IssuedBy = users.ID where invoices.totalReverse>0 and issuedDate between '" & fromDate & "' and '" & toDate & "' " & myCond & " order by " & order
	'response.write mySQL
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
				<TD dir=ltr align=right><a href="../AR/AccountReport.asp?act=showInvoice&invoice=<%=rs("invoice")%>"><%=rs("invoice")%></TD>
				<TD><a href="../order/TraceOrder.asp?act=show&order=<%=rs("order")%>"><%=rs("order")%></a></TD>
				<TD><%=rs("order_title")%></TD>
				<TD><A href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rs("customer")%>"><%=rs("customer_name") & "(" & rs("company_name") & ")"%></A></TD>
				<TD style='direction:LTR;text-align:right;'><%=Separate(rs("totalReceivable"))%></TD>
				<TD style='direction:LTR;text-align:right;'><%=Separate(rs("totalReverse"))%></TD>
				<TD title='›—Ê‘‰œÂ: <%=rs("salesperson")%>'><%=rs("issuedBy")%></TD>
			</TR>
			  
	<% 
		rs.moveNext
		Loop

		if TotalPages > 1 then
			pageCols=20
%>			
			<TR bgcolor="eeeeee" >
				<TD colspan=7 height=2 bgcolor=0></TD>
			</TR>
			<TR class="RepTableTitle">
				<TD bgcolor="#CCCCEE" height="30" colspan="7">
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

%>
<!--#include file="tah.asp" -->
