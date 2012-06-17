<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ÏÔÈæÑÏ ÇÓÊÚáÇã"
SubmenuItem=8
if not Auth(2 , 3) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<%
'Server.ScriptTimeout = 3600
%>
<STYLE>
	.CustTable {font-family:tahoma; width:100%; border:5 solid #C3DBEB; direction: RTL;background-color: #C3DBEB;color: black;}
	.CustTable td {padding:5;width: 50px;}
	.CustTable a {text-decoration:none;color:#000088;}
	.CustTable a:hover {text-decoration:underline;}
	.CustTable td:nth-child(odd) { background-color:#eee; }
	.CustTable td:nth-child(even) { background-color:#C3DBEB; }
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	div.Right {float: right;width: 110px;}
	div.rightHead {float: right;padding-left: 20px;}
	div.NewRow{clear: right;margin: 20px 10px 0 0;}
	a.link{margin: 0 15px 0 0;}
	td.empty {background-color: #C3DBEB !important;}
</STYLE>
<SCRIPT LANGUAGE='JavaScript'>
</SCRIPT>
<%
if request("act")="" then 
%>
<form action="" method="post">
	<div class="NewRow">
		<div class="RightHead">ãåáÊ ÊÍæíá</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="ÊÇííÏ" then
				if request("isDelay")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="isDelay" >
			<label>ÊÇÎíÑ ÏÇÑÏ</label>
		</div>
		<div class="Right">
			<input type="checkbox"<%if request("submit")="ÊÇííÏ" then 
				if request("today")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="today" >
			<label>ÇãÑæÒ</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="ÊÇííÏ" then 
				if request("tomorrow")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="tomorrow" >
			<label>İÑÏÇ</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="ÊÇííÏ" then 
				if request("nextWeek")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="nextWeek" >
			<label>2 ÊÇ 7 ÑæÒ ÏíÑ</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="ÊÇííÏ" then 
				if request("moreNextWeek")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="moreNextWeek" >
			<label>ÈíÔÊÑ ÇÒ 7 ÑæÒ</label>
		</div>
	</div>
	<div class="NewRow">
		<div class="RightHead">äæÚ ÓİÇÑÔ</div>
	<%
	set rs=Conn.Execute("select * from OrderTraceTypes where isActive=1")
	while not rs.eof
	%>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="ÊÇííÏ" then 
				if request("orderType-" & rs("id"))="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="orderType-<%=rs("id")%>"
			<label><%=rs("name")%></label>
		</div>
	<%
		rs.moveNext
	wend
	rs.close
	%>
	</div>
	<div class="NewRow">
		<input type="submit" name="submit" value="ÊÇííÏ">
		<a class="link" href="Inquiry.asp?act=advancedSearch&Submit=ÊÇííÏ&check_marhale=on&marhale_not_check=on&marhale_box=4&check_closed=on">ÑÏ æ ÓİÇÑÔ äÔÏååÇ</a>
		
	</div>
</form>
<div style="clear: both;margin:20px 0 0 0;">
<center>
<%
'set rs=Conn.Execute("select count(step) as stepCountLevel from QuoteSteps where IsActive=1 and step is not null group by step order by count(step) desc")
stepCountLevel = 1 'CInt(rs("stepCountLevel"))
'rs.close
set rs=Conn.Execute("select max(id) as stepCount from QuoteSteps where IsActive=1")
stepCount = CInt(rs("stepCount"))
rs.close
dim steps(1,10)
dim orderType(10)
set rs=Conn.Execute("select * from OrderTraceTypes where isActive=1")
i=0
while not rs.eof
	orderType(i)="orderType-"&rs("id")
	rs.moveNext
	i=i+1
wend
orderTypeCount=i-1
oldStep=-1

rs.close
%>
	<table class="CustTable">
		<tr>
	<%
set rs=Conn.Execute("select * from QuoteSteps where IsActive=1")
while not rs.eof

	fromDate=""
	toDate="9999/99/99" 'shamsiToday()
	orderTypes=""
	condition="" 
	if request("submit")="ÊÇííÏ" then 
		for ii=0 to orderTypeCount
			if request(orderType(ii))="on" then orderTypes = orderTypes & Split(orderType(ii),"-")(1) & ","
		next
		if len(orderTypes)>0 then 
			orderTypes = mid(orderTypes,1,len(orderTypes)-1)
			condition=" and Quotes.type in (" & orderTypes & ")"
		end if
		if request("isDelay")="on" or request("today")="on" or request("tomorrow")="on" or request("nextWeek")="on" or request("moreNextWeek")="on" then
		 	condition = condition & " and ( 0=1"
		 	if request("isDelay")="on" then 
		 		condition = condition & " or Quotes.return_date between '1389/01/01' and '" & shamsiDate(dateadd("d",-1,date())) & "'"
		 		fromDate = "1389/01/01"
		 		toDate = shamsiDate(dateadd("d",-1,date()))
		 	end if
			if request("today")="on" then 
				condition = condition & " or Quotes.return_date = '" & shamsiToday() & "'"
				if fromDate = "" then fromDate = shamsiToday()
				toDate = shamsiToday()
			end if
			if request("tomorrow")="on" then 
				condition = condition & " or Quotes.return_date = '" & shamsiDate(dateadd("d",1,date())) & "'"
				if fromDate = "" then fromDate = shamsiDate(dateadd("d",1,date()))
				toDate = shamsiDate(dateadd("d",1,date()))
			end if
			if request("nextWeek")="on" then 
				condition = condition & " or Quotes.return_date between '" & shamsiDate(dateadd("d",2,date())) & "' and '" & shamsiDate(dateadd("d",7,date())) & "'"
				if fromDate = "" then fromDate = shamsiDate(dateadd("d",2,date()))
				toDate = shamsiDate(dateadd("d",7,date()))
			end if
			if request("moreNextWeek")="on" then 
				condition = condition & " or Quotes.return_date > '" & shamsiDate(dateadd("d",7,date())) & "'"
				if fromDate = "" then fromDate = shamsiDate(dateadd("d",8,date()))
				toDate = "9999/99/99"
			end if
			condition = condition & ")"
		end if
	end if
	if fromDate="" then fromDate="1389/01/01"
	mySQL = "select isnull(count(id),0) as quotesCount from Quotes where step=" & rs("id") & condition
	set rs1=Conn.Execute(mySQL)
	%>
			<td title="ÈÑÇí ãÔÇåÏå ÌÒÆíÇÊ ßáíß ßäíÏ">
			<center>
				<img src="../images/folder1.gif" align="top">
				<br>
				<a href="inquiryFolow.asp?act=show&fromDate=<%=Server.URLEncode(fromDate)%>&toDate=<%=Server.URLEncode(toDate)%>&orderTypes=<%=Server.URLEncode(orderTypes)%>&step=<%=rs("id")%>"><%=rs("name") & " (" & rs1("quotesCount") & ")"%></a>
			</center>
			</td>
	<%
	rs1.close
	rs.moveNext
wend
		%>
		</tr>
	</table>
</center>
</div>
<%
rs.close
elseif request("act")="show" then
	myCriteria=""
	if request("fromDate")<>"" then myCriteria = " AND Quotes.return_date between '" & request("fromDate") & "' AND '" & request("toDate") & "'"
	if request("orderTypes")<>"" then myCriteria = myCriteria & " and Quotes.type in (" & request("orderTypes") & ")"
	myCriteria = myCriteria & " and Quotes.step=" & request("step")
	
	mySQL="SELECT Quotes.*, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon,DRV_Invoice.price,Quotes.customer FROM Quotes INNER JOIN  OrderTraceStatus ON Quotes.status = OrderTraceStatus.ID left outer join (select InvoiceQuoteRelations.quote,SUM(InvoiceLines.Price + InvoiceLines.Vat - InvoiceLines.Discount -InvoiceLines.Reverse) as price from InvoiceQuoteRelations inner join Invoices on InvoiceQuoteRelations.Invoice=Invoices.ID inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice where Invoices.Voided=0 group by InvoiceQuoteRelations.quote) DRV_Invoice on Quotes.ID=DRV_Invoice.quote WHERE (1=1 "& myCriteria & ") ORDER BY order_date DESC, Quotes.id DESC"	
	'response.write mysql
	set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		tmpCounter=0
'response.write mySQL
%>
	<div align="center" dir="LTR">
	<TABLE border="1" cellspacing="0" cellpadding="1" dir="RTL" borderColor="#555588">
		<TR bgcolor="#CCCCFF">
			<TD width="44"># ÇÓÊÚáÇã</TD>
			<TD width="46">ÊÇÑíÎ ÇÓÊÚáÇã</TD>
			<TD width="64">ÊÇÑíÎ ÊÍæíá</TD>
			<TD width="122">äÇã ÔÑßÊ</TD>
			<TD width="122">äÇã ãÔÊÑí</TD>
			<TD width="84">ÚäæÇä ßÇÑ</TD>
			<TD width="40">äæÚ</TD>
			<TD width="36">ÇÓÊÚáÇã íÑäÏå</TD>
			<TD width="18">æÖÚ</TD>
			<td width="50">ãÈáÛ ßá</td>
		</TR>
<%				Do while not RS1.eof 
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
		Else
			tmpColor="#DDDDDD"
		End If

		if RS1("Closed") then
			tmpStyle="background-color:#FFCCCC;"
		else
			tmpStyle=""
		End If
		
%>
		<TR bgcolor="<%=tmpColor%>" title="<%=RS1("StatusName")%>">
			<TD width="40" DIR="LTR"><A HREF="Inquiry.asp?act=show&quote=<%=RS1("id")%>" target="_blank"><%=RS1("id")%></A></TD>
			<TD DIR="LTR"><%=RS1("order_date")%></TD>
			<TD DIR="LTR"><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
			<TD><%=RS1("company_name") & "<br>Êáİä:("& RS1("telephone")& ")"%>&nbsp;</TD>
			<TD><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=RS1("customer")%>'><%=RS1("customer_name")%></a>&nbsp;</TD>
			<TD><%=RS1("order_title")%>&nbsp;</TD>
			<TD><%=RS1("order_kind")%></TD>
			<TD style="<%=tmpStyle%>"><%=RS1("marhale")%></TD>
			<TD><%=RS1("salesperson")%>&nbsp;</TD>
			<td><%if isnull(RS1("price")) then response.write "----" else response.write Separate(RS1("price")) end if %></td>
		</TR>
		<TR bgcolor="#FFFFFF">
			<TD colspan="10" style="height:10px"></TD>
		</TR>
<%					RS1.moveNext
		Loop

%>					<TR bgcolor="#ccccFF">
				<TD colspan="10">ÊÚÏÇÏ äÊÇíÌ ÌÓÊÌæ: <%=tmpCounter%></TD>
			</TR>
	</TABLE>
	</div>
	<BR>
<%			else
%>			<TABLE border="1" cellspacing="0" cellpadding="0" dir="RTL" align="center" width="600">
		<TR bgcolor="#FFFFDD">
			<TD align="center" style="height:40px;font-size:12pt;font-weight:bold;color:red">åí ÌæÇÈí äÏÇÑíã Èå ÔãÇ ÈÏåíã.</TD>
		</TR>
	</TABLE>
	<hr>
<%			End If

end if

%>
<!--#include file="tah.asp" -->