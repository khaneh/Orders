<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="œ‘»Ê—œ ”›«—‘« "
SubmenuItem=7
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
	div.Right {float: right;width: 95px;}
	div.rightHead {float: right;padding-left: 20px;}
	div.NewRow{clear: right;}
	td.empty {background-color: #C3DBEB !important;}
</STYLE>
<SCRIPT LANGUAGE='JavaScript'>
</SCRIPT>
<%
if request("act")="" then 
%>
<form action="" method="post">
	<div class="NewRow">
		<div class="RightHead">„Â·   ÕÊÌ·</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")=" «ÌÌœ" then
				if request("isDelay")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="isDelay" >
			<label> «ŒÌ— œ«—œ</label>
		</div>
		<div class="Right">
			<input type="checkbox"<%if request("submit")=" «ÌÌœ" then 
				if request("today")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="today" >
			<label>«„—Ê“</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")=" «ÌÌœ" then 
				if request("tomorrow")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="tomorrow" >
			<label>›—œ«</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")=" «ÌÌœ" then 
				if request("nextWeek")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="nextWeek" >
			<label>2  « 7 —Ê“ œÌê—</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")=" «ÌÌœ" then 
				if request("moreNextWeek")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="moreNextWeek" >
			<label>»Ì‘ — «“ 7 —Ê“</label>
		</div>
	</div>
	<div class="NewRow">
		<div class="RightHead">‰Ê⁄ ”›«—‘</div>
	<%
	set rs=Conn.Execute("select * from OrderTraceTypes where isActive=1")
	while not rs.eof
	%>
		<div class="Right">
			<input type="checkbox" <%if request("submit")=" «ÌÌœ" then 
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
	<div>
		<input type="submit" name="submit" value=" «ÌÌœ">
	</div>
</form>
<div style="clear: both;margin:20px 0 0 0;">
<center>
<%
set rs=Conn.Execute("select count(step) as stepCountLevel from orderTraceSteps where IsActive=1 and step is not null group by step order by count(step) desc")
stepCountLevel = CInt(rs("stepCountLevel"))
rs.close
set rs=Conn.Execute("select max(step) as stepCount from orderTraceSteps where IsActive=1 and step is not null")
stepCount = CInt(rs("stepCount"))
rs.close
dim steps(10,10)
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
set rs=Conn.Execute("select id,step from orderTraceSteps where IsActive=1 and step is not null order by step,ord")
while not rs.eof
	if oldStep=CInt(rs("step")) then 
		i=i+1
	else 
		i=0
	end if
	steps(rs("step"),i)=rs("id")
	oldStep=cint(rs("step"))
	rs.moveNext
wend
%>
	<table class="CustTable">
<%
	for i=0 to stepCountLevel - 1
	%>
		<tr>
	<%
		for s= 1 to stepCount
			if cint(steps(s,i))>0 then
				fromDate=""
				toDate=shamsiToday()
				orderTypes=""
				condition="" 
				if request("submit")=" «ÌÌœ" then 
					for ii=0 to orderTypeCount
						if request(orderType(ii))="on" then orderTypes = orderTypes & Split(orderType(ii),"-")(1) & ","
					next
					if len(orderTypes)>0 then 
						orderTypes = mid(orderTypes,1,len(orderTypes)-1)
						condition=" and orders_trace.type in (" & orderTypes & ")"
					end if
					if request("isDelay")="on" or request("today")="on" or request("tomorrow")="on" or request("nextWeek")="on" or request("moreNextWeek")="on" then
					 	condition = condition & " and ( 0=1"
					 	if request("isDelay")="on" then 
					 		condition = condition & " or orders_trace.return_date between '1389/01/01' and '" & shamsiDate(dateadd("d",-1,date())) & "'"
					 		fromDate = "1389/01/01"
					 		toDate = shamsiDate(dateadd("d",-1,date()))
					 	end if
						if request("today")="on" then 
							condition = condition & " or orders_trace.return_date = '" & shamsiToday() & "'"
							if fromDate = "" then fromDate = shamsiToday()
							toDate = shamsiToday()
						end if
						if request("tomorrow")="on" then 
							condition = condition & " or orders_trace.return_date = '" & shamsiDate(dateadd("d",1,date())) & "'"
							if fromDate = "" then fromDate = shamsiDate(dateadd("d",1,date()))
							toDate = shamsiDate(dateadd("d",1,date()))
						end if
						if request("nextWeek")="on" then 
							condition = condition & " or orders_trace.return_date between '" & shamsiDate(dateadd("d",2,date())) & "' and '" & shamsiDate(dateadd("d",7,date())) & "'"
							if fromDate = "" then fromDate = shamsiDate(dateadd("d",2,date()))
							toDate = shamsiDate(dateadd("d",7,date()))
						end if
						if request("moreNextWeek")="on" then 
							condition = condition & " or orders_trace.return_date > '" & shamsiDate(dateadd("d",7,date())) & "'"
							if fromDate = "" then fromDate = shamsiDate(dateadd("d",8,date()))
							toDate = "9999/99/99"
						end if
						condition = condition & ")"
					end if
				end if
				if fromDate="" then fromDate="1389/01/01"
				mySQL = "select orderTraceSteps.name,isnull(drv.orderCount,0) as orderCount from orderTraceSteps left outer join (select orders_trace.step, count(orders_trace.radif_sefareshat) as orderCount from orders_trace inner join Orders on orders_trace.radif_sefareshat=orders.id and orders.Closed=0 where 1=1 and orders_trace.return_date >'1389/01/01' " & condition & " group by orders_trace.step) drv on orderTraceSteps.id=drv.step where orderTraceSteps.id=" & steps(s,i)
				set rs=Conn.Execute(mySQL)
	%>
			<td title="»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ">
			<center>
				<img src="../images/folder1.gif" align="top">
				<br>
				<a href="orderFolow.asp?act=show&fromDate=<%=Server.URLEncode(fromDate)%>&toDate=<%=Server.URLEncode(toDate)%>&orderTypes=<%=Server.URLEncode(orderTypes)%>&step=<%=steps(s,i)%>"><%=rs("name") & " (" & rs("orderCount") & ")"%></a>
			</center>
			</td>
	<%
				rs.close
			else
			%>
			<td class="empty"></td>
			<%
			end if
		next
		%>
		</tr>
		<%
	next
%>
	</table>
</center>
</div>
<%
elseif request("act")="show" then
	myCriteria=""
	if request("fromDate")<>"" then myCriteria = " AND orders_trace.return_date between '" & request("fromDate") & "' AND '" & request("toDate") & "'"
	if request("orderTypes")<>"" then myCriteria = myCriteria & " and orders_trace.type in (" & request("orderTypes") & ")"
	myCriteria = myCriteria & " and orders_trace.step=" & request("step")
	
	mySQL="SELECT orders_trace.*, Orders.closed, OrderTraceStatus.Name AS StatusName, OrderTraceStatus.Icon,DRV_Invoice.price,orders.customer FROM Orders INNER JOIN  orders_trace ON Orders.ID = orders_trace.radif_sefareshat INNER JOIN  OrderTraceStatus ON orders_trace.status = OrderTraceStatus.ID left outer join (select InvoiceOrderRelations.[Order],SUM(InvoiceLines.Price + InvoiceLines.Vat - InvoiceLines.Discount -InvoiceLines.Reverse) as price from InvoiceOrderRelations inner join Invoices on InvoiceOrderRelations.Invoice=Invoices.ID inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice where Invoices.Voided=0 group by InvoiceOrderRelations.[Order]) DRV_Invoice on Orders.ID=DRV_Invoice.[Order] WHERE (orders.Closed=0 "& myCriteria & ") ORDER BY order_date DESC, radif_sefareshat DESC"	
	'response.write mysql
	set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		tmpCounter=0
'response.write mySQL
%>
	<div align="center" dir="LTR">
	<TABLE border="1" cellspacing="0" cellpadding="1" dir="RTL" borderColor="#555588">
		<TR bgcolor="#CCCCFF">
			<TD width="44"># ”›«—‘</TD>
			<TD width="46"> «—ÌŒ ”›«—‘</TD>
			<TD width="64"> «—ÌŒ  ÕÊÌ·</TD>
			<TD width="122">‰«„ ‘—ﬂ </TD>
			<TD width="122">‰«„ „‘ —Ì</TD>
			<TD width="84">⁄‰Ê«‰ ﬂ«—</TD>
			<TD width="40">‰Ê⁄</TD>
			<TD width="53">„—Õ·Â</TD>
			<TD width="36">”›«—‘ êÌ—‰œÂ</TD>
			<TD width="18">Ê÷⁄</TD>
			<td width="50">„»·€ ﬂ·</td>
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
			<TD width="40" DIR="LTR"><A HREF="TraceOrder.asp?act=show&order=<%=RS1("radif_sefareshat")%>" target="_blank"><%=RS1("radif_sefareshat")%></A></TD>
			<TD DIR="LTR"><%=RS1("order_date")%></TD>
			<TD DIR="LTR"><%=RS1("return_date") & " ("& RS1("return_time") & ")"%></TD>
			<TD><%=RS1("company_name") & "<br> ·›‰:("& RS1("telephone")& ")"%>&nbsp;</TD>
			<TD><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=RS1("customer")%>'><%=RS1("customer_name")%></a>&nbsp;</TD>
			<TD><%=RS1("order_title")%>&nbsp;</TD>
			<TD><%=RS1("order_kind")%></TD>
			<TD style="<%=tmpStyle%>"><%=RS1("marhale")%></TD>
			<TD><%=RS1("salesperson")%>&nbsp;</TD>
			<TD title="»—«Ì  €ÌÌ— „—Õ·Â ﬂ·Ìﬂ ﬂ‰Ìœ">
				<A HREF="../shopfloor/default.asp?orderNum=<%=rs1("radif_sefareshat")%>&marhale_box=<%=RS1("step")%>">
					<IMG SRC="<%=RS1("Icon")%>" WIDTH="20" HEIGHT="20" BORDER="0">
				</a>
			</TD>
			<td><%if isnull(RS1("price")) then response.write "----" else response.write Separate(RS1("price")) end if %></td>
		</TR>
		<TR bgcolor="#FFFFFF">
			<TD colspan="10" style="height:10px"></TD>
		</TR>
<%					RS1.moveNext
		Loop

%>					<TR bgcolor="#ccccFF">
				<TD colspan="10"> ⁄œ«œ ‰ «ÌÃ Ã” ÃÊ: <%=tmpCounter%></TD>
			</TR>
	</TABLE>
	</div>
	<BR>
<%			else
%>			<TABLE border="1" cellspacing="0" cellpadding="0" dir="RTL" align="center" width="600">
		<TR bgcolor="#FFFFDD">
			<TD align="center" style="height:40px;font-size:12pt;font-weight:bold;color:red">ÂÌç ÃÊ«»Ì ‰œ«—Ì„ »Â ‘„« »œÂÌ„.</TD>
		</TR>
	</TABLE>
	<hr>
<%			End If

end if

%>
<!--#include file="tah.asp" -->