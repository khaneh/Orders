<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' RFM
PageTitle="œ‘»Ê—œ „‘ —Ì«‰"
SubmenuItem=7
if not Auth(1 , 8) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.myTable {border-color:black;border-width: medium;border-style: solid;}
	.myTable tr:nth-child(1) {border-color: gray;border-width: thin;border-style: solid; background-color: lime ;text-align: center;border-bottom-width: thick;border-bottom-color: black;border-bottom-style: solid;}
	.myTable tr:nth-child(2n+2) { background-color:#eee; }
	.myTable tr:nth-child(2n+3) { background-color:#C3DBEB; }
</STYLE>
<BR>
<%
dim color(6)
dim text_r(4)
dim text_f(5)
dim text_v(6)
color(0)="red"
color(1)="Blue"
color(2)="white"
color(3)="gray"
color(4)="yellow"
color(5)="green"
text_r(0)="»Ì‘ «“ Ìﬂ ”«·"
text_r(1)="Ìﬂ ”«· ÅÌ‘"
text_r(2)="‘‘ „«Â ÅÌ‘"
text_r(3)="”Â „«Â ÅÌ‘"
text_f(0)="Â— »Ì‘ «“ Ìﬂ ”«·"
text_f(1)="Â— Ìﬂ ”«·"
text_f(2)="Â— ‘‘ „«Â"
text_f(3)="Â— ”Â „«Â"
text_f(4)="Â— Ìﬂ „«Â"
text_v(0)=" « ’œÂ“«— —Ì«·"
text_v(1)=" « Å«‰’œ Â“«— —Ì«·"
text_v(2)=" « Ìﬂ „·ÌÊ‰ —Ì«·"
text_v(3)=" « Å‰Ã „·ÌÊ‰ —Ì«·"
text_v(4)=" « œÂ „·ÌÊ‰ —Ì«·"
text_v(5)="œÂ „·ÌÊ‰ —Ì«· »Â »«·«"
Conn.Execute("update Invoices set issuedDate_en=dbo.udf_date_solarToDate(cast(substring(issuedDate,1,4) as int),cast(substring(issuedDate,6,2) as int),cast(substring(issuedDate,9,2) as int)) where Issued=1 and issuedDate_en is null")
if request("act")="" then 
	if request("ord")="" then 
		ord= " Recency,Frequency,Value"
	elseif request("ord")="1" then 
		ord= " Recency,Frequency,Value"
	elseif request("ord")="-1" then 
		ord= " Recency desc,Frequency desc,Value desc"
	elseif request("ord")="2" then 
		ord= " Frequency,Recency,Value"
	elseif request("ord")="-2" then 
		ord= " Frequency desc,Recency desc,Value desc"
	elseif request("ord")="3" then 
		ord= " Value,Recency,Frequency"
	elseif request("ord")="-3" then 
		ord= " Value desc,Recency desc,Frequency desc"	
	elseif request("ord")="4" then 
		ord= " count(rfv_value.customer),Recency,Frequency,Value"
	elseif request("ord")="-4" then 
		ord= " count(rfv_value.customer) desc,Recency,Frequency,Value"
	end if
		
%>
<select name="accountGroup" onchange="window.location='rfmModel.asp?accountGroup='+this.value;">
	<option <%if request("accountGroup")="" or request("accountGroup")="-1" then response.write "selected='selected'"%> value="-1">Â„ÂùÌ ’‰›ùÂ«</option>
<%
	set rs=Conn.Execute("select * from AccountGroups")
	while not rs.eof
%>
	<option <%if cint(request("accountGroup"))=cint(rs("id")) then response.write "selected='selected'"%> value="<%=rs("id")%>">
		<%=rs("name")%>
	</option>
<%
	rs.moveNext
	wend
	rs.close
%>
</select>
<table class="myTable" cellpadding="8px" cellspacing="0">
	<tr>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=request("accountGroup")%>&ord=<%if request("ord")="1" then response.write "-1" else response.write "1"%>">¬Œ—Ì‰ ”›«—‘</a>
		</td>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=request("accountGroup")%>&ord=<%if request("ord")="2" then response.write "-2" else response.write "2"%>"> ‰«Ê» ”›«—‘</a>
		</td>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=request("accountGroup")%>&ord=<%if request("ord")="3" then response.write "-3" else response.write "3"%>">„Ì«‰êÌ‰ —Ì«·</a>
		</td>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=request("accountGroup")%>&ord=<%if request("ord")="4" then response.write "-4" else response.write "4"%>"> ⁄œ«œ</a>
		</td>
	</tr>
<%
	if request("accountGroup")="" or request("accountGroup")="-1" then 
		myCond=""
		mySQLCool="select count(Accounts.id) as [count] from Accounts where Accounts.type=2 and Accounts.id not in (select distinct customer from Orders union select distinct customer from quotes )"
		mySQLWarm="select count(customer) as [count] from (select distinct customer from Quotes where Customer not in (select distinct customer from Orders)) drv"
		mySQLThreshold="select count(customer) as [count] from Orders where Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1))"
	else 
		myCond=" inner join AccountGroupRelations on AccountGroupRelations.account = rfv_value.customer  where AccountGroupRelations.accountGroup=" & request("accountGroup")
		mySQLCool="select count(Accounts.id) as [count] from Accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & request("accountGroup") & " and Accounts.type in (2,4) and Accounts.id not in (select distinct customer from Orders union select distinct customer from quotes )"
		mySQLWarm="select count(accounts.id) as [count] from accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & request("accountGroup") & " and id in (select customer from (select distinct customer from Quotes where Customer not in (select distinct customer from Orders)) drv)"
		mySQLThreshold="select count(customer) as [count] from Orders inner join AccountGroupRelations on AccountGroupRelations.account=orders.customer where AccountGroupRelations.accountGroup=" & request("accountGroup") & "  and Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1))"
	end if
	mySQL="select count(rfv_value.customer) as c,rfv_value.Value,rfv_recency.Recency,rfv_frequency.Frequency from rfv_frequency inner join rfv_recency on rfv_frequency.customer=rfv_recency.customer inner join rfv_value on rfv_recency.customer=rfv_value.customer " & myCond & " group by rfv_value.Value,rfv_recency.Recency,rfv_frequency.Frequency order by" & ord
	'response.write mySQL
	'response.end
	set rs=Conn.Execute(mySQL)
	while not rs.eof
%>
	<tr>
		<td style="background-color:<%=color(rs("Recency"))%>"><%=text_r(rs("Recency"))%></td>
		<td style="background-color:<%=color(rs("Frequency"))%>"><%=text_f(rs("Frequency"))%></td>
		<td style="background-color:<%=color(rs("Value"))%>"><%=text_v(rs("Value"))%></td>
		<td><a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=rs("Value")%>&recency=<%=rs("Recency")%>&frequency=<%=rs("Frequency")%>"><%=Separate(rs("c"))%></a></td>
	</tr>
<%
		rs.moveNext
	wend
	rs.close
	set rs=Conn.execute(mySQLCool)
%>
	<tr>
		<td title="„‘ —ÌùÂ«ÌÌ ﬂÂ ‰Â ”›«—‘ œ«—‰œ Ê ‰Â «” ⁄·«„" colspan="3">—«»ÿÂ ”—œ</td>
		<td><a href="rfmModel.asp?act=showCool&accountGroup=<%=request("accountGroup")%>"><%=rs("count")%></a></td>
	</tr>
<%
	rs.close
	set rs=Conn.execute(mySQLWarm)
%>
	<tr>
		<td title="„‘ —ÌùÂ«ÌÌ ﬂÂ «” ⁄·«„ œ«—‰œ «„« ”›«—‘ ‰œ«—‰œ" colspan="3">—«»ÿÂ ê—„</td>
		<td><a href="rfmModel.asp?act=showWarm&accountGroup=<%=request("accountGroup")%>"><%=rs("count")%></a></td>
	</tr>
<%
	rs.close
	set rs=Conn.execute(mySQLThreshold)
%>
	<tr>
		<td title="„‘ —ÌùÂ«ÌÌ ﬂÂ ”›«—‘ œ«—‰œ Ê ›«ﬂ Ê—  «ÌÌœ ‘œÂù«Ì ‰œ«—‰œ" colspan="3">„‘ —Ì ¬” «‰Â</td>
		<td><a href="rfmModel.asp?act=showThreshold&accountGroup=<%=request("accountGroup")%>"><%=rs("count")%></a></td>
	</tr>
</table>
<%
elseif request("act")="show" then 
%>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td width="30px">
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="1" then response.write "-1" else response.write "1"%>">¬Œ—Ì‰ ”›«—‘ (—Ê“)</a>
		</td>
		<td width="30px">
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="2" then response.write "-2" else response.write "2"%>">«Ê·Ì‰ ”›«—‘  « «„—Ê“ („«Â)</a>
		</td>
		<td width="30px">
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="3" then response.write "-3" else response.write "3"%>"> ⁄œ«œ ”›«—‘</a>
		</td>
		<td width="30px">
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="4" then response.write "-4" else response.write "4"%>">„Ì«‰êÌ‰ ”›«—‘ œ— „«Â</a>
		</td>
		<td width="50px">
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="5" then response.write "-5" else response.write "5"%>">Ã⁄⁄ ﬂ· ”›«—‘</a>
		</td>
		<td width="50px">
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="8" then response.write "-8" else response.write "8"%>">„Ì«‰êÌ‰ —Ì«·Ì ”›«—‘ œ—„«Â</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="6" then response.write "-6" else response.write "6"%>">„‘ —Ì</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=show&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&ord=<%if request("ord")="7" then response.write "-7" else response.write "7"%>">Ã„⁄ „«‰œÂ</a>
		</td>
	</tr>
<%
	if request("ord")="" then 
		ord= " [day]"
	elseif request("ord")="1" then 
		ord= " [day]"
	elseif request("ord")="-1" then 
		ord= " [day] desc"
	elseif request("ord")="2" then 
		ord= " [month]"
	elseif request("ord")="-2" then 
		ord= " [Month] desc"
	elseif request("ord")="3" then 
		ord= " [count]"
	elseif request("ord")="-3" then 
		ord= " [count] desc"	
	elseif request("ord")="4" then 
		ord= " [count]/cast([Month] as float)"
	elseif request("ord")="-4" then 
		ord= " [count]/cast([Month] as float) desc"
	elseif request("ord")="5" then 
		ord= " amount"
	elseif request("ord")="-5" then 
		ord= " amount desc"
	elseif request("ord")="6" then 
		ord= " accountTitle"
	elseif request("ord")="-6" then 
		ord= " accountTitle desc"
	elseif request("ord")="7" then 
		ord= " remain"
	elseif request("ord")="-7" then 
		ord= " remain desc"
	elseif request("ord")="8" then 
		ord= " [amount]/cast([Month] as float)"
	elseif request("ord")="-8" then 
		ord= " [amount]/cast([Month] as float) desc"
	end if
	
	if request("accountGroup")="" or request("accountGroup")="-1" then 
		myCond=""
	else 
		myCond=" and accounts.id in (select account from AccountGroupRelations where AccountGroupRelations.accountGroup=" & request("accountGroup") & ") "
	end if
	mySQL="select rfv_frequency.customer,rfv_frequency.[count],rfv_frequency.[Month],rfv_value.Amount,rfv_recency.[day],accounts.accountTitle, accounts.arBalance+accounts.aoBalance+accounts.apBalance as Remain from rfv_frequency inner join rfv_recency on rfv_frequency.customer=rfv_recency.customer inner join rfv_value on rfv_recency.customer=rfv_value.customer inner join Accounts on rfv_value.Customer = Accounts.id where Recency=" & request("Recency") & " and Value=" & request("Value") & " and Frequency=" & request("Frequency") & myCond & " order by" & ord
	'response.write mySQL
	'response.end
	set rs=Conn.Execute(mySQL)
	i=0
	while not rs.eof
		i=i+1
%>
	<tr>
		<td><%=rs("Day")%></td>
		<td><%=rs("Month")%></td>
		<td><%=rs("Count")%></td>
		<td><%=Round(CInt(rs("Count"))/Cint(rs("Month")),2)%></td>
		<td><%=Separate(rs("Amount"))%></td>
		<td><%=Separate(round(CDbl(rs("amount"))/cint(rs("month"))))%></td>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("Customer")%>"><%=rs("accountTitle")%></a></td>
		<td><%=Separate(rs("Remain"))%></td>
	</tr>
<%
		rs.moveNext
	wend
	rs.close
%>
	<tr>
		<td>Ã„⁄ ﬂ·:</td>
		<td colspan="7"><%=i%></td>
	</tr>
</table>
<%
elseif request("act")="showCool" then '---------------------------- C O O L --------------------------------------
	if request("ord")="" then 
		ord= " id"
	elseif request("ord")="1" then 
		ord= " id"
	elseif request("ord")="-1" then 
		ord= " id desc"
	elseif request("ord")="2" then 
		ord= " accountTitle"
	elseif request("ord")="-2" then 
		ord= " accountTitle desc"
	elseif request("ord")="3" then 
		ord= " createdDate"
	elseif request("ord")="-3" then 
		ord= " createdDate desc"
	end if
%>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td>
			<a href="rfmModel.asp?act=showCool&ord=<%if request("ord")="1" then response.write "-1" else response.write "1"%>">‘„«—Â „‘ —Ì</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showCool&ord=<%if request("ord")="2" then response.write "-2" else response.write "2"%>">‰«„ „‘ —Ì</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showCool&ord=<%if request("ord")="3" then response.write "-3" else response.write "3"%>"> «—ÌŒ «ÌÃ«œ</a>
		</td>
	</tr>
<%
	if request("accountGroup")="" or cint(request("accountGroup"))=-1 then 
		mySQL="select * from Accounts where type in (2,4) and id not in (select distinct customer from Orders union select distinct customer from quotes) order by " & ord
	else 
		mySQL="select Accounts.* from Accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & request("accountGroup") & " and type in (2,4) and id not in (select distinct customer from Orders union select distinct customer from quotes) order by " & ord
	end if
	set rs=Conn.Execute(mySQL)
	i=0
	while not rs.eof  
		i=i+1
%>
	<tr>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("id")%>"><%=rs("id")%></a></td>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("id")%>"><%=rs("accountTitle")%></a></td>
		<td><%=rs("createdDate")%></td>
	</tr>
<%		
		rs.moveNext
	wend
	rs.close
%>
	<tr>
		<td> ⁄œ«œ</td>
		<td colspan="2"><%=i%></td>
	</tr>
</table>
<%
elseif request("act")="showWarm" then '---------------------------------- W A R M -------------------------------
	if request("ord")="" then 
		ord= " id"
	elseif request("ord")="1" then 
		ord= " id"
	elseif request("ord")="-1" then 
		ord= " id desc"
	elseif request("ord")="2" then 
		ord= " accountTitle"
	elseif request("ord")="-2" then 
		ord= " accountTitle desc"
	elseif request("ord")="3" then 
		ord= " createdDate"
	elseif request("ord")="-3" then 
		ord= " createdDate desc"
	end if
%>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td>
			<a href="rfmModel.asp?act=showWarm&ord=<%if request("ord")="1" then response.write "-1" else response.write "1"%>">‘„«—Â „‘ —Ì</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showWarm&ord=<%if request("ord")="2" then response.write "-2" else response.write "2"%>">‰«„ „‘ —Ì</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showWarm&ord=<%if request("ord")="3" then response.write "-3" else response.write "3"%>"> «—ÌŒ «ÌÃ«œ</a>
		</td>
	</tr>
<%
	if request("accountGroup")="" or cint(request("accountGroup"))=-1 then 
		mySQL="select * from accounts where id in (select customer from (select distinct customer from Quotes where Customer not in (select distinct customer from Orders)) drv) order by " & ord
	else
		mySQL="select accounts.* from accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & request("accountGroup") & " and id in (select customer from (select distinct customer from Quotes where Customer not in (select distinct customer from Orders)) drv) order by " & ord
	end if
	'response.write mySQL
	'response.end
	set rs=Conn.Execute(mySQL)
	i=0
	while not rs.eof  
		i=i+1
%>
	<tr>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("id")%>"><%=rs("id")%></a></td>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("id")%>"><%=rs("accountTitle")%></a></td>
		<td><%=rs("createdDate")%></td>
	</tr>
<%		
		rs.moveNext
	wend
	rs.close
%>
	<tr>
		<td> ⁄œ«œ</td>
		<td colspan="2"><%=i%></td>
	</tr>
</table>
<%
elseif request("act")="showThreshold" then '------------------------------ T H R E S H O L D -------------------------------------
	if request("ord")="" then 
		ord= " orders.customer"
	elseif request("ord")="1" then 
		ord= " orders.customer"
	elseif request("ord")="-1" then 
		ord= " orders.customer desc"
	elseif request("ord")="2" then 
		ord= " accountTitle"
	elseif request("ord")="-2" then 
		ord= " accountTitle desc"
	elseif request("ord")="3" then 
		ord= " orders.createdDate"
	elseif request("ord")="-3" then 
		ord= " orders.createdDate desc"
	elseif request("ord")="4" then 
		ord= " invoiceCreatedDate"
	elseif request("ord")="-4" then 
		ord= " invoiceCreatedDate desc"
	elseif request("ord")="5" then 
		ord= " order_title"
	elseif request("ord")="-5" then 
		ord= " order_title desc"
	end if
%>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="1" then response.write "-1" else response.write "1"%>">‘„«—Â ”›«—‘</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="2" then response.write "-2" else response.write "2"%>">‰«„ „‘ —Ì</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="3" then response.write "-3" else response.write "3"%>"> «—ÌŒ «ÌÃ«œ ”›«—‘</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="4" then response.write "-4" else response.write "4"%>"> «—ÌŒ «ÌÃ«œ ›«ﬂ Ê—</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="5" then response.write "-5" else response.write "5"%>">⁄‰Ê«‰ ”›«—‘</a>
		</td>
	</tr>
<%
	if request("accountGroup")="" or cint(request("accountGroup"))=-1 then 
		mySQL="select orders.id,orders.createdDate, orders.customer, accounts.accountTitle, isnull(invoices.createdDate,-1) as invoiceCreatedDate, orders_trace.order_title, invoices.id as InvoiceID from Orders inner join accounts on orders.customer=accounts.id inner join orders_trace on orders.id = orders_trace.radif_sefareshat left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id where accounts.Status=1 and orders.Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1)) order by " & ord
	else
		mySQL= "select orders.id,orders.createdDate, orders.customer, accounts.accountTitle, isnull(invoices.createdDate,-1) as invoiceCreatedDate, orders_trace.order_title, invoices.id as InvoiceID from Orders inner join accounts on orders.customer=accounts.id inner join  AccountGroupRelations on AccountGroupRelations.account=accounts.id inner join orders_trace on orders.id = orders_trace.radif_sefareshat left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id where AccountGroupRelations.accountgroup=" & request("accountGroup") & " and accounts.Status=1 and orders.Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1)) order by " & ord
	end if
	set rs=Conn.Execute(mySQL)
	i=0
	while not rs.eof  
		i=i+1
%>
	<tr>
		<td><a href="../order/TraceOrder.asp?act=show&order=<%=rs("id")%>"><%=rs("id")%></a></td>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("customer")%>"><%=rs("accountTitle")%></a></td>
		<td><%=rs("createdDate")%></td>
		<td>
<%
		if rs("invoiceCreatedDate")="-1" then 
			response.write "›«ﬂ Ê— ‰œ«—œ"
		else
%>
			<a href="../AR/AccountReport.asp?act=showInvoice&invoice=<%=rs("invoiceID")%>"><%=rs("invoiceCreatedDate")%></a>
<%
		end if
%>
		</td>
		<td><%=rs("order_title")%></td>
	</tr>
<%		
		rs.moveNext
	wend
	rs.close
%>
	<tr>
		<td> ⁄œ«œ</td>
		<td colspan="4"><%=i%></td>
	</tr>
</table>
<%
end if
%>

<!--#include file="tah.asp" -->
