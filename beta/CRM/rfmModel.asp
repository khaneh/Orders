<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' RFM
PageTitle="������ �������"
SubmenuItem=7
if not Auth(1 , 8) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<script type="text/javascript" src="/js/jquery.ui.datepicker-cc.js"></script>
<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery.ui.datepicker-cc-ar.js"></script>
<script type="text/javascript" src="/js/jquery.ui.datepicker-cc-fa.js"></script>
<script type="text/javascript">
	$(document).ready(function () {
		console.log("i'm ready :)");
		//if ($("input#myDate").val()=="") $("input#myDate").val($.format.date(new Date(),'yyyy/MM/dd'));
		$("input#faDate").datepicker({
			onSelect: function(dateText,inst) {
				$("input#myDate").val($.format.date(new JalaliDate(inst['selectedYear'], inst['selectedMonth'], inst['selectedDay']).getGregorianDate(), "yyyy/MM/dd")); 	
			},
			dateFormat: "yy/mm/dd"
		});
	});
</script>
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->

<STYLE>
	.myTable {border-color:black;border-width: medium;border-style: solid;}
	.myTable tr:nth-child(1) {border-color: gray;border-width: thin;border-style: solid; background-color: lime ;text-align: center;border-bottom-width: thick;border-bottom-color: black;border-bottom-style: solid;}
	.myTable tr:nth-child(2n+2) { background-color:#eee; text-align: center;}
	.myTable tr:nth-child(2n+3) { background-color:#C3DBEB; text-align: center;}
</STYLE>
<BR>
<%
dim color(6)
dim text_r(5)
dim text_f(5)
dim text_v(6)
color(0)="#555"
color(1)="#777"
color(2)="#999"
color(3)="#bbb"
color(4)="#ddd"
color(5)="#fff"
text_r(0)="��� �� �� ���"
text_r(1)="�� ��� ���"
text_r(2)="�� ��� ���"
text_r(3)="�� ��� ���"
text_r(4)="�� ��� ���"
text_f(0)="�� ����� �� �� ���"
text_f(1)="�� ����� �� �� ���"
text_f(2)="�� ����� �� �� ���"
text_f(3)="��� �� �� ����� �� �� ���"
'text_f(4)="�� �� ���"
text_v(0)="�� ������ ����"
text_v(1)="�� ����� ���� ����"
text_v(2)="�� �� ����� ����"
text_v(3)="�� ��� ����� ����"
text_v(4)="�� �� ����� ����"
text_v(5)="�� ����� ���� �� ����"
Conn.Execute("update Invoices set issuedDate_en=dbo.udf_date_solarToDate(cast(substring(issuedDate,1,4) as int),cast(substring(issuedDate,6,2) as int),cast(substring(issuedDate,9,2) as int)) where Issued=1 and issuedDate_en is null")
if request("act")="" then 
	if request("ord")="" then 
		ord= " R.Recency,F.Frequency,V.Value"
	elseif request("ord")="1" then 
		ord= " R.Recency,F.Frequency,V.Value"
	elseif request("ord")="-1" then 
		ord= " R.Recency desc, F.Frequency desc, V.Value desc"
	elseif request("ord")="2" then 
		ord= " F.Frequency, R.Recency, V.Value"
	elseif request("ord")="-2" then 
		ord= " F.Frequency desc, R.Recency desc, V.Value desc"
	elseif request("ord")="3" then 
		ord= " V.Value, R.Recency, F.Frequency"
	elseif request("ord")="-3" then 
		ord= " V.Value desc, R.Recency desc, F.Frequency desc"	
	elseif request("ord")="4" then 
		ord= " count(R.customer), R.Recency, F.Frequency, V.Value"
	elseif request("ord")="-4" then 
		ord= " count(R.customer) desc, R.Recency, F.Frequency, V.Value"
	end if
	
	if request("accountGroup")="" then
		accountGroup = -1
	else
		accountGroup = cint(request("accountGroup"))
	end if
	
%>
<form action="rfmModel.asp" method="post" id='myForm'>
	<select name="accountGroup" onchange='$("form#myForm").sumbit();'>
		<option <%if accountGroup=-1 then response.write "selected='selected'"%> value="-1">���� ��ݝ��</option>
	<%
		set rs=Conn.Execute("select * from AccountGroups")
		while not rs.eof
	%>
		<option <%if accountGroup = cint(rs("id")) then response.write "selected='selected'"%> value="<%=rs("id")%>">
			<%=rs("name")%>
		</option>
	<%
		rs.moveNext
		wend
		rs.close
	%>
	</select>
	<span>�����:</span>
	<input id='faDate' type="text" name="faDate" dir="ltr" size="10" value='<%if request("faDate")<>"" then response.write request("faDate") end if%>'>
	<input id='myDate' type="hidden" value='<%if request("myDate")<>"" then response.write request("myDate") end if%>' name='myDate'>
	<input type="submit" value="�����">
</form>
<table class="myTable" cellpadding="8px" cellspacing="0">
	<tr>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=accountGroup%>&ord=<%if request("ord")="1" then response.write "-1" else response.write "1" %>">����� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=accountGroup%>&ord=<%if request("ord")="2" then response.write "-2" else response.write "2" %>">����� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=accountGroup%>&ord=<%if request("ord")="3" then response.write "-3" else response.write "3" %>">������ ����</a>
		</td>
		<td>
			<a href="rfmModel.asp?accountGroup=<%=accountGroup%>&ord=<%if request("ord")="4" then response.write "-4" else response.write "4" %>">�����</a>
		</td>
	</tr>
<%
	if request("myDate")="" then 
		myDate = "getdate()"
		faDate = shamsiToday()
	else
		myDate="'" & request("myDate") & "'"
		faDate= request("faDate")
	end if
	sDate = splitDate (faDate)
	if accountGroup = -1 then 
		myCond=""
		mySQLCool="declare @date datetime;set @date = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select count(Accounts.id) as [count] from Accounts where Accounts.type in (2,4) and Accounts.id not in (select distinct customer from Orders where createdDate < @date union select distinct relatedID from messages where RelatedTable='accounts' and msgDate<'" & faDate & "')"
		'select * from Accounts where type in (2,4) and id not in (select distinct customer from Orders where createdDate<'1390/10/21' union select distinct customer from quotes where order_date<'1390/10/21' union select distinct relatedID from messages where RelatedTable='accounts' and msgDate<'1390/10/21') order by id
		mySQLWarm="declare @date datetime;set @date = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select count(id) as [count] from accounts where id in (select distinct customer from (select customer from Orders where createdDate < @date and isOrder=0 union select distinct relatedID from messages where RelatedTable='accounts' and msgDate<'" & faDate & "') as ddr where Customer not in (select distinct customer from Orders where createdDate < @date and isOrder=1))"
		'select * from accounts where id in (select customer from (select distinct customer from (select customer from Quotes where order_date<'1390/10/21' union select relatedID as customer from messages where RelatedTable='accounts' and msgDate<'1390/10/21') as ddr where Customer not in (select distinct customer from Orders where createdDate<'1390/10/21')) drv) order by id
		mySQLThreshold="select count(customer) as [count] from Orders where Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1))"
	else 
		myCond=" inner join AccountGroupRelations on AccountGroupRelations.account = R.customer  where AccountGroupRelations.accountGroup=" & accountGroup
		mySQLCool="declare @date datetime;set @date = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select count(Accounts.id) as [count] from Accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & accountGroup & " and Accounts.type in (2,4) and Accounts.id not in (select distinct customer from Orders where createdDate< @date) and accounts.id not in (select distinct relatedID from messages where RelatedTable='accounts' and msgDate<'" & faDate & "')"
		mySQLWarm="declare @date datetime;set @date = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select count(accounts.id) as [count] from accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & accountGroup & " and id in (select customer from (select distinct customer from (select customer from Orders where createdDate < @date and isOrder=0 union select  relatedID as customer from messages where RelatedTable='accounts' and msgDate<'" & faDate & "') as ddr where Customer not in (select distinct customer from Orders where createdDate < @date and isOrder=1)) drv)"
		mySQLThreshold="declare @date datetime;set @date = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select count(customer) as [count] from Orders inner join AccountGroupRelations on AccountGroupRelations.account=orders.customer where orders.createdDate < @date and AccountGroupRelations.accountGroup=" & accountGroup & "  and Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1))"
	end if
	
	'response.write mydate
	'response.end 
	'mySQL="select count(R.customer) as c, V.Value, R.Recency, F.Frequency from dbo.rfm_recency('" & myDate & "') as R left outer join dbo.rfm_frequency('" & myDate & "') as F on F.customer = R.customer left outer join dbo.rfm_value('" & myDate & "') as V on R.customer = V.customer " & myCond & " group by V.Value, R.Recency, F.Frequency order by" & ord
	mySQL="declare @todate datetime; set @todate=" & myDate & "; select count(R.customer) as c, V.Value, R.Recency, F.Frequency from (select customer, case when datediff(day,max(issuedDate_en),@toDate) > 365 then 0 when datediff(day,max(issuedDate_en),@toDate) between 181 and 365 then 1 when datediff(day,max(issuedDate_en),@toDate) between 91 and 180 then 2 when datediff(day,max(issuedDate_en),@toDate) between 31 and 90 then 3 when datediff(day,max(issuedDate_en),@toDate) between 0 and 30 then 4 end as Recency, datediff(day,max(issuedDate_en),@toDate) as [Day] from Invoices where Voided=0 and Issued=1 and issuedDate_en <=@toDate group by Customer) as R left outer join ("
	mySQL=mySQL+"select Invoices.customer,case when count(issuedDate) = 1 then 0 when count(issuedDate) = 2 then 1 when count(issuedDate) = 3 then 2 when count(issuedDate) > 3 then 3 end as Frequency , count(issuedDate) as [Count] from Invoices inner join (select customer,max(issuedDate_en) as Issue from Invoices where voided=0 and issuedDate_en <=@todate group by Customer) as maxIssue on invoices.Customer=maxIssue.customer where Invoices.voided=0 and Invoices.issued=1 and dateadd(month,-6,maxIssue.issue) < issuedDate_en group by Invoices.Customer"
	mySQL=mySQL+") as F on F.customer = R.customer left outer join ("
	mySQL=mySQL+"select Invoices.customer,case when sum(totalReceivable)/count(issuedDate) <= 100000 then 0 when sum(totalReceivable)/count(issuedDate) <= 500000 and sum(totalReceivable)/count(issuedDate) > 100000 then 1 when sum(totalReceivable)/count(issuedDate) <= 1000000 and sum(totalReceivable)/count(issuedDate) > 500000 then 2 when sum(totalReceivable)/count(issuedDate) <= 5000000 and sum(totalReceivable)/count(issuedDate) > 1000000 then 3 when sum(totalReceivable)/count(issuedDate) <= 10000000 and sum(totalReceivable)/count(issuedDate) > 5000000 then 4 when sum(totalReceivable)/count(issuedDate) > 10000000 then 5 end as Value, sum(totalReceivable) as Amount from Invoices inner join (select customer,max(issuedDate_en) as Issue from Invoices where Invoices.voided=0 and Invoices.issued=1 and issuedDate_en <=@toDate group by Customer) as maxIssue on invoices.Customer=maxIssue.customer where Invoices.voided=0 and Invoices.issued=1 and dateadd(month,-6,maxIssue.issue) < issuedDate_en group by Invoices.Customer"
	mySQL=mySQL+") as V on R.customer = V.customer " & myCond & " group by V.Value, R.Recency, F.Frequency order by " & ord
	'response.write mySQL
	'response.end
	set rs=Conn.Execute(mySQL)
	while not rs.eof
		myRow = "<tr><td"
		if not isnull(rs("Recency")) then myRow = myRow & " style='background-color:" & color(rs("Recency")) & "'"
		myRow = myRow & ">"
		if not isnull(rs("Recency")) then 
			myRow = myRow & text_r(rs("Recency")) & "<br>(" & rs("Recency") & ")"
		else 
			myRow = myRow & "�����" 
		end if
		myRow = myRow & "</td><td"
		if not isnull(rs("Frequency")) then myRow = myRow & " style='background-color:" & color(rs("Frequency")) & "'"
		myRow = myRow & ">"
		if not isnull(rs("Frequency")) then 
			myRow = myRow & text_f(rs("Frequency")) & "<br>(" & rs("Frequency") & ")"
		else 
			myRow = myRow & "�����" 
		end if
		myRow = myRow & "</td><td"
		if not isnull(rs("Value")) then myRow = myRow & " style='background-color:" & color(rs("Value")) & "'"
		myRow = myRow & ">"
		if not isnull(rs("Value")) then 
			myRow = myRow & text_v(rs("Value")) & "<br>(" & rs("Value") & ")"
		else 
			myRow = myRow & "�����" 
		end if
		myRow = myRow & "</td><td><a href='rfmModel.asp?act=show&accountGroup=" & accountGroup & "&value=" & rs("Value") & "&recency=" & rs("Recency") & "&frequency=" & rs("Frequency") & "&myDate=" & Server.URLEncode(myDate) & "'>" & Separate(rs("c")) & "</a></td></tr>"
		Response.write myRow
		rs.moveNext
	wend
	rs.close
' 	Response.write mySQLCool

 	set rs=Conn.execute(mySQLCool)
%>
	<tr>
		<td title="��������� �� �� ����� ����� � �� ������� � �� �������� ���� ���� ��� ���" colspan="3">����� ���</td>
		<td><a href="rfmModel.asp?act=showCool&accountGroup=<%=accountGroup%>&faDate=<%=Server.URLEncode(faDate)%>"><%=Separate(rs("count"))%></a></td>
	</tr>
<%
	rs.close
' 	Response.end
	set rs=Conn.execute(mySQLWarm)
%>
	<tr>
		<td title="��������� �� ������� � �� �������� ����� ��� ����� ������" colspan="3">����� ���</td>
		<td><a href="rfmModel.asp?act=showWarm&accountGroup=<%=accountGroup%>&faDate=<%=Server.URLEncode(faDate)%>"><%=Separate(rs("count"))%></a></td>
	</tr>
<%
	rs.close
	set rs=Conn.execute(mySQLThreshold)
%>
	<tr>
		<td title="��������� �� ����� ����� � ������ ����� ����� ������" colspan="3">����� ������</td>
		<td><a href="rfmModel.asp?act=showThreshold&accountGroup=<%=accountGroup%>&faDate=<%=Server.URLEncode(faDate)%>"><%=Separate(rs("count"))%></a></td>
	</tr>
</table>
<%
elseif request("act")="show" then 
%>
<span>R: <%=request("recency")%>, F: <%=request("frequency")%>, V: <%=request("value")%></span>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td width="30px">
			<a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&myDate=<%=Server.URLEncode(request("myDate"))%>&ord=<%if request("ord")="1" then response.write "-1" else response.write "1"%>">����� ����� (���)</a>
		</td>
		<td width="30px">
			<a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&myDate=<%=Server.URLEncode(request("myDate"))%>&ord=<%if request("ord")="3" then response.write "-3" else response.write "3"%>">����� ����� �� �� ���</a>
		</td>
		<td width="30px">
			<a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&myDate=<%=Server.URLEncode(request("myDate"))%>&ord=<%if request("ord")="4" then response.write "-4" else response.write "4"%>">������ ������</a>
		</td>
		<td width="50px">
			<a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&myDate=<%=Server.URLEncode(request("myDate"))%>&ord=<%if request("ord")="5" then response.write "-5" else response.write "5"%>">��� ����� �� �ԝ���</a>
		</td>
		<td width="50px">
			<a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&myDate=<%=Server.URLEncode(request("myDate"))%>&ord=<%if request("ord")="8" then response.write "-8" else response.write "8"%>">������ �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&myDate=<%=Server.URLEncode(request("myDate"))%>&ord=<%if request("ord")="6" then response.write "-6" else response.write "6"%>">�����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=show&accountGroup=<%=request("accountGroup")%>&value=<%=request("value")%>&recency=<%=request("recency")%>&frequency=<%=request("frequency")%>&myDate=<%=Server.URLEncode(request("myDate"))%>&ord=<%if request("ord")="7" then response.write "-7" else response.write "7" %>">��� �����</a>
		</td>
	</tr>
<%
	if request("ord")="" then 
		ord= " R.[day]"
	elseif request("ord")="1" then 
		ord= " R.[day]"
	elseif request("ord")="-1" then 
		ord= " R.[day] desc"
	'elseif request("ord")="2" then 
	'	ord= " [month]"
	'elseif request("ord")="-2" then 
	'	ord= " [Month] desc"
	elseif request("ord")="3" then 
		ord= " F.[count]"
	elseif request("ord")="-3" then 
		ord= " F.[count] desc"	
	elseif request("ord")="4" then 
		ord= " F.Frequency"
	elseif request("ord")="-4" then 
		ord= " F.Frequency desc"
	elseif request("ord")="5" then 
		ord= " V.amount"
	elseif request("ord")="-5" then 
		ord= " V.amount desc"
	elseif request("ord")="6" then 
		ord= " accounts.accountTitle"
	elseif request("ord")="-6" then 
		ord= " accounts.accountTitle desc"
	elseif request("ord")="7" then 
		ord= " remain"
	elseif request("ord")="-7" then 
		ord= " remain desc"
	elseif request("ord")="8" then 
		ord= " V.amount / F.[count]"
	elseif request("ord")="-8" then 
		ord= " V.amount / F.[count] desc"
	end if
	
	if request("accountGroup")="" or request("accountGroup")="-1" then 
		myCond=""
	else 
		myCond=" and accounts.id in (select account from AccountGroupRelations where AccountGroupRelations.accountGroup=" & request("accountGroup") & ") "
	end if
	'mySQL="select rfv_frequency.customer,rfv_frequency.[count],rfv_frequency.[Month],rfv_value.Amount,rfv_recency.[day],accounts.accountTitle, accounts.arBalance+accounts.aoBalance+accounts.apBalance as Remain from rfv_frequency inner join rfv_recency on rfv_frequency.customer=rfv_recency.customer inner join rfv_value on rfv_recency.customer=rfv_value.customer inner join Accounts on rfv_value.Customer = Accounts.id where Recency=" & request("Recency") & " and Value=" & request("Value") & " and Frequency=" & request("Frequency") & myCond & " order by" & ord
	mySQL="declare @todate datetime; set @todate=" & request("myDate") & "; select R.customer, V.amount, V.Value, R.Recency, R.[day], F.[count], F.Frequency, accounts.accountTitle, accounts.arBalance+accounts.aoBalance+accounts.apBalance as Remain from (select customer, case when datediff(day,max(issuedDate_en),@toDate) > 365 then 0 when datediff(day,max(issuedDate_en),@toDate) between 181 and 365 then 1 when datediff(day,max(issuedDate_en),@toDate) between 91 and 180 then 2 when datediff(day,max(issuedDate_en),@toDate) between 31 and 90 then 3 when datediff(day,max(issuedDate_en),@toDate) between 0 and 30 then 4 end as Recency, datediff(day,max(issuedDate_en),@toDate) as [Day] from Invoices where Voided=0 and Issued=1 and issuedDate_en <=@toDate group by Customer) as R left outer join ("
	mySQL=mySQL+"select Invoices.customer,case when count(issuedDate) = 1 then 0 when count(issuedDate) = 2 then 1 when count(issuedDate) = 3 then 2 when count(issuedDate) > 3 then 3 end as Frequency , count(issuedDate) as [Count] from Invoices inner join (select customer,max(issuedDate_en) as Issue from Invoices where voided=0 and issued=1 and issuedDate_en <=@todate group by Customer) as maxIssue on invoices.Customer=maxIssue.customer where voided=0 and issued=1 and dateadd(month,-6,maxIssue.issue) < issuedDate_en group by Invoices.Customer"
	mySQL=mySQL+") as F on F.customer = R.customer left outer join ("
	mySQL=mySQL+"select Invoices.customer,case when sum(totalReceivable)/count(issuedDate) <= 100000 then 0 when sum(totalReceivable)/count(issuedDate) <= 500000 and sum(totalReceivable)/count(issuedDate) > 100000 then 1 when sum(totalReceivable)/count(issuedDate) <= 1000000 and sum(totalReceivable)/count(issuedDate) > 500000 then 2 when sum(totalReceivable)/count(issuedDate) <= 5000000 and sum(totalReceivable)/count(issuedDate) > 1000000 then 3 when sum(totalReceivable)/count(issuedDate) <= 10000000 and sum(totalReceivable)/count(issuedDate) > 5000000 then 4 when sum(totalReceivable)/count(issuedDate) > 10000000 then 5 end as Value, sum(totalReceivable) as Amount from Invoices inner join (select customer,max(issuedDate_en) as Issue from Invoices where voided=0 and issued=1 and issuedDate_en <=@toDate group by Customer) as maxIssue on invoices.Customer=maxIssue.customer where voided=0 and issued=1 and dateadd(month,-6,maxIssue.issue) < issuedDate_en group by Invoices.Customer"
	mySQL=mySQL+") as V on R.customer = V.customer inner join Accounts on R.customer = Accounts.id " & myCond & " where R.Recency=" & request("Recency") 
	if request("Value")<>"" then mySQL = mySQL & " and V.Value=" & request("Value")
	if request("Frequency")<>"" then mySQL = mySQL & " and F.Frequency=" & request("Frequency")
	mySQL = mySQL & " order by " & ord
	'response.write mySQL
	'response.end
	set rs=Conn.Execute(mySQL)
	i=0
	while not rs.eof
		i=i+1
%>
	<tr>
		<td><%=rs("Day")%></td>
		<td><%=rs("Count")%></td>
		<td><%=rs("Frequency")%></td>
		<td><%=Separate(rs("Amount"))%></td>
		<td><%=Separate(Round(cdbl(rs("amount")) / CInt(rs("count"))))%></td>
		<td><a href="AccountInfo.asp?act=show&selectedCustomer=<%=rs("Customer")%>"><%=rs("accountTitle")%></a></td>
		<td><%=Separate(rs("Remain"))%></td>
	</tr>
<%
		rs.moveNext
	wend
	rs.close
%>
	<tr>
		<td>��� ��:</td>
		<td colspan="6"><%=i%></td>
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
	sDate = splitDate(request("faDate"))
 %>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td>
			<a href="rfmModel.asp?act=showCool&ord=<%if request("ord")="1" then response.write "-1" else response.write "1" %>">����� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showCool&ord=<%if request("ord")="2" then response.write "-2" else response.write "2" %>">��� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showCool&ord=<%if request("ord")="3" then response.write "-3" else response.write "3" %>">����� �����</a>
		</td>
	</tr>
 <%
	if request("accountGroup")="" or request("accountGroup")="-1" then 
		mySQL="declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select * from Accounts where type in (2,4) and id not in (select distinct customer from Orders where createdDate < @myDate union select distinct relatedID from messages where RelatedTable='accounts' and msgDate<'" & request("faDate") & "') order by " & ord
	else 
		mySQL="declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&"); select Accounts.* from Accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & request("accountGroup") & " and type in (2,4) and id not in (select distinct customer from Orders where createdDate < @myDate union select distinct relatedID from messages where RelatedTable='accounts' and msgDate<'" & request("faDate") & "') order by " & ord
	end if
	'response.write mySQL
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
		<td>�����</td>
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
	sDate = splitDate(Request("faDate"))
%>
<table class="myTable" cellpadding="5px" cellspacing="0">
	<tr>
		<td>
			<a href="rfmModel.asp?act=showWarm&ord=<% if request("ord")="1" then response.write "-1" else response.write "1" %>">����� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showWarm&ord=<% if request("ord")="2" then response.write "-2" else response.write "2" %>">��� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showWarm&ord=<% if request("ord")="3" then response.write "-3" else response.write "3" %>">����� �����</a>
		</td>
	</tr>
 <%
	if request("accountGroup")="" or request("accountGroup")="-1" then 
'		mySQL="declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select * from accounts where id in (select customer from (select distinct customer from (select customer from orders where createdDate < @myDate and isOrder=0 union select relatedID as customer from messages where RelatedTable='accounts' and msgDate<'" & request("faDate") & "') as ddr where Customer not in (select distinct customer from Orders where createdDate < @myDate and isOrder=1)) drv) order by " & ord
		mySQL="declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select accounts.* from Accounts inner join (select distinct customer from Orders where isOrder=0 and createdDate < @myDate and customer not in (select distinct customer from Orders where isOrder=1 and createdDate < @myDate) union select distinct relatedID as customer from messages where RelatedTable='accounts' and msgDate<'" & request("faDate") & "' and relatedID not in (select distinct customer from Orders where isOrder=1 and createdDate < @myDate)) dd on dd.customer=accounts.id order by " & ord
	else
' 		mySQL="declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select accounts.* from accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id where AccountGroupRelations.accountGroup=" & request("accountGroup") & " and id in (select customer from (select distinct customer from ((select customer from Orders where createdDate < @myDate and isOrder=0 union select relatedID as customer from messages where RelatedTable='accounts' and msgDate<'" & request("faDate") & "') as ddr where Customer not in (select distinct customer from Orders where createdDate < @myDate and isOrder=1)) drv) order by " & ord
		mySQL="declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select accounts.* from Accounts inner join AccountGroupRelations on AccountGroupRelations.account=accounts.id inner join (select distinct customer from Orders where isOrder=0 and createdDate < @myDate and customer not in (select distinct customer from Orders where isOrder=1 and createdDate < @myDate) union select distinct relatedID as customer from messages where RelatedTable='accounts' and msgDate<'" & request("faDate") & "' and relatedID not in (select distinct customer from Orders where isOrder=1 and createdDate < @myDate)) dd on dd.customer=accounts.id where AccountGroupRelations.accountGroup=" & request("accountGroup") & " order by " & ord
	end if
' 	response.write mySQL
' 	response.end
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
		<td>�����</td>
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
			<a href="rfmModel.asp?act=showThreshold&ord=<% if request("ord")="1" then response.write "-1" else response.write "1" %>&faDate=<%=Request("faDate")%>">����� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="2" then response.write "-2" else response.write "2" %>&faDate=<%=Request("faDate")%>">��� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="3" then response.write "-3" else response.write "3" %>&faDate=<%=Request("faDate")%>">����� ����� �����</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="4" then response.write "-4" else response.write "4" %>&faDate=<%=Request("faDate")%>">����� ����� ������</a>
		</td>
		<td>
			<a href="rfmModel.asp?act=showThreshold&ord=<%if request("ord")="5" then response.write "-5" else response.write "5" %>&faDate=<%=Request("faDate")%>">����� �����</a>
		</td>
	</tr>
<%
	sDate = splitDate(Request("faDate"))
	if request("accountGroup")="" or request("accountGroup")="-1" then 
		mySQL="declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select orders.id,orders.createdDate, orders.customer, accounts.accountTitle, isnull(invoices.createdDate,-1) as invoiceCreatedDate, orders.orderTitle, invoices.id as InvoiceID from Orders inner join accounts on orders.customer=accounts.id left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id where orders.createdDate < @myDate and accounts.Status=1 and orders.Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1)) order by " & ord
	else
		mySQL= "declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate("&sDate(0)&","&sDate(1)&","&sDate(2)&");select orders.id,orders.createdDate, orders.customer, accounts.accountTitle, isnull(invoices.createdDate,-1) as invoiceCreatedDate, orders.orderTitle, invoices.id as InvoiceID from Orders inner join accounts on orders.customer=accounts.id inner join  AccountGroupRelations on AccountGroupRelations.account=accounts.id left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id where orders.createdDate < @myDate and AccountGroupRelations.accountgroup=" & request("accountGroup") & " and accounts.Status=1 and orders.Customer in (select customer from invoices where Issued=0 and voided=0 and customer not in (select customer from Invoices where Voided=0 and issued=1)) order by " & ord
	end if
' 	Response.write mySQL
' 	Response.End
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
			response.write "������ �����"
		else
%>
			<a href="../AR/AccountReport.asp?act=showInvoice&invoice=<%=rs("invoiceID")%>"><%=rs("invoiceCreatedDate")%></a>
<%
		end if
%>
		</td>
		<td><%=rs("orderTitle")%></td>
	</tr>
<%		
		rs.moveNext
	wend
	rs.close
%>
	<tr>
		<td>�����</td>
		<td colspan="4"><%=i%></td>
	</tr>
</table>
<%
end if
%>

<!--#include file="tah.asp" -->
