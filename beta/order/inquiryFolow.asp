<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="������ �������"
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
		<div class="RightHead">������</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="�����" then
				if request("isDelay")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="isDelay" >
			<label>����� ���</label>
		</div>
		<div class="Right">
			<input type="checkbox"<%if request("submit")="�����" then 
				if request("today")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="today" >
			<label>�����</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="�����" then 
				if request("tomorrow")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="tomorrow" >
			<label>����</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="�����" then 
				if request("nextWeek")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="nextWeek" >
			<label>2 �� 7 ��� ���</label>
		</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="�����" then 
				if request("moreNextWeek")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="moreNextWeek" >
			<label>����� �� 7 ���</label>
		</div>
	</div>
	<div class="NewRow">
		<div class="RightHead">��� �������</div>
	<%
	set rs=Conn.Execute("select * from OrderTypes where isActive=1")
	while not rs.eof
	%>
		<div class="Right">
			<input type="checkbox" <%if request("submit")="�����" then 
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
		<input type="submit" name="submit" value="�����">
		<a class="link" href="Inquiry.asp?act=advancedSearch&Submit=�����&check_marhale=on&marhale_not_check=on&marhale_box=4&check_closed=on">�� � ����� ������</a>
		
	</div>
</form>
<div style="clear: both;margin:20px 0 0 0;">
<center>
<%
dim orderType(10)
set rs=Conn.Execute("select * from OrderTypes where isActive=1")
i=0
while not rs.eof
	orderType(i)="orderType-"&rs("id")
	rs.moveNext
	i=i+1
wend
orderTypeCount=i-1
rs.close
%>
	<table class="CustTable">
		<tr>
	<%
	fromDate=""
	toDate="2100-12-30" 'shamsiToday()
	orderTypes=""
	condition="" 
	if request("submit")="�����" then 
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
		 		condition = condition & " or orders.returnDate between '2010-03-21' and '" & dateadd("d",-1,date()) & "'"
		 		fromDate = "2010-03-21"
		 		toDate = shamsiDate(dateadd("d",-1,date()))
		 	end if
			if request("today")="on" then 
				condition = condition & " or orders.returnDate = '" & date() & "'"
				if fromDate = "" then fromDate = shamsiToday()
				toDate = shamsiToday()
			end if
			if request("tomorrow")="on" then 
				condition = condition & " or orders.returnDate = '" & dateadd("d",1,date()) & "'"
				if fromDate = "" then fromDate = dateadd("d",1,date())
				toDate = dateadd("d",1,date())
			end if
			if request("nextWeek")="on" then 
				condition = condition & " or orders.returnDate between '" & dateadd("d",2,date()) & "' and '" & dateadd("d",7,date()) & "'"
				if fromDate = "" then fromDate = dateadd("d",2,date())
				toDate = dateadd("d",7,date())
			end if
			if request("moreNextWeek")="on" then 
				condition = condition & " or orders.returnDate > '" & dateadd("d",7,date()) & "'"
				if fromDate = "" then fromDate = dateadd("d",8,date())
				toDate = "2100-12-30"
			end if
			condition = condition & ")"
		end if
	end if
	if fromDate="" then fromDate="2010-03-21"
	mySQL = "select isnull(count(id),0) as quotesNew from orders where isClosed=0 and isOrder=0" & condition
	set rs=Conn.Execute(mySQL)
	%>
			<td title="���� ������ ������ ���� ����">
			<center>
				<img src="../images/folder1.gif" align="top">
				<br>
				<a href="inquiryFolow.asp?act=show&fromDate=<%=Server.URLEncode(fromDate)%>&toDate=<%=Server.URLEncode(toDate)%>&orderTypes=<%=Server.URLEncode(orderTypes)%>&isClose=0">��� ���<%=" (" & rs("quotesNew") & ")"%></a>
			</center>
			</td>
	<%
	mySQL = "select isnull(count(id),0) as quotesClose from orders where isClosed=1 and isOrder=0" & condition
	set rs=Conn.Execute(mySQL)
	%>
			<td title="���� ������ ������ ���� ����">
			<center>
				<img src="../images/folder1.gif" align="top">
				<br>
				<a href="inquiryFolow.asp?act=show&fromDate=<%=Server.URLEncode(fromDate)%>&toDate=<%=Server.URLEncode(toDate)%>&orderTypes=<%=Server.URLEncode(orderTypes)%>&isClose=1">�� ���<%=" (" & rs("quotesClose") & ")"%></a>
			</center>
			</td>
		</tr>
	</table>
</center>
</div>
<%
rs.close
elseif request("act")="show" then
%>
<div id='traceResult'></div>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		TransformXmlURL('/service/xml_getOrderTrace.asp?act=getQuoteFolow&isClose=<%=request("isClose")%>&fromDate=<%=request("fromDate")%>&toDate=<%=request("toDate")%>&orderTypes=<%=request("orderTypes")%>',"/xsl/orderShowList.xsl", function(result){
			$("#traceResult").html(result);
			$("#traceResult td.orderDates").each(function(i){
				var createdDate = $(this).find(".createdDate");
				var createdTime = $(this).find(".createdTime");
				var returnTime = $(this).find(".returnTime");
				var returnDate = $(this).find(".returnDate");
				if (returnDate.html()=="0"){
					returnDate.html("���� ����� ����!");
					returnTime.html("");
				} else {
					returnDate.html($.jalaliCalendar.gregorianToJalaliStr(returnDate.html()));
					var myTime = new Date(returnTime.html().replace(RegExp('-','g'),'/'));
					if (myTime.getHours()==0 && myTime.getMinutes()==0)
						returnTime.html("");
					else
					returnTime.html("("+((myTime.getHours()<10)?('0'+myTime.getHours()):myTime.getHours())+':'+((myTime.getMinutes() < 10)?('0'+myTime.getMinutes()):(myTime.getMinutes()))+")");
				}
				
				createdDate.html($.jalaliCalendar.gregorianToJalaliStr(createdDate.html()));
				
				var myTime = new Date(createdTime.html().replace(RegExp('-','g'),'/'));
				if (myTime.getHours()==0 && myTime.getMinutes()==0)
					createdTime.html("");
				else
				createdTime.html("("+((myTime.getHours()<10)?('0'+myTime.getHours()):myTime.getHours())+':'+((myTime.getMinutes() < 10)?('0'+myTime.getMinutes()):(myTime.getMinutes()))+")");
				
			});
		});
	});
</script>
<%

end if

%>
<!--#include file="tah.asp" -->