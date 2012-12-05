<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="œ‘»Ê—œ «” ⁄·«„"
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
		<div class="RightHead">«⁄ »«—</div>
		<div class="Right">
			<input type="checkbox" <%if request("submit")=" «ÌÌœ" then
				if request("isDelay")="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="isDelay" >
			<label>„‰ﬁ÷Ì ‘œÂ</label>
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
		<div class="RightHead">‰Ê⁄ «” ⁄·«„</div>
	<%
	set rs=Conn.Execute("select * from OrderTypes where isActive=1")
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
	<div class="NewRow">
		<input type="submit" name="submit" value=" «ÌÌœ" class="btn">
		
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
	if request("submit")=" «ÌÌœ" then 
		for ii=0 to orderTypeCount
			if request(orderType(ii))="on" then orderTypes = orderTypes & Split(orderType(ii),"-")(1) & ","
		next
		if len(orderTypes)>0 then 
			orderTypes = mid(orderTypes,1,len(orderTypes)-1)
			condition=" and Orders.type in (" & orderTypes & ")"
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
			<td title="»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ">
			<center>
				<img src="../images/folder1.gif" align="top">
				<br>
				<a href="inquiryFolow.asp?act=show&fromDate=<%=Server.URLEncode(fromDate)%>&toDate=<%=Server.URLEncode(toDate)%>&orderTypes=<%=Server.URLEncode(orderTypes)%>&isClose=0">À»  ‘œÂ<%=" (" & rs("quotesNew") & ")"%></a>
			</center>
			</td>
	<%
	mySQL = "select isnull(count(id),0) as quotesClose from orders where isClosed=1 and isOrder=0" & condition
	set rs=Conn.Execute(mySQL)
	%>
			<td title="»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ">
			<center>
				<img src="../images/folder1.gif" align="top">
				<br>
				<a href="inquiryFolow.asp?act=show&fromDate=<%=Server.URLEncode(fromDate)%>&toDate=<%=Server.URLEncode(toDate)%>&orderTypes=<%=Server.URLEncode(orderTypes)%>&isClose=1">—œ ‘œÂ<%=" (" & rs("quotesClose") & ")"%></a>
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
<div id='traceResult'>
	<center>
		<img style="margin:50px;" src="/images/ajaxLoad.gif"/>
	</center>
</div>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		TransformXmlURL('/service/xml_getOrderTrace.asp?act=getQuoteFolow&isClose=<%=request("isClose")%>&fromDate=<%=request("fromDate")%>&toDate=<%=request("toDate")%>&orderTypes=<%=request("orderTypes")%>',"/xsl.<%=version%>/orderShowList.xsl", function(result){
			$("#traceResult").html(result);
			$("#traceResult td.orderDates span.faDate").each(function(i){
				var myDate = $(this);
				var today = new Date();
				if (myDate.html()=="0"){
					myDate.html("----/--/--");
				} else {
					var myTime = new Date(myDate.html().replace(RegExp('-','g'),'/'));
					myDate.html($.jalaliCalendar.gregorianToJalaliStr(myDate.html()));
					var diff = daydiff(today,myTime);
					switch (true){
						case (diff < -1):
							myDate.attr("class","isRed");
							break;
						case (diff >=-1 && diff < 1):
							myDate.attr("class","isBlue");
							break;
						case (diff >= 1):
							myDate.attr("class","isBlack");
							break;
					}
					if (myTime.getHours()!=0 || myTime.getMinutes()!=0)
						myDate.attr("title", "("+((myTime.getHours()<10)?('0'+myTime.getHours()):myTime.getHours())+':'+((myTime.getMinutes() < 10)?('0'+myTime.getMinutes()):(myTime.getMinutes()))+")");
				}	
			});
			$(".list tr:not(.head):not(.sumTotal)").find("td:first").each(function(i,no){
				$(no).html(i+1);
			});
			var sumTotal = 0;
			$(".list tr:not(.head):not(.sumTotal)").find("td:last").each(function(i,no){
				sumTotal += getNum($(no).html());
			});
			$(".list .sumTotal td:last").html(echoNum(sumTotal));
		});
	});
</script>
<%

end if

%>
<!--#include file="tah.asp" -->