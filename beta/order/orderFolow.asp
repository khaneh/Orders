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
	set rs=Conn.Execute("select * from OrderTypes where isActive=1")
	while not rs.eof
	%>
		<div class="Right">
			<input type="checkbox" <%if request("submit")=" «ÌÌœ" then 
				if request("orderType-" & rs("id"))="on" then response.write " checked='checked' "
			else 
				response.write " checked='checked' "
			end if%> name="orderType-<%=rs("id")%>"/>
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
		<a class="btn btn-info" onclick="$('#openOrder').submit();"> „«Ì‘ ·Ì” Ì «“ ”›«—‘ùÂ«Ì »«“</a>
		<a class="btn btn-info" onclick="$('#orderHasNoRetDate').submit();">‰„«Ì‘ ·Ì” Ì «“ ”›«—‘ùÂ«ÌÌ ﬂÂ  «—ÌŒ  ÕÊÌ· ﬁ—«—œ«œ ‰œ«—œ</a>
	</div>
</form>
<form action="order.asp?act=advancedSearch" method="post" id="openOrder">
	<input name="checkClosed" value="on" type="hidden"/>
	<input name="checkIsOrder" value="on" type="hidden"/>
	<input name="isOrder" value="1" type="hidden"/>
	<input name="resultCount" value="500" type="hidden"/>
	<input type="hidden" name="submitBtn" value="Ã” ÃÊ"/>
</form>
<form action="order.asp?act=advancedSearch" method="post" id="orderHasNoRetDate">
	<input name="checkClosed" value="on" type="hidden"/>
	<input name="checkRetIsNull" value="on" type="hidden"/>
	<input name="checkIsOrder" value="on" type="hidden"/>
	<input name="isOrder" value="1" type="hidden"/>
	<input name="resultCount" value="500" type="hidden"/>
	<input type="hidden" name="submitBtn" value="Ã” ÃÊ"/>
</form>
<div style="clear: both;margin:20px 0 0 0;">
<center>
<%
set rs=Conn.Execute("select count(step) as stepCountLevel from orderSteps where IsActive=1 and step is not null group by step order by count(step) desc")
stepCountLevel = CInt(rs("stepCountLevel"))
rs.close
set rs=Conn.Execute("select max(step) as stepCount from orderSteps where IsActive=1 and step is not null")
stepCount = CInt(rs("stepCount"))
rs.close
dim steps(10,10)
dim orderType(10)
set rs=Conn.Execute("select * from OrderTypes where isActive=1")
i=0
while not rs.eof
	orderType(i)="orderType-"&rs("id")
	rs.moveNext
	i=i+1
wend
orderTypeCount=i-1
oldStep=-1
set rs=Conn.Execute("select id,step from orderSteps where IsActive=1 and step is not null order by step,ord")
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
				toDate=""
				orderTypes=""
				typeCondition="" 
				dateCondition="" 
				if request("submit")=" «ÌÌœ" then 
					for ii=0 to orderTypeCount
						if request(orderType(ii))="on" then orderTypes = orderTypes & Split(orderType(ii),"-")(1) & ","
					next
					if len(orderTypes)>0 then 
						orderTypes = mid(orderTypes,1,len(orderTypes)-1)
						typeCondition=" and orders.type in (" & orderTypes & ")"
					end if
					if request("isDelay")="on" or request("today")="on" or request("tomorrow")="on" or request("nextWeek")="on" or request("moreNextWeek")="on" then
					 	dateCondition = dateCondition & " and ( 0=1"
					 	if request("isDelay")="on" then 
					 		dateCondition = dateCondition & " or orders.returnDate between '2010-03-21' and '" & dateadd("d",-1,date()) & "'"
					 		fromDate = "2010-03-21"
					 		toDate = shamsiDate(dateadd("d",-1,date()))
					 	end if
						if request("today")="on" then 
							dateCondition = dateCondition & " or orders.returnDate = '" & Date() & "'"
							if fromDate = "" then fromDate = shamsiToday()
							toDate = shamsiToday()
						end if
						if request("tomorrow")="on" then 
							dateCondition = dateCondition & " or orders.returnDate = '" & dateadd("d",1,date()) & "'"
							if fromDate = "" then fromDate = dateadd("d",1,date())
							toDate = dateadd("d",1,date())
						end if
						if request("nextWeek")="on" then 
							dateCondition = dateCondition & " or orders.returnDate between '" & dateadd("d",2,date()) & "' and '" & dateadd("d",7,date()) & "'"
							if fromDate = "" then fromDate = dateadd("d",2,date())
							toDate = dateadd("d",7,date())
						end if
						if request("moreNextWeek")="on" then 
							dateCondition = dateCondition & " or orders.returnDate > '" & dateadd("d",7,date()) & "'"
							if fromDate = "" then fromDate = dateadd("d",8,date())
							toDate = "2100-12-30"
							'response.write toDate
						end if
						dateCondition = dateCondition & ")"
					end if
				end if
				'response.write request("moreNextWeek")
				if fromDate="" then fromDate="2010-03-21"
				if toDate="" then toDate="2100-12-30"
				mySQL = "select orderSteps.name,isnull(drv.orderCount,0) as orderCount from orderSteps left outer join (select step, count(id) as orderCount from orders where isClosed=0 and isOrder=1 and step = " & steps(s,i) & typeCondition & " and ((returnDate >'2010-03-21' " & dateCondition & ") or returnDate is null) group by step) drv on orderSteps.id=drv.step where orderSteps.id=" & steps(s,i)
				
				'if steps(s,i)=1 then response.write mySQL
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
%>
<div id='traceResult'>
	<center>
		<img style="margin:50px;" src="/images/ajaxLoad.gif"/>
	</center>
</div>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		TransformXmlURL('/service/xml_getOrderTrace.asp?act=getFolow&isOrder=1&fromDate=<%=request("fromDate")%>&toDate=<%=request("toDate")%>&orderTypes=<%=request("orderTypes")%>&step=<%=request("step")%>',"/xsl/orderShowList.xsl?v=<%=version%>", function(result){
			$("#traceResult").html(result);
			$("#traceResult td.orderDates").each(function(i){
				var createdDate = $(this).find(".createdDate");
				var createdTime = $(this).find(".createdTime");
				var returnTime = $(this).find(".returnTime");
				var returnDate = $(this).find(".returnDate");
				if (returnDate.html()=="0"){
					returnDate.html("›⁄·« „⁄·Ê„ ‰Ì” !");
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