<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
if request("isOrder")="yes" then
	PageTitle="”›«—‘"
else
	PageTitle="«” ⁄·«„"
end if

select case request("act")
	case "makeNewQoute"
		SubmenuItem=1
	case "customerSearch"
		SubmenuItem=1
	case "getType"
		SubmenuItem=1
	case "getNew"
		SubmenuItem=1
	case "getEdit"
		SubmenuItem=2
	case "edit"
		SubmenuItem=2
	case "getShow"
		SubmenuItem=4
	case "show"
		SubmenuItem=4
	case "trace"
		SubmenuItem=3
	case else
		SubmenuItem=3		
end select
if Request("act")<>"printQuote" then 
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->

<STYLE type="text/css">
	.CustTable {font-family:"b nazanin",tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.custTable tr:nth-child(even) {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.custTable tr:nth-child(odd) {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
	.rowTitle {margin: -5px 3px 0 0;color: #666;text-align: center;float: right;position: absolute;top: 0;}
	div.myRow {margin: 10px 0 20px 0;padding: 0 3px 5px 0;}
	div.exteraArea {margin: 5px 0;padding: 30px 3px 5px 0;position: relative;background-color: #DfDfDf;border-radius: 10px;border-width: 1px;border-color: #D1D1D1;border-style: solid;}
	label.key-label {margin: 0px 3px 0 0px;padding: 0 5px 0 5px;display: block;vertical-align: middle;width: 80px;font-size: 11pt;font-family: "b compset"}
	div.key-label{overflow: hidden;text-overflow: ellipsis;white-space: nowrap;}
	.myProp {font-weight: bold;color: #40F; margin: 0px 3px 0 0px;padding: 5px 0 5px 0;}
	div.btn label{background-color:yellow;color: blue;padding: 3px 30px 3px 30px;cursor: pointer;}
	div.btn {margin: -5px 250px 0px 5px;}
	div.btn img{margin: 0px 20px -5px 0;cursor: pointer;}
	div.btn-bar{text-align: center;}
	textarea.myDesc {width:170px;height: 45px;}
	span.sp1 {width: 80px;float: right;text-align: center;padding: 0;}
	span.sp2 {width: 30px;float: right;text-align: center;padding: 0;}
	div.funcBtn {width: 80px;margin: 0;float: left;position: absolute;top: 0;left: 0;}
	div.downBtn {width: 700px;margin-right: 26px;text-align: center;background-color: #CCC;padding: 20px 0 10px 0;}

	#errMsg {background-color: #ccc;width: 700px;color: red; font-size: 10pt;text-align: center;margin: 0 26px 0 0;}
	#totalPrice {background-color:#FED;border-width:0;}
	#totalVatedPrice {background-color:#FED;border-width:0;}
	#orderDetails {margin: 0 26px;}
	#orderOrgin{display: none;}
	.orderColor {background-color: black;color: yellow;}
	.quoteColor {background-color: #559;color: yellow;}
	.quoteColor td a:link{color: yellow;}
	.quoteColor td a:visited{color: #47FF00;}
	.quoteColor td a:hover{color: red;}
	.orderColor td a:link{color: yellow;}
	.orderColor td a:visited{color: #47FF00;}
	.orderColor td a:hover{color: red;}
	div.top1 {margin-top: 10px;margin-right: 20px;}
	.grayColor{background-color: #ccc;}
	.jsonSuggest li a img {float:left;margin-right:5px;}
	.jsonSuggest li a small {display:block;text-align:right;}
	.jsonSuggest { font-size:0.8em; }
	ul.jsonSuggest{width: 214px;}
	div.topBtn {margin: 10px 0 5px 0; text-align: center;position: relative;}
	div.stamp {; position: absolute;top: -20px;left: 10px;-webkit-filter: drop-shadow(9px 7px 13px #000);-webkit-transform: rotate(-20deg);-moz-filter: drop-shadow(9px 7px 13px #000);-moz-transform: rotate(-20deg);z-index: 999;border: 5px solid #559;padding: 10px;border-radius: 15px;}
	div.stamp h2 {color: yellow;}
	div.stamp div {margin: -10px 0 0 0;font-family: "b yekan";}
	div.stamp span {color: #636363;font-size: 7pt;}
	#approvedBy {color: black !important;font-size: 8pt !important;}
	
	div.mySearch{border-color: #555599;border-width: 5px;border-style: solid;background-color: #AAAAEE;padding: 10px;margin: 10px;}
	div.group {margin: 0px;padding: 0px 0 5px 0;min-height: 35px;position: relative;}
	fieldset.group {
		border-color: threedface;
		border-width: 10px 1px 1px 1px;
		border-style: solid;
		border-radius: 10px;
	}
	.offset0 {margin-right: 10px;}
	div.priceGroup {float: left;position: relative;top:0;left:0;}
	.groupTitle {
		margin: 0px -25px 0 0px;
		color: black;
		text-align: center;
		float: right;
/* 		background-color: #88F; */
		background: -webkit-gradient(linear, left top, left bottom, from(#aaF), to(#88F));
		background: -moz-linear-gradient(top, #55aaee, #003366);
		width: 25px;
		height: 100%;
		position: absolute;
	}
	span.groupName {
		font-weight: bold;
		font-size: 13pt;
		font-family: "B zar","B yekan";
		padding: 2px 10px;
		margin: 0 -3px;
	}
	i.groupIcon{
		display: block;
		position: absolute;
		right: 10px;
		top: 0px;
	}
	.disBtn{
	}
	.sec{float: right;width: 350px;height: 30px;position: relative;}
	.col0{width: 110px;float: right;margin-right: 20px;}
	.check{position: absolute;right: 0;}
	input.not{position: absolute;left: 0;}
	span.not{margin-left: 20px;float: left;}
	select.adv{width: 150px;height: 28px;}
	div.myBtn{text-align: center;}
	#orderStock{width: 350px;margin: 10px 20px 0 0;float: right;}
	#orderPurchase{width: 350px;margin: 10px 20px 0 0;float: right;}
	span.delCostBtn{float: left;opacity: .6;display: none;cursor: pointer;position: absolute;top:3px;left: 0;}
	td.delCost{position: relative;}
	div.customerHasInventory{position: relative;}
	div.customerHasInventory div{position: absolute;top:0;left: 0;opacity: .7;}
	textarea.orderNotes{width: 100%;height: 70px;}
	#voidedIcon{background-image: url('/images/voided.gif');width: 200px;height: 122px;position: absolute;top: 65px;left: 290px;}
	#printArea{position: relative;}
	#orderFiles{position: relative;margin:10px;width: 95%; direction: ltr;}
	#orderFiles div.fileName {font-size: small;padding: 3px 0;}
	#orderFiles span.fileSize{font-size: smaller;margin-left: 10px;}
	#orderLogs tr th {background-color: gray;color: yellow;cursor: pointer;font-family: "b zar","b yekan", tahoma;font-size: 11pt;}
	#orderLogs tr:nth-child(2n+1) {background-color: #CCC;}
	#orderLogs tr:nth-child(2n) {background-color: #FFF;}
	.pause{color: red;}
	.pause td {border-bottom: 1px solid red;border-top: 1px solid red;}
	.showLogedOrder {cursor: pointer;}
	#logStamp{font-family: "b yekan";font-size: 30pt;position: absolute;right: 300px;-webkit-transform: rotate(-20deg);z-index: 999;border: 2px solid red;padding: 10px;border-radius: 10px;display: none;}
	span.stName{padding: 0 5px;font-family: "b zar";font-weight: bold;font-size: 12px;color: #b04444}
	td.orderDates{direction: ltr;text-align: center;}
</STYLE>
<%
end if
if not Auth(2 , 9) then NotAllowdToViewThisPage() '«” ⁄·«„ 
isOrder=request("isOrder")
select case request("act")
		'-----------------------------------------------------------------------------------------------------	
	case "trace":				'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------	
		searchBox = request.form("search_box")
%>
<hr>
<FORM METHOD=POST ACTION="?act=trace" onSubmit="return checkValidation();">
	<div class="mySearch">
		<span>‰«„ „‘ —Ì Ì« ‘—ﬂ  Ì« ‘„«—Â ”›«—‘/«” ⁄·«„:</span>
		<INPUT TYPE="text" NAME="search_box" id="search_box" value="<%=searchBox%>" class="boot">
		<INPUT TYPE="submit" NAME="SubmitB" Value="Ã” ÃÊ" class="btn">
		<% if Auth(2 , 5) then %>
			<A class="btn offset1" HREF="order.asp?act=advancedSearch">Ã” ÃÊÌ ÅÌ‘—› Â</A>
		<% End If %>
	</div>
</FORM>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		$("#search_box").focus();
		TransformXmlURL("/service/xml_getOrderTrace.asp?act=search&searchBox=<%=Server.URLEncode(searchBox)%>","/xsl.<%=version%>/orderShowList.xsl", function(result){
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
<hr>
<div id='traceResult'>
	<center>
		<img style="margin:50px;" src="/images/ajaxLoad.gif"/>
	</center>
</div>
<%	
		'-----------------------------------------------------------------------------------------------------	
	case "":				'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------	
		Response.redirect "?act=getShow"

		'-----------------------------------------------------------------------------------------------------	
	case "search":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		
		'-----------------------------------------------------------------------------------------------------
	case "advancedSearch":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
%>
<hr>
<FORM METHOD=POST ACTION="?act=advancedSearch" onSubmit="return checkValidation();">
	<div class="mySearch">
		<div class="sec">
			<input type="checkbox" name="checkIsOrder" class="check" <%if request("checkIsOrder")="on" then response.write " checked='checked'"%>>
			<span class="col0">”›«—‘/«” ⁄·«„</span>
			<span>›ﬁÿ ”›«—‘</span>
			<input type="radio" name="isOrder" value="1" <%if request("isOrder")="1" then response.write " checked='checked'"%>>
			<span>›ﬁÿ «” ⁄·«„</span>
			<input type="radio" name="isOrder" value="0" <%if request("isOrder")="0" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkClosed" class="check" <%if request("checkClosed")="on" then response.write " checked='checked'"%>>
			<span class="col0">›ﬁÿ ”›«—‘ Â«Ì »«“</span>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkID" class="check" <%if request("checkID")="on" then response.write " checked='checked'"%>>
			<span class="col0">‘„«—Â ”›«—‘</span>
			<input type="text" name="fromID" value="<%=request("fromID")%>" size="10" class="boot num">
			<span> «</span>
			<input type="text" name="toID" value="<%=request("toID")%>" size="10" class="boot num">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkOrderType" class="check" <%if request("checkOrderType")="on" then response.write " checked='checked'"%>>
			<span class="col0">‰Ê⁄ ”›«—‘</span>
			<select name="orderType" class="btn adv">
		<%
		set rs=Conn.Execute("select * from orderTypes where IsActive=1")
		while not rs.eof
			sel=""
			if cint(request("orderType"))=cint(rs("id")) then sel=" selected='selected' "
			response.write "<option value='" & rs("id") & "' " & sel & ">" & rs("Name") & "</option>"
			rs.moveNext
		wend
		%>		
			</select>
			<span class="not">‰»«‘œ</span>
			<input type="checkbox" name="checkNotOrderType" class="not" <%if request("checkNotOrderType")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCreatedDate" class="check" <%if request("checkCreatedDate")="on" then response.write " checked='checked'"%>>
			<span class="col0"> «—ÌŒ «ÌÃ«œ</span>
			<input type="text" name="fromCreatedDate" value="<%=request("fromCreatedDate")%>" size="10" class="boot date">
			<span> «</span>
			<input type="text" name="toCreatedDate" value="<%=request("toCreatedDate")%>" size="10" class="boot date">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkStep" class="check" <%if request("checkStep")="on" then response.write " checked='checked'"%>>
			<span class="col0">„—Õ·Â</span>
			<select name="orderStep" class="btn adv">
		<%
		set rs=Conn.Execute("select * from orderSteps where IsActive=1")
		while not rs.eof
			sel=""
			if cint(request("orderStep"))=cint(rs("id")) then sel=" selected='selected' "
			response.write "<option value='" & rs("id") & "'" & sel & ">" & rs("Name") & "</option>"
			rs.moveNext
		wend
		%>		
			</select>
			<span class="not">‰»«‘œ</span>
			<input type="checkbox" name="checkNotStep" class="not" <%if request("checkNotStep")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkRetDate" class="check" <%if request("checkRetDate")="on" then response.write " checked='checked'"%>>
			<span class="col0"> «—ÌŒ  ÕÊÌ· ⁄„·Ì</span>
			<input type="text" name="fromRetDate" size="10" value="<%=request("fromRetDate")%>" class="boot date">
			<span> «</span>
			<input type="text" name="toRetDate" size="10" value="<%=request("toRetDate")%>" class="boot date">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkStatus" class="check" <%if request("checkStatus")="on" then response.write " checked='checked'"%>>
			<span class="col0">Ê÷⁄Ì </span>
			<select name="orderStatus" class="btn adv">
		<%
		set rs=Conn.Execute("select * from orderStatus where IsActive=1")
		while not rs.eof
			sel=""
			if cint(request("orderStatus"))=cint(rs("id")) then sel=" selected='selected' "
			response.write "<option value='" & rs("id") & "'" & sel & ">" & rs("Name") & "</option>"
			rs.moveNext
		wend
		%>		
			</select>
			<span class="not">‰»«‘œ</span>
			<input type="checkbox" name="checkNotStatus" class="not" <%if request("checkNotStatus")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkContractDate" class="check" <%if request("checkContractDate")="on" then response.write " checked='checked'"%>>
			<span class="col0"> «—ÌŒ  ÕÊÌ· ﬁ—«—œ«œ</span>
			<input type="text" name="fromContractDate" value="<%=request("fromContractDate")%>" size="10" class="boot date">
			<span> «</span>
			<input type="text" name="toContractDate" value="<%=request("toContractDate")%>" size="10" class="boot date">
		</div>
		
		<div class="sec">
			<input type="checkbox" name="checkRetIsNull" class="check" <%if request("checkRetIsNull")="on" then response.write " checked='checked'"%>>
			<span class="col0"> «—ÌŒ  ÕÊÌ· ﬁ—«—œ«œ ‰œ«—œ</span>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCompany" class="check" <%if request("checkCompany")="on" then response.write " checked='checked'"%>>
			<span class="col0">‰«„ ‘—ﬂ </span>
			<input type="text" name="companyName" value="<%=request("companyName")%>" size="20" class="boot">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCreatedBy" class="check" <%if request("checkCreatedBy")="on" then response.write " checked='checked'"%>>
			<span class="col0">”›«—‘ êÌ—‰œÂ</span>
			<select name="orderCreatedBy" class="btn adv">
		<%
		set rs=Conn.Execute("select * from users where display=1 order by realName")
		while not rs.eof
			sel=""
			if cint(request("orderCreatedBy"))=cint(rs("id")) then sel=" selected='selected' "
			response.write "<option value='" & rs("id") & "'" & sel	& ">" & rs("realName") & "</option>"
			rs.moveNext
		wend
		%>		
			</select>
			<span class="not">‰»«‘œ</span>
			<input type="checkbox" name="checkNotCreatedBy" class="not" <%if request("checkNotCreatedBy")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCustomer" class="check" <%if request("checkCustomer")="on" then response.write " checked='checked'"%>>
			<span class="col0">‰«„ „‘ —Ì</span>
			<input type="text" name="customerName" value="<%=request("customerName")%>" size="20" class="boot">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkTel" class="check" <%if request("checkTel")="on" then response.write " checked='checked'"%>>
			<span class="col0"> ·›‰</span>
			<input type="text" name="tel" value="<%=request("tel")%>" size="20" class="boot num">
		</div>
		
		<div class="sec">
			<input type="checkbox" name="checkTitle" class="check" <%if request("checkTitle")="on" then response.write " checked='checked'"%>>
			<span class="col0">⁄‰Ê«‰ ”›«—‘</span>
			<input type="text" name="orderTitle" value="<%=request("orderTitle")%>" size="20" class="boot">
		</div>
		<div class="myBtn">
			<span> ⁄œ«œ ‰ «ÌÃ</span>
			<input type="text" size="3" name="resultCount" class="myBtn num" value='<%if request("resultCount")="" then response.write "50" else response.write cint(request("resultCount"))%>'>
			<INPUT TYPE="submit" NAME="submitBtn" Value="Ã” ÃÊ" class="btn">
			
		</div>
	</div>
</FORM>
<hr>
<div id='traceResult'>
	<center>
		<img style="margin:50px;" src="/images/ajaxLoad.gif"/>
	</center>
</div>
<script type="text/javascript">
$(document).ready(function(){
	$(".sec input:not(.check)").each(function(){
		if($(this).prevAll(".check").is(":checked"))
			$(this).prop("disabled", false);
		else
			$(this).prop("disabled", true);
	});
	$(".sec select").each(function(){
		if($(this).prevAll(".check").is(":checked"))
			$(this).prop("disabled", false);
		else
			$(this).prop("disabled", true);
	});
});
$(".check").click(function(){
	if ($(this).is(":checked")){
		$(this).closest(".sec").find("input:not(.check)").prop("disabled", false);
		$(this).closest(".sec").find("select").prop("disabled", false);
	} else {
		$(this).closest(".sec").find("input:not(.check)").prop("disabled", true);
		$(this).closest(".sec").find("select").prop("disabled", true);
	}
});
</script>
<%	
		rs.close
		set rs=nothing
		if request("submitBtn")="Ã” ÃÊ" then
%>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		$("#search_box").focus();
		TransformXmlURL("/service/xml_getOrderTrace.asp?act=advanceSearch&<%=request.form%>","/xsl.<%=version%>/orderShowList.xsl", function(result){
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
		'-----------------------------------------------------------------------------------------------------
	case "getEdit":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
%>
<FORM METHOD=POST ACTION="?act=edit">
<div class="top1">
	<div class="rspan4">
		<span>‘„«—Â «” ⁄·«„/”›«—‘ —« Ê«—œ ﬂ‰Ìœ:</span>
		<INPUT Name="id" id="orderID" TYPE="text" maxlength="6" size="5" class="boot">
	</div>
	<div class="rspan1">
		<INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" class="btn">
	</div>
</div>
</FORM>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		$("#orderID").focus();
	});
</script>
<%
		
		'-----------------------------------------------------------------------------------------------------
	case "getShow":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
%>
<FORM METHOD=POST ACTION="?act=show">
<div class="top1">
	<div class="rspan4">
		<span>‘„«—Â «” ⁄·«„/”›«—‘ —« Ê«—œ ﬂ‰Ìœ:</span>
		<INPUT Name="id" id="orderID" TYPE="text" maxlength="6" size="5" class="boot">
	</div>
	<div class="rspan1">
		<INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" class="btn">
	</div>
</div>
</FORM>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		$("#orderID").focus();
	});
</SCRIPT>
<%
		
		'-----------------------------------------------------------------------------------------------------
	case "makeNewQoute":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
%>		
<br>
<FORM METHOD=POST ACTION="?act=customerSearch" onsubmit="if ($('#CustomerNameSearchBox').val()=='') return false;">
<div class="top1">
	<div class="rspan5"><B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox" id="CustomerNameSearchBox" class="boot">
	</div>
	<div class="rspan1">
		<INPUT TYPE="submit" value="Ã” ÃÊ" class="btn">
	</div>	
</div>
</FORM>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		$("#CustomerNameSearchBox").focus();
	});
</SCRIPT>
<%	
		'-----------------------------------------------------------------------------------------------------
	case "show":
		'-----------------------------------------------------------------------------------------------------
		orderID=request("id")
		if not IsNumeric(orderID) then 
			Response.redirect "?act=getShow&msg=" & Server.URLEncode("»Â ‰Ÿ— ‘„« ‘„«—Â ”›«—‘ ‰»«Ìœ ÌÂ ⁄œœ »«‘Âø")
		end if
%>
<script type="text/javascript" src="/js/jquery.printElement.min.js"></script>
<script type="text/javascript">
	function sleep(milliseconds) {
	  var start = new Date().getTime();
	  for (var i = 0; i < 1e7; i++) {
	    if ((new Date().getTime() - start) > milliseconds){
	      break;
	    }
	  }
	}
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=<%=orderID%>","/xsl.<%=version%>/orderShowHeader.xsl", function(result){
		$("#orderHeader").html(result);	
		$('td span[rel=tooltip]').tooltip();
		$('a#customerID').click(function(e){
			window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer='+$('a#customerID').attr("myID"), 'showCustomer');
			e.preventDefault();
		});
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=showOrder&id=<%=orderID%>","/xsl.<%=version%>/orderShowProperty.xsl", function(result){
		$("#orderDetails").html(result);
		$('div.priceGroup').tooltip({
	      selector: "div[rel=tooltip]"
	    });	
	    sleep(300);
	    var sum=0;
	    var sumTotal = getNum($("#sumTotal").text());
	    $(".dis").each(function(i,obj){
			if ($(obj).text().indexOf("%")>-1) {
				var per = parseInt($(obj).text().replace(/%/gi,''));
				var price = getNum($(obj).closest("div.priceGroup").find(".price").text());
				if (per>0 && per <100){
					sum += price / (100/per - 1);
				}	
			} else{
			   sum += getNum($(obj).text());
			}
	    });
	    if (sum>0){
	    	$("#sumDis").html(echoNum(sum));
	    	$("#sumDis").attr("data-original-title",$("#sumDis").attr("data-original-title")+" ("+ (sum/sumTotal*100).toFixed(2)+"%)");
	    }
	    sum=0;
	    $(".reverse").each(function(i,obj){
			if ($(obj).text().indexOf("%")>-1) {
				var per = parseInt($(obj).text().replace(/%/gi,''));
				var price = getNum($(obj).closest("div.priceGroup").find(".price").text());
				if (per>0 && per <100)
					sum += price / (100/per - 1);
			} else{
			   sum += getNum($(obj).text());
			}
	    });
	    if (sum>0){
	    	$("#sumReverse").html(echoNum(sum));
	    	$("#sumReverse").attr("data-original-title",$("#sumReverse").attr("data-original-title")+" ("+ (sum/sumTotal*100).toFixed(2)+"%)");
	    }
	    sum=0;
	    $(".over").each(function(i,obj){
			if ($(obj).text().indexOf("%")>-1) {
				var per = parseInt($(obj).text().replace(/%/gi,''));
				var price = getNum($(obj).closest("div.priceGroup").find(".price").text());
				if (per>0 && per <100)
					sum += price / (100/per - 1);
			} else{
			   sum += getNum($(obj).text());
			}
	    });
	    if (sum>0){
	    	$("#sumOver").html(echoNum(sum));
	    	$("#sumOver").attr("data-original-title",$("#sumOver").attr("data-original-title")+" ("+ (sum/sumTotal*100).toFixed(2)+"%)");
	    }
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=showLog&id=<%=orderID%>","/xsl.<%=version%>/orderShowlogs.xsl", function(result){
		$("#orderLogs").html(result);	
		$("#orderLogs td.date").each(function(i){
			$(this).html($.jalaliCalendar.gregorianToJalaliStr($(this).html()));
		});
		$("#orderLogs td.time").each(function(i){
			myTime = new Date($(this).html().replace(RegExp('-','g'),'/'));
			$(this).html(((myTime.getHours()<10)?('0'+myTime.getHours()):myTime.getHours())+':'+((myTime.getMinutes() < 10)?('0'+myTime.getMinutes()):(myTime.getMinutes())));
		});
		$("td.showLogedOrder").click(function(){
			var logID = $(this).parent().attr("logID");
			$("#logStamp").css("display","block");
			$("#orderHeader").html('<center><img style="margin:50px;" src="/images/ajaxLoad.gif"></center>');
			$("#orderDetails").html('<center><img style="margin:50px;" src="/images/ajaxLoad.gif"></center>');
			loadXMLDoc("/service/xml_getOrderProperty.asp?act=showHead&logID=" + logID, function(orderXML){
				TransformXml(orderXML, "/xsl.<%=version%>/orderShowHeader.xsl", function(result){
					$("#orderHeader").html(result);	
					$('a#customerID').click(function(e){
						window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer='+$('a#customerID').attr("myID"), 'showCustomer');
						e.preventDefault();
					});
				});
			});
			loadXMLDoc("/service/xml_getOrderProperty.asp?act=showOrder&logID=" + logID, function(orderXML){
				TransformXml(orderXML, "/xsl.<%=version%>/orderShowProperty.xsl", function(result){
					$("#orderDetails").html(result);
					$('div.priceGroup').tooltip({
				      selector: "div[rel=tooltip]"
				    });
				});
			});
		});
	});
	TransformXmlURL("/service/xml_getMessage.asp?act=related&table=orders&id=<%=orderID%>","/xsl.<%=version%>/showRelatedMessage.xsl", function(result){
		$("#orderMessages").html(result);
		$("td.msgBody").each(function(i){
			$(this).html($(this).html().replace(/\n/gi,"<br/>"));
			});
			$("input[name=addNewMessage]").click(function(){
				document.location="../home/message.asp?RelatedTable=orders&RelatedID=" + $("#orderID").val() + "&retURL=" + escape("../Order/order.asp?act=show&id=" + $("#orderID").val());
		});
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=stock&id=<%=orderID%>","/xsl.<%=version%>/orderShowRequestInventory.xsl", function(result){
		$("#orderStock").html(result);	
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=purchase&id=<%=orderID%>","/xsl.<%=version%>/orderShowRequestPurchase.xsl", function(result){
		$("#orderPurchase").html(result);	
	});
	TransformXmlURL("/service/xml_getCosts.asp?act=orderCost&id=<%=orderID%>","/xsl.<%=version%>/costListShow.xsl", function(result){
		$("#orderCosts").html(result);
		$(".delCost").mouseover(function(event){
			$(this).find(".delCostBtn").css("display","block");
		});
		$(".delCost").mouseout(function(event){
			$(this).find(".delCostBtn").css("display","none");
		});	
		$(".delcostBtn").click(function(){
			$("#costDelItemID").val($(this).attr("costID"));
			$(".delCostBtn[costID=" + $(this).attr("costID") + "]").closest("tr").find("td").each(function(i,e){
				$("#costDelComfirmList").append("<li>" + $(e).text() + "</li>");
			});
			$("#costDelComfirm").dialog("open");
		});
	});
	$(document).ready(function(){
		$("a[name=jobTicketPrint]").click(function(){
			var group = $(this).attr("group");
			$("#printArea").find("div.group").css("display","none");
			$("#printArea").find("div.group[groupname=" + group + "]").css("display","block");
			var myRow = $("#printArea").find("div.myRow div.group:visible").closest("div.myRow");
			$("#printArea").find("div.myRow").css("display","none");
			myRow.css("display","block");
			$("hr").css("display","none");
			$('#printArea').printElement({overrideElementCSS:['/css/order_property.css','/css/jame.css'],pageTitle:'›—„ ”›«—‘ - ' + $("#orderID").val(),printMode:'popup'});
			$("#printArea").find("div.group").css("display","block");
			$("#printArea").find("div.myRow").css("display","block");
			myRow.css("display","block");
		});
		$("#costDelComfirm").dialog({
			autoOpen: false,
			buttons: {" «ÌÌœ":function(){
				$.ajax({
					type:"POST",
					url:"/service/json_editCost.asp",
					data:{act:"del",id:$("#costDelItemID").val()},
					dataType:"json"
				}).done(function (data){
					if (data.status=="ok")
						$(".delCostBtn[costID=" + $("#costDelItemID").val() + "]").closest("tr").remove();
				});
				$(this).dialog("close");
			}},
			title: " «ÌÌœ"
		});
		$("#approveOrder").dialog({
			autoOpen: false,
			buttons: {" «ÌÌœ":function(){
				$(this).dialog("close");
				document.location = '?act=approve&id=<%=orderID%>';
			}},
			title: " «ÌÌœ"
		});
		$("#approveQuote").dialog({
			autoOpen: false,
			buttons: {" «ÌÌœ":function(){
				document.location='?act=approve&id=<%=orderID%>';
				$(this).dialog("close");
			}},
			title: " «ÌÌœ"
		});
		$('#approvedBtn').click(function(){
			if ($("#isOrder").val()=='False')
				$("#approveQuote").dialog("open");
			else
				$("#approveOrder").dialog("open");
		});
		$("#changeCustomer").dialog({
			autoOpen: false,
			title: " €ÌÌ— Õ”«» ”›«—‘",
			height: 350
		});
		$("#cancelBtn").click(function(){
			$("#cancelDlg").dialog("open");
		});
		$("#cancelDlg").dialog({
			autoOpen: false,
			title: "—œ «” ⁄·«„/”›«—‘",
			buttons: {" «ÌÌœ":function(){
				if ($("#cancelDlg input[name=cancelReason]").val()!=''){
					document.location='?act=cancel&id=<%=orderID%>&reason=' + escape($("#cancelDlg input[name=cancelReason]").val());
					$(this).dialog("close");
				}
			}}
		});
		$('#changeCustomerBtn').click(function(){
			$("#changeCustomer").dialog("open");
		});
		$("input[name=search]").jsonSuggest({
			url: '/service/json_getAccount.asp?act=account' + escape($("input[name=search]").val()),
			minCharacters: 5,
			maxResults: 20,
			width: 214,
			caseSensitive: false,
			onSelect: function(sel){
				document.location='?act=changeCustomer&id=<%=orderID%>&customer='+sel.id;
			}
		});
		$("#printBtn").click(function(){
			if ($("#isOrder").val()=='False' || ($("#isOrder").val()=='True' && $("#step").val()=='40')){
				$("#printQuoteDlg").dialog("open");
			} else {
				$.getJSON('/service/json_orderStatus.asp',{act:"print",orderID:<%=orderID%>},function(data){
					if(data.status=="ok")
						$('#printArea').printElement({overrideElementCSS:['/css/order_property.css','/css/jame.css'],pageTitle:'›—„ ”›«—‘ - ' + $("#orderID").val(),printMode:'popup'});
				});
			}
			
		});
		$("#printQuoteDlg").dialog({
			autoOpen: false,
			title: "‰ÕÊÂ ç«Å «” ⁄·«„",
			buttons: {" «ÌÌœ": function(){
				$("#printQuoteDlg").dialog("close");
				$.getJSON('/service/json_orderStatus.asp',{act:"print",orderID:<%=orderID%>},function(data){
					if(data.status=="ok")
						document.location='?act=printQuote&id=<%=orderID%>&by=<%=session("id")%>&inShort=' + $("#inShort").prop("checked");
				});
			}}
		});
		$("#convert2orderBtn").click(function(){
			$("#convert2orderDlg").dialog("open");
		});
		$("#convert2orderDlg").dialog({
			autoOpen:false,
			title: " »œÌ· »Â ”›«—‘",
			buttons: {" «ÌÌœ": function(){
				var loc='';
				loc = '?act=convertToOrder&id=<%=orderID%>';
				if ($("#isOrder").val()=="True")
					loc +="&isOrder=1";
				if ($("#convert2orderDlg input[name=orderReturnDate]").val()!='')
					loc += '&returnDate=' + escape($.jalaliCalendar.jalaliToGregorianStr($("#convert2orderDlg input[name=orderReturnDate]").val()));
				if ($("#convert2orderDlg input[name=orderReturnTime]").val()!='' && $("#convert2orderDlg input[name=orderReturnDate]").val()!='')
					loc += escape(" " + $("#convert2orderDlg input[name=orderReturnTime]").val());
				if ($("#convert2orderDlg input[name=customerApprovedDate]").val()!='' && $("#convert2orderDlg select[name=customerApprovedType] option:selected").val()!='-1'){
					loc += '&customerApprovedType=' + $("#convert2orderDlg select[name=customerApprovedType] option:selected").val(); 
					loc += '&customerApprovedDate=' + escape($.jalaliCalendar.jalaliToGregorianStr($("#convert2orderDlg input[name=customerApprovedDate]").val()));
					$(this).dialog("close");
					document.location = loc;
				} else {
					$("#errMsg").html(" «—ÌŒ Ê ‰Ê⁄ «—«∆Â «” ⁄·«„ »Â „‘ —Ì —Ê «⁄·«„ ﬂ‰Ìœ")
				}
			}}
		});
		$("#orderEditBtn").click(function(){
			if ($("#isOrder").val()=='False')
				document.location = '?act=edit&id=<%=orderID%>'
			else
				$("#orderEditDlg").dialog("open");
		});
		$("#orderEditDlg").dialog({
			autoOpen:false,
			title:" ÊÃÂ",
			buttons:{" «ÌÌœ": function(){
				$(this).dialog("close");
				document.location = '?act=edit&id=<%=orderID%>';
			}}
			
		});
		$("#changeReturnDateBtn").click(function(){
			$("#changeReturnDateDlg").dialog("open");
		});
		$("#changeReturnDateDlg").dialog({
			autoOpen:false,
			title:"«’·«Õ  «—ÌŒ  ÕÊÌ·",
			buttons:{" «ÌÌœ": function(){
				$(this).dialog("close");
				$.ajax({
					type:"POST",
					url:"/service/json_orderStatus.asp",
					data:{act:"returnDate",orderID:$("#orderID").val(),retDate:$.jalaliCalendar.jalaliToGregorianStr($("#changeReturnDateInput").val()) + " " + $("#changeReturnTimeInput").val()},
					dataType:"json"
				}).done(function (data){
					if (data.status=="ok")
						document.location = '?act=show&id=<%=orderID%>&msg=' + escape(' €ÌÌ—  «—ÌŒ  ÕÊÌ· «‰Ã«„ ‘œ.');
				});	
			}}
			
		});
		$("#pauseBtn").click(function(){
			$("#pauseDlg").dialog("open");
		});
		$("#pauseDlg").dialog({
			autoOpen:false,
			title:"‘—Õ Œÿ«",
			buttons: {" «ÌÌœ": function(){
				$("#pauseDlg").dialog("close");
				document.location='?act=pause&id=<%=orderID%>&createdBy=' + $("#createdBy").val() + '&reason=' + escape($("#pauseReason").val());
			}}
		});
		$("#orderLogs").dialog({
			autoOpen:false,
			title:" «—ÌŒçÂ  €ÌÌ—«  ”›«—‘",
			buttons:{"»” ‰":function(){
				$(this).dialog("close");
			}},
			height: 450,
			width: 500
		});
		$("#orderLogsBtn").click(function(){
			$("#orderLogs").dialog("open");
		});
/*
		if ($("#isClosed").val()!='False')
			$("#voidedIcon").css("display","block");
		else
			$("#voidedIcon").css("display","none");
*/
		$.getJSON("/service/json_getFile.asp",
			{act:"list",orderID:"<%=orderID%>"},
			function (js){
				$.each(js, function(i,s){
					$('<div title='+s.itemName+' class="fileName"><a href="'+s.realPath+'">'+s.itemPath+'</a><span class="fileSize">'+getReadableFileSizeString(s.itemSize)+' </span></div>').appendTo('#orderFiles');
				});
			});
	});
	function getReadableFileSizeString(fileSizeInBytes) {
	    var i = -1;
	    var byteUnits = [' kB', ' MB', ' GB', ' TB', 'PB', 'EB', 'ZB', 'YB'];
	    do {
	        fileSizeInBytes = fileSizeInBytes / 1024;
	        i++;
	    } while (fileSizeInBytes > 1024);
	
	    return Math.max(fileSizeInBytes, 0.1).toFixed(1) + byteUnits[i];
	};
</script>
<%
	set rs = Conn.Execute("SELECT orders.type,orders.step,orders.createdBy, orders.id, orders.isClosed,orders.isApproved,orders.isOrder, isnull(invoices.Approved,0) as Approved,isnull(invoices.Issued,0) as Issued, users.realName as approvedByName FROM orders left outer join InvoiceOrderRelations on InvoiceOrderRelations.[Order] = orders.id left outer join Invoices on Invoices.ID=InvoiceOrderRelations.Invoice left outer join users on orders.approvedBy=users.id WHERE orders.ID=" & orderID)
%>
<div id="costDelComfirm">
	<h3>¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ ⁄„·ﬂ—œ –Ì· Õ–› ‘Êœø</h3>
	<input type="hidden" id="costDelItemID" />
	<lu id="costDelComfirmList"></lu>
</div>
<div id="approveOrder">
	<h3> «ÌÌœ ”›«—‘ »Â „‰“·Â «„÷«¡ ¬‰ «“ Ã«‰» ‘„« „Ìù»«‘œ. Ê „⁄‰Ì ¬‰ «Ì‰ ŒÊ«Âœ »Êœ ﬂÂ ‘„« »‰œÂ«Ì –Ì· —« çﬂ ﬂ—œÂù«Ìœ</h3>
	<lu>
		<li><span>”›«—‘ »Â ’Ê—  ﬂ«„· À»  ‘œÂ «” </span></li>
		<li><span> «—ÌŒ  ÕÊÌ· »« „‘ —Ì çﬂ ‘œÂ «” </span></li>
	</lu>
</div>
<div id="approveQuote">
	<h3>’œÊ— «” ⁄·«„ »Â „‰“·Â  «ÌÌœ ‘„« „Ìù»«‘œ Ê „⁄‰Ì ¬‰ «Ì‰ ŒÊ«Âœ »Êœ ﬂÂ ‘„« »‰œÂ«Ì –Ì· —« çﬂ ﬂ—œÂù«Ìœ</h3>
	<lu>
		<li>«” ⁄·«„ »Â ’Ê—  ﬂ«„· À»  ‘œÂ «” </li>
		<li>‘„« œ—ŒÊ«”  „‘ —Ì —« »Â ’Ê—  ﬂ«„· À»  ﬂ—œÂù«Ìœ</li>
		<li>«” ⁄·«„ ¬„«œÂ «—”«· »—«Ì „‘ —Ì ŒÊ«Âœ »Êœ</li>
		<li>„Ì“«‰ ÅÌ‘ Å—œ«Œ  Ê ”——”Ìœ »Â œ—” Ì Ê«—œ ‘œÂ</li>
	</lu>
</div>
<div id="changeCustomer">
	<h5>‰«„ Ê Ì« ‘„«—Â „‘ —Ì —« Ê«—œ ﬂ‰Ìœ</h5>
	<input type="text" size="40" name="search"/>
	<input type="hidden" name="customerID"/>
</div>
<div id="cancelDlg">
	<h3>·ÿ›« œ·Ì· ŒÊœ —« ‘—Õ œÂÌœ</h3>
	<input type="text" size="30" name="cancelReason"/>
</div>
<div id="printQuoteDlg">
	<input type="checkbox" id="inShort"/>	
	<span>Œ·«’Â »«‘œ</span>
</div>
<div id="pauseDlg">
	<h5>œ·Ì·  Êﬁ› —« ‘—Õ œÂÌœ</h5>
	<input type="text" size="30" id="pauseReason"/>
</div>
<div id="convert2orderDlg">
	<h5> »œÌ· »Â ”›«—‘ »Â „‰“·Â  «ÌÌœ «” ⁄·«„ «“ Ã«‰» „‘ —Ì „Ìù»«‘œ</h5>
	<span>·ÿ›« ‰ÕÊÂ Ê  «—ÌŒ  «ÌÌœ «” ⁄·«„ —« »Ì«‰ ﬂ‰Ìœ</span>
	<select name="customerApprovedType">
		<option value="-1">--- «‰ Œ«» ﬂ‰Ìœ ---</option>
<%
	set customerApproveType = conn.Execute("select * from orderCustomerApprovedTypes")
	while not customerApproveType.EOF
		Response.write "<option value='" & customerApproveType("id") & "'>" & customerApproveType("name") & "</option>"
		customerApproveType.MoveNext
	wend
%>								
	</select>
	<br/>
	<span> «—ÌŒ  «ÌÌœ  Ê”ÿ „‘ —Ì</span>
	<input type="text" size="10" name="customerApprovedDate" class="date"/>
	<br/> «—ÌŒ  ÕÊÌ·: <input name="orderReturnDate" class="date" size="10"/>
	<input name="orderReturnTime" class="time" size="5"/>
	<br/>
	<span id="errMsg"></span>
</div>
<div id="orderEditDlg">
	<h6> ÊÃÂ œ«‘ Â »«‘Ìœ ﬂÂ ÊÌ—«Ì‘ Ìﬂ ”›«—‘ »Â „⁄‰Ì  €ÌÌ— œ— ﬁ—«—œ«œ »« „‘ —Ì ŒÊ«Âœ »Êœ! Ê «Ì‰ »œ«‰ „⁄‰« ŒÊ«Âœ »Êœ ﬂÂ ”›«—‘ ‘„« „ Êﬁ› ŒÊ«Âœ ‘œ  «  «ÌÌœ „‘ —Ì »—”œ.</h6>
</div>
<div id='changeReturnDateDlg'>
	<span> «—ÌŒ  ÕÊÌ·</span>
	<input type="text" size="10" id='changeReturnDateInput' class="date">
	<span>”«⁄   ÕÊÌ·</span>
	<input type="text" size="5" id='changeReturnTimeInput' class="time">
</div>
<input type="hidden" id='isOrder' value="<%=rs("isOrder")%>">
<input type="hidden" id='isClosed' value="<%=rs("isClosed")%>">
<input type="hidden" id='orderID' value="<%=rs("id")%>">
<input type="hidden" id='createdBy' value="<%=rs("createdBy")%>">
<input type="hidden" id='step' value="<%=rs("step")%>">
<div class="topBtn">
<%	
	orderType = CInt(rs("type"))
	isOrder = cbool(rs("isOrder"))
	isApproved = cbool(rs("isApproved"))
	step = CInt(rs("step"))
	if not isApproved and isOrder and CInt(rs("step"))=40 then 
		Response.write "<div id='pauseOrderEditStamp' class='stamp'><h2>„ Êﬁ›</h2><div><span>„‰ Ÿ—  «ÌÌœ «” ⁄·«„ „Ãœœ</span></div></div>"
	elseif CInt(rs("step"))=40 then
		Response.write "<div id='pauseOrderEditStamp' class='stamp'><h2>„ Êﬁ›</h2><div><span>„‰ Ÿ— «’·«Õ ”›«—‘</span></div></div>" 
	elseif isApproved then
		Response.write "<div id='isApprovedStamp' class='stamp'><h2>’«œ— ‘œÂ</h2><div><span>’«œ— ‘œÂ  Ê”ÿ: </span><span id='approvedBy'>" & rs("approvedByName")  & "</span></div></div>"
	end if
	if (auth(2,2) and not CBool(rs("isClosed")) and CBool(rs("isOrder"))) then Response.write "<a class='btn btn-inverse' href='#' id='changeReturnDateBtn'> €ÌÌ—  «—ÌŒ  ÕÊÌ·</a>"
	if (auth(2,4) and not CBool(rs("isClosed")) and not CBool(rs("issued"))) then Response.write "<a class='btn btn-primary' href='#' id='changeCustomerBtn'> €ÌÌ— Õ”«»</a>"
	if (auth(2,2) and not CBool(rs("isClosed")) and Not CBool(rs("approved")) and not CBool(rs("issued"))) then Response.write "<a class='btn' href='#' id='orderEditBtn'>ÊÌ—«Ì‘</a>"
	if (auth(2,1) and not CBool(rs("isOrder")) and CBool(rs("isApproved"))) then Response.write "<a class='btn btn-success' href='#' id='convert2orderBtn'> »œÌ· »Â ”›«—‘</a>"
	if (auth(2,1) and CBool(rs("isOrder")) and not CBool(rs("isApproved")) and cint(rs("step")))=40 then Response.write "<a class='btn btn-success' href='#' id='convert2orderBtn'>‘—Ê⁄ „Ãœœ</a>"
	if (auth(2,"A") and not CBool(rs("isApproved")) and not CBool(rs("isClosed"))) and not CBool(rs("isOrder")) then Response.write "<a class='btn btn-primary' href='#' id='approvedBtn'>’œÊ— «” ⁄·«„</a>"
	if (auth(2,"D") and not CBool(rs("isClosed")) and CBool(rs("isApproved")))  then 
		Response.write "<a class='btn' href='#' id='pauseBtn'>"
		if CInt(rs("step"))=40 then 
			Response.write "—›⁄  Êﬁ›"
		else
			Response.write "œ—ŒÊ«”   Êﬁ›"
		end if
		response.write "</a>"
	end if
	if (auth(2,"C") and not rs("isClosed") and rs("isOrder")) then Response.write "<a class='btn btn-danger' href='#' id='cancelBtn'>ﬂ‰”·</a>"
	if (not rs("isClosed") and not rs("isOrder")) then Response.write "<a class='btn btn-danger' href='#' id='cancelBtn'>—œ «” ⁄·«„</a>"
	if (rs("isApproved")) then
		set rs1 = Conn.Execute("select * from InvoiceOrderRelations where [order]=" & orderID)
		while not rs1.eof
			response.write "<a class='btn btn-inverse' href='../AR/AccountReport.asp?act=showInvoice&invoice=" & rs1("invoice") & "' title='‘„«—Â " & rs1("invoice") & "'>’Ê— Õ”«»</a>"
			rs1.MoveNext
		wend
		rs1.close
		set rs1=Nothing
	end if
	Response.write "<br/>"
	if (auth(2,"B") and (isApproved or (not isApproved and step=40))) then Response.write "<a class='btn' href='#' id='printBtn'>ç«Å</a>" '--- print 
	if (Auth(3 , 1) and isOrder and step<>40) then Response.write "<a class='btn' href='../shopfloor/default.asp?orderID=" & orderID & "'> €ÌÌ— „—Õ·Â</a>"
	Response.write "<a class='btn' href='#' id='orderLogsBtn'>‰„«Ì‘  «—ÌŒçÂ</a>"
	if (orderType=2 and auth(2,"E") and CBool(isOrder) and step<>40) then 
		Response.write "<a class='btn' href='#' name='jobTicketPrint' group='plate'>œ” Ê— Å·Ì </a>"
		Response.write "<a class='btn' href='#' name='jobTicketPrint' group='print'>œ” Ê— ç«Å</a>"
		Response.write "<a class='btn' href='#' name='jobTicketPrint' group='binding'>œ” Ê— ’Õ«›Ì</a>"
	end if
%>
</div>
<div id="orderLogs"></div>
<div id="printArea">
<!-- 	<div id="voidedIcon"></div> -->
	<div id="orderHeader"></div>
	<div id="logStamp">»«Ìê«‰Ì ‘œÂ</div>
	<div id="orderDetails"></div>
</div>
<div id="orderMessages"></div>
<div id="orderStock"></div>
<div id="orderPurchase"></div>
<div id="orderCosts"></div>
<div id="orderFiles" class="well well-small"></div>
<%		
		'-----------------------------------------------------------------------------------------------------
	case "changeCustomer":		'-----------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID = cdbl(Request("id"))
		customerID = cdbl(Request("customer"))
		if not Auth(2 , 4) then NotAllowdToViewThisPage()
		conn.Execute("update orders set customer = " & customerID & ", isPrinted=0,LastUpdatedDate=getDate(),LastUpdatedBy=" &session("ID")& " where id = " & orderID)
		conn.Execute("update invoices set customer = " & customerID & " from invoices inner join invoiceOrderRelations on invoiceOrderRelations.invoice = invoices.id where invoiceOrderRelations.[order] = " & orderID)
		response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("„‘ —Ì «Ì‰ ”›«—‘  €ÌÌ— ÅÌœ« ﬂ—œ")
		'-----------------------------------------------------------------------------------------------------
	case "submitNew":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderTypeID = request("orderType")
		CustomerID=request.form("CustomerID") 
		mySQL="SET NOCOUNT ON;INSERT INTO orders (CreatedBy, Customer, isClosed, returnDate, orderTitle, Type, status, step, LastUpdatedBy, Notes, property,productionDuration, Qtty, price, PaperSize,lastUpdatedDate, deposit, dueDate, chequeDueDate) VALUES ("& session("ID") & ", "& CustomerID & ", 0, N'"& sqlSafe(request.form("ReturnDateTime")) & "', N'"& sqlSafe(request.form("OrderTitle")) & "', "& orderTypeID & ", 1, 1, "& session("ID") & ", N'"& sqlSafe(request.form("Notes")) & "',N'" & request.form("myXML") & "', " & sqlSafe(request.form("productionDuration")) & "," & sqlSafe(request.form("qtty")) & ",'" & sqlSafe(request.form("totalPrice")) & "', N'" & sqlSafe(request.form("PaperSize")) & "',getdate()," & CDbl(Request.Form("deposit")) & "," & CInt(Request.Form("dueDate")) & ", " & Request.Form("chequeDueDate") & "); SELECT SCOPE_IDENTITY() AS NewQuote;SET NOCOUNT OFF"
		response.write mySQL
				
		set RS1 = Conn.execute(mySQL)
		QuoteID = RS1 ("NewQuote")
		RS1.close
		conn.close
		response.redirect "?act=show&id=" & QuoteID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  «” ⁄·«„ À»  ‘œ")
		'-----------------------------------------------------------------------------------------------------
	case "submitEdit":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		set rs = Conn.Execute("select orders.customer, orders.ver, orders.isClosed, orders.isApproved, orders.isOrder, isnull(invoices.approved,0) as approved, isnull(invoices.issued,0) as issued,accounts.arBalance, accounts.creditLimit, accounts.status as accountStatus, accounts.maxCreditDay, isnull(ar.firstDebit,0) as firstDebit from orders inner join accounts on orders.customer=accounts.id left outer join (select Account,datediff(day, dbo.udf_date_solarToDate(cast(substring(min(effectiveDate),1,4) as int),cast(substring(min(effectiveDate),6,2) as int),cast(substring(min(effectiveDate),9,2) as int)),getDate()) as firstDebit from ARItems where FullyApplied=0 and IsCredit=0 and voided=0 group by account) as ar on orders.customer=ar.account left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice = invoices.id where orders.id = " & orderID)
		msg=""
		errMsg=""
		if rs.eof then 
			errMsg = "ç‰Ì‰ ‘„«—Â ”›«—‘/«” ⁄·«„ «Ì ÅÌœ« ‰‘œ!!!!"
		elseif rs("isClosed") then 
			errMsg = "”›«—‘ »” Â ‘œÂ! ›·–« «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ"
		elseif rs("issued") then 
			errMsg = "Œÿ«Ì ⁄ÃÌ»! ”›«—‘ »” Â ‰‘œÂ «„« ’«œ— ‘œÂ!!!"
		elseif rs("approved") then 
			errMsg = "’Ê— Õ”«»  «ÌÌœ ‘œÂ"
		elseif CBool(rs("isApproved")) and not CBool(rs("isOrder")) then 
			msg = "«” ⁄·«„ ’«œ— ‘œÂ —« ÊÌ—«Ì‘ ‰„ÊœÌœ°<br> ›·–« «“ ’œÊ— Œ«—Ã ‘œ<br>"
		elseif CBool(rs("isOrder")) then 
			msg = "”›«—‘ —« ÊÌ—«Ì‘ ‰„ÊœÌœ<br>›·–« „ Êﬁ› ŒÊ«Âœ ‘œ  « ‘„«  «ÌÌœ „‘ —Ì —« œ—Ì«›  ﬂ‰Ìœ<br>"
		end if
		if (cdbl(rs("arBalance"))+cdbl(rs("creditLimit")) < 0) then 
			CustomerID = cdbl(rs("customer"))
			msg = msg & "<br>»œÂÌ «Ì‰ Õ”«» «“ „Ì“«‰ «⁄ »«— ¬‰ »Ì‘ — ‘œÂ°<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>"
		elseif (CDbl(rs("firstDebit")) > CDbl(rs("maxCreditDay"))) then
			CustomerID = cdbl(rs("customer"))
			firstDebit = rs("firstDebit")
			msg = msg & "<br>»œÂÌ «Ì‰ „‘ —Ì „—»Êÿ »Â " & firstDebit & " —Ê“ ê–‘ Â »ÊœÂ°<br> ﬂÂ «“ ”——”Ìœ Å—œ«Œ  ¬‰ ê–‘ Â<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br>(„„ﬂ‰ «”  ﬂÂ «‘ﬂ«·Ì œ— œÊŒ ‰ »«‘œ)<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>"
		end if
		if errMsg="" then
			ver = CInt(rs("ver"))
			if Request.form("returnDateNull")="on" then 
				returnDateTime="null"
			else
				returnDateTime = "N'" & request.form("returnDateTime") & "'"
			end if
			isOrder = ""
			if CBool(rs("isOrder")) then 
				isOrder=", step=40,customerApprovedType=null, customerApprovedDate=null"
				ver = ver + 1
			end if
			if request.form("myXML")<>"" then 
				Conn.Execute ("update orders set ver = " & ver & ", property=N'" & request.form("myXML") & "', productionDuration=" &cint(request.form("productionDuration"))& ", Notes=N'" &sqlSafe(request.form("Notes"))& "', returnDate=" & returnDateTime & ", paperSize=N'" &sqlSafe(request.form("paperSize"))& "', OrderTitle=N'" &sqlSafe(request.form("OrderTitle"))& "', qtty=" &cdbl(request.form("qtty"))& ", price='" & sqlSafe(request.form("totalPrice")) & "',LastUpdatedDate=getDate(), LastUpdatedBy=" &session("ID")& ", Customer=" &sqlSafe(request.form("customerID"))& ",isApproved=0, isPrinted=0 " & isOrder & ", approvedBy=null, deposit=" & CDbl(Replace(Request.form("deposit"),",","")) & ", dueDate=" & CInt(Request.Form("dueDate")) & ", chequeDueDate=" & CInt(Request.Form("chequeDueDate")) & " where id="&orderID)
				conn.Execute("delete InvoiceLines where Invoice in (select invoice from InvoiceOrderRelations where [Order]=" & orderID & ")")
				response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode(msg & "«ÿ·«⁄«  »Â —Ê“ ‘œ")
			else	
				response.redirect "?act=show&id=" & orderID & "&errMsg=" & Server.URLEncode("Œÿ«ÌÌ —Œ œ«œÂ!")
			end if
		else
			response.redirect "?act=show&id=" & orderID & "&errMsg=" & Server.URLEncode(errMsg)
		end if
		'-----------------------------------------------------------------------------------------------------
	case "convertToOrder":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		isOrder = false
		if Request("isOrder")="1" then isOrder = true
		orderID=cdbl(request("id"))
		returnDate = Request("returnDate")
		customerApprovedType = Request("customerApprovedType")
		customerApprovedDate = Request("customerApprovedDate")
		if not Auth(2 , 1) then NotAllowdToViewThisPage()
		set rs = Conn.Execute("select orders.*, accounts.arBalance, accounts.creditLimit, accounts.status as accountStatus, accounts.maxCreditDay, isnull(ar.firstDebit,0) as firstDebit from orders inner join accounts on orders.customer=accounts.id left outer join (select Account,datediff(day, dbo.udf_date_solarToDate(cast(substring(min(effectiveDate),1,4) as int),cast(substring(min(effectiveDate),6,2) as int),cast(substring(min(effectiveDate),9,2) as int)),getDate()) as firstDebit from ARItems where FullyApplied=0 and IsCredit=0 and voided=0 group by account) as ar on orders.customer=ar.account where orders.id=" & orderID)
		if not isOrder and rs("isOrder") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("ﬁ»·« »Â ”›«—‘  »œÌ· ‘œÂ!")
		elseif not isOrder and not rs("isApproved") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("«» œ« «Ì‰ «” ⁄·«„ —«  «ÌÌœ ﬂ‰Ìœ!")
		elseif not isOrder and (rs("step")="40" or rs("status")<>"1") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("«Ì‰ «” ⁄·«„ „ Êﬁ› „Ìù»«‘œ. Å” «“  »œÌ· ¬‰ „⁄–Ê—Ì„<br>·ÿ›« ‰”»  »Â —›⁄  ÊﬁÌ› «ﬁœ«„ ›—„«ÌÌœ.")
		elseif rs("isClosed") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("ŒÌ·Ì ŒÌ·Ì ⁄ÃÌ»Â! «Ì‰ «” ⁄·«„ ﬂ‰”· ‘œÂ!!")
		elseif (cdbl(rs("arBalance"))+cdbl(rs("creditLimit")) < 0) then 
			CustomerID = cdbl(rs("customer"))
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("»œÂÌ «Ì‰ Õ”«» «“ „Ì“«‰ «⁄ »«— ¬‰ »Ì‘ — ‘œÂ°<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>")
		elseif (CDbl(rs("firstDebit")) > CDbl(rs("maxCreditDay"))) then
			CustomerID = cdbl(rs("customer"))
			firstDebit = rs("firstDebit")
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("»œÂÌ «Ì‰ „‘ —Ì „—»Êÿ »Â " & firstDebit & " —Ê“ ê–‘ Â »ÊœÂ°<br> ﬂÂ «“ ”——”Ìœ Å—œ«Œ  ¬‰ ê–‘ Â<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br>(„„ﬂ‰ «”  ﬂÂ «‘ﬂ«·Ì œ— œÊŒ ‰ »«‘œ)<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>")
		elseif CInt(rs("accountStatus"))<>1 then 
			CustomerID = cdbl(rs("customer"))
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("«Ì‰ Õ”«» ›⁄«· ‰Ì” !<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>")
		else
			set rs=Conn.Execute("select Orders.Customer, orders.isOrder, orders.isApproved, orders.property as data, orderTypes.property as keys from Orders inner join orderTypes on orders.Type=orderTypes.ID where orders.id=" & orderID & " and lastUpdatedDate>lastUpdate")
			if rs.eof then 
				response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b>ﬁÌ„ ùÂ« »—Ê“ ‘œÂ! </b><br>ﬁ»· «“  «ÌÌœ° ·ÿ›« «” ⁄·«„ —« »—Ê“ ﬂ‰Ìœ")
			else
				isOrder = CBool(rs("isOrder"))
				set data=server.createobject("MSXML2.DomDocument")
				set keys=server.createobject("MSXML2.DomDocument")
				data.loadXML(rs("data"))
				keys.loadXML(rs("keys"))
				CustomerID = rs("Customer")
				InvoiceDate = shamsiToday()
				requestID=""
				purchaseID=""
				set rsr = Conn.Execute("select * from InventoryItemRequests where rowID is not null and orderID=" & orderID)
				while not rsr.eof
					requestID = requestID & rsr("id") & ","
					rsr.MoveNext
				wend
				set rsp = Conn.Execute("select * from PurchaseRequests where rowID is not null and orderID=" & orderID)
				while not rsp.eof
					purchaseID = purchaseID & rsp("id") & ","
					rsp.MoveNext
				wend
				errMSG=""
				'**************************** Check Insert or Update Invoice ****************
				mySQL = "select * from InvoiceOrderRelations where [Order]=" & orderID
				set rs = Conn.Execute(mySQL)
				if rs.eof then 
					'**************************** Inserting new Invoice ****************
					newInvoice = true
					mySQL="INSERT INTO Invoices (CreatedDate, CreatedBy, Customer, isA) VALUES (N'"& InvoiceDate & "', "& session("ID") & ", "& CustomerID & ", 0);SELECT @@Identity AS NewInvoice"
					set RS = Conn.execute(mySQL).NextRecordSet
					InvoiceID = RS ("NewInvoice")
				else
					'**************************** Updating new Invoice ****************
					newInvoice = false
					InvoiceID = RS ("Invoice")
					Conn.Execute("delete invoiceLines where invoice = " & invoiceID)
					mySQL = "update invoices set CreatedDate=N'" & InvoiceDate & "', CreatedBy=" & Session("ID") & ", Customer=" & CustomerID & " where id = " & invoiceID
					Conn.Execute(mySQL)
				end if
				mySQL = "select * from accounts where id=" & customerID
				set rs = Conn.Execute(mySQL)
				isADefault = 0
				if rs("isADefault")="True" then isADefault = 1
				vatRate = CLng(session("VatRate"))
				TotalPrice = 0 
				TotalDiscount = 0 
				TotalReverse = 0
				TotalReceivable = 0
				totalVat = 0
				for each row in data.selectNodes("//row")
					rowName = row.GetAttribute("name")
					rowDesc = ""
					rowID = row.GetAttribute("id")
					if row.selectNodes("./key[@name='description']").length>0 then 
						set tmp = row.selectNodes("./key[@name='description']")
						rowDesc = tmp(0).text
					end if
					set sRow = keys.selectNodes("//rows[@name='" & rowName & "']")(0)
					for each sGroup in sRow.selectNodes("./row/group[@isPrice!='no']")
						l=0
						w=0
						groupName=sGroup.GetAttribute("name")
						item = sGroup.GetAttribute("item")
						desc = sGroup.GetAttribute("label") 
						
						hasStock = sGroup.GetAttribute("hasStock")
						hasPurchase = sGroup.GetAttribute("hasPurchase")
						hasKey = 0
						stockName = desc
						stockDesc = ""
						stockOldItemID = -1
						for each sKey in sGroup.selectNodes("./key")
							keyName=sKey.GetAttribute("name")
							if row.selectNodes("./key[@name='" & keyName & "']").length>0 then hasKey = hasKey + 1
						next
						if hasKey>0 then 
							select case groupName 
								case "plate"
									desc = row.selectSingleNode("./key[@name='plate-stockName']").text 
								case "paper"
									desc = row.selectSingleNode("./key[@name='paper-stockName']").text 
								case "verni"
									set tmp = sGroup.selectSingleNode("//key[@name='verni-mat']/option[.=" & row.selectSingleNode("./key[@name='verni-mat']").text & "]")
									desc = desc & " " & tmp.GetAttribute("label") 
									set tmp = sGroup.selectSingleNode("//key[@name='verni-wat']/option[.=" & row.selectSingleNode("./key[@name='verni-wat']").text & "]")
									desc = desc & " " & tmp.GetAttribute("label")
								case "print"
									desc = desc & " " & row.selectSingleNode("./key[@name='print-color']").text & " —‰ê "
									if row.selectSingleNode("./key[@name='print-sp-color']").text <> "0" then 
										desc = desc & " " & row.selectSingleNode("./key[@name='print-sp-color']").text & " —‰ê Œ«’ "
									end if
									set tmp = sGroup.selectSingleNode("//key[@name='print-d-type']/option[.=" & row.selectSingleNode("./key[@name='print-d-type']").text & "]")
									desc = desc & " " & tmp.GetAttribute("label")
									set tmp = sGroup.selectSingleNode("//key[@name='print-control']/option[.=" & row.selectSingleNode("./key[@name='print-control']").text & "]")
									desc = desc & " " & tmp.GetAttribute("label")
								case "selefon"
									desc = row.selectSingleNode("./key[@name='selefon-purchaseDesc']").text
								case "uv"
									desc = row.selectSingleNode("./key[@name='uv-purchaseDesc']").text
								case "binding"
									set tmp = sGroup.selectSingleNode("//key[@name='binding-type']/option[.=" & row.selectSingleNode("./key[@name='binding-type']").text & "]")
									desc = desc & " " & tmp.GetAttribute("label") 
									set tmp = sGroup.selectSingleNode("//key[@name='binding-size']/option[.=" & row.selectSingleNode("./key[@name='binding-size']").text & "]")
									desc = desc & " " & tmp.GetAttribute("label")
									if row.selectSingleNode("./key[@name='binding-orient']").text <> "1" then 
										set tmp = sGroup.selectSingleNode("//key[@name='binding-orient']/option[.=" & row.selectSingleNode("./key[@name='binding-orient']").text & "]")
										desc = desc & " " & tmp.GetAttribute("label")
									end if
								case "service"
									desc = desc & " " & row.selectSingleNode("./key[@name='service-purchaseName']").text
								case "serviceIN"
									desc = desc & " " & row.selectSingleNode("./key[@name='serviceIN-item']").text
								case "packing"
									set tmp = sGroup.selectSingleNode("//key[@name='packing-type']/option[.=" & row.selectSingleNode("./key[@name='packing-type']").text & "]")
									desc = desc & " " & tmp.GetAttribute("label")
								case "invRequest"
									desc = desc & " " & row.selectSingleNode("./key[@name='invRequest-name']").text
							end select
							if rowDesc<>"" then 
								if Len(rowDesc)>50 then
									rowDesc = left(rowDesc,50) & "..."
								end if
								desc = desc & "(" & rowDesc & ")"
							end if
							'-----------------------STOCK
							if hasStock="yes" or hasStock="option" then
								customerHasStock = "0"
								set tmp = row.selectNodes("./key[@name='" & groupName & "-stockName']")
								if tmp.length>0 then 
									stockName =  tmp(0).text
									if stockName = "" then 
										hasStock = "no" 
									else
										hasStock = "yes" 
									end if
									set tmp = row.selectNodes("./key[@name='" & groupName & "-stockDesc']")
									if tmp.length>0 then
										stockDesc = tmp(0).text
									else
										stockDesc = ""
									end if
									set tmp = row.selectNodes("./key[@name='" & groupName & "-customer']")
									if tmp.length>0 then
										if tmp(0).text="on" then customerHasStock = "1"
									end if
									set tmp = row.selectNodes("./key[@name='" & groupName & "-itemCode']")
									if tmp.length>0 then
										if IsNumeric(tmp(0).text) then stockOldItemID = CDbl(tmp(0).text)
									end if
								end if
							end if 
							'-----------------------PURCHACE
							if hasPurchase="yes" then 
								set tmp = row.selectNodes("./key[@name='" & groupName & "-purchaseTypeID']")
								if tmp.length>0 then 
									purchaseTypeID = tmp(0).text
									set tmp = row.selectNodes("./key[@name='" & groupName & "-purchaseName']")
									purchaseName =  tmp(0).text
									set tmp = row.selectNodes("./key[@name='" & groupName & "-purchaseDesc']")
									purchaseDesc =  tmp(0).text
								end if
							end if
							'-----------------------INVOICE
							set tmp = row.selectSingleNode("./key[@name='" & groupName & "-price']")
	' 						Response.write "<br>" & groupName & "<br>"
							thePrice = cdbl(Replace(tmp.text,",",""))
							set tmp = row.selectSingleNode("./key[@name='" & groupName & "-dis']")
							dis = tmp.text
							set tmp = row.selectSingleNode("./key[@name='" & groupName & "-over']")
							over = tmp.text
							set tmp = row.selectNodes("./key[@name='" & groupName & "-reverse']")
							if tmp.length>0 then 
								reverse = tmp(0).text
							else
								reverse = 0
							end if
							if row.selectNodes("./key[@name='" & groupName & "-qtty']").length>0 then 
								set tmp = row.selectSingleNode("./key[@name='" & groupName & "-qtty']")
								qtty = CDbl(Replace(tmp.text,",",""))
							elseif row.selectNodes("./key[@name='" & groupName & "-count']").length>0 then 
								set tmp = row.selectSingleNode("./key[@name='" & groupName & "-count']")
								qtty = CDbl(Replace(tmp.text,",",""))
							else
								qtty = 1
							end if
							if row.selectNodes("./key[@name='" & groupName & "-form']").length>0 then 
								set tmp = row.selectSingleNode("./key[@name='" & groupName & "-form']")
								sets = CDbl(Replace(tmp.text,",",""))
							elseif groupName="plate" then 
								set tmp = row.selectSingleNode("./key[@name='plate-color-count']")
								sets = CDbl(Replace(tmp.text,",",""))
							else
								sets = 1
							end if
							if row.selectNodes("./key[@name='" & groupName & "-l']").length>0 then 
								set tmp = row.selectSingleNode("./key[@name='" & groupName & "-l']")
								l = CDbl(tmp.text)
								set tmp = row.selectSingleNode("./key[@name='" & groupName & "-w']")
								w = CDbl(tmp.text)
							else
								if row.selectNodes("./key[@name='" & groupName & "-size']").length>0 then 
									set tmp = row.selectSingleNode("./key[@name='" & groupName & "-size']")
									if InStr(tmp.text,"X")>0 then
										l = CDbl(split(replace(tmp.text,"other:",""),"X")(0))
										w = CDbl(split(replace(tmp.text,"other:",""),"X")(1))
									elseif InStr(tmp.GetAttribute("label"),"3")>0 then 
										l=42
										w=30
									elseif InStr(tmp.GetAttribute("label"),"4")>0 then
										l=21
										w=30
									else
										sizeID=tmp.text
										if InStr(sizeID,"X")>0 then 
											l = CDbl(split(tmp.GetAttribute("label"),"X")(0))
											w = CDbl(split(tmp.GetAttribute("label"),"X")(1))
										else
											if sGroup.selectNodes("./key[@name='" & groupName & "-size']/option[" & sizeID & "]").length>0 then 
												set tmp = sGroup.selectSingleNode("./key[@name='" & groupName & "-size']/option[" & sizeID & "]")
												if InStr(tmp.GetAttribute("label"),"X")>0 then 
													l = CDbl(split(tmp.GetAttribute("label"),"X")(0))
													w = CDbl(split(tmp.GetAttribute("label"),"X")(1))
												end if
											end if
										end if
									end if
								else
									l=0
									w=0
								end if
							end if
	 						appQtty = qtty * sets
	 						
	 						if InStr(dis,"%")<>0 then
		 						dis = CDbl(Replace(dis,"%",""))
								if dis<>100 then 
									if InStr(reverse,"%")<>0 then
										price = thePrice * (100/(100-dis-CDbl(Replace(reverse,"%",""))))
									else
										price = thePrice * (100/(100-dis)) + CDbl(reverse)
									end if
								else
									if thePrice = 0 then 
										price = 0
									else
										if InStr(over,"%")<>0 then
											over = CDbl(Replace(over,"%",""))
											price = thePrice * (100/over)
										else
											over = CDbl(over)
											price = thePrice + over
										end if
									end if
								end if
								discount = price * dis / 100
								if InStr(reverse,"%")<>0 then
									theReverse = price * CDbl(Replace(reverse,"%","")) / 100
								else
									theReverse = CDbl(reverse)
								end if
							else
								discount = CDbl(dis)
								if InStr(reverse,"%")<>0 then
									price = (thePrice + discount) * (100/(100-CDbl(Replace(reverse,"%","")))) 
									theReverse = price * CDbl(Replace(reverse,"%","")) / 100
								else
									theReverse = CDbl(reverse)
									price = thePrice + discount + theReverse
								end if
							end if
							if isADefault then 
								if groupName<>"paper" then
									vat = thePrice * vatRate / 100
								else
									vat = 0
								end if
							else
								vat = 0
							end if
							hasVat = 1
							if groupName="paper" then hasVat=0
							
							'-------------------- INVOICE Insert
							mySQL = "INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, [Reverse], Vat, hasVat) VALUES (" & InvoiceID & ", " & item & ", N'" & left(desc,100) & "', " & l & ", " & w & ", " & qtty & ", " & sets & ", " & appQtty & ", " & price & ", " & discount & ", " & theReverse & ", " & vat & ", " & hasVat & ")"
							'Response.write mySQL
							Conn.Execute(mySQL)
							TotalPrice = TotalPrice + price
							TotalDiscount = TotalDiscount + discount
							TotalReverse = TotalReverse + theReverse
							TotalReceivable = TotalReceivable + thePrice + vat
							totalVat = totalVat + vat
							'-------------------- STOCK Insert
							if hasStock = "yes" then 
								if requestID="" then
									if stockOldItemID>0 then 
										set rsInv = conn.Execute("select * from inventoryItems where oldItemID=" & stockOldItemID)
										mySQL = "INSERT INTO InventoryItemRequests (orderID, ItemName, comment, qtty, createdBy, invoiceItem, rowID, customerHaveInvItem, itemID) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & "," & rsInv("id") & ")"
										rsInv.Close
									else
										mySQL = "INSERT INTO InventoryItemRequests (orderID, ItemName, comment, qtty, createdBy, invoiceItem, rowID, customerHaveInvItem) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & ")"
									end if
									Conn.Execute(mySQL)
								else
									'----------- Check UPDATE or INSERT ?
									mySQL = "select * from InventoryItemRequests where orderID=" & orderID & " and invoiceItem = " & item & " and rowID = " & rowID & " order by ID desc"
	' 								Response.write mySQL
									set rsr = Conn.Execute(mySQL)
									if rsr.eof then 
										if stockOldItemID>0 then 
											set rsInv = conn.Execute("select * from inventoryItems where oldItemID=" & stockOldItemID)
											mySQL = "INSERT INTO InventoryItemRequests (orderID, ItemName, comment, qtty, createdBy, invoiceItem, rowID, customerHaveInvItem, itemID) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & "," & rsInv("id") & ")"
											rsInv.Close
										else
											mySQL = "INSERT INTO InventoryItemRequests (orderID, ItemName, comment, qtty, createdBy, invoiceItem, rowID, customerHaveInvItem) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & ")"
										end if
										Conn.Execute(mySQL)
									else
										requestID = Replace(requestID,rsr("id") & ",","")
										if rsr("status")="new" and (cdbl(rsr("qtty")) <> cdbl(appQtty) or Trim(rsr("itemName")) <> Trim(stockName) or Trim(rsr("comment")) <> Trim(stockDesc) or CBool(rsr("customerHaveInvItem"))<> CBool(customerHasStock)) then 
											'------- IF CHANGE UPDATE
											if stockOldItemID>0 then 
												set rsInv = conn.Execute("select * from inventoryItems where oldItemID=" & stockOldItemID)
												mySQL = "UPDATE InventoryItemRequests SET ItemName=N'" & stockName & "', comment=N'" & stockDesc & "', qtty=" & appQtty & ", createdBy=" & Session("ID") & ", customerHaveInvItem=" & customerHasStock & ", itemID= " & rsInv("id") & " where id=" & rsr("id")
												rsInv.Close
											else
												mySQL = "UPDATE InventoryItemRequests SET ItemName=N'" & stockName & "', comment=N'" & stockDesc & "', qtty=" & appQtty & ", createdBy=" & Session("ID") & ", customerHaveInvItem=" & customerHasStock & " where id=" & rsr("id")
											end if
											Conn.Execute(mySQL)
										elseif rsr("status")="del" then 
											'------- IF delete Insert new request
											if stockOldItemID>0 then 
												set rsInv = conn.Execute("select * from inventoryItems where oldItemID=" & stockOldItemID)
												mySQL = "INSERT INTO InventoryItemRequests (orderID, ItemName, comment, qtty, createdBy, invoiceItem, rowID, customerHaveInvItem, itemID) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & ", " & rsInv("id") & ")"
												rsInv.Close
											else
												mySQL = "INSERT INTO InventoryItemRequests (orderID, ItemName, comment, qtty, createdBy, invoiceItem, rowID, customerHaveInvItem) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & ")"
											end if
											Conn.Execute(mySQL)
										elseif CDbl(rsr("qtty")) <> CDbl(appQtty) or Trim(rsr("itemName"))<> Trim(stockName) or Trim(rsr("comment")) <> Trim(stockDesc) or CBool(rsr("customerHaveInvItem"))<> CBool(customerHasStock) then
											errMSG = errMSG & "<span title='" & CDbl(rsr("qtty")) &":"& CDbl(appQtty) &","& Trim(rsr("itemName"))&":"& Trim(stockName) &","& Trim(rsr("comment")) &":"& Trim(stockDesc) &","& CBool(rsr("customerHaveInvItem"))&":"& CBool(customerHasStock) &"'>»—«Ì "  & stockName & " ÕÊ«·Â ’«œ— ‘œÂ. Å”  €ÌÌ— œ— ¬‰ „„ﬂ‰ ‰Ì” </span><br>"
										end if
									end if
								end if
							end if
						
							'------------------ PURCHASE Insert
							if hasPurchase = "yes" then 
								if purchaseID="" then 
									mySQL="INSERT INTO PurchaseRequests (orderID,typeName,comment,typeID,qtty,createdBy,price,isService,rowID) VALUES (" & orderID & ",N'" & purchaseName & "',N'" & purchaseDesc & "'," & purchaseTypeID & "," & appQtty & "," & Session("ID") &"," & thePrice & ",1," & rowID & ")"
		' 							Response.write mySQL
									Conn.Execute(mySQL)
								else
									'----------- Check UPDATE or INSERT ?
									mySQL = "select * from PurchaseRequests where orderID=" & orderID & " and typeID = " & purchaseTypeID & " and rowID = " & rowID & " order by ID desc"
										Response.write mySQL & "<br>"
									set rsp = Conn.Execute(mySQL)
									if rsp.eof then 
										mySQL="INSERT INTO PurchaseRequests (orderID,typeName,comment,typeID,qtty,createdBy,price,isService,rowID, invoiceItem) VALUES (" & orderID & ",N'" & purchaseName & "',N'" & purchaseDesc & "'," & purchaseTypeID & "," & appQtty & "," & Session("ID") &"," & thePrice & ",1," & rowID & "," & item & ")"
										Conn.Execute(mySQL)
									else
										purchaseID = Replace(purchaseID,rsp("id") & ",","")
										if rsp("status")="new" and (cdbl(rsp("qtty")) <> cdbl(appQtty) or trim(rsp("typeName")) <> trim(purchaseName) or trim(rsr("comment")) <> trim(purchaseDesc) or cdbl(rsp("typeID")) <> cdbl(purchaseTypeID) or cdbl(rsp("price")) <> cdbl(thePrice)) then
											mySQL = "UPDATE PurchaseRequests SET typeName = N'" & purchaseName & "',comment = N'" & purchaseDesc & "', typeID = " & purchaseTypeID & ", qtty = " & appQtty & ", price=" & thePrice & " WHERE id=" & rsp("id")
											Conn.Execute(mySQL)
										elseif rsp("status")="del" then 
											mySQL="INSERT INTO PurchaseRequests (orderID,typeName,comment,typeID,qtty,createdBy,price,isService,rowID) VALUES (" & orderID & ",N'" & purchaseName & "',N'" & purchaseDesc & "'," & purchaseTypeID & "," & appQtty & "," & Session("ID") &"," & thePrice & ",1," & rowID & ")"
											Conn.Execute(mySQL)
										elseif cdbl(rsp("qtty")) <> cdbl(appQtty) or cdbl(rsp("typeID")) <> cdbl(purchaseTypeID) then
											errMSG = errMSG & "»—«Ì "  & purchaseName & " ”›«—‘ Œ—Ìœ «ÌÃ«œ ‘œÂ. Å”  €ÌÌ— œ— ¬‰ „„ﬂ‰ ‰Ì” <br>"
										end if
									end if
								end if
							end if
						end if
					next
				next
				RFD = TotalReceivable - fix(TotalReceivable / 5000) * 5000
				TotalReceivable = TotalReceivable - RFD
				TotalDiscount = TotalDiscount + RFD
				if RFD > 0 then 
					mySQL = "INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, [Reverse], Vat, hasVat) VALUES (" & InvoiceID & ", 39999, N' Œ›Ì› —‰œ ›«ﬂ Ê—', 0, 0, 0, 0, 0, 0, " & RFD & ", 0, 0, 0)"
					'Response.write mySQL
					Conn.Execute(mySQL)
				end if
				mySQL = "update invoices set totalPrice=" & totalPrice & ", totalDiscount=" & totalDiscount & ", totalReverse=" & totalReverse & ", TotalReceivable=" & TotalReceivable & ", totalVat=" & totalVat & ",isA=" & isADefault & "  where id=" & invoiceID
	' 				Response.write mySQL
				Conn.Execute(mySQL)
				if newInvoice then 
					mySQL = "INSERT INTO InvoiceOrderRelations (Invoice, [Order]) VALUES (" & InvoiceID & ", " & orderID & ")"
					Conn.Execute(mySQL)
				end if
				if len(purchaseID)>0 then 
					Response.write (purchaseID)
					for each id in Split(purchaseID,",")
						if id<>"" then Conn.Execute("update PurchaseRequests set status='del' where id=" & id)
					Next
				end if
				if len(requestID)>0 then 
					Response.write (requestID)
					for each id in Split(requestID,",")
						if id<>"" then
							set rs = Conn.Execute("select * from InventoryItemRequests where id=" & id)
							if rs("status")="pick" then 
								errMSG = errMSG & "»—«Ì <a href='../inventory/default.asp?show=" & rs("id") & "'>"  & rs("itemName") & "</a> ÕÊ«·Â ’«œ— ‘œÂ. Ê ‘„« ¬‰—« Õ–› ﬂ—œÂù«Ìœ!<br>"
							else
								Conn.Execute("update InventoryItemRequests set status='del' where id=" & id)
							end if
						end if
					Next
				end if
				if errMSG="" then 
					if isOrder then 
						mySQL = "update orders set isApproved=1,approvedBy=" & Session("id") & ", isPrinted=0, step=1, customerApprovedDate='" & customerApprovedDate & "', customerApprovedType=" & customerApprovedType 
					else
						mySQL = "update orders set orderDate=getDate(), isOrder=1, isPrinted=0, customerApprovedDate='" & customerApprovedDate & "', customerApprovedType=" & customerApprovedType 
					end if
					
					if returnDate<>"" then 
						mySQL = mySQL & " ,returnDate='" & returnDate & "'"
' 					else
' 						mySQL = mySQL & " ,returnDate=null"
					end if
					mySQL = mySQL & ",LastUpdatedDate=getDate(),LastUpdatedBy=" &session("ID")& " where id=" & orderID
					Conn.Execute(mySQL)
					if isOrder then
						response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b> «ÌÌœ „Ãœœ „‘ —Ì À»  ‘œ Ê ”›«—‘ „Ãœœ »Â „—Õ·Â ‘—Ê⁄ —› </b>")
					else
						response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b>»Â ”›«—‘  »œÌ· ‘œ!</b>")
					end if
				else
					response.redirect "?act=show&id=" & orderID & "&errmsg=" & Server.URLEncode(errMSG)
				end if
				rs.Close
				set rs=Nothing
				rsr.Close
				set rsr=Nothing
				rsp.Close
				set rsp = Nothing
			end if
			Conn.Close
		end if
		'-----------------------------------------------------------------------------------------------------
	case "approve":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		set rs=Conn.Execute("select * from Orders where id=" & orderID )
		if (rs("isApproved")) then 
			response.redirect "?act=show&id=" & orderID & "&errmsg=" & Server.URLEncode("ﬁ»·«  «ÌÌœ ‘œÂ!")
		elseif rs("step")="40" or rs("status")<>"1" then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("«Ì‰ «” ⁄·«„ „ Êﬁ› „Ìù»«‘œ. Å” «“  »œÌ· ¬‰ „⁄–Ê—Ì„<br>·ÿ›« ‰”»  »Â —›⁄  ÊﬁÌ› «ﬁœ«„ ›—„«ÌÌœ.")
		elseif rs("isClosed") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("ŒÌ·Ì ŒÌ·Ì ⁄ÃÌ»Â! «Ì‰ «” ⁄·«„ ﬂ‰”· ‘œÂ!!")
		else
			if (not rs("isOrder")) then 
				ver = CInt(rs("ver")) + 1
				mySQL = "update orders set ver = " & ver & ", isApproved=1, isPrinted=0, approvedDate=getDate(), LastUpdatedDate=getDate(), approvedBy=" & Session("id") & ", LastUpdatedBy=" &session("ID")& " where id="& orderID
				Conn.execute(mySQL)
				response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("«” ⁄·«„ ’«œ— ‘œ")
			end if
		end if
		'-----------------------------------------------------------------------------------------------------
	case "pause":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		if not Auth(2 , "D") then NotAllowdToViewThisPage()
		orderID=cdbl(request("id"))
		
		Conn.Execute("update orders set LastUpdatedDate=getDate(), isPrinted=0,LastUpdatedBy=" &session("ID")& ", step = (case step when 40 then 1 else 40 end) where id=" & orderID)
		msgFrom=session("id")
		msgTo= Request("createdBy")
		msgTitle		= " Êﬁ› ”›«—‘"
		msgBody			= Request("reason")
		RelatedTable	= "orders"
		relatedID		= orderID
		replyTo			= "null"
		IsReply			= 0
		urgent			= 1
		MsgDate			= shamsiToday()
		MsgTime			= currentTime10()
		msgType			= 0
		MySQL = "INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent, type) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", '"& relatedID & "', '"& RelatedTable & "', "& urgent & ", " & msgType & ")"
		conn.Execute MySQL
		response.redirect "?act=show&id=" & orderID
		'-----------------------------------------------------------------------------------------------------
	case "cancel":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		set rs = Conn.Execute("select orders.createdBy, orders.isClosed, orders.isApproved, orders.isOrder, isnull(invoices.approved,0) as approved, isnull(invoices.issued,0) as issued from orders left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice = invoices.id where orders.id = " & orderID)
		if CBool(rs("isOrder")) then 
			if not Auth(2 , "C") then NotAllowdToViewThisPage()
		else
			if not Auth(2 , 2) then NotAllowdToViewThisPage()
		end if
		errMsg = ""
		msg = ""
		msgTo = rs("createdBy")
		if Request("reason")="" then
			errMsg = "Õ „« »«Ìœ œ·Ì· œ«‘ Â »«‘Â"
		elseif rs.eof then 
			errMsg = "ç‰Ì‰ ”›«—‘Ì ÅÌœ« ‰‘œ!"
		elseif CBool(rs("issued")) then 
			errMsg = "«Ì‰ ”›«—‘ ›«ﬂ Ê— ’«œ— ‘œÂ œ«—œ! ›·–« ﬂ‰”·Ì ¬‰ „„ﬂ‰ ‰Ì” "
		elseif CBool(rs("approved")) then 
			errMsg = "«Ì‰ ”›«—‘ ›«ﬂ Ê—  «ÌÌœ ‘œÂ œ«—œ! ›·–« ﬂ‰”·Ì ¬‰ „„ﬂ‰ ‰Ì” ° „ê— ›«ﬂ Ê— „—»ÊÿÂ «“  «ÌÌœ Œ«—Ã ‘Êœ"
		elseif CBool(rs("isClosed")) then
			errMsg = "«Ì‰ ”›«—‘ ﬁ»·« ﬂ‰”· ‘œÂ"
		elseif not CBool(rs("isOrder")) then
			conn.Execute("update orders set status = 2, isPrinted=0,isClosed=1,LastUpdatedDate=getDate(),LastUpdatedBy=" &session("ID")& " where id=" & orderID)
			msgFrom=session("id")
' 			msgTo=rs(-1)
			msgTitle		= "—œ «” ⁄·«„"
			msgBody			= Request("reason")
			RelatedTable	= "orders"
			relatedID		= orderID
			replyTo			= "null"
			IsReply			= 0
			urgent			= 1
			MsgDate			= shamsiToday()
			MsgTime			= currentTime10()
			msgType			= 0
			MySQL = "INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent, type) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", '"& relatedID & "', '"& RelatedTable & "', "& urgent & ", " & msgType & ")"
			conn.Execute MySQL
			msg = "«” ⁄·«„ —œ ‘œ"
		elseif CBool(rs("isOrder")) then
			set rs = conn.Execute("select * from InventoryItemRequests where status<>'del' and OrderID=" & orderID)
			while not rs.EOF
				if rs("status") = "pick" then errMsg = errMsg & "»—«Ì " & rs("itemName") & " ÕÊ«·Â ’«œ— ‘œÂ<br>"
				rs.MoveNext
			wend
			set rs = conn.Execute("select * from PurchaseRequests where status<>'del' and OrderID=" & orderID)
			while not rs.EOF
				if rs("status") = "ord" then errMsg = errMsg & "»—«Ì " & rs("typeName") & " ”›«—‘ Œ—Ìœ «ÌÃ«œ ‘œÂ<br>"
				rs.MoveNext
			wend
			if errMsg="" then 
				Conn.Execute("update InventoryItemRequests set status='del' where status='new' and orderID = " & orderID)
				Conn.Execute("update PurchaseRequests set status='del' where status='new' and orderID = " & orderID)
				conn.Execute("update invoices set voided = 1, voidedBy = " & Session("id") & ", voidedDate = '" & shamsitoday() & "' where id in (select invoice from InvoiceOrderRelations where [order]=" & orderID & ")")
				conn.Execute("update orders set status = 2,isClosed=1, isPrinted=0,LastUpdatedDate=getDate(),LastUpdatedBy=" &session("ID")& " where id=" & orderID)
				msgFrom=session("id")
' 				msgTo=0
				msgTitle		= "ﬂ‰”· ”›«—‘"
				msgBody			= Request("reason")
				RelatedTable	= "orders"
				relatedID		= orderID
				replyTo			= "null"
				IsReply			= 0
				urgent			= 1
				MsgDate			= shamsiToday()
				MsgTime			= currentTime10()
				msgType			= 0
				MySQL = "INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent, type) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", '"& relatedID & "', '"& RelatedTable & "', "& urgent & ", " & msgType & ")"
				conn.Execute MySQL
				msg = "”€«—‘ ﬂ‰”· ‘œ"
			end if
		end if
		if errMsg<>"" then 
			response.redirect "?act=show&id=" & orderID & "&errMsg=" & Server.URLEncode(errMsg)
		else
			response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode(msg)
		end if
		'-----------------------------------------------------------------------------------------------------
	case "customerSearch":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		if isnumeric(request("CustomerNameSearchBox")) then
			response.redirect "?act=getType&isOrder=" & isOrder & "&selectedCustomer=" & request("CustomerNameSearchBox")
		elseif request("CustomerNameSearchBox") <> "" then
			SA_TitleOrName=request("CustomerNameSearchBox")
			mySQL="SELECT * FROM Accounts WHERE (REPLACE(AccountTitle, ' ', '') LIKE REPLACE(N'%"& sqlSafe(SA_TitleOrName) & "%', ' ', '') ) ORDER BY AccountTitle"
			Set RS1 = conn.Execute(mySQL)
			if (RS1.eof) then 
				conn.close
				response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
			end if 

			SA_TitleOrName=request("CustomerNameSearchBox")
			SA_Action="return true;"
			SA_SearchAgainURL="order.asp"
			SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<br>
		<FORM METHOD=POST ACTION="?act=getType&isOrder=<%=isOrder%>">
			<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
		end if
		'-----------------------------------------------------------------------------------------------------
	case "getType":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		customerID=request("selectedCustomer")
		mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
		Set RS1 = conn.Execute(mySQL)
		if (RS1.eof) then 
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='..//CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
		end if 
%>
<script type="text/javascript">
	$(document).ready(function(){
		$("select[name=OrderType]").focus();
	});
</script>
<br>
<div dir='rtl'>
	<B>ê«„ ”Ê„ : ê—› ‰ ‰Ê⁄ «” ⁄·«„</B>
</div>
<form method="post" action="?act=getNew&isOrder=<%=isOrder%>" onSubmit="return checkValidation();">
	<input name="selectedCustomer" type="hidden" value="<%=customerID%>">
	<div style="margin: 50px 0 0 0;text-align: center;">
		<SELECT NAME="OrderType" style='height: 28px;' class="btn">
			<OPTION value="-1" style='color:red;'>«‰ Œ«» ﬂ‰Ìœ</option>
<%
		Set RS2 = conn.Execute("SELECT [User] as ID, DefaultOrderType FROM UserDefaults WHERE ([User] = "& session("ID") & ") OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")
		defaultOrderType=RS2("DefaultOrderType")
		RS2.close
		Set RS2 = Nothing

		set RS_TYPE=Conn.Execute ("SELECT ID, Name FROM OrderTypes WHERE (IsActive=1) ORDER BY ID")
		Do while not RS_TYPE.eof	
%>
			<OPTION value="<%=RS_TYPE("ID")%>" <%if RS_TYPE("ID")=defaultOrderType then response.write "selected"%>><%=RS_TYPE("Name")%></option>
<%
			RS_TYPE.moveNext
		loop
		RS_TYPE.close
		set RS_TYPE = nothing
%>		
		</SELECT>
		<input type="submit" value="«œ«„Â" class="btn">
	</div>
</form>
<%		'-----------------------------------------------------------------------------------------------------
	case "getNew":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		if isOrder="yes" then 
			response.redirect "?errmsg=" & Server.URLEncode("‘—„‰œÂ°ù «Ê· «” ⁄·«„ À»  ﬂ‰Ìœ Ê ”Å” ¬‰—« »Â ”›«—‘  »œÌ· ﬂ‰Ìœ")
		end if
		customerID=request("selectedCustomer")
		mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
		Set RS1 = conn.Execute(mySQL)
	
		if (RS1.eof) then 
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='..//CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
		end if 
		rs1.close
		set rs1=nothing
		set rs=Conn.Execute("select * from OrderTypes where id=" & request("orderType"))
		if rs.eof then 
			conn.close
			response.redirect "?errmsg=" & server.urlEncode("‰Ê⁄ «” ⁄·«„ —«  ⁄ÌÌ‰ ﬂ‰Ìœ")
		end if
		orderTypeID=rs("id")
		rs.close
		set rs=nothing
%>
<script type="text/javascript" src="calcOrder.11.js"></script>
<script type="text/javascript">
	function checkForceHead(e){
		if ($(e).val()=="") 
			$(e).addClass("forceErr");
		else 
			$(e).removeClass("forceErr");
	}
	function checkValidation() {
		if ($("input[name=ReturnDate]").val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("„Ê⁄œ «⁄ »«— —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=ReturnDate]').focus();
			return false;
		} else if ($('input[name="ReturnTime"]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("“„«‰ (”«⁄ ) «⁄ »«— —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=ReturnTime]').focus();
			return false;
		} else if ($('input[name=OrderTitle]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì· —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=OrderTitle]').focus();
			return false;
		} else if ($('input[name=paperSize]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("”«Ì“ —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=paperSize]').focus();
			return false;
		} else if ($('input[name=qtty]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html(" Ì—«é —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=qtty]').focus();
			return false;
		} else if ($("[name=productionDuration]").size()==1 && $("[name=productionDuration]").val()==''){
			$("#errMsg").html("“„«‰  Ê·Ìœ —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=productionDuration]').focus();
			return false;
		} else {
			return makeOutXML();
		} 
	}
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=getProperty&type=<%=orderTypeID%>","/xsl.<%=version%>/orderEditProperty.xsl", function(result){
		$("#orderDetails").html(result);	
		readyForm();
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=getNew&isOrder=&typeID=<%=orderTypeID%>&id=<%=CustomerID%>","/xsl.<%=version%>/orderEditHeader.xsl", function(result){
		$("#orderHeader").html(result);	
	});
</script>
<input type="hidden" id="vatRate" value="<%=Session("vatRate")%>"/>
<%
if auth(2,"G") then 
	Response.write "<input type='hidden' id='disLimitP' value='100'/><input type='hidden' id='disLimit' value='100000000'/>"
elseif auth(2,"F") then
	Response.write "<input type='hidden' id='disLimitP' value='10'/><input type='hidden' id='disLimit' value='1000000'/>"
else
	Response.write "<input type='hidden' id='disLimitP' value='0'/><input type='hidden' id='disLimit' value='0'/>"
end if
if auth(6,4) then 
	response.write "<input type='hidden' id='reverseLimit' value='1'/>"
else
	response.write "<input type='hidden' id='reverseLimit' value='0'/>"
end if
 %>
<br>
</div>
<br>
<!-- ê—› ‰ «” ⁄·«„ -->
<FORM METHOD=POST ACTION="?act=submitNew" onSubmit="return checkValidation();">
	<div id="orderHeader"></div>
	<div class="downBtn">
		<hr/>
		<INPUT TYPE="submit" id='Submit' Name="Submit" Value="–ŒÌ—Â" class="btn" style="width:100px;">
		<INPUT TYPE="button" Name="Cancel" Value="«‰’—«›" style="width:100px;" class="btn" onClick="window.location='order.asp';" >
	</div>
</FORM>
<div id="orderOrgin"></div>
<div id='orderDetails'></div>
<%		'-----------------------------------------------------------------------------------------------------
	case "edit":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=request("id")
		if not Auth(2 , 2) then NotAllowdToViewThisPage()
%>
<script type="text/javascript" src="calcOrder.11.js"></script>
<script type="text/javascript">
function checkValidation() {
	if (!$("[name=returnDateNull]").is(":checked")){
		if ($("input[name=ReturnDate]").val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("„Ê⁄œ  ÕÊÌ· ⁄„·Ì —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=ReturnDate]').focus();
			return false;
		} else if ($('input[name="ReturnTime"]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("“„«‰ (”«⁄ )  ÕÊÌ· ⁄„·Ì —« Ê«—œ ﬂ‰Ìœ");
			$('input[name=ReturnTime]').focus();
			return false;
		} 
	} 
	if ($('input[name=OrderTitle]').val().replace(/^\s*|\s*$/g,'')==''){
		$("#errMsg").html("⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì· —« Ê«—œ ﬂ‰Ìœ");
		$('input[name=OrderTitle]').focus();
		return false;
	} else if ($('input[name=paperSize]').val().replace(/^\s*|\s*$/g,'')==''){
		$("#errMsg").html("”«Ì“ —« Ê«—œ ﬂ‰Ìœ");
		$('input[name=paperSize]').focus();
		return false;
	} else if ($('input[name=qtty]').val().replace(/^\s*|\s*$/g,'')==''){
		$("#errMsg").html(" Ì—«é —« Ê«—œ ﬂ‰Ìœ");
		$('input[name=qtty]').focus();
		return false;
	} else {
		return makeOutXML();
	} 
}	
TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=<%=orderID%>","/xsl.<%=version%>/orderEditHeader.xsl", function(result){
		$("#orderHeader").html(result);	
		$('a#customerID').click(function(e){
			window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer='+$('a#customerID').attr("myID"), 'showCustomer');
			e.preventDefault();
		});
		$("#orderHeader .forceErr").each(function(i,e){
			if ($(e).val()=="") 
				$(e).addClass("forceErr");
			else 
				$(e).removeClass("forceErr");
		});
		$("[name=returnDateNull]").click(function(){
			$("[name=ReturnDate]").val("");
			$("[name=ReturnTime]").val("");
		});
	});
	loadXMLDoc("/service/xml_getOrderProperty.asp?act=editOrder&id=<%=orderID%>", function(orderXML){
		TransformXml(orderXML,"/xsl.<%=version%>/orderEditProperty.xsl", function(result){
			$("#orderDetails").html(result);
			readyForm();	
		});
		if ($(orderXML).find("keys").attr("canUpdate")=="yes"){
			$(".downBtn").append("<input type='button' value='»—Ê“ —”«‰Ì ﬁÌ„ ùÂ«' class='btn btn-success' onclick='updatePrice();'>")
		}
	});
		
	
function updatePrice(){
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=editOrder&id=<%=orderID%>&update=yes","/xsl.<%=version%>/orderEditProperty.xsl", function(result){
		$("#orderDetails").html(result);
		readyForm();
		alert("ﬁÌ„ ùÂ« »—Ê“ —”«‰Ì ‘œ");
	});
}
</script>
<input type="hidden" id="vatRate" value="<%=Session("vatRate")%>"/>
<%
if auth(2,"G") then 
	Response.write "<input type='hidden' id='disLimitP' value='100'/><input type='hidden' id='disLimit' value='100000000'/>"
elseif auth(2,"F") then
	Response.write "<input type='hidden' id='disLimitP' value='10'/><input type='hidden' id='disLimit' value='1000000'/>"
else
	Response.write "<input type='hidden' id='disLimitP' value='0'/><input type='hidden' id='disLimit' value='0'/>"
end if
if auth(6,4) then 
	response.write "<input type='hidden' id='reverseLimit' value='1'/>"
else
	response.write "<input type='hidden' id='reverseLimit' value='0'/>"
end if
 %>
<br><br>
<br>
<!-- ÊÌ—«Ì‘ -->
<hr>
<FORM METHOD=POST ACTION="?act=submitEdit&id=<%=orderID%>" onSubmit="return checkValidation();">
	<div id="orderHeader"></div>
	<div class="downBtn">
		<hr/>
		<INPUT TYPE="submit" id='Submit' Name="Submit" Value="–ŒÌ—Â" class="btn" style="width:100px;">
		<INPUT TYPE="button" Name="Cancel" Value="«‰’—«›" style="width:100px;" class="btn" onClick="window.location='order.asp?act=show&id=<%=orderID%>';" >
	</div>
</FORM>
<div id='orderDetails'></div>
<%		'-----------------------------------------------------------------------------------------------------
	case "printQuote":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
%>		
<!--#include File="../version.asp" -->	
<%
		orderID=CDbl(request("id"))
%>	

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
	<meta http-equiv="Content-Language" content="fa">
	<title></title>
	<link rel="StyleSheet" href="/css/jame.css" type="text/css">
	<link rel="StyleSheet" href="/css/font.css" type="text/css">
	<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
	<script type="text/javascript" src="/js/xslTransform.js"></script>
	<style>
		.logo{margin: 3px 0px 0 0; float: right;}
		.logo img{width: 220px;height: 60px;}
		.head {padding: 15px 20px 0 20px;}
		body{font-family: "b zar","b yaghut", "b compset", tahoma;font-size: 12pt; direction: rtl;}
		.sign{position: relative;}
		.sign p{position:absolute;left: 20px;top:30px;text-align: center;}
		.inthename{font-weight: bold;text-align: center;margin: 10px 0 0 0;}
		.date{text-align: left;}
		.no{text-align: left;}
		.tail{position: fixed;bottom: 5px;right: -130px; text-align: right;width: 100%; font-size: 12pt;}
	</style>
	<script type="text/javascript">
		var queries = {};
		$.each(document.location.search.substr(1).split('&'),function(c,q){
			var i = q.split('=');
			queries[i[0].toString()] = i[1].toString();
		});
		TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=" + queries['id'],"/xsl.<%=version%>/quotePrintHead.xsl", function(result){
			$("#head").html(result);
			TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=" + queries['id'],"/xsl.<%=version%>/quotePrintTail.xsl", function(result){
				$("#tail").html(result);
				
				var xslURL;
				if (queries['inShort']=="false")
					xslURL = "/xsl.<%=version%>/quotePrintProperty.xsl";
				else
					xslURL = "/xsl.<%=version%>/quotePrintPropertyShort.xsl";
				TransformXmlURL("/service/xml_getOrderProperty.asp?act=showOrder&id=" + queries['id'], xslURL, function(result){
					$("#quote").html(result);
					window.print();
					document.location='?act=show&id=<%=orderID%>';
				});
			});
		});
		 
		
	</script>
</head>

<body>
	<div id="head"></div>
	<div id="quote"></div>
	<div id="tail"></div>
</body>
</html>
<%			
end select
%>
