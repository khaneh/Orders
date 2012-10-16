<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
if request("isOrder")="yes" then
	PageTitle="�����"
else
	PageTitle="�������"
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
	#isApprovedStamp {background-image: url("/images/approved.gif"); position: absolute;top: 0px;left: -30px;height: 92px;width: 199px;-webkit-filter: hue-rotate(120deg) drop-shadow(20px 15px 20px #000);-webkit-transform: rotate(-40deg);-moz-filter: hue-rotate(120deg) drop-shadow(20px 15px 20px #000);-moz-transform: rotate(-40deg);}
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
</STYLE>
<%
end if
if not Auth(2 , 9) then NotAllowdToViewThisPage() '������� 
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
		<span>��� ����� �� ���� �� ����� �����/�������:</span>
		<INPUT TYPE="text" NAME="search_box" id="search_box" value="<%=searchBox%>" class="boot">
		<INPUT TYPE="submit" NAME="SubmitB" Value="�����" class="btn">
		<% if Auth(2 , 5) then %>
			<A class="btn offset1" HREF="order.asp?act=advancedSearch">������ �������</A>
		<% End If %>
	</div>
</FORM>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		$("#search_box").focus();
		TransformXmlURL("/service/xml_getOrderTrace.asp?act=search&searchBox=<%=Server.URLEncode(searchBox)%>","/xsl/orderShowList.xsl?v=<%=version%>", function(result){
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
			<span class="col0">�����/�������</span>
			<span>��� �����</span>
			<input type="radio" name="isOrder" value="1" <%if request("isOrder")="1" then response.write " checked='checked'"%>>
			<span>��� �������</span>
			<input type="radio" name="isOrder" value="0" <%if request("isOrder")="0" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkClosed" class="check" <%if request("checkClosed")="on" then response.write " checked='checked'"%>>
			<span class="col0">��� ����� ��� ���</span>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkID" class="check" <%if request("checkID")="on" then response.write " checked='checked'"%>>
			<span class="col0">����� �����</span>
			<input type="text" name="fromID" value="<%=request("fromID")%>" size="10" class="boot num">
			<span>��</span>
			<input type="text" name="toID" value="<%=request("toID")%>" size="10" class="boot num">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkOrderType" class="check" <%if request("checkOrderType")="on" then response.write " checked='checked'"%>>
			<span class="col0">��� �����</span>
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
			<span class="not">�����</span>
			<input type="checkbox" name="checkNotOrderType" class="not" <%if request("checkNotOrderType")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCreatedDate" class="check" <%if request("checkCreatedDate")="on" then response.write " checked='checked'"%>>
			<span class="col0">����� �����</span>
			<input type="text" name="fromCreatedDate" value="<%=request("fromCreatedDate")%>" size="10" class="boot date">
			<span>��</span>
			<input type="text" name="toCreatedDate" value="<%=request("toCreatedDate")%>" size="10" class="boot date">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkStep" class="check" <%if request("checkStep")="on" then response.write " checked='checked'"%>>
			<span class="col0">�����</span>
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
			<span class="not">�����</span>
			<input type="checkbox" name="checkNotStep" class="not" <%if request("checkNotStep")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkRetDate" class="check" <%if request("checkRetDate")="on" then response.write " checked='checked'"%>>
			<span class="col0">����� ����� ����</span>
			<input type="text" name="fromRetDate" size="10" value="<%=request("fromRetDate")%>" class="boot date">
			<span>��</span>
			<input type="text" name="toRetDate" size="10" value="<%=request("toRetDate")%>" class="boot date">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkStatus" class="check" <%if request("checkStatus")="on" then response.write " checked='checked'"%>>
			<span class="col0">�����</span>
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
			<span class="not">�����</span>
			<input type="checkbox" name="checkNotStatus" class="not" <%if request("checkNotStatus")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkContractDate" class="check" <%if request("checkContractDate")="on" then response.write " checked='checked'"%>>
			<span class="col0">����� ����� �������</span>
			<input type="text" name="fromContractDate" value="<%=request("fromContractDate")%>" size="10" class="boot date">
			<span>��</span>
			<input type="text" name="toContractDate" value="<%=request("toContractDate")%>" size="10" class="boot date">
		</div>
		
		<div class="sec">
			<input type="checkbox" name="checkRetIsNull" class="check" <%if request("checkRetIsNull")="on" then response.write " checked='checked'"%>>
			<span class="col0">����� ����� ������� �����</span>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCompany" class="check" <%if request("checkCompany")="on" then response.write " checked='checked'"%>>
			<span class="col0">��� ����</span>
			<input type="text" name="companyName" value="<%=request("companyName")%>" size="20" class="boot">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCreatedBy" class="check" <%if request("checkCreatedBy")="on" then response.write " checked='checked'"%>>
			<span class="col0">����� ������</span>
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
			<span class="not">�����</span>
			<input type="checkbox" name="checkNotCreatedBy" class="not" <%if request("checkNotCreatedBy")="on" then response.write " checked='checked'"%>>
		</div>
		<div class="sec">
			<input type="checkbox" name="checkCustomer" class="check" <%if request("checkCustomer")="on" then response.write " checked='checked'"%>>
			<span class="col0">��� �����</span>
			<input type="text" name="customerName" value="<%=request("customerName")%>" size="20" class="boot">
		</div>
		<div class="sec">
			<input type="checkbox" name="checkTel" class="check" <%if request("checkTel")="on" then response.write " checked='checked'"%>>
			<span class="col0">����</span>
			<input type="text" name="tel" value="<%=request("tel")%>" size="20" class="boot num">
		</div>
		
		<div class="sec">
			<input type="checkbox" name="checkTitle" class="check" <%if request("checkTitle")="on" then response.write " checked='checked'"%>>
			<span class="col0">����� �����</span>
			<input type="text" name="orderTitle" value="<%=request("orderTitle")%>" size="20" class="boot">
		</div>
		<div class="myBtn">
			<span>����� �����</span>
			<input type="text" size="3" name="resultCount" class="myBtn num" value='<%if request("resultCount")="" then response.write "50" else response.write cint(request("resultCount"))%>'>
			<INPUT TYPE="submit" NAME="submitBtn" Value="�����" class="btn">
			
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
		if request("submitBtn")="�����" then
%>
<SCRIPT type="text/javascript">
	$(document).ready(function(){
		$("#search_box").focus();
		TransformXmlURL("/service/xml_getOrderTrace.asp?act=advanceSearch&<%=request.form%>","/xsl/orderShowList.xsl?v=<%=version%>", function(result){
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
		<span>����� �������/����� �� ���� ����:</span>
		<INPUT Name="id" id="orderID" TYPE="text" maxlength="6" size="5" class="boot">
	</div>
	<div class="rspan1">
		<INPUT TYPE="submit" Name="Submit" Value="�����" class="btn">
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
		<span>����� �������/����� �� ���� ����:</span>
		<INPUT Name="id" id="orderID" TYPE="text" maxlength="6" size="5" class="boot">
	</div>
	<div class="rspan1">
		<INPUT TYPE="submit" Name="Submit" Value="�����" class="btn">
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
	<div class="rspan5"><B>��� ��� : ����� ���� ��� ����</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox" id="CustomerNameSearchBox" class="boot">
	</div>
	<div class="rspan1">
		<INPUT TYPE="submit" value="�����" class="btn">
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
%>
<script type="text/javascript" src="/js/jquery.printElement.min.js"></script>
<script type="text/javascript">
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=<%=orderID%>","/xsl/orderShowHeader.xsl?v=<%=version%>", function(result){
		$("#orderHeader").html(result);	
		$('a#customerID').click(function(e){
			window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer='+$('a#customerID').attr("myID"), 'showCustomer');
			e.preventDefault();
		});
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=showOrder&id=<%=orderID%>","/xsl/orderShowProperty.xsl?v=<%=version%>", function(result){
		$("#orderDetails").html(result);
		$('div.priceGroup').tooltip({
	      selector: "div[rel=tooltip]"
	    });	
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=showLog&id=<%=orderID%>","/xsl/orderShowlogs.xsl?v=<%=version%>", function(result){
		$("#orderLogs").html(result);	
		$("#orderLogs td.date").each(function(i){
			$(this).html($.jalaliCalendar.gregorianToJalaliStr($(this).html()));
		});
		$("#orderLogs td.time").each(function(i){
			myTime = new Date($(this).html().replace(RegExp('-','g'),'/'));
			$(this).html(((myTime.getHours()<10)?('0'+myTime.getHours()):myTime.getHours())+':'+((myTime.getMinutes() < 10)?('0'+myTime.getMinutes()):(myTime.getMinutes())));
		});
	});
	TransformXmlURL("/service/xml_getMessage.asp?act=related&table=orders&id=<%=orderID%>","/xsl/showRelatedMessage.xsl?v=<%=version%>", function(result){
		$("#orderMessages").html(result);
		$("td.msgBody").each(function(i){
			$(this).html($(this).html().replace(/\n/gi,"<br/>"));
			});
			$("input[name=addNewMessage]").click(function(){
				document.location="../home/message.asp?RelatedTable=orders&RelatedID=" + $("#orderID").val() + "&retURL=" + escape("../Order/order.asp?act=show&id=" + $("#orderID").val());
		});
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=stock&id=<%=orderID%>","/xsl/orderShowRequestInventory.xsl?v=<%=version%>", function(result){
		$("#orderStock").html(result);	
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=purchase&id=<%=orderID%>","/xsl/orderShowRequestPurchase.xsl?v=<%=version%>", function(result){
		$("#orderPurchase").html(result);	
	});
	TransformXmlURL("/service/xml_getCosts.asp?act=orderCost&id=<%=orderID%>","/xsl/costListShow.xsl?v=<%=version%>", function(result){
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
		$("#costDelComfirm").dialog({
			autoOpen: false,
			buttons: {"�����":function(){
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
			title: "�����"
		});
		$("#approveOrder").dialog({
			autoOpen: false,
			buttons: {"�����":function(){
				document.location='?act=approve&id=<%=orderID%>';
				$(this).dialog("close");
			}},
			title: "�����"
		});
		$("#approveQuote").dialog({
			autoOpen: false,
			buttons: {"�����":function(){
				document.location='?act=approve&id=<%=orderID%>';
				$(this).dialog("close");
			}},
			title: "�����"
		});
		$('#approvedBtn').click(function(){
			if ($("#isOrder").val()=='False')
				$("#approveQuote").dialog("open");
			else
				$("#approveOrder").dialog("open");
		});
		$("#changeCustomer").dialog({
			autoOpen: false,
			title: "����� ���� �����",
			height: 350
		});
		$("#cancelBtn").click(function(){
			$("#cancelDlg").dialog("open");
		});
		$("#cancelDlg").dialog({
			autoOpen: false,
			title: "�� �������/�����",
			buttons: {"�����":function(){
				if ($("#cancelDlg input[name=cancelReason]").val()!=''){
					document.location='?act=cancel&id=<%=orderID%>&reason=' + escape($("#cancelDlg input[name=cancelReason]").val());
					$(this).dialog("close");
				}
			}},
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
			if ($("#isOrder").val()=='False'){
				$("#printQuoteDlg").dialog("open");
			} else {
				$('#printArea').printElement({overrideElementCSS:['/css/order_property.css','/css/jame.css'],pageTitle:'��� ����� - " & OrderID & "',printMode:'popup'});
			}
			
		});
		$("#printQuoteDlg").dialog({
			autoOpen: false,
			title: "���� �ǁ �������",
			buttons: {"�����": function(){
				$("#printQuoteDlg").dialog("close");
				document.location='?act=printQuote&id=<%=orderID%>&inShort=' + $("#inShort").prop("checked");
			}}
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
	set rs = Conn.Execute("SELECT orders.id, orders.isClosed,orders.isApproved,orders.isOrder, isnull(invoices.Approved,0) as Approved,isnull(invoices.Issued,0) as Issued FROM orders left outer join InvoiceOrderRelations on InvoiceOrderRelations.[Order] = orders.id left outer join Invoices on Invoices.ID=InvoiceOrderRelations.Invoice WHERE orders.ID=" & orderID)
%>
<div id="costDelComfirm">
	<h3>��� ����� ����� �� ������ ��� ��� ��Ͽ</h3>
	<input type="hidden" id="costDelItemID" />
	<lu id="costDelComfirmList"></lu>
</div>
<div id="approveOrder">
	<h3>����� ����� �� ����� ����� �� �� ���� ��� ������. � ���� �� ��� ����� ��� �� ��� ������ ��� �� �� �������</h3>
	<lu>
		<li>����� �� ���� ���� ��� ��� ���</li>
		<li>������� � ����� �����  �� ����� �� ��� ���</li>
		<li>������� �� ����� ����� ��� � ����� ��� ���</li>
	</lu>
</div>
<div id="approveQuote">
	<h3>����� ������� �� ����� ����� ��� ������ � ���� �� ��� ����� ��� �� ��� ������ ��� �� �� �������</h3>
	<lu>
		<li>������� �� ���� ���� ��� ��� ���</li>
		<li>������� ���� ����� ����� ��� ���</li>
	</lu>
</div>
<div id="changeCustomer">
	<h5>��� � �� ����� ����� �� ���� ����</h5>
	<input type="text" size="40" name="search"/>
	<input type="hidden" name="customerID"/>
</div>
<div id="cancelDlg">
	<h3>���� ���� ��� �� ��� ����</h3>
	<input type="text" size="30" name="cancelReason"/>
</div>
<div id="printQuoteDlg">
	<input type="checkbox" id="inShort"/>	
	<span>����� ����</span>
</div>
<input type="hidden" id='isOrder' value="<%=rs("isOrder")%>">
<input type="hidden" id='isClosed' value="<%=rs("isClosed")%>">
<input type="hidden" id='orderID' value="<%=rs("id")%>">
<div class="topBtn">
<%	
	isOrder = rs("isOrder")
	if rs("isApproved") then Response.write "<div id='isApprovedStamp'></div>"
	if (auth(2,4) and not CBool(rs("isClosed")) and not CBool(rs("issued"))) then Response.write "<a class='btn btn-primary' href='#' id='changeCustomerBtn'>����� ����</a>"
	if (auth(2,2) and not CBool(rs("isClosed")) and Not CBool(rs("approved")) and not CBool(rs("issued"))) then Response.write "<a class='btn' href='?act=edit&id=" & orderID & "'>������</a>"
	if (auth(2,1) and not CBool(rs("isOrder")) and CBool(rs("isApproved"))) then Response.write "<a class='btn btn-success' href='?act=convertToOrder&id=" & orderID & "'>����� �� �����</a>"
	if (auth(2,"A") and not CBool(rs("isApproved")) and not CBool(rs("isClosed"))) then Response.write "<a class='btn btn-primary' href='#' id='approvedBtn'>�����</a>"
	if (auth(2,2) and not CBool(rs("isClosed")) and CBool(rs("isOrder"))) then Response.write "<a class='btn' href='?act=pause&id=" & orderID & "'>����</a>"
	if (auth(2,"C") and not rs("isClosed") and rs("isOrder")) then Response.write "<a class='btn btn-danger' href='#' id='cancelBtn'>����</a>"
	if (not rs("isClosed") and not rs("isOrder")) then Response.write "<a class='btn btn-danger' href='#' id='cancelBtn'>�� �������</a>"
	if (rs("isApproved")) then
		set rs = Conn.Execute("select * from InvoiceOrderRelations where [order]=" & orderID)
		while not rs.eof
			response.write "<a class='btn btn-inverse' href='../AR/AccountReport.asp?act=showInvoice&invoice=" & rs("invoice") & "' title='����� " & rs("invoice") & "'>��������</a>"
			rs.MoveNext
		wend
	end if
	if (auth(2,"B")) then Response.write "<a class='btn' href='#' id='printBtn'>�ǁ</a>" '--- print 
%>
</div>
<div id="orderLogs"></div>
<div id="printArea">
<!-- 	<div id="voidedIcon"></div> -->
	<div id="orderHeader"></div>
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
		conn.Execute("update orders set customer = " & customerID & " where id = " & orderID)
		conn.Execute("update invoices set customer = " & customerID & " from invoices inner join invoiceOrderRelations on invoiceOrderRelations.invoice = invoices.id where invoiceOrderRelations.[order] = " & orderID)
		response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("����� ��� ����� ����� ���� ���")
		'-----------------------------------------------------------------------------------------------------
	case "submitNew":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderTypeID = request("orderType")
		CustomerID=request.form("CustomerID") 
		mySQL="SET NOCOUNT ON;INSERT INTO orders (CreatedBy, Customer, isClosed, returnDate, orderTitle, Type, status, step, LastUpdatedBy, Notes, property,productionDuration, Qtty, price, PaperSize,lastUpdatedDate) VALUES ("& session("ID") & ", "& CustomerID & ", 0, N'"& sqlSafe(request.form("ReturnDateTime")) & "', N'"& sqlSafe(request.form("OrderTitle")) & "', "& orderTypeID & ", 1, 1, "& session("ID") & ", N'"& sqlSafe(request.form("Notes")) & "',N'" & request.form("myXML") & "', " & sqlSafe(request.form("productionDuration")) & "," & sqlSafe(request.form("qtty")) & ",'" & sqlSafe(request.form("totalPrice")) & "', N'" & sqlSafe(request.form("PaperSize")) & "',getdate()); SELECT SCOPE_IDENTITY() AS NewQuote;SET NOCOUNT OFF"
		response.write mySQL
				
		set RS1 = Conn.execute(mySQL)
		QuoteID = RS1 ("NewQuote")
		RS1.close
		conn.close
		response.redirect "?act=show&id=" & QuoteID & "&msg=" & Server.URLEncode("������� ������� ��� ��")
		'-----------------------------------------------------------------------------------------------------
	case "submitEdit":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		set rs = Conn.Execute("select orders.isClosed, orders.isApproved, orders.isOrder, isnull(invoices.approved,0) as approved, isnull(invoices.issued,0) as issued from orders left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice = invoices.id where orders.id = " & orderID)
		msg=""
		errMsg=""
		if rs.eof then 
			errMsg = "���� ����� �����/������� �� ���� ���!!!!"
		elseif rs("isClosed") then 
			errMsg = "����� ���� ���! ���� ����� ������ �� ���� �����"
		elseif rs("issued") then 
			errMsg = "���� ����! ����� ���� ���� ��� ���� ���!!!"
		elseif rs("approved") then 
			errMsg = "�������� ����� ���"
		elseif rs("isApproved") then 
			msg = "�����/������� ����� ��� �� ������ �����ϡ<br> ���� �� ����� ���� ��<br>"
		end if
		if errMsg="" then
			if Request.form("returnDateNull")="on" then 
				returnDateTime="null"
			else
				returnDateTime = "N'" & request.form("returnDateTime") & "'"
			end if
			if request.form("myXML")<>"" then 
				Conn.Execute ("update orders set property=N'" & request.form("myXML") & "', productionDuration=" &cint(request.form("productionDuration"))& ", Notes=N'" &sqlSafe(request.form("Notes"))& "', returnDate=" & returnDateTime & ", paperSize=N'" &sqlSafe(request.form("paperSize"))& "', OrderTitle=N'" &sqlSafe(request.form("OrderTitle"))& "', qtty=" &cdbl(request.form("qtty"))& ", price='" & sqlSafe(request.form("totalPrice")) & "',LastUpdatedDate=getDate(), LastUpdatedBy=" &session("ID")& ", Customer=" &sqlSafe(request.form("customerID"))& ",isApproved=0 where id="&orderID)
				conn.Execute("delete InvoiceLines where Invoice in (select invoice from InvoiceOrderRelations where [Order]=" & orderID & ")")
				response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode(msg & "������� �� ��� ��")
			else	
				response.redirect "?act=show&id=" & orderID & "&errMsg=" & Server.URLEncode("����� �� ����!")
			end if
		else
			response.redirect "?act=show&id=" & orderID & "&errMsg=" & Server.URLEncode(errMsg)
		end if
		'-----------------------------------------------------------------------------------------------------
	case "convertToOrder":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		if not Auth(2 , 1) then NotAllowdToViewThisPage()
		set rs = Conn.Execute("select orders.*, accounts.arBalance, accounts.creditLimit, accounts.status as accountStatus from orders inner join accounts on orders.customer=accounts.id where orders.id=" & orderID)
		if rs("isOrder") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("���� �� ����� ����� ���!")
		elseif not rs("isApproved") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("����� ��� ������� �� ����� ����!")
		elseif (cdbl(rs("arBalance"))+cdbl(rs("creditLimit")) < 0) then 
			CustomerID = cdbl(rs("customer"))
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("���� ��� ���� �� ����� ������ �� ����� ���<br> ���� �� �с��� ���� ����� ����.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>����� ����</a>")
		elseif CInt(rs("accountStatus"))<>1 then 
			CustomerID = cdbl(rs("customer"))
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("��� ���� ���� ����!<br> ���� �� �с��� ���� ����� ����.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>����� ����</a>")
		else
			Conn.Execute("update orders set isOrder=1,isApproved=0,returnDate=null where id=" & orderID)
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&msg=" & Server.URLEncode("�� ����� ����� ��!")
		end if
		'-----------------------------------------------------------------------------------------------------
	case "approve":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		set rs=Conn.Execute("select Orders.Customer, orders.isOrder, orders.isApproved, orders.property as data, orderTypes.property as keys from Orders inner join orderTypes on orders.Type=orderTypes.ID where orders.id=" & orderID & " and lastUpdatedDate>lastUpdate")
		if rs.eof then 
			response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b>���ʝ�� ���� ���! </b><br>��� �� ����ϡ ���� �����/������� �� ���� ����")
		elseif (rs("isApproved")) then 
			response.redirect "?act=show&id=" & orderID & "&errmsg=" & Server.URLEncode("���� ����� ���!")
		else
			if (not rs("isOrder")) then 
				Conn.execute("update orders set isApproved=1,approvedDate=getDate() where id="& orderID)
				response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("������� ����� ��")
			else
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
						for each sKey in sGroup.selectNodes("./key")
							keyName=sKey.GetAttribute("name")
							if row.selectNodes("./key[@name='" & keyName & "']").length>0 then hasKey = hasKey + 1
						next
						if rowDesc<>"" then desc = desc & "(" & rowDesc & ")"
						if hasKey>0 then 
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
										if sGroup.selectNodes("./key[@name='" & groupName & "-size']/option[" & sizeID & "]").length>0 then 
											set tmp = sGroup.selectSingleNode("./key[@name='" & groupName & "-size']/option[" & sizeID & "]")
											if InStr(tmp.GetAttribute("label"),"X")>0 then 
												l = CDbl(split(tmp.GetAttribute("label"),"X")(0))
												w = CDbl(split(tmp.GetAttribute("label"),"X")(1))
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
									mySQL = "INSERT INTO InventoryItemRequests (orderID,ItemName,comment,qtty,createdBy,invoiceItem,rowID, customerHaveInvItem) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & ")"
									Conn.Execute(mySQL)
								else
									'----------- Check UPDATE or INSERT ?
									mySQL = "select * from InventoryItemRequests where orderID=" & orderID & " and invoiceItem = " & item & " and rowID = " & rowID & " order by ID desc"
	' 								Response.write mySQL
									set rsr = Conn.Execute(mySQL)
									if rsr.eof then 
										mySQL = "INSERT INTO InventoryItemRequests (orderID,ItemName,comment,qtty,createdBy,invoiceItem,rowID, customerHaveInvItem) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & ")"
										Conn.Execute(mySQL)
									else
										requestID = Replace(requestID,rsr("id") & ",","")
										if rsr("status")="new" and (cdbl(rsr("qtty")) <> cdbl(appQtty) or Trim(rsr("itemName")) <> Trim(stockName) or Trim(rsr("comment")) <> Trim(stockDesc) or CBool(rsr("customerHaveInvItem"))<> CBool(customerHasStock)) then 
											'------- IF CHANGE UPDATE
											mySQL = "UPDATE InventoryItemRequests SET ItemName=N'" & stockName & "', comment=N'" & stockDesc & "', qtty=" & appQtty & ", createdBy=" & Session("ID") & ", customerHaveInvItem=" & customerHasStock & " where id=" & rsr("id")
											Conn.Execute(mySQL)
										elseif rsr("status")="del" then 
											'------- IF delete Insert new request
											mySQL = "INSERT INTO InventoryItemRequests (orderID,ItemName,comment,qtty,createdBy,invoiceItem,rowID,customerHaveInvItem) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & "," & customerHasStock & ")"
										Conn.Execute(mySQL)
										elseif CDbl(rsr("qtty")) <> CDbl(appQtty) or Trim(rsr("itemName"))<> Trim(stockName) or Trim(rsr("comment")) <> Trim(stockDesc) or CBool(rsr("customerHaveInvItem"))<> CBool(customerHasStock) then
											errMSG = errMSG & "���� "  & stockName & " ����� ���� ���. �� ����� �� �� ���� ����<br>"
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
									mySQL = "select * from PurchaseRequests where orderID=" & orderID & " and invoiceItem = " & item & " and rowID = " & rowID & " order by ID desc"
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
											errMSG = errMSG & "���� "  & purchaseName & " ����� ���� ����� ���. �� ����� �� �� ���� ����<br>"
										end if
									end if
								end if
							end if
						end if
					next
				next
				RFD = TotalReceivable - fix(TotalReceivable / 1000) * 1000
				TotalReceivable = TotalReceivable - RFD
				TotalDiscount = TotalDiscount + RFD
				if RFD > 0 then 
					mySQL = "INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, [Reverse], Vat, hasVat) VALUES (" & InvoiceID & ", 39999, N'����� ��� ������', 0, 0, 0, 0, 0, 0, " & RFD & ", 0, 0, 0)"
					'Response.write mySQL
					Conn.Execute(mySQL)
				end if
				mySQL = "update invoices set totalPrice=" & totalPrice & ", totalDiscount=" & totalDiscount & ", TotalReceivable=" & TotalReceivable & ", totalVat=" & totalVat & ",isA=" & isADefault & "  where id=" & invoiceID
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
								errMSG = errMSG & "���� "  & stockName & " ����� ���� ���. � ��� ���� ��� �������!<br>"
							else
								Conn.Execute("update InventoryItemRequests set status='del' where id=" & id)
							end if
						end if
					Next
				end if
				if errMSG="" then 
					Conn.execute("update orders set isApproved=1,approvedDate=getDate() where id="& orderID)
					'response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b>����� ��</b>")
' 					Response.write "<a href='?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b>����� ��</b>")&"'>��� ��</a>"
					response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b>����� ��</b>")
				else
					response.redirect "?act=show&id=" & orderID & "&errmsg=" & Server.URLEncode(errMSG)
' 					Response.write "<a href='?act=show&id=" & orderID & "&errmsg=" & Server.URLEncode(errMSG)&"'>��� ��</a>"
				end if
				rs.Close
				set rs=Nothing
				rsr.Close
				set rsr=Nothing
				rsp.Close
				set rsp = Nothing
			end if
		end if
		'-----------------------------------------------------------------------------------------------------
	case "pause":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		
		'-----------------------------------------------------------------------------------------------------
	case "cancel":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		if not Auth(2 , "C") then NotAllowdToViewThisPage()
		orderID=cdbl(request("id"))
		set rs = Conn.Execute("select orders.createdBy, orders.isClosed, orders.isApproved, orders.isOrder, isnull(invoices.approved,0) as approved, isnull(invoices.issued,0) as issued from orders left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice = invoices.id where orders.id = " & orderID)
		errMsg = ""
		msg = ""
		msgTo = rs("createdBy")
		if Request("reason")="" then
			errMsg = "���� ���� ���� ����� ����"
		elseif rs.eof then 
			errMsg = "���� ������ ���� ���!"
		elseif CBool(rs("issued")) then 
			errMsg = "��� ����� ������ ���� ��� ����! ���� ����� �� ���� ����"
		elseif CBool(rs("approved")) then 
			errMsg = "��� ����� ������ ����� ��� ����! ���� ����� �� ���� ���ʡ �� ������ ������ �� ����� ���� ���"
		elseif CBool(rs("isClosed")) then
			errMsg = "��� ����� ���� ���� ���"
		elseif not CBool(rs("isOrder")) then
			conn.Execute("update orders set isClosed=1 where id=" & orderID)
			msgFrom=session("id")
' 			msgTo=rs(-1)
			msgTitle		= "�� �������"
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
			msg = "������� �� ��"
		elseif CBool(rs("isOrder")) then
			set rs = conn.Execute("select * from InventoryItemRequests where status<>'del' and OrderID=" & orderID)
			while not rs.EOF
				if rs("status") = "pick" then errMsg = errMsg & "���� " & rs("itemName") & " ����� ���� ���<br>"
				rs.MoveNext
			wend
			set rs = conn.Execute("select * from PurchaseRequests where status<>'del' and OrderID=" & orderID)
			while not rs.EOF
				if rs("status") = "ord" then errMsg = errMsg & "���� " & rs("typeName") & " ����� ���� ����� ���<br>"
				rs.MoveNext
			wend
			if errMsg="" then 
				Conn.Execute("update InventoryItemRequests set status='del' where status='new' and orderID = " & orderID)
				Conn.Execute("update PurchaseRequests set status='del' where status='new' and orderID = " & orderID)
				conn.Execute("update invoices set voided = 1, voidedBy = " & Session("id") & ", voidedDate = '" & shamsitoday() & "' where id in (select invoice from InvoiceOrderRelations where [order]=" & orderID & ")")
				conn.Execute("update orders set isClosed=1 where id=" & orderID)
				msgFrom=session("id")
' 				msgTo=0
				msgTitle		= "���� �����"
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
				msg = "����� ���� ��"
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
				response.redirect "?errmsg=" & Server.URLEncode("���� ����� ���� ���.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>���� ���Ͽ</a>")
			end if 

			SA_TitleOrName=request("CustomerNameSearchBox")
			SA_Action="return true;"
			SA_SearchAgainURL="order.asp"
			SA_StepText="��� ��� : ������ ����"
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
			response.redirect "?errmsg=" & Server.URLEncode("���� ����� ���� ���.<br><a href='..//CRM/AccountEdit.asp?act=getAccount'>���� ���Ͽ</a>")
		end if 
%>
<script type="text/javascript">
	$(document).ready(function(){
		$("select[name=OrderType]").focus();
	});
</script>
<br>
<div dir='rtl'>
	<B>��� ��� : ����� ��� �������</B>
</div>
<form method="post" action="?act=getNew&isOrder=<%=isOrder%>" onSubmit="return checkValidation();">
	<input name="selectedCustomer" type="hidden" value="<%=customerID%>">
	<div style="margin: 50px 0 0 0;text-align: center;">
		<SELECT NAME="OrderType" style='height: 28px;' class="btn">
			<OPTION value="-1" style='color:red;'>������ ����</option>
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
		<input type="submit" value="�����" class="btn">
	</div>
</form>
<%		'-----------------------------------------------------------------------------------------------------
	case "getNew":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		if isOrder="yes" then 
			response.redirect "?errmsg=" & Server.URLEncode("�����塝 ��� ������� ��� ���� � Ӂ� ���� �� ����� ����� ����")
		end if
		customerID=request("selectedCustomer")
		mySQL="SELECT * FROM Accounts WHERE (ID='"& CustomerID & "')"
		Set RS1 = conn.Execute(mySQL)
	
		if (RS1.eof) then 
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("���� ����� ���� ���.<br><a href='..//CRM/AccountEdit.asp?act=getAccount'>���� ���Ͽ</a>")
		end if 
		rs1.close
		set rs1=nothing
		set rs=Conn.Execute("select * from OrderTypes where id=" & request("orderType"))
		if rs.eof then 
			conn.close
			response.redirect "?errmsg=" & server.urlEncode("��� ������� �� ����� ����")
		end if
		orderTypeID=rs("id")
		rs.close
		set rs=nothing
%>
<script type="text/javascript" src="calcOrder.js"></script>
<script type="text/javascript">
	function checkValidation() {
		if ($("input[name=ReturnDate]").val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("���� ������ �� ���� ����");
			$('input[name=ReturnDate]').focus();
			return false;
		} else if ($('input[name="ReturnTime"]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("���� (����) ������ �� ���� ����");
			$('input[name=ReturnTime]').focus();
			return false;
		} else if ($('input[name=OrderTitle]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("����� ��� ���� ���� �� ���� ����");
			$('input[name=OrderTitle]').focus();
			return false;
		} else if ($('input[name=paperSize]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("���� �� ���� ����");
			$('input[name=paperSize]').focus();
			return false;
		} else if ($('input[name=qtty]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("���ǎ �� ���� ����");
			$('input[name=qtty]').focus();
			return false;
		} else if ($("[name=productionDuration]").size()==1 && $("[name=productionDuration]").val()==''){
			$("#errMsg").html("���� ����� �� ���� ����");
			$('input[name=productionDuration]').focus();
			return false;
		} else {
			return makeOutXML();
		} 
	}
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=getProperty&type=<%=orderTypeID%>","/xsl/orderEditProperty.xsl?v=<%=version%>", function(result){
		$("#orderDetails").html(result);	
		readyForm();
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=getNew&isOrder=&typeID=<%=orderTypeID%>&id=<%=CustomerID%>","/xsl/orderEditHeader.xsl?v=<%=version%>", function(result){
		$("#orderHeader").html(result);	
	});
</script>
<input type="hidden" id="vatRate" value="<%=Session("vatRate")%>"/>
<br>
</div>
<br>
<!-- ����� ������� -->
<FORM METHOD=POST ACTION="?act=submitNew" onSubmit="return checkValidation();">
	<div id="orderHeader"></div>
	<div class="downBtn">
		<hr/>
		<INPUT TYPE="submit" id='Submit' Name="Submit" Value="�����" class="btn" style="width:100px;">
		<INPUT TYPE="button" Name="Cancel" Value="������" style="width:100px;" class="btn" onClick="window.location='order.asp';" >
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
<script type="text/javascript" src="calcOrder.js"></script>
<script type="text/javascript">
function checkValidation() {
	if (!$("[name=returnDateNull]").is(":checked")){
		if ($("input[name=ReturnDate]").val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("���� ����� ���� �� ���� ����");
			$('input[name=ReturnDate]').focus();
			return false;
		} else if ($('input[name="ReturnTime"]').val().replace(/^\s*|\s*$/g,'')==''){
			$("#errMsg").html("���� (����) ����� ���� �� ���� ����");
			$('input[name=ReturnTime]').focus();
			return false;
		} 
	} 
	if ($('input[name=OrderTitle]').val().replace(/^\s*|\s*$/g,'')==''){
		$("#errMsg").html("����� ��� ���� ���� �� ���� ����");
		$('input[name=OrderTitle]').focus();
		return false;
	} else if ($('input[name=paperSize]').val().replace(/^\s*|\s*$/g,'')==''){
		$("#errMsg").html("���� �� ���� ����");
		$('input[name=paperSize]').focus();
		return false;
	} else if ($('input[name=qtty]').val().replace(/^\s*|\s*$/g,'')==''){
		$("#errMsg").html("���ǎ �� ���� ����");
		$('input[name=qtty]').focus();
		return false;
	} else {
		return makeOutXML();
	} 
}	
TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=<%=orderID%>","/xsl/orderEditHeader.xsl?v=<%=version%>", function(result){
		$("#orderHeader").html(result);	
		$('a#customerID').click(function(e){
			window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer='+$('a#customerID').attr("myID"), 'showCustomer');
			e.preventDefault();
		});
		$("[name=returnDateNull]").click(function(){
			$("[name=ReturnDate]").val("");
			$("[name=ReturnTime]").val("");
		});
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=editOrder&id=<%=orderID%>","/xsl/orderEditProperty.xsl?v=<%=version%>", function(result){
		$("#orderDetails").html(result);
		readyForm();	
	});
</script>
<input type="hidden" id="vatRate" value="<%=Session("vatRate")%>"/>
<br><br>
<br>
<!-- ������ -->
<hr>
<FORM METHOD=POST ACTION="?act=submitEdit&id=<%=orderID%>" onSubmit="return checkValidation();">
	<div id="orderHeader"></div>
	<div class="downBtn">
		<hr/>
		<INPUT TYPE="submit" id='Submit' Name="Submit" Value="�����" class="btn" style="width:100px;">
		<INPUT TYPE="button" Name="Cancel" Value="������" style="width:100px;" class="btn" onClick="window.location='order.asp?act=show&id=<%=orderID%>';" >
	</div>
</FORM>
<div id='orderDetails'></div>
<%		'-----------------------------------------------------------------------------------------------------
	case "printQuote":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
%>		<!--#include File="../version.asp" -->	<%
		orderID=request("id")
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
		.logo{margin: 10px 0px 0 0; }
		.logo img{width: 220px;height: 60px;}
		body{font-family: "b zar","b yaghut", "b compset", tahoma;font-size: 14pt; direction: rtl;}
		.sign{position: relative;}
		.sign p{position:absolute;left: 20px;top:30px;text-align: center;}
		.inthename{font-weight: bold;text-align: center;margin: 30px 0 0 0;}
		.date{text-align: left;}
		.tail{position: fixed;bottom: 5px;right: -130px; text-align: right;width: 100%; font-size: 12pt;}
	</style>
	<script type="text/javascript">
		var queries = {};
		$.each(document.location.search.substr(1).split('&'),function(c,q){
			var i = q.split('=');
			queries[i[0].toString()] = i[1].toString();
		});
		TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=" + queries['id'],"/xsl/quotePrintHead.xsl?v=<%=version%>", function(result){
			$("#head").html(result);
			TransformXmlURL("/service/xml_getOrderProperty.asp?act=showHead&id=" + queries['id'],"/xsl/quotePrintTail.xsl?v=<%=version%>", function(result){
				$("#tail").html(result);
				
				var xslURL;
				if (queries['inShort']=="false")
					xslURL = "/xsl/quotePrintProperty.xsl?v=<%=version%>";
				else
					xslURL = "/xsl/quotePrintPropertyShort.xsl?v=<%=version%>";
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
