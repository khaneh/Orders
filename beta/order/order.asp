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
	label.key-label {margin: 0px 3px 0 0px;white-space: nowrap;padding: 0 5px 0 5px;display: inline;vertical-align: middle;overflow: hidden;text-overflow: ellipsis;font-size: 11pt;font-family: "b compset"}
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
	.orderColor td a:link{color: yellow;}
	div.top1 {margin-top: 10px;margin-right: 20px;}
	.grayColor{background-color: #ccc;}
	
	div.topBtn {margin: 10px 0 5px 0; text-align: center;}
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
		font-size: 9pt;
		font-family: "B titr","B yekan";
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
</STYLE>
<%
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
		TransformXmlURL("/service/xml_getOrderTrace.asp?act=search&searchBox=<%=Server.URLEncode(searchBox)%>","/xsl/orderShowList.xsl?v=<%=version%>", function(result){
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
		});
	});
</script>
<hr>
<div id='traceResult'></div>
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
			<span class="col0"> «—ÌŒ ”›«—‘</span>
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
<div id='traceResult'></div>
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
		TransformXmlURL("/service/xml_getOrderTrace.asp?act=advanceSearch&<%=request.form%>","/xsl/orderShowList.xsl?v=<%=version%>", function(result){
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
%>
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
				document.location='?act=approve&id=<%=orderID%>';
				$(this).dialog("close");
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
	});
	
</script>
<%
	set rs = Conn.Execute("SELECT orders.id, orders.isClosed,orders.isApproved,orders.isOrder, isnull(invoices.Approved,0) as Approved,isnull(invoices.Issued,0) as Issued FROM orders left outer join InvoiceOrderRelations on InvoiceOrderRelations.[Order] = orders.id left outer join Invoices on Invoices.ID=InvoiceOrderRelations.Invoice WHERE orders.ID=" & orderID)
%>
<div id="costDelComfirm">
	<h3>¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ ⁄„·ﬂ—œ –Ì· Õ–› ‘Êœø</h3>
	<input type="hidden" id="costDelItemID" />
	<lu id="costDelComfirmList"></lu>
</div>
<div id="approveOrder">
	<h3> «ÌÌœ ”›«—‘ »Â „‰“·Â «„÷«¡ ¬‰ «“ Ã«‰» ‘„« „Ìù»«‘œ. Ê „⁄‰Ì ¬‰ «Ì‰ ŒÊ«Âœ »Êœ ﬂÂ ‘„« »‰œÂ«Ì –Ì· —« çﬂ ﬂ—œÂù«Ìœ</h3>
	<lu>
		<li>”›«—‘ »Â ’Ê—  ﬂ«„· À»  ‘œÂ «” </li>
		<li>ﬁ—«—œ«œ Ê  «—ÌŒ  ÕÊÌ·  »« „‘ —Ì çﬂ ‘œÂ «” </li>
		<li>«” ⁄·«„ »Â „‘ —Ì «—«∆Â ‘œÂ Ê «„÷«¡ ‘œÂ «” </li>
	</lu>
</div>
<div id="approveQuote">
	<h3> «ÌÌœ «” ⁄·«„ »Â „‰“·Â  «ÌÌœ ‘„« „Ìù»«‘œ Ê „⁄‰Ì ¬‰ «Ì‰ ŒÊ«Âœ »Êœ ﬂÂ ‘„« »‰œÂ«Ì –Ì· —« çﬂ ﬂ—œÂù«Ìœ</h3>
	<lu>
		<li>«” ⁄·«„ »Â ’Ê—  ﬂ«„· À»  ‘œÂ «” </li>
		<li>«” ⁄·«„ »—«Ì „‘ —Ì «—”«· ‘œÂ «” </li>
	</lu>
</div>
<input type="hidden" id='isOrder' value="<%=rs("isOrder")%>">
<input type="hidden" id='orderID' value="<%=rs("id")%>">
<div class="topBtn">
<%	
	
	if (not CBool(rs("isClosed")) and Not CBool(rs("approved")) and not CBool(rs("issued"))) then Response.write "<a class='btn' href='?act=edit&id=" & orderID & "'>ÊÌ—«Ì‘</a>"
	if (not CBool(rs("isOrder")) and CBool(rs("isApproved"))) then Response.write "<a class='btn btn-success' href='?act=convertToOrder&id=" & orderID & "'> »œÌ· »Â ”›«—‘</a>"
	if (not rs("isApproved")) then Response.write "<a class='btn btn-primary' href='#' id='approvedBtn'> «ÌÌœ</a>"
	if (not CBool(rs("isClosed")) and CBool(rs("isOrder"))) then Response.write "<a class='btn' href='?act=pause&id=" & orderID & "'> Êﬁ›</a>"
	if (not rs("isClosed") and rs("isOrder")) then Response.write "<a class='btn btn-danger' href='?act=cancel&id=" & orderID & "'>ﬂ‰”·</a>"
	if (not rs("isClosed") and not rs("isOrder")) then Response.write "<a class='btn btn-danger' href='?act=cancel&id=" & orderID & "'>—œ «” ⁄·«„</a>"
	if (rs("isApproved")) then
		set rs = Conn.Execute("select * from InvoiceOrderRelations where [order]=" & orderID)
		while not rs.eof
			response.write "<a class='btn btn-inverse' href='../AR/AccountReport.asp?act=showInvoice&invoice=" & rs("invoice") & "' title='‘„«—Â " & rs("invoice") & "'>’Ê— Õ”«»</a>"
			rs.MoveNext
		wend
	end if
%>
</div>
<div id="orderLogs"></div>
<div id="orderHeader"></div>
<div id="orderDetails"></div>
<div id="orderMessages"></div>
<div id="orderStock"></div>
<div id="orderPurchase"></div>
<div id="orderCosts"></div>
<%		
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
		response.redirect "?act=show&id=" & QuoteID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  «” ⁄·«„ À»  ‘œ")
		'-----------------------------------------------------------------------------------------------------
	case "submitEdit":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		Conn.Execute ("update orders set property=N'" & request.form("myXML") & "', productionDuration=" &cint(request.form("productionDuration"))& ", Notes=N'" &sqlSafe(request.form("Notes"))& "', returnDate=N'" &request.form("returnDateTime")& "', paperSize=N'" &sqlSafe(request.form("paperSize"))& "', OrderTitle=N'" &sqlSafe(request.form("OrderTitle"))& "', qtty=" &cdbl(request.form("qtty"))& ", price='" & sqlSafe(request.form("totalPrice")) & "',LastUpdatedDate=getDate(), LastUpdatedBy=" &session("ID")& ", Customer=" &sqlSafe(request.form("customerID"))& ",isApproved=0 where id="&orderID)
		conn.Execute("delete InvoiceLines where Invoice in (select invoice from InvoiceOrderRelations where [Order]=" & orderID & ")")
		response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  »Â —Ê“ ‘œ")
		'-----------------------------------------------------------------------------------------------------
	case "convertToOrder":	'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		set rs = Conn.Execute("select orders.*, accounts.arBalance, accounts.creditLimit from orders inner join accounts on orders.customer=accounts.id where orders.id=" & orderID)
		if rs("isOrder") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("ﬁ»·« »Â ”›«—‘  »œÌ· ‘œÂ!")
		elseif not rs("isApproved") then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("«» œ« «Ì‰ «” ⁄·«„ —«  «ÌÌœ ﬂ‰Ìœ!")
		elseif (cdbl(rs("arBalance"))+cdbl(rs("creditLimit")) < 0) then 
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&errmsg=" & Server.URLEncode("»œÂÌ «Ì‰ Õ”«» «“ „Ì“«‰ «⁄ »«— ¬‰ »Ì‘ — ‘œÂ°<br> ·ÿ›« »« ”—Å—”  ›—Ê‘ Â„«Â‰ê ﬂ‰Ìœ.<br><a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=" & CustomerID & "'>‰„«Ì‘ Õ”«»</a>")
		else
			Conn.Execute("update orders set isOrder=1,isApproved=0,returnDate=null where id=" & orderID)
			Conn.Close
			response.redirect "?act=show&id=" & orderID &"&msg=" & Server.URLEncode("»Â ”›«—‘  »œÌ· ‘œ!")
		end if
		'-----------------------------------------------------------------------------------------------------
	case "approve":			'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=cdbl(request("id"))
		set rs=Conn.Execute("select Orders.Customer, orders.isOrder, orders.isApproved, orders.property as data, orderTypes.property as keys from Orders inner join orderTypes on orders.Type=orderTypes.ID where orders.id=" & orderID & " and lastUpdatedDate>lastUpdate")
		if rs.eof then 
			response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b>ﬁÌ„ ùÂ« »—Ê“ ‘œÂ! </b><br>ﬁ»· «“  «ÌÌœ° ·ÿ›« ”›«—‘/«” ⁄·«„ —« »—Ê“ ﬂ‰Ìœ")
		elseif (rs("isApproved")) then 
			response.redirect "?act=show&id=" & orderID & "&errmsg=" & Server.URLEncode("ﬁ»·«  «ÌÌœ ‘œÂ!")
		else
			if (not rs("isOrder")) then 
				Conn.execute("update orders set isApproved=1,approvedDate=getDate() where id="& orderID)
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
							if hasStock="yes" then
								set tmp = row.selectNodes("./key[@name='" & groupName & "-stockName']")
								if tmp.length>0 then 
									stockName =  tmp(0).text
									set tmp = row.selectNodes("./key[@name='" & groupName & "-stockDesc']")
									if tmp.length>0 then
										stockDesc = tmp(0).text
									else
										stockDesc = ""
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
							dis = CDbl(Replace(tmp.text,"%",""))
							set tmp = row.selectSingleNode("./key[@name='" & groupName & "-over']")
							over = CDbl(Replace(tmp.text,"%",""))
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
							if row.selectNodes("./key[@name='" & groupName & "-size']").length>0 then 
								set tmp = row.selectSingleNode("./key[@name='" & groupName & "-size']")
								if InStr(tmp.text,"X")>0 then
									l = CDbl(split(tmp.text,"X")(0))
									w = CDbl(split(tmp.text,"X")(1))
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
	 						appQtty = qtty * sets
							if dis<>100 then 
								price = thePrice * (100/(100-dis))
							else
								if thePrice = 0 then 
									price = 0
								else
									price = thePrice * (100/over)
								end if
							end if
							discount = price * dis / 100
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
							mySQL = "INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, [Reverse], Vat, hasVat) VALUES (" & InvoiceID & ", " & item & ", N'" & left(desc,100) & "', " & l & ", " & w & ", " & qtty & ", " & sets & ", " & appQtty & ", " & price & ", " & discount & ", 0, " & vat & ", " & hasVat & ")"
							'Response.write mySQL
							Conn.Execute(mySQL)
							TotalPrice = TotalPrice + price
							TotalDiscount = TotalDiscount + discount
							TotalReverse = TotalReverse + 0
							TotalReceivable = TotalReceivable + thePrice + vat
							totalVat = totalVat + vat
							'-------------------- STOCK Insert
							if hasStock = "yes" then 
								if requestID="" then
									mySQL = "INSERT INTO InventoryItemRequests (orderID,ItemName,comment,qtty,createdBy,invoiceItem,rowID) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & ")"
									Conn.Execute(mySQL)
								else
									'----------- Check UPDATE or INSERT ?
									mySQL = "select * from InventoryItemRequests where orderID=" & orderID & " and invoiceItem = " & item & " and rowID = " & rowID & " order by ID desc"
	' 								Response.write mySQL
									set rsr = Conn.Execute(mySQL)
									if rsr.eof then 
										mySQL = "INSERT INTO InventoryItemRequests (orderID,ItemName,comment,qtty,createdBy,invoiceItem,rowID) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & ")"
										Conn.Execute(mySQL)
									else
										requestID = Replace(requestID,rsr("id") & ",","")
										if rsr("status")="new" and (rsr("qtty")<>appQtty or rsr("itemName")<>stockName or rsr("comment")<>stockDesc) then 
											'------- IF CHANGE UPDATE
											mySQL = "UPDATE InventoryItemRequests SET ItemName=N'" & stockName & "',comment=N'" & stockDesc & "',qtty=" & appQtty & ",createdBy=" & Session("ID") & " where id=" & rsr("id")
											Conn.Execute(mySQL)
										elseif rsr("status")="del" then 
											'------- IF delete Insert new request
											mySQL = "INSERT INTO InventoryItemRequests (orderID,ItemName,comment,qtty,createdBy,invoiceItem,rowID) VALUES (" & orderID & ",N'" & stockName & "',N'" & stockDesc & "'," & appQtty & "," & Session("ID") &"," & item & "," & rowID & ")"
										Conn.Execute(mySQL)
										elseif rsr("qtty")<>appQtty or rsr("itemName")<>stockName or rsr("comment")<>stockDesc then
											errMSG = errMSG & "»—«Ì "  & stockName & " ÕÊ«·Â ’«œ— ‘œÂ. Å”  €ÌÌ— œ— ¬‰ „„ﬂ‰ ‰Ì” <br>"
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
										if rsp("status")="new" and (rsp("qtty")<>appQtty or rsp("typeName")<>purchaseName or rsr("comment")<>purchaseDesc or rsp("typeID")<>purchaseTypeID or rsp("price")<>thePrice) then
											mySQL = "UPDATE PurchaseRequests SET typeName = N'" & purchaseName & "',comment = N'" & purchaseDesc & "', typeID = " & purchaseTypeID & ", qtty = " & appQtty & ", price=" & thePrice & " WHERE id=" & rsp("id")
											Conn.Execute(mySQL)
										elseif rsp("status")="del" then 
											mySQL="INSERT INTO PurchaseRequests (orderID,typeName,comment,typeID,qtty,createdBy,price,isService,rowID) VALUES (" & orderID & ",N'" & purchaseName & "',N'" & purchaseDesc & "'," & purchaseTypeID & "," & appQtty & "," & Session("ID") &"," & thePrice & ",1," & rowID & ")"
											Conn.Execute(mySQL)
										elseif rsp("qtty")<>appQtty or rsp("typeName")<>purchaseName or rsr("comment")<>purchaseDesc or rsp("typeID")<>purchaseTypeID then
											errMSG = errMSG & "»—«Ì "  & purchaseName & " ”›«—‘ Œ—Ìœ «ÌÃ«œ ‘œÂ. Å”  €ÌÌ— œ— ¬‰ „„ﬂ‰ ‰Ì” <br>"
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
					mySQL = "INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, [Reverse], Vat, hasVat) VALUES (" & InvoiceID & ", 39999, N' Œ›Ì› —‰œ ›«ﬂ Ê—', 0, 0, 0, 0, 0, 0, " & RFD & ", 0, 0, 0)"
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
								errMSG = errMSG & "»—«Ì "  & stockName & " ÕÊ«·Â ’«œ— ‘œÂ. Ê ‘„« ¬‰—« Õ–› ﬂ—œÂù«Ìœ!<br>"
							else
								Conn.Execute("update InventoryItemRequests set status='del' where id=" & id)
							end if
						end if
					Next
				end if
				if errMSG="" then 
					Conn.execute("update orders set isApproved=1,approvedDate=getDate() where id="& orderID)
					'response.redirect "?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b> «ÌÌœ ‘œ</b>")
					Response.write "<a href='?act=show&id=" & orderID & "&msg=" & Server.URLEncode("<b> «ÌÌœ ‘œ</b>")&"'>»—Ê »Â</a>"
				else
					Response.write "<a href='?act=show&id=" & orderID & "&errmsg=" & Server.URLEncode(errMSG)&"'>»—Ê »Â</a>"
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
<script type="text/javascript" src="calcOrder.js"></script>
<script type="text/javascript">
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
<!-- ê—› ‰ «” ⁄·«„ -->
<FORM METHOD=POST ACTION="?act=submitNew" onSubmit="return checkValidation();">
	<div id="orderHeader"></div>
	<div class="downBtn">
		<hr/>
		<INPUT TYPE="submit" id='Submit' Name="Submit" Value=" «ÌÌœ" class="btn" style="width:100px;">
		<INPUT TYPE="button" Name="Cancel" Value="«‰’—«›" style="width:100px;" class="btn" onClick="window.location='order.asp';" >
	</div>
</FORM>
<div id="orderOrgin"></div>
<div id='orderDetails'></div>
<%		'-----------------------------------------------------------------------------------------------------
	case "edit":		'---------------------------------------------------------------------------------
		'-----------------------------------------------------------------------------------------------------
		orderID=request("id")
%>
<script type="text/javascript" src="calcOrder.js"></script>
<script type="text/javascript">
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
	});
	TransformXmlURL("/service/xml_getOrderProperty.asp?act=editOrder&id=<%=orderID%>","/xsl/orderEditProperty.xsl?v=<%=version%>", function(result){
		$("#orderDetails").html(result);
		readyForm();	
	});
</script>
<input type="hidden" id="vatRate" value="<%=Session("vatRate")%>"/>
<br><br>
<br>
<!-- ÊÌ—«Ì‘ -->
<hr>
<FORM METHOD=POST ACTION="?act=submitEdit&id=<%=orderID%>" onSubmit="return checkValidation();">
	<div id="orderHeader"></div>
	<div class="downBtn">
		<hr/>
		<INPUT TYPE="submit" id='Submit' Name="Submit" Value=" «ÌÌœ" class="btn" style="width:100px;">
		<INPUT TYPE="button" Name="Cancel" Value="«‰’—«›" style="width:100px;" class="btn" onClick="window.location='order.asp?act=show&id=<%=orderID%>';" >
	</div>
</FORM>
<div id='orderDetails'></div>
<%			
end select
%>
