<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " ê“«—‘ ›⁄«·Ì  —Ê“«‰Â ﬂ«—»—«‰"
SubmenuItem=8
if not Auth(0 , 9) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
div.allMessage{color: #3AC;}
div.allQuote{color: #77F;}
div.allOrder{color: #aa0;}
div.msg{display: none;clear: both;font-weight: 100;cursor: pointer; }
div.date{float: right;padding: 5px 10px 5px 10px;text-align: center;border-bottom: 2px solid #CCC;white-space: nowrap;overflow: scroll;max-width: 150px;}
div.body{float: right;padding: 5px 10px 5px 10px;border-bottom: 2px solid #CCC;color: black;overflow: scroll;max-width: 570px;white-space: nowrap;}
div.empty{clear: both;}
span.toName{color:red;}
span.typeName{color: blue;font-weight: bold;}
div.user{text-align: center;font-weight: bold;padding: 3px 0 3px 0;color: yellow;margin: 8px 0 2px 0;display: none;}
div.userName{background-color: black;}
div.quote{display: none;clear: both;font-weight: 100;cursor: pointer;}
div.order{display: none;clear: both;font-weight: 100;cursor: pointer;}
</style>
<form method="post">
	<input type="hidden" name="msgID" id='msgID' value="0">
	<input type="hidden" name="quoteID" id='quoteID' value="0">
	<input type="hidden" name="orderID" id='orderID' value="0">
	<input type="text" name="date" id='date' dir="ltr" value="<% if request("date")<>"" then response.write request("date") else response.write shamsiToday() end if%>" size="10" maxlength="10">
	<input type="hidden" name="today" id='today' value="<%=shamsiToday()%>">
	<input type="submit" name="submit" value="‰„«Ì‘">
	<div id='result'></div>
</form>
<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
<script type="text/javascript" src="/js/jalaliCalendar.js"></script>
<script type="text/javascript" src="/js/jquery.dateFormat-1.0.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$.ajaxSetup({
		cache: false
	});
	$("input#date").blur(function(){
		acceptDate($(this));
	});
	var loadUrl="activityLog_ajax.asp";
	$.getJSON(loadUrl,
		{act:'message',id:$("input#msgID").val(),date:$("#date").val()},
		function(json){
			echoMsg(json);
		}
	);
	$.getJSON(loadUrl,
		{act:'quote',id:$("input#quoteID").val(),date:$("#date").val()},
		function(json){
			echoQuote(json);
		}
	);
	$.getJSON(loadUrl,
		{act:'order',id:$("input#orderID").val(),date:$("#date").val()},
		function(json){
			echoOrder(json);
		}
	);
	if ($("input#date").val()==$("input#today").val()){
		setInterval(function() {
			$.getJSON(loadUrl,
				{act:'message',id:$("input#msgID").val(),date:$("#date").val()},
				function(json){
					echoMsg(json);
				}
			);
			$.getJSON(loadUrl,
				{act:'quote',id:$("input#quoteID").val(),date:$("#date").val()},
				function(json){
					echoQuote(json);
				}
			);
			$.getJSON(loadUrl,
				{act:'order',id:$("input#orderID").val(),date:$("#date").val()},
				function(json){
					echoOrder(json);
				}
			);
		}, 30000);	
	}
function echoMsg(e){
	$.each(e,function(i,msg){
		if ($("div#user-" + msg.from).is(":hidden"))
			$("div#user-" + msg.from).show();
		$("div#message-" + msg.from).children("div:first").before(
			"<div class='msg' onClick='openMsg(" + msg.id + ");' title='»Â " + msg.toName + "'>" +
			"<div class='date'><span class='typeName'>" + msg.typeName + "</span> ”«⁄  " + msg.time.substring(0,5) +
			"</div><div class='body'>" + msg.body + "</div></div>"
		);
		$("div.msg:hidden").slideDown("slow");
		if ($("input#msgID").val()<msg.id)
			$("input#msgID").val(msg.id);
	});
}	
function echoQuote(e){
	$.each(e,function(i,msg){
		if ($("div#user-" + msg.createdBy).is(":hidden"))
			$("div#user-" + msg.createdBy).show();
		$("div#quote-" + msg.createdBy).children("div:first").before(
			"<div class='quote' onClick='openQuote(" + msg.id + ");' title='ÃÂ  " + msg.customer + "(" + msg.company + 
			")'><div class='date'><span class='typeName'>«” ⁄·«„ " +msg.kind + "</span> œ— ”«⁄  " + msg.time + "</div><div class='body' title='" + 
			msg.note + "'>" +msg.title+" »Â  ⁄œ«œ " + msg.qtty+ " œ— ”«Ì“ " + msg.size + " Ê ﬁÌ„  " + msg.price + "</div></div>"
		);
		$("div.quote:hidden").slideDown("slow");
		if ($("input#quoteID").val()<msg.id)
			$("input#quoteID").val(msg.id);
	});
}
function echoOrder(e){
	$.each(e,function(i,msg){
		if ($("div#user-" + msg.createdBy).is(":hidden"))
			$("div#user-" + msg.createdBy).show();
		$("div#order-" + msg.createdBy).children("div:first").before(
			"<div class='order' onClick='openOrder(" + msg.id + ");' title='ÃÂ  " + msg.customer + "(" + msg.company + 
			")'><div class='date'><span class='typeName'>”›«—‘ " +msg.kind + "</span> œ— ”«⁄  " + msg.time + "</div><div class='body'>" +msg.title+" »Â  ⁄œ«œ " +
			msg.qtty+ " œ— ”«Ì“ " + msg.size + " Ê ﬁÌ„  " + msg.price + "</div></div>"
		);
		$("div.order:hidden").slideDown("slow");
		if ($("input#orderID").val()<msg.id)
			$("input#orderID").val(msg.id);
	});
}
});
function acceptDate(obj){ 
	if (obj.val()=="") {
		$("#result").html("·ÿ›«  «—ÌŒ —« Ê«—œ ﬂ‰Ìœ");
		obj.focus();
	}
	else if (obj.val()=="//") {
		var today = new Date();
		obj.val($.format.date(today,"yyyy/MM/dd"));
	} else {
		//var myRegExp = new RegExp("^(13)?[7-9][0-9]/[0-1]?[0-9]/[0-3]?[0-9]$");
		var rege=/^(13)?[7-9][0-9]\/[0-1]?[0-9]\/[0-3]?[0-9]$/;

		if( rege.test(obj.val()) ) {
			var SP = obj.val().split("/");
			if (SP[0].length == 2) SP[0] = "13" + SP[0] ;
			if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
			if (SP[2].length == 1) SP[2] = "0"  + SP[2] ;
			obj.val(SP.join("/"));
			if (obj.attr('id')=="startDate"){
				$("#start_time").val($.jalaliCalendar.jalaliToGregorianStr($("#startDate").val()) + " " + $("#startTime").val());
				checkBut(obj);
			} else if (obj.attr('id')=="endDate"){
				$("#end_time").val($.jalaliCalendar.jalaliToGregorianStr($("#endDate").val()) + " " + $("#endTime").val());
				checkBut(obj);
			}
		}
		if(!rege.test(obj.val())||( SP[0]<'1376' || SP[1]>'12' || SP[2]>'31' )) {
			$("#result").html("›—„   «—ÌŒ »«Ìœ YYYY/MM/DD »«‘œ.");
			//obj.select();
			obj.focus();
		};
	}
	
}
function openOrder(id){
	window.open('/beta/order/TraceOrder.asp?act=show&order=' + id, '_blank');
}
function openQuote(id){
	window.open('/beta/order/Inquiry.asp?act=show&quote=' + id, '_blank');
}
function openMsg(id){
	window.open('message.asp?act=show&id=' + id, '_blank');
}

</script>
<%
set rs=Conn.Execute("select * from users where display=1 order by realName")
while not rs.eof 
%>
	<div id='user-<%=rs("id")%>' class='user'>
		<div class="userName"><%=rs("realName")%></div>
		<div id='message-<%=rs("id")%>' class='allMessage'><div class="empty"></div></div>
		<div id='quote-<%=rs("id")%>' class='allQuote'><div class="empty"></div></div>
		<div id='order-<%=rs("id")%>' class='allOrder'><div class="empty"></div></div>
	</div>
<%	
	rs.moveNext
wend

%>
<!--#include file="tah.asp" -->