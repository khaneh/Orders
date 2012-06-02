<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="Ê—Êœ ⁄„·ﬂ—œ"
SubmenuItem=7
if not Auth(3 , 7) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	li{margin: 8px 10px 0 0;}
</STYLE>
<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
<script type="text/javascript" src="/js/jalaliCalendar.js"></script>
<script type="text/javascript" src="/js/jquery.dateFormat-1.0.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$.ajaxSetup({
		cache: false
	});
	function dateDiff( str1, str2 ) {
	    var diff = Date( str2 ) - Date( str1 ); 
	    return isNaN( diff ) ? NaN : {
	        diff : diff,
	        ms : Math.floor( diff            % 1000 ),
	        s  : Math.floor( diff /     1000 %   60 ),
	        m  : Math.floor( diff /    60000 %   60 ),
	        h  : Math.floor( diff /  3600000 %   24 ),
	        d  : Math.floor( diff / 86400000        )
	    };
	}
	
	function checkBut(obj){
		var ret = true;
		if ($("#is_count").val()=="True"){
			if ($("#is_countinuous").val()=="True"){
				ret = ret && ($.isNumeric($("#startCounter").val()) && $.isNumeric($("#endCounter").val()) && $.isNumeric($("#order").val()))
			} else {
				ret = ret && ($.isNumeric($("#qtty").val()) && $.isNumeric($("#order").val()));
			}
		}
		if ($("#is_time").val()=="True"){
			if ($("#is_countinuous").val()=="True") {
				var lastTime = new Date($("#lastTime").val());
				var startTime = new Date($("#start_time").val());
				if (lastTime > startTime) {
					ret = false;
					$("#result").html("‘„« ﬁ»·« œ— «Ì‰ “„«‰ ›⁄«·Ì Ì Ê«—œ ﬂ—œÂù«Ìœ!");
					$("#startTime").val(lastTime.getHours()+':'+lastTime.getMinutes());
					$("#startDate").val($.format.date(lastTime,"yyyy/MM/dd"));
					$("#startDate").focus();
				}
			}
			if ($("#startDate").val()!="" && $("#startTime").val()!="" && $("#endDate").val()!="" && $("#endTime").val()!="") {
				var diff = dateDiff(Date($("#start_time").val()+":00"), Date($("#end_time").val()+":00"));
				if (diff.diff>0){
					var result = "⁄„·ﬂ—œ ‘„«: ";
					result += (diff.d>0)? diff.d + "—Ê“ ":"";
					result += (diff.h>0)? diff.h + "”«⁄  ":"";
					result += (diff.m>0)? diff.m + "œﬁÌﬁÂ ":"";
					if (obj.attr('id')!="order") 
						$("#result").html(result);
					console.log(obj.attr('id'));
					if ($("#order_found").val()=="0"){
						ret = false;
					} else {
						ret = ret && true;
					}
				} else {
					if (obj.attr('id')!="order") 
						$("#result").html(" «—ÌŒ Ê ”«⁄  Å«Ì«‰ »«Ìœ «“ ‘—Ê⁄ »“—ê — »«‘œ.");
					ret=false;
				}
			}
		}
		
		//console.log('checked!' + ret);
		if (ret)
			$("#save").prop("disabled", false);
		else
			$("#save").prop("disabled", true);
	}
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
	function acceptTime(obj) {
		if (obj.val()=="") {
			$("#result").html("·ÿ›« ”«⁄  —« Ê«—œ ﬂ‰Ìœ.");
			obj.focus();
		} else if (obj.val()==":"){
			var now = new Date();
			obj.val(now.getHours()+':'+now.getMinutes());
			//var myRegExp = new RegExp("^[0-2]?[0-9]:[0-5]?[0-9]$");
			var rege = /^[0-2]?[0-9]:[0-5]?[0-9]$/;
			if( rege.test(obj.val()) ) {
				var SP = obj.val().split(":");
				if (SP[0].length == 1) SP[0] = "0" + SP[0] ;
				if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
				obj.val(SP.join(":"));
	
				if (obj.attr('id')=="startTime"){
					$("#start_time").val($.jalaliCalendar.jalaliToGregorianStr($("#startDate").val()) + " " + $("#startTime").val());
					checkBut(obj);
				} else if (obj.attr('id')=="endTime"){
					$("#end_time").val($.jalaliCalendar.jalaliToGregorianStr($("#endDate").val()) + " " + $("#endTime").val());
					checkBut(obj);
				}
			}
		} else {
			//var myRegExp = new RegExp("^[0-2]?[0-9]:[0-5]?[0-9]$");
			var rege = /^[0-2]?[0-9]:[0-5]?[0-9]$/;
			if( rege.test(obj.val()) ) {
				var SP = obj.val().split(":");
				if (SP[0].length == 1) SP[0] = "0" + SP[0] ;
				if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
				obj.val(SP.join(":"));
	
				if (obj.attr('id')=="startTime"){
					$("#start_time").val($.jalaliCalendar.jalaliToGregorianStr($("#startDate").val()) + " " + $("#startTime").val());
					checkBut(obj);
				} else if (obj.attr('id')=="endTime"){
					$("#end_time").val($.jalaliCalendar.jalaliToGregorianStr($("#endDate").val()) + " " + $("#endTime").val());
					checkBut(obj);
				}
			}
			if(!rege.test(obj.val())||( SP[0]>23 || SP[1]>59)) {
				$("#result").html("›—„  ”«⁄  »«Ìœ HH:MM »«‘œ.");
				obj.focus();
			};
		}
	}
	if ($("#is_countinuous").val()=="True" && $("#is_count").val()=="True"){
		$("#result").html(ajax_load);
		$("#startCounter").ready(function(){
			$.getJSON(loadUrl,
				{act:"counter",operation_type:$("input[name=operation_type]:checked").val(),driver_id:$("#driver_id").val()},
				function(json){
					if (json.lastCounter>0) {
						$("#result").html("¬Œ—Ì‰ ﬂ‰ Ê— ÅÌœ« ‘œ");
						$("#startCounter").val(json.lastCounter+1);
						$("#endCounter").focus();
					} else{
						$("#result").html("”«»ﬁÂù«Ì «“ ¬Œ—Ì‰ ﬂ‰ Ê— ÅÌœ« ‰‘œ");
						$("#startCounter").val(json.lastCounter);
						$("#startCounter").focus();
					}
				}
			);
			
		});
	}  
	if ($("#is_countinuous").val()=="True" && $("#is_time").val()=="True"){
		$("#result").html(ajax_load);
		$("#start_time").ready(function(){
			$.getJSON(loadUrl,
				{act:"time",operation_type:$("input[name=operation_type]:checked").val(),driver_id:$("#driver_id").val()},
				function(json){
					if (json.foundLastTime=='0') {
						$("#result").html("‘„«  «ﬂ‰Ê‰ »—«Ì «Ì‰ „—ﬂ“ Â“Ì‰Â “„«‰Ì Ê«—œ ‰ﬂ—œÂù«Ìœ!");
						today =new Date();
						//$("#start_time").val($.jalaliCalendar.gregorianToJalaliStr('2012/4/17'));
						//$("#startDate").val($.jalaliCalendar.jalaliToGregorianStr('1391/1/29'));
						
						$("#startDate").val($.format.date(today,"yyyy/MM/dd"));
						$("#start_time").val($.jalaliCalendar.jalaliToGregorianStr($("#startDate").val()));
						$("#endDate").val($.format.date(today,"yyyy/MM/dd"));
						$("#end_time").val($.jalaliCalendar.jalaliToGregorianStr($("#endDate").val()));
						$("#startDate").focus();
						console.log('today');
					} else{
						if (json.isNewDay=='1'){
							$("#result").html("»Â ‰Ÿ— „Ì«œ ﬂÂ ‘„« —Ê“ ÃœÌœÌ —Ê ¬€«“ ﬂ—œÌœ° —Ê“ ŒÊ»Ì œ«‘ Â »«‘Ìœ.");	
							$("#startDate").focus();
							$("#lastTime").val(json.lastTime.replace(RegExp('-','g'),'/'));
							
						} else {
							$("#lastTime").val(json.lastTime.replace(RegExp('-','g'),'/'));
							var lastTime = new Date($("#lastTime").val());
							$("#startDate").val($.jalaliCalendar.gregorianToJalaliStr(json.lastTime));
							$("#startTime").val(lastTime.getHours()+':'+lastTime.getMinutes());
							$("#startTime").blur();
							$("#endDate").focus();
							$("#result").html("¬Œ—Ì‰ “„«‰ Ê«—œ ‘œÂ œ— «Ì‰ „—ﬂ“ Â“Ì‰Â œ— «—ÌŒ <b>" +$("#startDate").val() + "</b> Ê œ— ”«⁄  <b>" + $("#startTime").val() + "</b> „Ìù»«‘œ° Õ „« »«Ìœ «œ«„Â «Ì‰ “„«‰ —« Ê«—œ ﬂ‰Ìœ");
							
						}
						
					}
				}
			);
			$("#endCounter").focus();
		});
	}
	var loadUrl="cost_ajax_check.asp";
	var ajax_load = "<img src=\'/css/images/ajax-loader.gif\' width='20px' alt='Loading Ö'>";
	$("#startCounter").blur(function(){
		//$("#result").html(ajax_load).load(loadUrl,"act=counter");
		if ($.isNumeric($("#startCounter").val()) && parseFloat($("#startCounter").val())>0){
			$("#endCounter").focus();
		} else {
			$("#startCounter").val("");
			$("#startCounter").focus();
		}
	});
	$("#startDate").blur(function(){
		acceptDate($(this));
	});
	$("#endDate").blur(function(){
		acceptDate($(this));
	});
	$("#startTime").blur(function(){
		acceptTime($(this));
	});
	$("#endTime").blur(function(){
		acceptTime($(this));
	});
	$("#endCounter").blur(function(){
		if ($(this).val()!=''){
			if ($.isNumeric($(this).val())) {
				var endCounter = parseFloat($("#endCounter").val()); 
				var startCounter = parseFloat($("#startCounter").val());
				if (startCounter >= endCounter){
					$(this).val("");
					$("#result").html("ﬂ‰ Ê— Å«Ì«‰ »«Ìœ »“—ê — «“ ﬂ‰ Ê— ‘—Ê⁄ »«‘œ!");
					$(this).focus();
				} else {
					var myCounter = parseFloat($("#endCounter").val()) - parseFloat($("#startCounter").val()); 
					$("#result").html("⁄„·ﬂ—œ ‘„«: " + myCounter + " ﬂ·Ìﬂ");
					checkBut($(this));
				}
			} else {
				$(this).focus();
				$("#result").html("·ÿ›« ⁄œœ Ê«—œ ﬂ‰Ìœ!");
				$("#save").prop("disabled", true);
			}
		} else {
			$(this).focus();
		}
	});
	$("#order").blur(function(){
		$("#result").html(ajax_load);
		var id=$("#order").val();
		$.getJSON(loadUrl,
			{act:"order",orderID:id},
			function(json){
				if (json.order>0){
					var result = "”›«—‘: " + json.orderKind + "° " + "<b>" + json.orderTitle + "</b>" + "<br>";
					result += " Ê”ÿ: " + "<b>" + json.customerName + "</b>" + "<br>" ;
					result += "œ— „—Õ·Â: " + "<b>" + json.orderStep + "</b>";
					$("#result").html(result);
					$("#order_found").val("1");
					checkBut($("#order"));
				} else {
					$("#result").html("ç‰Ì‰ ”›«—‘Ì ÊÃÊœ ‰œ«œ° ·ÿ›« œﬁ  ﬂ‰Ìœ!");
					$("#order").val("");
					$("#order").focus();
					$("#save").prop("disabled", true);
					$("#order_found").val("0");
				}
			}
		);
	});
});
</script>
<br>
<%
if request("act")="" then 
	mySQL="select cost_centers.name as centerName,cost_drivers.name as driverName,cost_drivers.ID from cost_drivers inner join cost_centers on cost_drivers.cost_center_id=cost_centers.id inner join cost_user_relations on cost_user_relations.driver_id=cost_drivers.id where cost_user_relations.user_id=" & session("id")
	set rs=Conn.Execute(mySQL)
	while not rs.eof
	%>
	<li><a href='costs.asp?act=add&ID=<%=rs("id")%>'><%="<b>" & rs("centerName") &"</b> - "& rs("driverName") %></a><br>
	<%
		rs.MoveNext
	wend
	rs.close
elseif request("act")="add" then '-------------------------------- A D D
	today=shamsiToday()
	mySQL="select * from cost_drivers where id=" & request("id")
	set rs=Conn.Execute(mySQL)
	if rs.eof then 
		msg="ÌÂ « ›«ﬁ ⁄ÃÌ» —Œ œ«œÂ!"
		response.redirect "?errmsg=" & Server.URLEncode(msg)
	end if
	mySQL = "select * from cost_operation_type where driver_id=" & request("id")
	set rsOP=Conn.Execute(mySQL)
	if rs.eof then 
		msg="ÂÌç ›⁄«·Ì Ì »—«Ì «Ì‰ œ—«ÌÊ— À»  ‰‘œÂ"
		response.redirect "?errmsg=" & Server.URLEncode(msg)
	end if
	currTime=now()
	startTime=FormatDateTime(dateAdd("n",-30,currTime),4)
	endDate=FormatDateTime(currTime,4)
	%>
<a href="costs.asp">»«“ê‘ </a><br>
<form method="post" action="?act=insert">
<b><%=rs("Name")%></b>
<%
i=0
while not rsOP.eof
	i=i+1
%>
	<input type="radio" name="operation_type" value="<%=rsOP("id")%>" <%if i=1 then response.write "checked='checked'"%>>
	<span><%=rsOP("name")%></span>
<%
	rsOP.moveNext
wend
%>
<br><br>

	<input type="hidden" name="driver_id" value="<%=rs("id")%>" id='driver_id'>
	<input type="hidden" name="is_count" id='is_count' value="<%=rs("is_count")%>">
	<input type="hidden" name="is_time" id='is_time' value="<%=rs("is_time")%>">
	<input type="hidden" name="is_direct" id='is_direct' value="<%=rs("is_direct")%>">
	<input type="hidden" name="is_countinuous" id='is_countinuous' value="<%=rs("is_countinuous")%>">
	<input type="hidden" name="lastTime" id='lastTime'>
<%	if rs("is_count")="True" then 
		if rs("is_countinuous")="True" then 
	%>
	<span>ﬂ‰ Ê— ‘—Ê⁄:</span>
	<input type="text" name="startCounter" size="7" id='startCounter'>
	<span>ﬂ‰ Ê— Å«Ì«‰:</span>
	<input type="text" name="endCounter" size="7" id='endCounter'><br>
	<%	
		else
	%>
	<span>„›œ«— —« Ê«—œ ﬂ‰Ìœ:</span>
	<input type="text" name="qtty" size="3" id='qtty'><br>
	<%
		end if
	end if 
	if rs("is_time")="True" then 
	%>
	<input type="hidden" name="start_time" id='start_time'>
	<input type="hidden" name="end_time" id='end_time'>
	<span> «—ÌŒ ‘—Ê⁄:</span>
	<input type="text" name="startDate" id='startDate' size="10" maxlength="10" value="" >
	<span>”«⁄  ‘—Ê⁄:</span>
	<input type="text" name="startTime" id='startTime' size="5" value="">
	<span> «—ÌŒ Å«Ì«‰:</span>
	<input type="text" name="endDate" id='endDate' size="10" maxlength="10" value="">
	<span>”«⁄  Å«Ì«‰:</span>
	<input type="text" name="endTime" id='endTime' size="5" value=""><br>
	<%		
	end if
	if rs("is_direct")="True" then 
	%>
	<span>‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ:</span>
	<input type="text" name="order" id='order' size="6">
	<input type="hidden" name="order_found" id='order_found' value="0">
	<%
	end if
	%>
	<br>
	<span> Ê÷ÌÕ: </span>
	<input type="text" name="description">
	<input type="submit" value="–ŒÌ—Â" disabled="disabled" id='save'>
</form>
<div id='result'></div>
	<%
	rs.close
	rsOP.close
elseif request("act")="edit" then '-------------------------------- E D I T

elseif request("act")="insert" then '-------------------------------- I N S E R T
	startCounter = "null"
	endCounter = "null"
	startTime = "null"
	endTime = "null"
	order = "null"
	if request("startCounter")<>"" then startCounter=request("startCounter")
	if request("endCounter")<>"" then endCounter=request("endCounter")
	if request("start_time")<>"" then startTime= "'" & request("start_time") & "'"
	if request("end_time")<>"" then endTime= "'" & request("end_time") & "'"
	if request("order")<>"" then order=request("order")
	if request("qtty")<>"" then 
		startCounter = 0
		endCounter = CDbl(request("qtty"))
	end if
	operationType = request("operation_type")
	mySQL = "select * from cost_drivers where id=" & request("driver_id")
	set rs=Conn.Execute(mySQL)
	if not rs.eof then 
		rate = CDbl(rs("rate"))
		mySQL="INSERT INTO costs (operation_type, start_counter, end_counter, rate, start_time, end_time, [order], user_id, description) VALUES (" & operationType & ", " & startCounter & ", " & endCounter & ", " & rate & ", " & startTime & ", "& endTime &", " & order & ", " & session("id") & ", N'" & request("description") & "');"
		response.write(mySQL)
		Conn.Execute(mySQL)
		response.redirect "?msg="&Server.URLEncode("»Â ”·«„ Ì° ⁄„·ﬂ—œ Ê‰ —Ê Ê«—œ ﬂ—œÌœ")
	else
		response.redirect "?errmsg="&Server.URLEncode("ÌÂ Œÿ«Ì ‰«„„ﬂ‰ —Œ œ«œÂ° »Â ’„Ì„Ì «ÿ·«⁄ »œÌœ!!!")
	end if
	rs.close
	set rs=nothing
elseif request("act")="update" then '-------------------------------- U P D A T E
end if

%>


<!--#include file="tah.asp" -->