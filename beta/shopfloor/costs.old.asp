<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="Ê—Êœ ⁄„·ﬂ—œ"
SubmenuItem=7
if not Auth(3 , 6) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<link type="text/css" href="/css/jquery-ui-1.8.16.custom.css" rel="stylesheet" />

<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.8.16.custom.min.js"></script>
<!--script type="text/javascript" src="/js/jquery.ui.core.js"></script-->
<script type="text/javascript" src="/js/jquery.dateFormat-1.0.js"></script>
<script type="text/javascript" src="/js/jquery.ui.datepicker-cc.js"></script>
<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery.ui.datepicker-cc-ar.js"></script>
<script type="text/javascript" src="/js/jquery.ui.datepicker-cc-fa.js"></script>

<script type="text/javascript" src="/js/jquery-ui-timepicker-addon.js"></script>

<script type="text/javascript">
		//alert(new Date(Date.parse("2011/11/26 03:00 ")));
		//2011/11/26 03:00 
	$(document).ready(function() {
		$.ajaxSetup({
			cache: false
		});
		addFunctions2Row(0);
		$("button#addRow").click(function () {
			$('span#msg').html("");
			myRow = $("table#myTable tr:last").html();
			//checkRow();
			//myID = $("table#myTable tr:last").attr("id");
			myID = $("table#myTable tr:last").index();
			//myID = myIDtext.substr(myIDtext.indexOf("-n-"), myIDtext.indexOf("-", myIDtext.indexOf("-n-") + 4));
			disableOldRow(myID);
			newID= parseInt(myID)+1;
			newID = "-" + newID;
			myID = "-" + myID;
			newRow="<tr>" + $("table#myTable tr:last").html().replace(RegExp(myID, "g"), newID).replace(RegExp('class="hasDatepicker"', "g"), '') + "</tr>";
			//console.log(newRow);
			$("table#myTable tr:last").after(newRow);
			myID=$("table#myTable tr:last").index();
			addFunctions2Row(myID);
			$(this).prop('disabled', true);
			return false;
			
		});
		//-------------------------------- DISABLE OLD ROW ---------------------------
		function disableOldRow(rowID){
			//$("input#startDate-fa-"+rowID).readOnly=true;
			$("input#startDate-fa-" + rowID).prop("disabled", true);
	    	$("input#startDateTime-" + rowID).prop("disabled", true);
	    	$("input#startTime-" + rowID).prop("disabled", true);
	    	$("input#endDate-fa-" + rowID).prop("disabled", true);
	    	$("input#endDateTime-" + rowID).prop("disabled", true);
	    	$("input#endTime-" + rowID).prop("disabled", true);
	    	$("input#startCounter-" + rowID).prop("disabled", true);
	    	$("input#endCounter-" + rowID).prop("disabled", true);
			$("select#operationType-" + rowID).prop("disabled", true);
			$("select#costCenter-" + rowID).prop("disabled", true);
		}
		// --------------------------------- ADD EVENTS ------------------------------
		function addFunctions2Row(rowID) {
			//if (rowID>0)  disableOldRow(rowID-1)
			$("select#costCenter-" + rowID).prop("disabled", false);
			console.log("ID: "+ rowID);
			$("input#startDate-fa-" + rowID).datepicker({
				onSelect: function(dateText,inst) {
					$("input#endDate-fa-" + rowID).datepicker('option', 'minDate', new JalaliDate(inst['selectedYear'], inst['selectedMonth'], inst['selectedDay']));
					$("input#startDateTime-" + rowID).val($.format.date(new JalaliDate(inst['selectedYear'], inst['selectedMonth'], inst['selectedDay']).getGregorianDate(), "yyyy/MM/dd")+ ' ' + $('input#startTime-'+rowID).val());
					checkRow(rowID);
					//$('#startDate-0').val($.datepicker.formatDate("dd/MM/yyyy",'2009-12-18 10:54:50.546'));
	
				},
				dateFormat: "d MM yy"
			});
			
			$("input#endDate-fa-" + rowID).datepicker({
				dateFormat: "d MM yy",
				onSelect: function(dateText,inst) { 
					$("input#endDateTime-" + rowID).val($.format.date(new JalaliDate(inst['selectedYear'], inst['selectedMonth'], inst['selectedDay']).getGregorianDate(), "yyyy/MM/dd")+' '+$('input#endTime-' + rowID).val());
					checkRow(rowID);
				}
			});
		    $('input#startTime-' + rowID).timepicker({
		    	timeOnlyTitle: "“„«‰ ‘—Ê⁄",
		    	timeText: "“„«‰",
		    	hourText: "”«⁄ ",
		    	minuteText: "œﬁÌﬁÂ",
		    	currentText: "Õ«·«",
		    	closeText: "«‰Ã«„",
		    	stepMinute: 10,
		    	onOpen: function() {$(this.val(''))},
		    	onClose: function() {
		    		if ($('input#startDateTime-' + rowID).val().length > 10)
		    			$('input#startDateTime-' + rowID).val($('input#startDateTime-' + rowID).val().substring(0,10) + ' ' + $('input#startTime-' + rowID).val());
		    			checkRow(rowID);
		    	}
		    });
		    $('input#endTime-' + rowID).timepicker({
		    	timeOnlyTitle: "“„«‰ Å«Ì«‰",
		    	timeText: "“„«‰",
		    	hourText: "”«⁄ ",
		    	minuteText: "œﬁÌﬁÂ",
		    	currentText: "Õ«·«",
		    	closeText: "«‰Ã«„",
		    	stepMinute: 10,
		    	onClose: function() {
		    		if ($('input#endDateTime-' + rowID).val().length > 10)
		    			$('input#endDateTime-' + rowID).val($('input#endDateTime-' + rowID).val().substring(0,10) + ' ' + $('input#endTime-' + rowID).val());
		    			checkRow(rowID);
		    	}
		    }); 
		    $("select#operationType-" + rowID).ready(function () {
		    	//$(this).change();
				$("select#operationType-" + rowID).change(function () {
					console.log('operation change!');
					//checkRow();
				});
			});
		    //$('input#startDate-fa-' + rowID).change(function () {checkRow(rowID);});
		    //$('input#endDate-fa-' + rowID).change(function () {checkRow(rowID);});
		    //$('input#startTime-' + rowID).change(function () {checkRow(rowID);});
		    //$('input#endTime-' + rowID).change(function () {checkRow(rowID);});
		    $('input#order-' + rowID).change(function () {
		    	checkRow(rowID);
		    });
		    $('input#description-'+rowID).focus(function() {
		    	checkRow(rowID);
		    });
		    $("select#costCenter-" + rowID).change(function (){
		    	$("div#costType-" + rowID)
		    		.html("<img src='/css/images/ajax-loader.gif' width='20px' alt='Loading ...'>")
		    		.load("costCenterType.asp", {centerID: $("select#costCenter-" + rowID).val(), row: rowID});
		    });
		    $("select#costCenter-" + rowID).ready(function (){
		    	$("div#costType-" + rowID)
		    		.html("<img src='/css/images/ajax-loader.gif' width='20px' alt='Loading ...'>")
		    		.load("costCenterType.asp", {centerID: $("select#costCenter-" + rowID).val(),row: rowID});
		    });
		    $("input#isDirect-" + rowID).change(function (){
		    	$("input#order-" + rowID).prop("disabled", $("input#isDirect-" + rowID).val()=='False');
		    });
		    $("input#type-" + rowID).change(function (){
		    	$("input#startDate-fa-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='1' || $("input#type-" + rowID).val()=='3');
		    	$("input#startDateTime-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='1' || $("input#type-" + rowID).val()=='3');
		    	$("input#startTime-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='1' || $("input#type-" + rowID).val()=='3');
		    	$("input#endDate-fa-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='1');
		    	$("input#endDateTime-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='1');
		    	$("input#endTime-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='1');
		    	$("input#startCounter-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='2');
		    	$("input#endCounter-" + rowID).prop("disabled", $("input#type-" + rowID).val()=='2');
		    	
		    });
		    $("input#isCountiuous-" + rowID).change(function (){ 
		    	//----------------------------- is countinuos flag set ------------------------------
		    	if ($("input#isCountiuous-" + rowID).val()=="True") {
		    		//-------------------------- check type is count --------------------------------
		    		if ($("input#type-" + rowID).val()=='1') { 
		    			hasAny=false;
		    			$('input[id*="endCounter-"]').each(function (i){
				    		if ($("input#driverID-" + rowID).val()==$('input#' + $(this).attr('id').replace('endCounter', 'driverID')).val()) {
				    			if ($(this).attr('id').replace('endCounter-','') != rowID) hasAny=true;
				    		}});
				    	if (!hasAny) { 
				    		// ------------- if first row of this DRIVER -------------
				    		console.log("first row of this driver, temporary do nothing!");
				    		$('input#startDate-fa-'+rowID).val('');
				    		$('input#startTime-'+rowID).val('');
				    		$('input#startDateTime-'+rowID).val('');
				    		$('input#startDate-fa-'+rowID).prop('disabled', false);
				    		$('input#startTime-'+rowID).prop('disabled', false);
				    		$('input#startDateTime-'+rowID).prop('disabled', false);
				    	} else {
				    		var lastCounter=0;
				    		$('input[id*="endCounter-"]').each(function (i){
				    			if ($("input#driverID-" + rowID).val()==$('input#' + $(this).attr('id').replace('endCounter', 'driverID')).val()) {
				    				if ($(this).attr('id').replace('endCounter-','') < rowID) lastCounter = $(this).val();
				    				console.log('last: ' + i + '-' + lastCounter);
				    			} else {
				    				console.log('nothing!!');
				    			}
				    		});
				    		$('input#startCounter-'+rowID).val(lastCounter);
				    		$('input#startCounter-'+rowID).prop('disabled', true);
				    		$('input#startDate-fa-'+rowID).val('');
				    		$('input#startTime-'+rowID).val('');
				    		$('input#startDateTime-'+rowID).val('');
				    		
				    	}
				    }
				    //------------------------- check type is time --------------------------------------
				    if ($("input#type-" + rowID).val()=='2') { 
				    	console.log('this is time :)');
				    	hasAny=false;
		    			$('input[id*="endCounter-"]').each(function (i){
				    		if ($("input#driverID-" + rowID).val()==$('input#' + $(this).attr('id').replace('endCounter', 'driverID')).val()) {
				    			if ($(this).attr('id').replace('endCounter-','') != rowID) hasAny=true;
				    		}});
				    	if (!hasAny) { 
				    		// ------------- if first row of this DRIVER -------------
				    		console.log("first row of this driver, temporary do nothing!");
				    		$('input#startCounter-'+rowID).val('');
				    		$('input#startDate-fa-'+rowID).prop('disabled', false);
				    		$('input#startTime-'+rowID).prop('disabled', false);
				    		$('input#startDateTime-'+rowID).prop('disabled', false);
				    	} else {
				    		var lastDate = '';
				    		var lastTime = '';
				    		var lastDateTime=new Date();
				    		$('input[id*="endDateTime-"]').each(function (i){
				    			if ($("input#driverID-" + rowID).val()==$('input#' + $(this).attr('id').replace('endDateTime', 'driverID')).val()) {
				    				if ($(this).attr('id').replace('endDateTime-','') < rowID) {
				    					lastDateTime = $(this).val();
				    					lastDate =	$('input#' + $(this).attr('id').replace('endDateTime', 'endDate-fa')).val();
				    					lastTime =	$('input#' + $(this).attr('id').replace('endDateTime', 'endTime')).val();
				    					console.log($(this).attr('id'));
				    				}
				    				console.log('last: ' + i + '-' + lastDate);
				    				
				    			} else {
				    				console.log('nothing found!!');
				    			}
				    		});
				    		$('input#startDate-fa-'+rowID).val(lastDate);
				    		$('input#startTime-'+rowID).val(lastTime);
				    		$('input#startDateTime-'+rowID).val(lastDateTime);
				    		$('input#startCounter-'+rowID).val('');
				    		$('input#startDate-fa-'+rowID).prop('disabled', true);
				    		$('input#startTime-'+rowID).prop('disabled', true);
				    		$('input#startDateTime-'+rowID).prop('disabled', true);
				    	}
				    }
				};
		    });
		};
		
		function checkRow(rowID){
			var result = true;
			if ($('input[id$="-' + rowID + '"]').length=0) {
				result = false;
			}
			$('input[id$="-' + rowID + '"]').each(function (i) {
				if ($(this).attr('disabled')!='disabled') {
					if ($(this).val().length > 0) {
						result = result && true;
					} else {
						result = false;
					}
				} 
			
			});
			if ($("input#type-" + rowID).val()=='2')
				if (result) {
					var startDate = new Date(Date.parse($('input#startDateTime-' + rowID).val()));
					var endDate = new Date(Date.parse($('input#endDateTime-' + rowID).val()));
					console.log("start: " + startDate);
					console.log("endDate: " + endDate);
					if ((endDate - startDate)<=0) result=false;
					//$('span#msg').html("test! " + ((endDate - startDate)/3600000));
				}
			if (result) $('span#msg').html("·ÿ›« ﬁ»· «“ «Ì‰ﬂÂ Œÿ ÃœÌœ —« «÷«›Â ﬂ‰Ìœ° Ìﬂ »«— œÌê— çﬂ ﬂ‰Ìœ. çÊ‰ «Ì‰ Œÿ »⁄œ «“ «÷«›Â ‘œ‰ Œÿ »⁄œÌ €Ì— ﬁ«»· ÊÌ—«Ì‘ ŒÊ«Âœ »Êœ!");
			$("button#addRow").prop('disabled', !result);
			console.log('---------------------------------------');
		};
	});
</script>
<style>
	.ltrInput {direction: ltr;}
	
</style>
<%
if request("act")="" then 
%>
<form>
	<table id="myTable">
		<thead>
			<td colspan="3" align="center" bgcolor="#F77"><b>‘—Ê⁄</b></td>
			<td colspan="3" align="center" bgcolor="#7F7"><b>Å«Ì«‰</b></td>
			<td colspan="3" align="center" bgcolor="#AAF"><b> Œ’Ì’ Â“Ì‰Â</b></td>
			<td align="center" bgcolor="#77F"><b> Ê÷ÌÕ« </b></td>
		</thead>
		<thead>
			<td align="center" bgcolor="#F99"> «—ÌŒ</td>
			<td align="center" bgcolor="#FBB">”«⁄ </td>
			<td align="center" bgcolor="#F55">ﬂ‰ Ê—</td>
			<td align="center" bgcolor="#9F9"> «—ÌŒ</td>
			<td align="center" bgcolor="#BFB">”«⁄ </td>
			<td align="center" bgcolor="#5F5">ﬂ‰ Ê—</td>
			<td align="center" bgcolor="#99F">„—ﬂ“ Â“Ì‰Â</td>
			<td align="center" bgcolor="#BBF">‰Ê⁄ ⁄„·Ì« </td>
			<td align="center" bgcolor="#88F">‘„«—Â ”›«—‘</td>
			<td align="center" bgcolor="#77F"> Ê÷ÌÕ« </td>
		</thead>
		<tr>
			<td>
				<input name="startDate-fa-0" id='startDate-fa-0' type="text" size="10">
				<input name="startDateTime-0" id='startDateTime-0' type="hidden">
				<input name="driverID-0" id='driverID-0' type="hidden">
				<input name="type-0" id='type-0' type="hidden">
				<input name="isDirect-0" id='isDirect-0' type="hidden">
				<input name="isCountiuous-0" id='isCountiuous-0' type="hidden">
			</td>
			<td><input name="startTime-0" id='startTime-0' type="text" size="5"></td>
			<td><input name="startCounter-0" id='startCounter-0' type="text" size="6"></td>
			<td>
				<input name="endDate-fa-0" id='endDate-fa-0' type="text" size="10">
				<input name="endDateTime-0" id='endDateTime-0' type="hidden" size="10">
			</td>
			<td><input name="endTime-0" id='endTime-0' type="text" size="5"></td>
			<td><input name="endCounter-0" id='endCounter-0' type="text" size="6"></td>
			<td>
				<select name="costCenter-0" id="costCenter-0">
<%
mySQL="select distinct cost_centers.* from cost_drivers inner join cost_centers on cost_drivers.cost_center_id=cost_centers.id inner join cost_user_relations on cost_user_relations.driver_id=cost_drivers.id inner join cost_operation_type on cost_drivers.id=cost_operation_type.driver_id where cost_user_relations.user_id=" & session("id")
set rs=Conn.Execute(mySQL)
while not rs.eof
%>
					<option value="<%=rs("id")%>"><%=rs("name")%></option>
<%
	rs.MoveNext
wend
rs.close
%>
				</select>
			</td>
			<td><div id="CostType-0"></div></td>
			<td><input name="order-0" id='order-0' type="text" size="6"></td>
			<td><input name="description-0" type="text" size="20"></td>
		</tr>
	</table>
</form>

<button id='addRow' disabled="disabled">«÷«›Â</button>
<span id='msg'></span>
<%	
elseif request("act")="setDriver" then 
	
	set rs=Conn.Execute("select * from cost_drivers where id="&request.form("driver"))
	%>
<form method=post action='?act=addCost'>
	<span><%=rs("name")%></span>
	<input name='diverValue' type=text>
	<%
	if rs("has_order")="True" then response.write("<input name='order' type=text>")
	%>
	<span><%=rs("unitSize")%></span>
</form>
	<%
end if
%>
<!--#include file="tah.asp" -->
