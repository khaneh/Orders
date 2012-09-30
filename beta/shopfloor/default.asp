<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="Ê—Êœ „—«Õ·"
SubmenuItem=1
if not Auth(3 , 1) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<div class="inPage">
	<span>‘„«—Â ”›«—‘: </span>
	<input id="orderID" size="6" class="boot"/>
	<select id="step" class="btn">
<%
	set RS=Conn.Execute ("SELECT * FROM OrderSteps WHERE (IsActive=1) order by ord")
	while not rs.eof
		Response.write "<OPTION value='" & RS("ID") & "'>" & RS("name") & "</option>"
		rs.MoveNext
	wend
%>		
	</select>
	<input id="submit" name="submit" type="button" value=" «ÌÌœ" class="btn"/>
	<div id="message" class="well well-small"></div>
</div>
<div id="orderHeader" class="well"></div>

<script type="text/javascript">
$(document).ready(function(){
	$("#submit").click(function(){
		$("#step").prop("disabled", true);
		$("#submit").prop("disabled", true);
		$.getJSON("/service/json_orderStatus.asp",
		{act:"set",orderID:Number($("#orderID").val()),step:$("#step option:selected").val()},
		function(json){
			if (json.status=='done')
				$("#message").html("Ê÷⁄Ì  ”›«—‘ »Â " + json.stepName + "  €ÌÌ— ÅÌœ« ﬂ—œ!");
			else
				$("#message").html("Œÿ«ÌÌ —Œ œ«œÂ! Ê÷⁄Ì  ”›«—‘  €ÌÌ— ÅÌœ« ‰ﬂ—œ!");
		});
		$("#step").prop("disabled", false);
		$("#submit").prop("disabled", false);
		
	});
	$("#orderID").blur(function(){
		var orderID = Number($("#orderID").val());
		if (!isNaN(orderID) && orderID!='') {
			loadXMLDoc("/service/xml_getOrderProperty.asp?act=showHead&id=" + orderID, function(orderXML){
				console.log(orderXML);
				var isOrder = $(orderXML).find("status isOrder").text();
				var isClosed = $(orderXML).find("status isClosed").text();
				var isApproved = $(orderXML).find("status isApproved").text();
				var step = $(orderXML).find("status step").text();
				/*
if (isOrder=='0'){
					$("#message").html("«” ⁄·«„ —Ê  €ÌÌ— „—Õ·Â ‰„ÌùœÌ„");
					$("#step").prop("disabled", true);
				} else
*/ if (isClosed!='0'){
					$("#message").html("”›«—‘ »” Â ‘œÂ!");
					$("#step").prop("disabled", true);
					$("#submit").prop("disabled", true);
				} else if (isApproved=='0'){
					$("#message").html("”›«—‘  «ÌÌœ ‰‘œÂ");
					//$("#step").prop("disabled", true);
					//$("#submit").prop("disabled", true);
				} else {
					console.log(step);
					$("#step").prop("disabled", false);
					$("#submit").prop("disabled", false);
					$("#step option[value=" + step + "]").prop("selected", true);
				}	
				
				TransformXml(orderXML, "/xsl/orderShowHeader.xsl", function(result){
					$("#orderHeader").html(result);	
					$('a#customerID').click(function(e){
						window.open('../CRM/AccountInfo.asp?act=show&selectedCustomer='+$('a#customerID').attr("myID"), 'showCustomer');
						e.preventDefault();
					});
				});
			});
		} else{
			$("#step").prop("disabled", true);
			$("#submit").prop("disabled", true);
		}
	});
});
</script>

<!--#include file="tah.asp" -->
