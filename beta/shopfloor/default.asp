<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="Ê—Êœ „—«Õ·"
SubmenuItem=1
if not Auth(3 , 1) then NotAllowdToViewThisPage()
orderID=0
if IsNumeric(Request("orderID")) then orderID = CDbl(Request("orderID"))
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
	span.stName{padding: 0 5px;font-family: "b zar";font-weight: bold;font-size: 12px;color: #b04444}
	.orderColor {background-color: black;color: yellow;}
	.quoteColor {background-color: #559;color: yellow;}
	.quoteColor td a:link{color: yellow;}
	.quoteColor td a:visited{color: #47FF00;}
	.quoteColor td a:hover{color: red;}
	.orderColor td a:link{color: yellow;}
	.orderColor td a:visited{color: #47FF00;}
	.orderColor td a:hover{color: red;}
</style>

<div class="inPage">
	<span>‘„«—Â ”›«—‘: </span>
	<input id="orderID" size="6" class="boot" <%if orderID>0 then Response.write " value='" & orderID & "'"%>/>
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
		checkOrder();
	});
	if (parseInt($("#orderID").val())>0)
		checkOrder();
});
function checkOrder(){
	var orderID = Number($("#orderID").val());
	if (!isNaN(orderID) && orderID!='') {
		loadXMLDoc("/service/xml_getOrderProperty.asp?act=showHead&id=" + orderID, function(orderXML){
			
			var isOrder = $(orderXML).find("status isOrder").text();
			var isClosed = $(orderXML).find("status isClosed").text();
			var isApproved = $(orderXML).find("status isApproved").text();
			var step = $(orderXML).find("status step").text();
			if (isOrder=='0'){
				$("#message").html("<b>«” ⁄·«„</b> —Ê  €ÌÌ— „—Õ·Â ‰„ÌùœÌ„° ‘„« „ÊŸ› Â” Ìœ ﬂÂ  Ê·Ìœ —« „ Êﬁ› ﬂ‰Ìœ.  « «Ì‰ «” ⁄·«„ »Â ”›«—‘  »œÌ· ‘Êœ");
				$("#step").prop("disabled", true);
			} else
			if (isClosed!='0'){
				$("#message").html("”›«—‘ »” Â ‘œÂ!");
				$("#step").prop("disabled", true);
				$("#submit").prop("disabled", true);
			} else if (isApproved=='0'){
				$("#message").html("”›«—‘  «ÌÌœ ‰‘œÂ");
			} else if (step=='40'){
				$("#message").html("<B>”›«—‘ „ Êﬁ› ‘œÂ.</B> Å” ‰»«Ìœ »—«Ì «Ì‰ ”›«—‘ ﬂ«—Ì ﬂ—œ!");
				$("#step").prop("disabled", true);
				$("#submit").prop("disabled", true);
			} else {
				//console.log(step);
				$("#step").prop("disabled", false);
				$("#submit").prop("disabled", false);
				$("#step option[value=" + step + "]").prop("selected", true);
			}	
			
			TransformXml(orderXML, "/xsl.<%=version%>/orderShowHeader.xsl", function(result){
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
}
</script>

<!--#include file="tah.asp" -->
