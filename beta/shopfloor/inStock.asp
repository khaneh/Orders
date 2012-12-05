<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'shopfloor (3)
PageTitle="áíÓÊ ãæÌæÏí ßÇÛĞ ÇİÓÊ ÏÑ ÇäÈÇÑ"
SubmenuItem=8
if not Auth(3 , 8) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<style>
	.w{background-color: white;}
	.w td{padding: 5px;}
	.g{background-color: #CCC;}
	.g td{padding: 5px;}
	th{background: #dd5;font-size: 13px;}
	td {position: relative;}
	span.chgPriceBtn{float: left;opacity: .6;display: none;cursor: pointer;position: absolute;top:2px;left: 0;}
	#purchasePrice{direction: ltr;};
</style>
<script type="text/javascript">
$(document).ready(function(){
	$.ajaxSetup({
		cache: false
	});	
	$(".chgPrice").mouseover(function(event){
		$(this).find(".chgPriceBtn").css("display","block");
	});
	$(".chgPrice").mouseout(function(event){
		$(this).find(".chgPriceBtn").css("display","none");
	});
	$(".chgPriceBtn").click(function(){
		$("#paperType").val($(this).closest("tr").attr("typeID"));
		$("#chgPriceDlg h5").html("ÈåÇí ÎÑíÏ " + $($(this).closest("tr").find("td")[0]).html() + " ÑÇ æÇÑÏ ßäíÏ");
		$("#purchasePrice").val($(this).closest("tr").find(".paperPrice").html());
		$("#chgPriceDlg").dialog("open");
	});
	$("#purchasePrice").blur(function(){
		$("#calcPrice").html(echoNum(Math.ceil(getNum($(this).val()) * 1.25)));
	});
	$("#chgPriceDlg").dialog({
		autoOpen: false,
		buttons: {"ÊÇííÏ":function(){
			$.ajax({
				type:"POST",
				url:"/service/json_getInventory.asp",
				data:{act:"updatePaperPrice",
					paperType:$("#paperType").val(),
					price: Math.ceil(getNum($("#purchasePrice").val()) * 1.25),
					cost: Math.ceil(getNum($("#purchasePrice").val()) * 1.05),
					purchasePrice:getNum($("#purchasePrice").val())},
				dataType:"json"
			}).done(function (data){
				if (parseInt(data.price)>0)
					$("tr.chgPrice[typeid=" + $("#paperType").val() + "]").find("span.paperPrice").html(echoNum(data.price));
			});
			$(this).dialog("close");
		}},
		title: "æÑæÏ ŞíãÊ ÌÏíÏ"
	});
});
</script>
<div id="chgPriceDlg">
	<h5></h5>
	<input id="purchasePrice" type="text" size="10"/>
	<span id="calcPrice"></span>
	<input id="paperType" type="hidden"/>
	<div>* áØİÇ ŞíãÊ ÎÑíÏ ÑÇ æÇÑÏ ßäíÏ.</div>
</div>
<br>
<br>
<div class="rspan5">
	<table cellspacing="0">
	<tr>
		<th>ßÏ ÇäÈÇÑ</th>
		<th>ÚäæÇä</th>
		<th>ãæÌæÏí</th>
	</tr>
	<%
	set rs=Conn.Execute("select * from InventoryItems inner join InventoryItemCategoryRelations on InventoryItemCategoryRelations.Item_ID=InventoryItems.id where Cat_ID=1 and enabled=1 and qtty>0 and owner=-1 order by InventoryItems.name ")
	cls="w"
	while not rs.eof 
		if cls="w" then
			cls="g"
		else
			cls="w"
		end if
	%>
	<tr class="<%=cls%>">
		<td><%=rs("oldItemID")%></td>
		<td><%=rs("name")%></td>
		<td><%=Separate(rs("qtty"))%></td>
	</tr>
	<%
		rs.moveNext
	wend
	%>
	</table>
</div>
<div class="rspan4">
	<table cellspacing="0">
		<tr>
			<th>äæÚ ßÇÛĞ</th>
			<th>ÈåÇí İÑæÔ¡ <small>åÑßíáæ</small></th>
		</tr>
		<%
	set rs = conn.Execute("select * from orderTypes where id=2")
	set prop=server.createobject("MSXML2.DomDocument")
	prop.loadXML(rs("property"))
	rs.close
	set rs = Nothing
	set paper = prop.SelectNodes("//key[@name='paper-type']/option")
	for each paperType in paper
		if cls="w" then
			cls="g"
		else
			cls="w"
		end if
		Response.write "<tr class='" & cls & " chgPrice' typeID='"_ 
			& paperType.text & "'><td>"_ 
			& paperType.GetAttribute("label") & "</td><td><span class='paperPrice'>"_ 
			& Separate(paperType.GetAttribute("price")) & "</span>"
		if auth(3,"B") then response.write "<span class='chgPriceBtn label'>ÊÛííÑ</span>"
		Response.write "</td></tr>"
	next
	
		%>
	</table>	
</div>
<!--#include file="tah.asp" -->