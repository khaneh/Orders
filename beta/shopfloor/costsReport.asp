<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
	PageTitle="����� ������"
	SubmenuItem=9
%>
<!--#include file="top.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
if not Auth(3 , "A") then NotAllowdToViewThisPage()
%>
<style >
	span.delCostBtn{float: left;opacity: .6;display: none;cursor: pointer;position: absolute;top:3px;left: 0;}
	td.delCost{position: relative;}
</style>
<script type="text/javascript">
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
		$("#check").click(function(){
			var fromdate = escape($("#dateFrom").val());
			var todate = escape($("#dateTo").val());
			var url = "/service/xml_getCosts.asp?act=costCenter&id=" + $("#costCenters").val()+"&fromdate="+fromdate+"&todate="+todate;
			if ($("#insertDate").is(":checked"))
				url += "&insertdate=1"
			TransformXmlURL(url,"/xsl.<%=version%>/costListShow.xsl", function(result){
				$("#report").html(result);
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
		});
	});
</script>
<div>
	<span>����</span>
	<select id="costCenters" class="btn">
		<option value="-1">������ ����</option>
<%
	set rs=Conn.Execute("select * from cost_centers")
	while Not rs.EOF
		Response.write "<option value='" & rs("id") & "'>" & rs("name") & "</option>"
		rs.MoveNext
	wend
%>		
	</select>
	<span>�����</span>
	<input type="text" dir="ltr" class="date boot" size="10" id="dateFrom" value="<%=shamsitoday()%>"/>
	<span>���</span>
	<input type="text" dir="ltr" class="date boot" size="10" id="dateTo" value="<%=shamsitoday()%>"/>
	<input type="button" class="btn" value="�����" id="check"/>
	<br>
	<input type="checkbox" id="insertDate" checked="checked"/>
	<span>����� ��� �� ����� Ȑ�� <small class="comment">(�� ����� �� ��� ���� ���������� �� ����� ������ ����� ���� ������� �� � ����� ���� ���� ����� ����� ���)</small></span>
</div>
<div id="report"></div>
<div id="costDelComfirm">
	<h3>��� ����� ����� �� ������ ��� ��� ��Ͽ</h3>
	<input type="hidden" id="costDelItemID" />
	<lu id="costDelComfirmList"></lu>
</div>