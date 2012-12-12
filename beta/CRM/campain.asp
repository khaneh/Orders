<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%

PageTitle="ÏÔÈæÑÏ ßãíä"
SubmenuItem=4
if not Auth(1 , 9) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<script type="text/javascript">
	TransformXmlURL("/service/xml_campain.asp?act=all","/xsl.<%=version%>/campain.xsl", function(result){
		$("#campains").html(result);
		$(".exten").each(function(i,e){
			$.getJSON('/service/json_pbx.asp',{exten:$(e).html()},function(data){
				$(e).closest('tr').find('.totalCallied').html(data.total);
				$(e).closest('tr').find('.answerdCallid').html(data.answered);
			});
		});
		
	});
</script>
<div id="campains"></div>
<!--#include file="tah.asp" -->
