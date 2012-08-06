<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="ÒÇÑÔ ÇÚáÇãíååÇ"
SubmenuItem=11
if not Auth("C" , 9) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {padding:5;border:1pt solid gray;}
	.RepTable tr:nth-child(2n) {background-color: #FFFFFF;}
	.RepTable tr:nth-child(2n+3) {background-color: #DDDDDD;}
	.RepTable tr:nth-child(1) {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTableTitle {background-color: #CCCCFF; text-align: center; font-weight:bold; height:50;}
	.RepTableHeader {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.RepTableFooter {background-color: #BBBBBB; direction: LTR; }
	.myBtn {text-align: center;}
	.voided {background-color: #fee !important;color: #999;}
</STYLE>
<BR>
<div class="myBtn">
	<select id="gl" class="btn">
<%
set rs=Conn.Execute("select * from gls")
while not rs.eof
	selected = ""
	if rs("isDefault") then selected=" selected='selected' "
	response.write "<option " & selected & " value='" & rs("id") & "'>" & rs("fiscalYear") & "</option>"
	rs.moveNext
wend
%>		 
	</select>
	<select id="type" class="btn">
<%
set rs=Conn.Execute("select * from ARMemoTypes where name<>'-'")
while not rs.eof
	selected = ""
	if cint(rs("id"))=2 then selected=" selected='selected' "
	response.write "<option " & selected & " value='" & rs("id") & "'>" & rs("name") & "</option>"
	rs.moveNext
wend
rs.close
set rs=nothing
%>	
	</select>
</div>
<div id="Memos"></div>

<script type="text/javascript">
$(document).ready(function(){
	showMemo();
	
	$("select").change(function(){
		showMemo();
	});
});
function showMemo(){
	TransformXmlURL("/service/xml_getMemo.asp?act=showType&gl="+$("#gl").val()+"&type="+$("#type").val(),"/xsl/showMemoList.xsl", function(result){
		$("#Memos").html(result);	
		var sum=0;
		$(".credit").each(function(){
			sum+=parseInt($(this).html());
			$(this).html(echoNum($(this).html()));
		});
		$("#sumCredit").html(echoNum(sum));
		sum=0;
		$(".debit").each(function(){
			sum+=parseInt($(this).html());
			$(this).html(echoNum($(this).html()));
		});
		$("#sumDebit").html(echoNum(sum));
		sum = 0;
		$("#Memos tr:not(.voided)").each(function(){
			if (sum!=0)
				$(this).children("td:first").html(sum);
			sum++;
		});
	});
}
function echoNum(str){
	var regex = /(-?[0-9]+)([0-9]{3})/;
	str = Math.floor(str);
    str += '';
    while (regex.test(str)) {
        str = str.replace(regex, '$1,$2');
    }
    //str += ' kr';
    return str;
}
</script>