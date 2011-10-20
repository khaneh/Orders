<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
response.buffer=true
PageTitle="Ç íÔ ÝÇßÊæÑ"
SubmenuItem=6
if not Auth(6 , 6) then NotAllowdToViewThisPage()

fixSys = "AR"
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<%if Auth(6 , 6) then 
ReportLogRow=request("r")
%>
<CENTER>
	<INPUT TYPE="button" value=" Ç " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">
</CENTER>
<BR>
<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no >
</iframe>
<%else
	NotAllowdToViewThisPage()
end if%>

<!--#include file="tah.asp" -->
