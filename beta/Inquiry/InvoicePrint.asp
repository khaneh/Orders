<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
response.buffer=true
PageTitle="چاپ پیش فاکتور"
SubmenuItem=6
if not Auth(6 , 6) then NotAllowdToViewThisPage()

fixSys = "AR"
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%if Auth(6 , 6) then 
ReportLogRow=request("r")
%>
<BR>
<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no >
</iframe>
<%else
	NotAllowdToViewThisPage()
end if%>

<!--#include file="tah.asp" -->