<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
response.buffer=true
PageTitle="ÒÇÑÔ ÍÓÇÈ"
SubmenuItem=6
if not Auth(6 , 6) then NotAllowdToViewThisPage()

fixSys = "AR"
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%if Auth(6 , 6) then %>
<!--#include File="../AO/include_AccountReport.asp"-->
<%else
	NotAllowdToViewThisPage()
end if%>

<!--#include file="tah.asp" -->