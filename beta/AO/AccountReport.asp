<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'A0 (11 [=B])
PageTitle= "ÒÇÑÔ ÍÓÇÈ"
SubmenuItem=1
if not Auth("B" , 1) then NotAllowdToViewThisPage()

fixSys = "-"
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include File="include_AccountReport.asp"-->


<!--#include file="tah.asp" -->