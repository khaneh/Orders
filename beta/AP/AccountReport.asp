<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= "ÒÇÑÔ ÍÓÇÈ"
SubmenuItem=5
if not Auth(7 , 5) then NotAllowdToViewThisPage()

fixSys = "AP"
SA_IsVendor = 1
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include File="../AO/include_AccountReport.asp"-->


<!--#include file="tah.asp" -->