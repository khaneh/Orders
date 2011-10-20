<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= "ÏæÎÊä"
SubmenuItem=6
if not Auth(7 , 6) then NotAllowdToViewThisPage()

fixSys = "AP"
SA_IsVendor = 1
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include file="../AO/include_ItemsRelation.asp" -->

<!--#include file="tah.asp" -->