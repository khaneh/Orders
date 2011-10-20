<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="ÏæÎÊä"
SubmenuItem=7
if not Auth(6 , 7) then NotAllowdToViewThisPage()

fixSys = "AR"
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include file="../AO/include_ItemsRelation.asp" -->

<!--#include file="tah.asp" -->