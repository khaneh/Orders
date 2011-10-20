<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'A0 (11 [=B])
PageTitle= "ÏæÎÊä"
SubmenuItem=4
if not Auth("B" , 4) then NotAllowdToViewThisPage()

fixSys = "-"
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<!--#include file="include_ItemsRelation.asp" -->

<!--#include file="tah.asp" -->