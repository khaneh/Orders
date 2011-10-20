<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
if (session("ID")="") then
	session.abandon
	response.redirect "/login.asp"
else
response.redirect "home"
end if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
open("home/default.asp","PDH","scrollbars=yes; menubar=no; width=770; height=580")
//-->
</SCRIPT>
<HTML>
<HEAD>
<TITLE> PDH Co </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	body { font-family: tahoma; font-size: 9pt;}
</style>
</HEAD>
<center>
<BR><BR><H1>سلام</H1>