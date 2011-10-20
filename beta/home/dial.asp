<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Server.ScriptTimeout = 600
Response.Expires= -1
if (session("ID")="") then
	session.abandon
	response.redirect "../login.asp?err=براي ديدن اين صفحه بايد وارد سيستم شويد"
end if
%>
<!--#Include file='../config.asp' -->
<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
	<meta http-equiv="Content-Language" content="fa">
	<style>
		Table { font-size: 9pt;}
		Input { font-family:tahoma; font-size: 9pt;}
	</style>
	<TITLE>تلفن</TITLE>
	<script language="JavaScript">
	
	</SCRIPT>
</HEAD>
<BODY leftmargin=0 topmargin=0 bgcolor='#DDDDFF' style="direction:rtl;">

<%
' makeCall(telNumber,exten)
if IsNumeric(request("tel")) and IsNumeric(request("exten")) then
	theNumber = request("tel")
	exten = request("exten")
else
	response.write "خطايي رخ داده"
	response.end
end if

loginURL = "http://192.168.10.10:8088/pdhco/rawman?action=login&username=mypdhco&secret=66042700"
dialURL = "http://192.168.10.10:8088/pdhco/rawman?action=Originate&Channel=DAHDI/R1/" & theNumber & "&Exten=" & exten & "&Context=pdhco-cos4&Priority=1&Async=true" 
' Create the xml object 
Set loginConnection = CreateObject("Microsoft.XMLHTTP") 
' Conect to voip 4 login to AMI  
loginConnection.Open "get", loginURL, False 
loginConnection.Send 
'response.write loginConnection.responseText
set loginConnection = nothing
'response.write "<br>" & dialURL & "<br>"
set dialConnection = CreateObject("Microsoft.XMLHTTP") 
dialConnection.open "get", dialURL, False
dialConnection.Send 
'response.write dialConnection.responseText
set dialConnection = nothing
'ResponsePage = GetConnection.responseText
'response.write dialURL & "<br>"
%>
	<span>داخلي </span>
	<span><%=exten%></span>
	<span> شماره </span>
	<span><%=theNumber%></span>
	<span> شماره گيري مي‌كند.</span>
</BODY>
</HTML>
