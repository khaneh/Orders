<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Server.ScriptTimeout = 600
Response.Expires= -1
if (session("ID")="") then
	session.abandon
	response.redirect "../login.asp?err=»—«Ì œÌœ‰ «Ì‰ ’›ÕÂ »«Ìœ Ê«—œ ”Ì” „ ‘ÊÌœ"
end if
Function WaitFor(SecDelay,ShowMsg) 
    timeStart = Timer() 
    timeEnd = timeStart + SecDelay 
 
    Msg = "Timer started at " & timeStart & "<br>" 
    Msg = Msg & "Script will continue in " 
 
    i = SecDelay 
    Do While timeStart < timeEnd 
        If i = Int(timeEnd) - Int(timeStart) Then 
        Msg = Msg & i 
        If i <> 0 Then Msg = Msg & ", " 
        If ShowMsg = 1 Then Response.Write Msg 
%> 
 
<%         Response.Flush() %> 
 
<% 
        Msg = "" 
        i = i - 1 
        End if 
        timeStart = Timer() 
    Loop 
    Msg = "...<br>Slept for " & SecDelay & " seconds (" & _ 
        Timer() & ")" 
    If ShowMsg = 1 Then Response.Write Msg 
End Function 
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
	<TITLE> ·›‰</TITLE>
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
	response.write "Œÿ«ÌÌ —Œ œ«œÂ"
	response.end
end if

loginURL =	"http://192.168.10.10:8088/pdhco/rawman?action=login&username=mypdhco&secret=66042700"
logoffURL= 	"http://192.168.10.10:8088/pdhco/rawman?action=logoff"
dialURL = 	"http://192.168.10.10:8088/pdhco/rawman?action=Originate&Channel=DAHDI/R1/" & theNumber & "&Exten=" & exten & "&Context=pdhco-cos4&Priority=1&Async=true"
' Create the xml object 
'Set pbxConnection = CreateObject("Microsoft.XMLHTTP") 
Set pbxConnection = CreateObject("MSXML2.ServerXMLHTTP.4.0") 
' Conect to voip 4 login to AMI  
pbxConnection.open "GET", loginURL, False 
pbxConnection.Send ""
'response.write pbxConnection.responseText
'set loginConnection = nothing
'call WaitFor(5,0)
'response.write i & "<br>" & dialURL & "<br>"
'set dialConnection = CreateObject("Microsoft.XMLHTTP") 
pbxConnection.open "GET", dialURL, False
pbxConnection.Send ""
'response.write pbxConnection.responseText& "<br>"
'set dialConnection = nothing
'set logoffConnection = CreateObject("Microsoft.XMLHTTP") 
pbxConnection.open "GET", logoffURL, False
pbxConnection.Send ""
'response.write pbxConnection.responseText& "<br>"
set pbxConnection = nothing
'ResponsePage = GetConnection.responseText
'response.write dialURL & "<br>"
%>
	<span>œ«Œ·Ì </span>
	<span><%=exten%></span>
	<span> ‘„«—Â </span>
	<span><%=theNumber%></span>
	<span> ‘„«—Â êÌ—Ì „Ìùﬂ‰œ.</span>
</BODY>
</HTML>
