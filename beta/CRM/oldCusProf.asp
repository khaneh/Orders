<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
if (session("ID")="") then
	session.abandon
	response.redirect "default.asp?err=session expired"
end if

'conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"

Set conn = Server.CreateObject("ADODB.Connection")
conn.open conStr

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	Table { font-size: 9pt;}
</style>
<TITLE>اطلاعات مشتري</TITLE>
</HEAD>

<BODY>
<font face="tahoma">
<TABLE border=0 width="100%">
<TR>
	<TD align="left"><A HREF="logout.asp" style="font-size: 7pt;">Logout</A></TD>
	<TD align="right">&nbsp;&nbsp;</TD>
</TR>
</TABLE>
<BR>
<TABLE border="2" width="100%" cellpadding="0" cellspacing="0" bordercolor="#000000">
<tr><td>
<% OLD_ACID = request("acID") %>
<!--#include File="include_OldCusData.asp"-->
</td></tr>
</TABLE>
</font>
</BODY>
</HTML>
