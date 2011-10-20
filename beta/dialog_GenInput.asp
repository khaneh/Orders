<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Server.ScriptTimeout = 600
Response.Expires= -1
if (session("ID")="") then
	session.abandon
	response.redirect "../login.asp?err=براي ديدن اين صفحه بايد وارد سيستم شويد"
end if

'conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"
%>
<!--#Include file='config.asp' -->
<%
'Set conn = Server.CreateObject("ADODB.Connection")
'conn.open conStr
%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	Table { font-size: 9pt;}
	Input { font-family:tahoma; font-size: 9pt;}
</style>
<TITLE>وارد كنيد</TITLE>
<script language="JavaScript">
<!--
var Arguments = new Array(2)

function documentKeyDown() {
	var theKey = event.keyCode;
	if (theKey == 27) { 
		window.close();
	}
	else if (theKey == 13) {
		if (document.getElementById('dlgSearchBox').value != ""){
			window.dialogArguments.value=document.getElementById('dlgSearchBox').value;
			window.close();
		}
	}
}

document.onkeydown = documentKeyDown;

function onMyLoad(){
	document.getElementById('dlgDispTxt').innerText=window.dialogArguments.value
	window.dialogArguments.value=""
	document.getElementById('dlgSearchBox').focus();
	document.getElementById('dlgSearchBox').select();
}

//-->
</SCRIPT>
</HEAD>
<BODY leftmargin=0 topmargin=0 bgcolor='#DDDDFF' onload="onMyLoad()">
<font face="tahoma">
	<TABLE border="0" width="100%" height="100%" cellspacing="0" cellpadding="0">
	<TR><TD valign="middle">
		<table border="2" cellspacing="0" cellpadding="10" align="center" valign="middle" dir="RTL" bordercolor='#224488' bgcolor='#C3C3FF'>
		<tr><td><TABLE>
		<TR>
			<TD Height='30'><b><span id='dlgDispTxt'>هيچ سوالي نيست</span><b></TD>
		</TR>
		<TR>
			<TD><input type="text" id="dlgSearchBox" style="font-family:tahoma;font-size:9pt;width:100%;" value=""></TD>
		</TR>
		</TABLE>
		</td></tr>
		</table>
	</TD></TR>
	</TABLE>
</font>
</BODY>
</HTML>
