<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Server.ScriptTimeout = 600
Response.Expires= -1
if (session("ID")="") then
	session.abandon
	response.redirect "../login.asp?err=براي ديدن اين صفحه بايد وارد سيستم شويد"
end if

'conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
'conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=jame;Data Source=.\sqlexpress;PWD=5tgb;"
'Set conn = Server.CreateObject("ADODB.Connection")
'conn.open conStr
%>
<!--#include File="../config.asp"-->
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
</script></HEAD>
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

		<!--- Iz inja By Alix - 82-07-16 -->
		<TABLE width=100%>
		<TR>
			<TD><A onclick="setSearchString('ديجيتال')">[ديجيتال]</A></TD>
			<TD><A onclick="setSearchString('فيلم')">[فيلم]</A></TD>
			<TD><A onclick="setSearchString('پيرامون')">[پيرامون]</A></TD>
			<TD><A onclick="setSearchString('بيروني')">[بيروني]</A></TD>
			<TD><A onclick="setSearchString('افست')">[افست]</A></TD>
			<TD><A onclick="setSearchString('اغذ')">[کاغذ]</A></TD>
		</TR>
		<TR>
			<TD><A onclick="setSearchString('چسب')">[چسب]</A></TD>
			<TD><A onclick="setSearchString('برش')">[برش]</A></TD>
			<TD><A onclick="setSearchString('لمينيت')">[لمينيت]</A></TD>
			<TD colspan=2><A onclick="setSearchString('لارج فرمت')">[لارج فرمت]</A></TD>
			<TD><!--A onclick="setSearchString('افست')">[افست]</A--></TD>
			<TD><!--A onclick="setSearchString('اغذ')">[کاغذ]</A--></TD>
		</TR>
		</TABLE>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			function setSearchString(st)
			{	
				document.getElementById('dlgSearchBox').value = st;
				window.dialogArguments.value= st;
				window.close();
			}
		//-->
		</SCRIPT>
		<!----- Ta inja -------------------->

		</td></tr>
		</table>
	</TD></TR>
	</TABLE>
</font>
</SCRIPT>
</BODY>
</HTML>
