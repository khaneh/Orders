<%
response.buffer = true
'server.transfer "beta/login.asp"
response.redirect "beta"
function sqlSafe (s)
  st=s
  st=replace(St,"'","`")
  st=replace(St,chr(34),"`")
  sqlSafe=st
end function
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	Table { font-size: 10pt;}
</style>
<TITLE>Login </TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var tempKeyBuffer;
function myKeyDownHandler(){
	tempKeyBuffer=window.event.keyCode;
}
function myKeyPressHandler(){
//	alert (tempKeyBuffer)
	if (tempKeyBuffer>=65 && tempKeyBuffer<=90){
		window.event.keyCode=tempKeyBuffer+32;
	}
	else if(tempKeyBuffer==186){
		window.event.keyCode=59;
	}
	else if(tempKeyBuffer==188){
		window.event.keyCode=44;
	}
	else if(tempKeyBuffer==190){
		window.event.keyCode=46;
	}
	else if(tempKeyBuffer==191){
		window.event.keyCode=47;
	}
	else if(tempKeyBuffer==192){
		window.event.keyCode=96;
	}
	else if(tempKeyBuffer>=219 && tempKeyBuffer<=221){
		window.event.keyCode=tempKeyBuffer-128;
	}
	else if(tempKeyBuffer==222){
		window.event.keyCode=39;
	}
}
//-->
</SCRIPT>
</HEAD>

<BODY onLoad="document.all.UserName.focus();">

<font face="tahoma">
<FORM METHOD=POST ACTION="default.asp">
<div dir='rtl' align = "center" >
<!--IMG SRC="images/khaneh.jpg" WIDTH="350" HEIGHT="20" BORDER=0 ALT=""-->

<TABLE>
<TR>
	<TD colspan="2" align="center"></TD>
</TR>
<TR>
	<TD> نام كاربر </TD>
	<TD><INPUT TYPE="text" NAME="UserName" onkeyDown="return myKeyDownHandler();" onKeyPress="return myKeyPressHandler();"></TD>
</TR>
<TR>
	<TD> رمز عبور </TD>
	<TD><INPUT TYPE="password" NAME="Password" onkeyDown="return myKeyDownHandler();" onKeyPress="return myKeyPressHandler();"></TD>
</TR>
<TR>
	<TD></TD>
	<TD><INPUT style="font-family:tahoma; width:100%;" TYPE="submit" name="act" value="ورود"></TD>
</TR>
</TABLE>
 <br>
&nbsp; 
<br>
<hr>
</div>
</FORM>
<br>
<%
if request("act")="ورود" then
'	conStr="dsn=sefareshat; uid=sefadmin; pwd=5tgb; DATABASE=sefareshat"
	'	conStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("kid.mdb")
    conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"
    Set conn = Server.CreateObject("ADODB.Connection")
	conn.open conStr
'	conn.Execute ("DELETE FROM Temp_Moein")
    
    Set RS1 = conn.Execute("SELECT * FROM [tblCSR] WHERE [UserName]='" & sqlSafe(request("UserName")) & "' AND  [Password]='" & sqlSafe(request("Password")) & "' ")
    If (RS1.EOF) Then
		session.abandon
	%>
		<TABLE border='1' width='70%' align='center' cellpadding='10' >
		<TR>
			<td align='center' dir='rtl' bgcolor='#FF8888'>Login Incorrect&nbsp;</td>
		</TR>
		</TABLE>
	<%else
		session("ID")=RS1("CSRID")
		conn.Close
		response.redirect "search.asp"
    End If

	conn.Close
elseif request.querystring("err")<>"" then 
%>
		<TABLE border='1' width='70%' align='center' cellpadding='10' >
		<TR>
			<td align='center' dir='rtl' bgcolor='#FF8888'><%=request.querystring("err")%>&nbsp;</td>
		</TR>
		</TABLE>
<%
end if
%>
</font>
</BODY>
</HTML>
