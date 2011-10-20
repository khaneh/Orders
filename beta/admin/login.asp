<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<% 'Response.Addheader "WWW-Authenticate", "BASIC" %>
<%
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
<TABLE cellspacing=0 cellpadding=0 width=300 height=150 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR>
	<TD>
		<font face="tahoma">
		<%
		if request("act")="ورود" then
'			conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=jame;Data Source=.\sqlexpress;PWD=5tgb;"
			'	conStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("kid.mdb")
			
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.open conStr

			u = sqlSafe(request("UserName"))
			p = sqlSafe(request("Password"))
			Set RS1 = conn.Execute("SELECT * FROM [Users] WHERE [UserName]='" & u & "' AND  [Password]='" & p & "' ")
			If (RS1.EOF) or (u<>"samimi" and u<>"mohammad" and u<>"kid" and u<>"vazehi" and u<>"zamani") Then
				session.abandon
			%>
					<div align=center style='background-color: #FF8888;width:300' >كدكاربري يا كلمه عبور اشتباه است&nbsp;</div><br><br>
			<%else
				session("AID")=RS1("ID")
				session("ID")=RS1("ID")
				session("CSRName") = RS1("RealName")

				Set RS2 = conn.Execute("SELECT GLs.Name, GLs.ID, GLs.FiscalYear, UserDefaults.[User] FROM GLs INNER JOIN UserDefaults ON GLs.ID = UserDefaults.WorkingGL WHERE (UserDefaults.[User] = '"& RS1("ID") & "') OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")
				session("OpenGL")=RS2("id")
				session("FiscalYear")=RS2("FiscalYear")
				session("OpenGLName")=RS2("name")
				RS2.close

				RS1.close
				conn.Close
				
				response.redirect "default.asp"
			End If

			conn.Close
		elseif request.querystring("err")<>"" then 
		%>
					<div align=center style='background-color: #FF8888;width:300'><%=request.querystring("err")%>&nbsp;</div><br><br>
		<%
		end if
		%>
	</TD>
</TR>
<TR>
	<TD>

		<FORM METHOD=POST ACTION="login.asp">
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
		</div>
		</FORM>
		<br>
		</font>
	</TD>
</TR>
</TABLE>

</BODY>
</HTML>
