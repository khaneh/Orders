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
	Table { font-family:tahoma; font-size: 9pt;}
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
'			conStr="dsn=sefareshat; uid=sefadmin; pwd=5tgb; DATABASE=sefareshat"
'			conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
'constr="dsn=SQLSysDSN;uid=sefadmin;pwd=5tgb;DATABASE=sefareshat;APP=ASP Script"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=jame;Data Source=.\sqlexpress;PWD=5tgb;"
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.open conStr
			
			Set RS1 = conn.Execute("SELECT * FROM [Users] WHERE [UserName]='" & sqlSafe(request("UserName")) & "' AND  [Password]='" & sqlSafe(request("Password")) & "' ")

			If (RS1.EOF) Then
				session.abandon
				rs1.close
				set rs1=conn.execute("select * from [Users] where [UserName]='" & sqlSafe(request("UserName")) & "' AND  [oldPassword]='" & sqlSafe(request("Password")) & "' ")
				if rs1.eof then 
			%>
					<div align=center style='background-color: #FF8888;width:300' >كدكاربري يا كلمه عبور اشتباه است&nbsp;</div><br><br>
			<%
				else
			%>
					<div align=center style='background-color: #008888;width:300' >به دلايل امنيتي رمز عبور شما تغيير كرده، لطفا رمز جديد را وارد تماييد و در صورتي كه رمز نداريد با داخلي 31 تماس بگيريد&nbsp;</div><br><br>
			<%
				end if
				rs1.close
			else
				session("ID")=RS1("ID")
				session("CSRName") = RS1("RealName")
				session("Permission") = RS1("Permission")
				session("exten")= RS1("Extention")
				Set RS2 = conn.Execute("SELECT GLs.*, UserDefaults.[User] FROM GLs INNER JOIN UserDefaults ON GLs.ID = UserDefaults.WorkingGL WHERE (UserDefaults.[User] = '"& RS1("ID") & "') OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")
				session("VatRate")=RS2("Vat")
				session("OpenGL")=RS2("ID")
				session("FiscalYear")=RS2("FiscalYear")
				session("OpenGLName")=RS2("name")
				session("OpenGLStartDate")=RS2("StartDate")
				session("OpenGLEndDate")=RS2("EndDate")
				RS2.movenext
				session("differentGL") = False
				if not RS2.EOF then
					temp=RS2("id")
					if temp <> session("OpenGL") then
						session("differentGL") = True
					end if
				end if
				RS2.close

				RS1.close
				conn.Close
				
				' Added By kid 820910
				if session("ID")=16 OR session("ID")=17 then ' shahami = 16  dehghan = 17
					session.Timeout=240
				end if

				if request.cookies("OldURL")<>"" then
					aa = request.cookies("OldURL")
					response.cookies("OldURL") = ""
					'response.form = request.cookies("OldForm") 
					'response.redirect split(aa,"?")(0)
					response.redirect aa
				else
					response.redirect "default.asp"
				end if
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

		<FORM METHOD=POST ACTION="/beta/login.asp">
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
