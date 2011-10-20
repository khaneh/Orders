<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<HTML>
<HEAD>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<TITLE>Logout </TITLE>
</HEAD>

<BODY>
<font face="tahoma">
<br>
<%
		session.abandon
%>
<TABLE cellspacing=0 cellpadding=0 width=300 height=150 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR>
	<td align='center' dir='rtl' bgcolor='#88FF88'><H3>خروج از سيستم با موفقيت انجام شد</H3>
	
	<p align='center' ><A HREF="login.asp" style="font-size: 12pt;">ورود مجدد</A></p>
	</td>
</TR>
</TABLE>
<br>
</font>
</BODY>
</HTML>
