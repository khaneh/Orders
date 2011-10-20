<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree, myConnStr
	myStrTree = getTreeStr("1_2", Array())

	If Request.Form("mla_conn_submit") <> "" Then
		' saves connection info into a cookie
		If Request.Form("mla_conn_cookie") <> "" Then
			Response.Cookies("ConnStrDsn")("Dsn") = Request.Form("mla_conn_dsn")
			Response.Cookies("ConnStrDsn")("User") = Request.Form("mla_conn_user")
			Response.Cookies("ConnStrDsn")("Password") = Request.Form("mla_conn_password")
			Response.Cookies("ConnStrDsn")("Level") = Request.Form("mla_user_level")
			Response.Cookies("ConnStrDsn").Expires = #1/1/2020#
		End If
		' creates connection string
		myConnStr = mla_getConnStr_dsn(Request.Form("mla_conn_dsn"), Request.Form("mla_conn_user"), Request.Form("mla_conn_password"))
		Session("isConnected") = True
		Session("ConnStr") = myConnStr
		Response.Redirect "../hlp/connected.asp?refresh=1"
		Response.End
	End If
%>

<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
<P CLASS="treeinfo"><% = myStrTree %></P>

<FORM NAME="mla_conn" METHOD=POST ACTION="dsn.asp">
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form">
		<TR><TD CLASS="caption" COLSPAN=4><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/server.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Connection"> <% = myTObj.getTerm(2) %></TD></TR>
		<TR>
			<TD CLASS="formlabel"><% = myTObj.getTerm(4) %> :</TD>
			<TD><INPUT TYPE="text" NAME="mla_conn_dsn" CLASS="alphanumeric" TITLE="<% = myTObj.getTerm(4) %>"></TD>
			<TD COLSPAN=2>&nbsp;</TD>
		</TR>
		<TR>
			<TD CLASS="formlabel"><% = myTObj.getTerm(106) %> :</TD>
			<TD COLSPAN=3><INPUT TYPE="text" NAME="mla_conn_user" CLASS="alphanumeric" TITLE="<% = myTObj.getTerm(106) %>"></TD>
		</TR>
		<TR>
			<TD CLASS="formlabel"><% = myTObj.getTerm(107) %> :</TD>
			<TD COLSPAN=3><INPUT TYPE="password" NAME="mla_conn_password" CLASS="alphanumeric" TITLE="<% = myTObj.getTerm(107) %>"></TD>
		</TR>
		<TR><TD COLSPAN=4>&nbsp;</TD></TR>
		<TR>
			<TD COLSPAN=3 CLASS="formlabel"><% = myTObj.getTerm(108) %> :</TD><TD><INPUT TYPE="checkbox" NAME="mla_conn_cookie"></TD>
		</TR>
		<TR><TD COLSPAN=4>&nbsp;</TD></TR>
		<TR>
			<TD COLSPAN=4 ALIGN=CENTER>
				<INPUT TYPE="reset" VALUE="<% = myTObj.getTerm(50) %>" NAME="mla_conn_reset"> &nbsp;
				<INPUT TYPE="submit" VALUE="<% = myTObj.getTerm(109) %>" NAME="mla_conn_submit">
			</TD>
		</TR>
	</TABLE>
</FORM>

<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
<!--
	document.mla_conn.mla_conn_dsn.value = "<% = remdquote(Request.Cookies("ConnStrDsn")("Dsn")) %>";
	document.mla_conn.mla_conn_user.value = "<% = remdquote(Request.Cookies("ConnStrDsn")("User")) %>";
	document.mla_conn.mla_conn_password.value = "<% = remdquote(Request.Cookies("ConnStrDsn")("Password")) %>";

	if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M1'); 
//-->
</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->