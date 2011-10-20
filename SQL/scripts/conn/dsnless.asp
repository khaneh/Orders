<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->

<%
	Dim myStrTree, myConnStr
	myStrTree = getTreeStr("1_1", Array())

	If Request.Form("mla_conn_submit") <> "" Then
		' saves connection info into a cookie
		If Request.Form("mla_conn_cookie") <> "" Then
			Response.Cookies("ConnStr")("Provider") = Request.Form("mla_conn_provider")
			Response.Cookies("ConnStr")("Datasource") = Request.Form("mla_conn_datasource")
			Response.Cookies("ConnStr")("Portnumber") = Request.Form("mla_conn_portnumber")
			Response.Cookies("ConnStr")("Networklibrary") = Request.Form("mla_conn_networklibrary")
			Response.Cookies("ConnStr")("Initialcatalog") = Request.Form("mla_conn_initialcatalog")
			Response.Cookies("ConnStr")("Trusted") = Request.Form("mla_conn_trusted")
			Response.Cookies("ConnStr")("User") = Request.Form("mla_conn_user")
			Response.Cookies("ConnStr")("Password") = Request.Form("mla_conn_password")
			Response.Cookies("ConnStr").Expires = #1/1/2020#
		End If
		' creates connection string
		myConnStr = mla_getConnStr(Request.Form("mla_conn_provider"), Request.Form("mla_conn_datasource"), Request.Form("mla_conn_portnumber"), Request.Form("mla_conn_networklibrary"), Request.Form("mla_conn_initialcatalog"), Request.Form("mla_conn_trusted"), Request.Form("mla_conn_user"), Request.Form("mla_conn_password"))
		Session("isConnected") = True
		Session("ConnStr") = myConnStr
		Response.Redirect "../hlp/connected.asp?refresh=1"
		Response.End
	End If
%>

<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
<P CLASS="treeinfo"><% = myStrTree %></P>

<FORM NAME="mla_conn" METHOD=POST ACTION="dsnless.asp">
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form">
		<TR><TD CLASS="caption" COLSPAN=4><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/server.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Connection"> <% = myTObj.getTerm(2) %></TD></TR>
		<TR>
			<TD CLASS="formlabel"><% = myTObj.getTerm(100) %> :</TD>
			<TD><INPUT TYPE="text" NAME="mla_conn_datasource" CLASS="alphanumeric" TITLE="<% = myTObj.getTerm(100) %>"></TD>
			<TD CLASS="formlabel"><% = myTObj.getTerm(101) %> :</TD>
			<TD><INPUT TYPE="text" NAME="mla_conn_portnumber" CLASS="numeric" TITLE="<% = myTObj.getTerm(101) %>"></TD>
		</TR>
		<TR>
			<TD CLASS="formlabel"><% = myTObj.getTerm(102) %> :</TD>
			<TD COLSPAN=3><INPUT TYPE="text" NAME="mla_conn_initialcatalog" CLASS="alphanumeric" TITLE="<% = myTObj.getTerm(102) %>"></TD>
		</TR>
		<TR>
			<TD CLASS="formlabel"><% = myTObj.getTerm(103) %> :</TD>
			<TD COLSPAN=3>
				<INPUT TYPE="radio" NAME="mla_conn_provider" VALUE="SQLOLEDB" ID="mla_conn_provider_1" CHECKED><LABEL FOR="mla_conn_provider_1">SQLOLEDB</LABEL> &nbsp; 
				<INPUT TYPE="radio" NAME="mla_conn_provider" VALUE="ODBC" ID="mla_conn_provider_2"><LABEL FOR="mla_conn_provider_2">ODBC</LABEL>
			</TD>
		</TR>
		<TR>
			<TD CLASS="formlabel"><% = myTObj.getTerm(104) %> :</TD>
			<TD>
				<SELECT NAME="mla_conn_networklibrary">
					<OPTION VALUE="">Default</OPTION>
					<OPTION VALUE="dbnmpntw">Named Pipes</OPTION>
					<OPTION VALUE="dbmssocn">Winsock TCP/IP</OPTION>
					<OPTION VALUE="dbmsspxn">SPX/IPX</OPTION>
					<OPTION VALUE="dbmsvinn">Banyan Vines</OPTION>
					<OPTION VALUE="dbmsrpcn">Multi-Protocol (RPC)</OPTION>
				</SELECT>
			</TD>
			<TD CLASS="formlabel"><% = myTObj.getTerm(105) %> :</TD>
			<TD><INPUT TYPE="checkbox" NAME="mla_conn_trusted"></TD>
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
	document.mla_conn.mla_conn_datasource.value = "<% = remdquote(Request.Cookies("ConnStr")("Datasource")) %>";
	document.mla_conn.mla_conn_portnumber.value = "<% = remdquote(Request.Cookies("ConnStr")("Portnumber")) %>";
	setRadioValue(document.mla_conn.mla_conn_provider, "<% = Request.Cookies("ConnStr")("Provider") %>", "");
	setSelectValue(document.mla_conn.mla_conn_networklibrary, "<% = Request.Cookies("ConnStr")("Networklibrary") %>", "");
	document.mla_conn.mla_conn_initialcatalog.value = "<% = remdquote(Request.Cookies("ConnStr")("Initialcatalog")) %>";
	document.mla_conn.mla_conn_trusted.checked = <% If Request.Cookies("ConnStr")("Trusted") = "" Then Response.Write "false" Else Response.Write "true" End If%>;
	document.mla_conn.mla_conn_user.value = "<% = remdquote(Request.Cookies("ConnStr")("User")) %>";
	document.mla_conn.mla_conn_password.value = "<% = remdquote(Request.Cookies("ConnStr")("Password")) %>";

	if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M1'); 
//-->
</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->