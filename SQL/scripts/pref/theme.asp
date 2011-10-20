<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree

	If Request.Form("mla_cfg_cancel") <> "" Then
		Response.Redirect "default.asp"
		Response.End
	End If

	If Request.Form("mla_cfg_submit") <> "" Then
		Response.Cookies("Option")("Theme") = Request.Form("mla_cfg_theme")
		Response.Cookies("Option").Expires = #1/1/2020#
		Response.Redirect "default.asp?refresh=1"
		Response.End
	End If

	myStrTree = getTreeStr("3_1", Array())
%>

<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>
	<FORM NAME="mla_cfg" METHOD=POST ACTION="theme.asp">
		<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Option Form">
			<TR><TD CLASS="caption" COLSPAN=2><% = myTObj.getTerm(22) & " \ " &  myTObj.getTerm(23) %></TD></TR>
			<TR>
				<TD CLASS="formlabel"><% = myTObj.getTerm(492) %>  :</TD>
				<TD><% = getListBox("mla_cfg_theme", mla_cfg_themes, 0, 1, "", mla_cfg_theme, "alphanumeric", "") %></TD>
			</TR>
			<TR><TD COLSPAN=2>&nbsp;</TD></TR>
			<TR>
				<TD COLSPAN=2 ALIGN=CENTER>
					<INPUT TYPE="submit" VALUE="<% = myTObj.getTerm(51) %>" NAME="mla_cfg_cancel" onClick="document.nocheck = true;"> &nbsp;
					<INPUT TYPE="submit" VALUE="<% = myTObj.getTerm(54) %>" NAME="mla_cfg_submit">
				</TD>
			</TR>
		</TABLE>
	</FORM>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->