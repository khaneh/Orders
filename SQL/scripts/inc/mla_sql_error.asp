<%
	Sub mla_displayError(pErr, pStr)
		Dim myStrHTML
		closeConnection()
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P>&nbsp;</P>
	<FORM METHOD=POST ACTION=#>
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" WIDTH=50% ALIGN=CENTER SUMMARY="Error Information">
		<TR><TD CLASS="caption" COLSPAN=2><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/error.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Error">&nbsp;<% = myTObj.getTerm(66) & " " & pErr.Number %></TD></TR>
		<TR>
			<TD><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/obj32/error.gif" WIDTH="32" HEIGHT="32" BORDER=0 ALT="Error"></TD>
			<TD CLASS="forminfo"><B><% = myTObj.getTerm(66) & " " & pErr.Number %></B></TD>
		</TR>
		<TR>
			<TD>&nbsp;</TD>
			<TD>
				<P><% = pErr.Description %></P>
				<P><% = txt2html(pStr) %></P>
			</TD>
		</TR>
		<TR><TD COLSPAN=2>&nbsp;</TD></TR>
		<TR><TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE="button" VALUE="<% = myTObj.getTerm(67) %>" onClick="history.go(-1);"></TD></TR>
	</TABLE>
	</FORM>
	<P>&nbsp;</P>
</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->
<%
		Response.End
	End Sub
%>