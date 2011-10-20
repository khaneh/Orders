<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	If Session("isConnected") Then
		Response.Redirect "connected.asp"
		Response.End
	End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>

	<P ALIGN=CENTER>
	<IMG SRC="../../themes/mylittletools.gif" WIDTH="383" HEIGHT="50" BORDER=0 ALT="myLittleTools.net"><BR>
	© 2000-2003, Elian Chrebor, <A HREF="http://www.mylittletools.net" TARGET=_blank>myLittleTools.net</A> - All rights reserved<BR>
	&nbsp;<BR>
	</P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form" WIDTH=320>
		<TR><TD CLASS="caption" COLSPAN=2><% = myTObj.getTerm(72) %></TD></TR>
		<TR><TD CLASS="formlabel"><% = myTObj.getTerm(73) %> :</TD><TD>myLittleAdmin For SQL Server and MSDE</TD></TR>
		<TR><TD CLASS="formlabel"><% = myTObj.getTerm(74) %> :</TD><TD><% = mla_version_number %></TD></TR>
		<TR><TD CLASS="formlabel"><% = myTObj.getTerm(75) %> :</TD><TD><% = mla_release_number %></TD></TR>
	</TABLE>
	&nbsp;<BR>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form" WIDTH=320>
		<TR><TD CLASS="caption"><% = myTObj.getTerm(513) %></TD></TR>
		<TR><TD>
			<% = myTObj.getTerm(512) %><BR>
			<% = myTObj.getTerm(70) %><BR>
			<UL>
			<LI><A CLASS="contentLink" HREF="../conn/default.asp"><% = myTObj.getTerm(2) %></A></LI>
			<LI><A CLASS="contentLink" HREF="../pref/default.asp"><% = myTObj.getTerm(22) %></A></LI>
			</UL>
		</TD></TR>
	</TABLE>
	&nbsp;<BR>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form" WIDTH=320>
		<TR><TD CLASS="caption"><% = myTObj.getTerm(515) %></TD></TR>
		<TR><TD>
			<UL>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/new.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="New" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(53) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/property.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="Structure" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(153) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/rename.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="Rename" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(65) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/drop.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="Drop" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(63) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/edit.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="Edit" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(52) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/sql.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="Generate SQL" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(68) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/content.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="Content" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(152) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/search.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="Search" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(154) %></LI>			
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/xml.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="XML Export" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(450) %></LI>
			<LI><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/csv.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALT="CSV Export" ALIGN="MIDDLE">&nbsp; <% = myTObj.getTerm(451) %></LI>		
			</UL>
		</TD></TR>
	</TABLE>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->