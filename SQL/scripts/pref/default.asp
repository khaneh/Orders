<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree
	myStrTree = getTreeStr("3", Array())
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form" WIDTH=320>
		<TR><TD CLASS="caption" COLSPAN=4><% = myTObj.getTerm(22) %></TD></TR>
		<TR><TD>
			<% = myTObj.getTerm(70) %><BR>
			<UL>
			<% If mla_auth(14, 0) Then %><LI><A CLASS="contentLink" HREF="theme.asp"><% = myTObj.getTerm(23) %></A></LI><% End If %>
			<% If mla_auth(14, 1) Then %><LI><A CLASS="contentLink" HREF="language.asp"><% = myTObj.getTerm(24) %></A></LI><% End If %>
			<% If mla_auth(14, 2) Then %><LI><A CLASS="contentLink" HREF="display.asp"><% = myTObj.getTerm(25) %></A></LI><% End If %>
			</UL>
		</TD></TR>
	</TABLE>


	<% If Request.QueryString("Refresh") = "1" Then %>
	<P>&nbsp;</P>
	<H2><% = myTObj.getTerm(514)  %></H2>
	<% End If %>

	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
		<!--
		<% If Request.QueryString("refresh") <> "" Then %>
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].document.location = "../inc/tree_con.asp";
		<% Else %>
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M10'); 
		<% End If %>
		//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->