<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree
	myStrTree = getTreeStr("1", Array())
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form" WIDTH=320>
		<TR><TD CLASS="caption"><% = myTObj.getTerm(2) %></TD></TR>
		<TR><TD>
			<% = myTObj.getTerm(70) %><BR>
			<UL>
			<LI><A CLASS="contentLink" HREF="../conn/dsnless.asp"><% = myTObj.getTerm(3) %></A></LI>
			<LI><A CLASS="contentLink" HREF="../conn/dsn.asp"><% = myTObj.getTerm(4) %></A></LI>
			<LI><A CLASS="contentLink" HREF="../../restart.asp" TARGET="_top"><% = myTObj.getTerm(32) %></A></LI>
			</UL>
		</TD></TR>
	</TABLE>

	<% If Session("isConnected") Then %>
	<P>
		&nbsp;<BR>
		<% = "<B>" & myTObj.getTerm(510) & " :</B><BR>" & Session("ConnStr") %>
	</P>
	<% End If %>

	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
		<!--
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M1'); 
		//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->