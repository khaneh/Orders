<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree
	openConnection
	myStrTree = getTreeStr("2_4", Array())
	closeConnection
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form" WIDTH=320>
		<TR><TD CLASS="caption"><% = myTObj.getTerm(30) %></TD></TR>
		<TR><TD>
			<% = myTObj.getTerm(70) %><BR>
			<UL>
			<LI><A CLASS="contentLink" HREF="query.asp"><% = myTObj.getTerm(31) %></A></LI>
			</UL>
		</TD></TR>
	</TABLE>

	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
		<!--
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M2_4'); 
		//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->