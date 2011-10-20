<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree
	openConnection
	myStrTree = getTreeStr("2_3", Array())
	closeConnection
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Connection Form" WIDTH=320>
		<TR><TD CLASS="caption"><% = myTObj.getTerm(19) %></TD></TR>
		<TR><TD>
			<% = myTObj.getTerm(70) %><BR>
			<UL>
			<% If mla_auth(11, 0) Then %><LI><A CLASS="contentLink" HREF="srvcon.asp"><% = myTObj.getTerm(230) %></A></LI><% End If %>
			<% If mla_auth(12, 0) Then %><LI><A CLASS="contentLink" HREF="srvrole.asp"><% = myTObj.getTerm(231) %></A></LI><% End If %>
			</UL>
		</TD></TR>
	</TABLE>

	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
		<!--
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M2_3'); 
		//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->