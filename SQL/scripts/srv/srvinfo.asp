<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myArrInfo, myStrTree
	openConnection
	myStrTree = getTreeStr("2", Array())
	myArrInfo = getSrvInfo()
	closeConnection
%>

<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Server Information">
		<TR><TD CLASS="caption" COLSPAN=2><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/server.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Server"> <% = myTObj.getTerm(110) & " " & myArrInfo(0, 0) %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(111) %> :</TD><TD><%  = myArrInfo(0, 0) %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(112) %> :</TD><TD><%  = Replace(myArrInfo(1, 0), vbLf, "<BR>") %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(113) %> :</TD><TD><%  = myArrInfo(2, 0) %></TD></TR>
	</TABLE>

	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
		<!--
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M2'); 
		//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->