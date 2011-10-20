<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrInfo, myArrFileInfo, myFileCount, i, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2), Array(myDbName))
	myArrInfo = getDbInfo(myDbName)
	myArrFileInfo = getDbFileInfo(myDbName)
	closeConnection
	If isArray(myArrFileInfo) Then myFileCount = UBound(myArrFileInfo, 2) Else myFileCount = 0 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Database Information">
		<TR><TD CLASS="caption" COLSPAN=2><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/database.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Database"> <% = myTObj.getTerm(120) & " [ " & myDbName & " ]" %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(121) %> :</TD><TD><%  =myArrInfo(0, 0) %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(122) %> :</TD><TD><%  =myArrInfo(1, 0) %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(123) %> :</TD><TD><%  =myArrInfo(2, 0) %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(131) %> :</TD><TD><%  =myArrInfo(3, 0) %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(132) %> :</TD><TD><%  =myArrInfo(4, 0) %></TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(133) %> :</TD><TD><%  =myArrInfo(5, 0) %></TD></TR>
		<TR><TD CLASS="info" COLSPAN=2>&nbsp;</TD></TR>
		<TR><TD CLASS="info"><% = myTObj.getTerm(124) %> :</TD><TD>&nbsp;</TD></TR>
		<% For i = 0 To myFileCount %>
		<TR><TD CLASS="formlabel"><% = Trim(myArrFileInfo(0, i)) %></TD><TD>&nbsp;</TD></TR>
		<TR>
			<TD>&nbsp;</TD>
			<TD>
				<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 ALIGN=CENTER SUMMARY="Database Files Information">
					<TR><TD CLASS="info"><% = myTObj.getTerm(125) %> :</TD><TD><% = myArrFileInfo(2, i) %></TD></TR>
					<TR><TD CLASS="info"><% = myTObj.getTerm(126) %> :</TD><TD><% = myArrFileInfo(3, i) %></TD></TR>
					<TR><TD CLASS="info"><% = myTObj.getTerm(127) %> :</TD><TD><% = myArrFileInfo(4, i) %></TD></TR>
					<TR><TD CLASS="info"><% = myTObj.getTerm(128) %> :</TD><TD><% = myArrFileInfo(5, i) %></TD></TR>
					<TR><TD CLASS="info"><% = myTObj.getTerm(129) %> :</TD><TD><% = myArrFileInfo(6, i) %></TD></TR>
					<TR><TD CLASS="info"><% = myTObj.getTerm(130) %> :</TD><TD><% = myArrFileInfo(7, i) %></TD></TR>
				</TABLE>
			</TD>
		</TR>
		<% Next %>
	</TABLE>

	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
		<!--
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('<% = myNodeId %>'); 
		//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->