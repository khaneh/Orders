<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myArrDb, myDbCount, myStrTree, myObjStr, myNodeId, i, myStrClass
	openConnection
	myStrTree = getTreeStr("2_1", Array())
	myArrDB = getDbList()
	closeConnection
	If isArray(myArrDb) Then myDbCount = UBound(myArrDb, 2) Else myDbCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<% = myTObj.getTerm(5) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Database List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(56) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myDbCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myNodeId = "M2_1_" & (i+1) 
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/database.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrDb(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<A HREF=""dbinfo.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myArrDb(0, i)) & """>"
							myObjStr.strAppend "<B>" & myArrDb(0, i) & "</B>"
							myObjStr.strAppend "</A></TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
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
		if (window.parent.frames['Tree']) window.parent.frames['Tree'].oTree.expandNode('M2_1'); 
		<% End If %>
		//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->