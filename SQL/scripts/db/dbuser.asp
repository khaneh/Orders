<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrUser, myUserCount, i, myObjStr, myStrClass, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_4", Array(myDbName))
	myArrUser = getDbUser(myDbName)
	closeConnection
	If isArray(myArrUser) Then myUserCount = UBound(myArrUser, 2) Else myUserCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/user.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Users"> <% = myTObj.getTerm(10) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="User List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(210) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(211) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myUserCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/user.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrUser(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<B>" & myArrUser(0, i) & "</B>"
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrUser(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
		</TD></TR>
	</TABLE>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->