<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrRole, myRoleCount, i, myObjStr, myStrClass, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_5", Array(myDbName))
	myArrRole = getDbRole(myDbName)
	closeConnection
	If isArray(myArrRole) Then myRoleCount = UBound(myArrRole, 2) Else myRoleCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/role.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Roles"> <% = myTObj.getTerm(11) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Role List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(220) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(221) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myRoleCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/role.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrRole(0, i) & """ ALIGN=""MIDDLE""> <B>" & myArrRole(0, i) & "</B></TD>" & vbCrLf
							If myArrRole(2, i) Then
								myObjStr.strAppend "<TD>" & myTObj.getTerm(222) & "</TD>" & vbCrLf
							Else
								myObjStr.strAppend "<TD>" & myTObj.getTerm(223) & "</TD>" & vbCrLf
							End If
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