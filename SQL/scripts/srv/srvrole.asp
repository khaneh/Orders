<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myArrRole, myRoleCount, i, myObjStr, myStrClass, myStrTree
	openConnection
	myStrTree = getTreeStr("2_3_2", Array())
	myArrRole = getSrvRole()
	closeConnection
	If isArray(myArrRole) Then myRoleCount = UBound(myArrRole, 2) Else myRoleCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/srvrole.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Server Role"> <% = myTObj.getTerm(231) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Server Role List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(250) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(251) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myRoleCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/srvrole.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrRole(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<B>" & myArrRole(0, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrRole(1, i) & "</TD>" & vbCrLf
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