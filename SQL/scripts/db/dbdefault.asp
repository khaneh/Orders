<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrDefault, myDefaultCount, i, myObjStr, myStrClass, myStrTree, myDefaultName
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_7", Array(myDbName))
	myArrDefault = getDbDefault(myDbName)
	closeConnection
	If isArray(myArrDefault) Then myDefaultCount = UBound(myArrDefault, 2) Else myDefaultCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P	 CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/default.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Defaults"> <% = myTObj.getTerm(13) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Default List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(190) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(191) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myDefaultCount
							myDefaultName = "[" & rembracket(myArrDefault(1, i)) & "].[" & rembracket(myArrDefault(0, i)) & "]"
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/default.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrDefault(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<B>" & myArrDefault(0, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrDefault(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrDefault(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrDefault(3, i) & "</TD>" & vbCrLf
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