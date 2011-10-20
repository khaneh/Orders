<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myNodeId, myDbName, myTblName, myTblShortName, myArrCol, myColCount, myArrIx, myIxCount, myCurrentIx, myArrFk, myFkCount, myCurrentFk, myArrTr, myTrCount, myTrName, myTrShortName, i, j, myObjStr, myStrClass, myStrTree, myArrCnt, myCntCount, myCntName, myCntShortName, myIxName, myFkName
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	myTblName = Request.QueryString("tbl")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_1_2", Array(myDbName, myTblName))
	myTblShortName = getObjShortName(myDbName, myTblName)
	myArrCol = getTblCol(myDbName, myTblName)
	myArrIx = getTblIx(myDbName, myTblName)
	myArrFk = getTblForeignKey(myDbName, myTblName)
	myArrTr = getTblTrigger(myDbName, myTblName)
	myArrCnt = getTblConstraint(myDbName, myTblName)
	closeConnection
	If isArray(myArrCol) Then myColCount = UBound(myArrCol, 2) Else myColCount = -1 End If
	If isArray(myArrIx) Then myIxCount = UBound(myArrIx, 2) Else myIxCount = -1 End If
	If isArray(myArrFk) Then myFkCount = UBound(myArrFk, 2) Else myFkCount = -1 End If
	If isArray(myArrTr) Then myTrCount = UBound(myArrTr, 2) Else myTrCount = -1 End If
	If isArray(myArrCnt) Then myCntCount = UBound(myArrCnt, 2) Else myCntCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/table.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Columns"> <% = myTObj.getTerm(331) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Table Stucture">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(311) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(300) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(301) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(302) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(303) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(304) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(305) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(306) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(323) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(307) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(308) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(309) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(310) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myColCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>"
							If myArrCol(12, i) Then
								myObjStr.strAppend "<IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/pkey.gif"" WIDTH=""15"" HEIGHT=""15"" BORDER=0 ALT=""Primary Key"" ALIGN=""MIDDLE"">"
							Else
								myObjStr.strAppend "&nbsp;"
							End If
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP><B>" & myArrCol(0, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCol(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>" & myArrCol(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(5, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCol(6, i) & "&nbsp;</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCol(7, i) & "&nbsp;</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(8, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(9, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(10, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(11, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
		</TD></TR>
	</TABLE>

	&nbsp;<BR>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption" COLSPAN=2><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/index.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Indexes"> <% = myTObj.getTerm(312) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Table Indexes">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(313) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(314) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(315) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(316) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(319) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(317) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(318) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(322) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(320) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(321) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						myCurrentIx = 0
						j = 0
						Set myObjStr = New mlt_string
						For i = 0 To myIxCount
							If myArrIx(3, i) Then
								myIxName = "[" & rembracket(myArrIx(1, i)) & "]"								
							Else
								myIxName = "[" & rembracket(myTblShortName) & "].[" & rembracket(myArrIx(1, i)) & "]"
							End If
							If j MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myCurrentIx = myArrIx(0, i)
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP><B>" & myArrIx(1, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>"
							Do While myArrIx(0, i) = myCurrentIx
								myObjStr.strAppend myArrIx(2, i) & "<BR>"
								i = i + 1
								If i > myIxCount Then Exit Do
							Loop
							i = i - 1
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrIx(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrIx(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrIx(7, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrIx(5, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>" & myArrIx(6, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrIx(8, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrIx(9, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>&nbsp;" & myArrIx(10, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
							j = j + 1
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
		</TD></TR>
	</TABLE>

	&nbsp;<BR>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption" COLSPAN=2><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/relation.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Relationship"> <% = myTObj.getTerm(324) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Table Relationship">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(325) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(326) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(327) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						myCurrentFk = 0
						Set myObjStr = New mlt_string
						For i = 0 To myFkCount
							If myCurrentFk <> myArrFk(0, i) Then
								myFkName = "[" & rembracket(myArrFk(1, i)) & "]"
								myStrClass = "odd"
								myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
								myObjStr.strAppend "<TD NOWRAP><B>" & myArrFk(1, i) & "</B></TD>" & vbCrLf
								myObjStr.strAppend "<TD NOWRAP><B>" & myArrFk(2, i) & "</B></TD>" & vbCrLf
								myObjStr.strAppend "<TD NOWRAP><B>" & myArrFk(3, i) & "</B></TD>" & vbCrLf
								myObjStr.strAppend "</TR>" & vBCrLf
								myCurrentFk = myArrFk(0, i)
							End If
							myStrClass = "even"
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD NOWRAP>&nbsp;</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrFk(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrFk(5, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>&nbsp;</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
		</TD></TR>
	</TABLE>

	&nbsp;<BR>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption" COLSPAN=2><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/trigger.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Triggers"> <% = myTObj.getTerm(328) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Trigger List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(329) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myTrCount
							myTrName = "[" & rembracket(myDbName) & "].[" & rembracket(myArrTr(1, i)) & "].[" & rembracket(myArrTr(0, i)) & "]"
							myTrShortName = "[" & rembracket(myArrTr(1, i)) & "].[" & rembracket(myArrTr(0, i)) & "]"
							If j MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD NOWRAP><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/trigger.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrTr(0, i) & """ ALIGN=""MIDDLE""> <A HREF=""trstruct.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&tbl=" & Server.URLEncode(myTblName) & "&tr=" & Server.URLEncode(myTrName) & """><B>" & myArrTr(0, i) & "</B></A></TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrTr(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrTr(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
		</TD></TR>
	</TABLE>

	&nbsp;<BR>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption" COLSPAN=2><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/constraint.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Constraints"> <% = myTObj.getTerm(347) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Constraint List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(348) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(349) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myCntCount
							myCntName = "[" & rembracket(myDbName) & "].[" & rembracket(myArrCnt(1, i)) & "].[" & rembracket(myArrCnt(0, i)) & "]"
							myCntShortName = "[" & rembracket(myArrCnt(0, i)) & "]"
							If j MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD NOWRAP><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/constraint.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrCnt(0, i) & """ ALIGN=""MIDDLE""> <B>" & myArrCnt(0, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCnt(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCnt(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & txt2html(myArrCnt(3, i)) & "</TD>" & vbCrLf
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