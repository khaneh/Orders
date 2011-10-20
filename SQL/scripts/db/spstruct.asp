<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myNodeId, myDbName, mySpName, myArrCol, myColCount, i, myObjStr, myStrClass, mySpTxt, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	mySpName = Request.QueryString("sp")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_3_1", Array(myDbName, mySpName))
	myArrCol = getSpParam(myDbName, mySpName)
	mySpTxt = txt2html(getObjTxt(myDbName, mySpName))
	closeConnection
	If isArray(myArrCol) Then myColCount = UBound(myArrCol, 2) Else myColCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><% = myTObj.getTerm(174) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="SP Stucture">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(172) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(301) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(302) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(303) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(304) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(305) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(173) %></TD></TR>
				</THEAD>
				<TBODY>
					<TR CLASS="odd">
						<TD NOWRAP><B>@RETURN_VALUE</B></TD>
						<TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD>
					</TR>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myColCount
							If i MOD 2 = 0 Then myStrClass = "even" Else myStrClass = "odd" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD NOWRAP><B>" & myArrCol(0, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCol(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>" & myArrCol(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(5, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>" & myArrCol(6, i) & "&nbsp;</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
			
			&nbsp;<BR>

			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="SP Definition">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(171) %></TD></TR>
				</THEAD>
				<TBODY>
					<TR><TD><% = mySpTxt %></TD></TR>
				</TBODY>
			</TABLE>
			
		</TD></TR>
	</TABLE>
</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->