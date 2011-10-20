<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myNodeId, myDbName, myViewName, myArrCol, myColCount, i, myObjStr, myStrClass, myViewTxt, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	myViewName = Request.QueryString("view")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_2_2", Array(myDbName, myViewName))
	myArrCol = getViewCol(myDbName, myViewName)
	myViewTxt = txt2html(getObjTxt(myDbName, myViewName))
	closeConnection
	If isArray(myArrCol) Then myColCount = UBound(myArrCol, 2) Else myColCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><% = myTObj.getTerm(331) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="View Stucture">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(300) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(301) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(302) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(303) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(304) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(305) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(307) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(308) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(309) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(310) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myColCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD NOWRAP><B>" & myArrCol(0, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCol(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>" & myArrCol(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(5, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(6, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(7, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(8, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(9, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
			
			&nbsp;<BR>

			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="View Definition">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(161) %></TD></TR>
				</THEAD>
				<TBODY>
					<TR><TD><% = myViewTxt %></TD></TR>
				</TBODY>
			</TABLE>
			
		</TD></TR>
	</TABLE>
</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->