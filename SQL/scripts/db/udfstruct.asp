<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myNodeId, myDbName, myUdfName, myArrCol, myColCount, i, myObjStr, myStrClass, myUdfTxt, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	myUdfName = Request.QueryString("udf")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_9_1", Array(myDbName, myUdfName))
	myArrCol = getSpParam(myDbName, myUdfName)
	myUdfTxt = txt2html(getObjTxt(myDbName, myUdfName))
	closeConnection
	If isArray(myArrCol) Then myColCount = UBound(myArrCol, 2) Else myColCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><% = myTObj.getTerm(174) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="UDF Stucture">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(172) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(301) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(302) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(303) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(304) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(305) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myColCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD NOWRAP><B>"
							If myArrCol(0, i) = "" Then myObjStr.strAppend "@RETURN_VALUE" Else myObjStr.strAppend myArrCol(0, i) End If
							myObjStr.strAppend "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP>" & myArrCol(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>" & myArrCol(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP ALIGN=RIGHT>&nbsp;" & myArrCol(5, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vBCrLf
						Next
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
			</TABLE>
			
			&nbsp;<BR>

			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="UDF Definition">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(175) %></TD></TR>
				</THEAD>
				<TBODY>
					<TR><TD><% = myUdfTxt %></TD></TR>
				</TBODY>
			</TABLE>
			
		</TD></TR>
	</TABLE>
</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->