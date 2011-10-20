<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrUdt, myUdtCount, i, myObjStr, myStrClass, myStrTree, myUdtName
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_8", Array(myDbName))
	myArrUdt = getDbUdt(myDbName)
	closeConnection
	If isArray(myArrUdt) Then myUdtCount = UBound(myArrUdt, 2) Else myUdtCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P	 CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/udt.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="User Defined Types"> <% = myTObj.getTerm(14) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="User Defined Type List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(200) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(201) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(202) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(203) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(190) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(180) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myUdtCount
							myUdtName = "[" & rembracket(myArrUdt(1, i) ) & "].[" & rembracket(myArrUdt(0, i)) & "]"
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/udt.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrUdt(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<B>" & myArrUdt(0, i) & "</B>"
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrUdt(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrUdt(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD ALIGN=RIGHT>" & myArrUdt(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD ALIGN=RIGHT>" & myArrUdt(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrUdt(5, i) & "&nbsp;</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrUdt(6, i) & "&nbsp;</TD>" & vbCrLf
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