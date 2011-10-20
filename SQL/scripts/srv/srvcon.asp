<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myArrCon, myConCount, i, myObjStr, myStrClass, myStrImg, myStrTree
	openConnection
	myStrTree = getTreeStr("2_3_1", Array())
	myArrCon = getSrvCon()
	closeConnection
	If isArray(myArrCon) Then myConCount = UBound(myArrCon, 2) Else myConCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/connection.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Connections"> <% = myTObj.getTerm(230) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Connection List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(240) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(241) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(242) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(243) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myConCount
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							Select Case myArrCon(1, i) 
								Case 1: myStrImg = "connection_ntgroup.gif"
								Case 2 : myStrImg = "connection_ntuser.gif"
								Case Else : myStrImg = "connection.gif"
							End Select
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vbCrLf
							myObjStr.strAppend "<TD NOWRAP><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/" & myStrImg & """ WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrCon(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<B>" & myArrCon(0, i) & "</B>"
							myObjStr.strAppend "</TD>" & vbCrLf
							Select Case myArrCon(1, i) 
								Case 1 : myObjStr.strAppend "<TD>" & myTObj.getTerm(244) & "</TD>" & vbCrLf
								Case 2 : myObjStr.strAppend "<TD>" & myTObj.getTerm(245) & "</TD>" & vbCrLf
								Case Else : myObjStr.strAppend "<TD>" & myTObj.getTerm(223) & "</TD>" & vbCrLf
							End Select
							myObjStr.strAppend "<TD>" & myArrCon(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrCon(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrCon(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "</TR>" & vbCrLf
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