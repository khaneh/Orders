<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrSp, mySpCount, i, myObjStr, myStrClass, mySpName, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_3", Array(myDbName))
	myArrSp = getDbSp(myDbName)
	closeConnection
	If isArray(myArrSp) Then mySpCount = UBound(myArrSp, 2) Else mySpCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/sp.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Stored Procedures"> <% = myTObj.getTerm(9) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Stored Procedure List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(170) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To mySpCount
							mySpName ="[" & rembracket(myArrSp(1, i)) & "].[" & rembracket(myArrSp(0, i)) & "]"
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/sp.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrSp(0, i) & """ ALIGN=""MIDDLE""> "
							If mla_auth(3, 4) Then myObjStr.strAppend "<A HREF=""spstruct.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&sp=" & Server.URLEncode(mySpName) & """>"
							If Trim(myArrSp(3, i)) <> 0 Then
								myObjStr.strAppend "<B><I>" & myArrSp(0, i) & "</I></B>"
							Else
								myObjStr.strAppend "<B>" & myArrSp(0, i) & "</B>"
							End If
							If mla_auth(3, 4) Then myObjStr.strAppend "</A>"
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrSp(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrSp(2, i) & "</TD>" & vbCrLf
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