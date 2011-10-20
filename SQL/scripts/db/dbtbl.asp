<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrTbl, myTblCount, i, myObjStr, myStrClass, myTblName, myTblShortName, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_1", Array(myDbName))
	myArrTbl = getDbTbl(myDbName)
	closeConnection
	If isArray(myArrTbl) Then myTblCount = UBound(myArrTbl, 2) Else myTblCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/table.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Tables"> <% = myTObj.getTerm(7) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Table List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(150) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(126) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(151) %></TD><TD CLASS="collabel">&nbsp;</TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myTblCount
							myTblName = "[" & rembracket(myDbName) & "].[" & rembracket(myArrTbl(1, i)) & "].[" & rembracket(myArrTbl(0, i)) & "]"
							myTblShortName = "[" & rembracket(myArrTbl(1, i)) & "].[" & rembracket(myArrTbl(0, i)) & "]"
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/table.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrTbl(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<A HREF=""tblstruct.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&tbl=" & Server.URLEncode(myTblName) & """>"
							If Trim(myArrTbl(5, i)) = "S" Then
								myObjStr.strAppend "<B><I>" & myArrTbl(0, i) & "</I></B>"
							Else
								myObjStr.strAppend "<B>" & myArrTbl(0, i) & "</B>"
							End If
							If mla_auth(1, 4) Or mla_auth(1, 5) Then myObjStr.strAppend "</A>"
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrTbl(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrTbl(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrTbl(3, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD ALIGN=RIGHT>" & myArrTbl(4, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>"
							If mla_auth(1, 5) Then myObjStr.strAppend "&nbsp;<A HREF=""content.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&obj=" & Server.URLEncode(myTblName) & "&type=tbl""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/content.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myTObj.getTerm(152) & """></A>"
							If mla_auth(1, 6) Then myObjStr.strAppend "&nbsp;<A HREF=""search.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&obj=" & Server.URLEncode(myTblName) & "&type=tbl""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/search.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myTObj.getTerm(154) & """></A>"
							myObjStr.strAppend "&nbsp;" & vbCrLf
							If mla_auth(1, 4) Then myObjStr.strAppend "&nbsp;<A HREF=""tblstruct.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&tbl=" & Server.URLEncode(myTblName) & """><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/property.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myTObj.getTerm(153) & """></A>"							
							myObjStr.strAppend "&nbsp;</TD>" & vbCrLf
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