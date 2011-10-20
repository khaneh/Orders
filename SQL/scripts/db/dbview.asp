<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrView, myViewCount, i, myObjStr, myStrClass, myViewName, myStrTree, myShortViewName
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_2", Array(myDbName))
	myArrView = getDbView(myDbName)
	closeConnection
	If isArray(myArrView) Then myViewCount = UBound(myArrView, 2) Else myViewCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/view.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Views"> <% = myTObj.getTerm(8) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="View List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(160) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD><TD CLASS="collabel">&nbsp;</TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myViewCount
							myViewName = "[" & rembracket(myDbName) & "].[" & rembracket(myArrView(1, i)) & "].[" & rembracket(myArrView(0, i)) & "]"
							myShortViewName = "[" & rembracket(myArrView(1, i)) & "].[" & rembracket(myArrView(0, i)) & "]"
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/view.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrView(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<A HREF=""viewstruct.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&view=" & Server.URLEncode(myViewName) & """>"
							If Trim(myArrView(3, i)) <> 0 Then
								myObjStr.strAppend "<B><I>" & myArrView(0, i) & "</I></B>"
							Else
								myObjStr.strAppend "<B>" & myArrView(0, i) & "</B>"
							End If
							If mla_auth(2, 4) Or mla_auth(2, 5) Then myObjStr.strAppend "</A>"
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrView(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrView(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>"
							If mla_auth(2, 5) Then myObjStr.strAppend "&nbsp;<A HREF=""content.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&obj=" & Server.URLEncode(myViewName) & "&type=view""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/content.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myTObj.getTerm(152) & """></A>"
							If mla_auth(2, 6) Then myObjStr.strAppend "&nbsp;<A HREF=""search.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&obj=" & Server.URLEncode(myViewName) & "&type=view""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/search.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myTObj.getTerm(154) & """></A>"
							myObjStr.strAppend "&nbsp;" & vbCrLf
							If mla_auth(2, 4) Then myObjStr.strAppend "&nbsp;<A HREF=""viewstruct.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&view=" & Server.URLEncode(myViewName) & """><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/property.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myTObj.getTerm(153) & """></A>"
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