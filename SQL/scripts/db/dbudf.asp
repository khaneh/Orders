<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrUdf, myUdfCount, i, myObjStr, myStrClass, myUdfName, myStrTree, myUdfShortName
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_9", Array(myDbName))
	myArrUdf = getDbFunction(myDbName)
	closeConnection
	If isArray(myArrUdf) Then myUdfCount = UBound(myArrUdf, 2) Else myUdfCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/udf.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="User defined functions"> <% = myTObj.getTerm(16) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="UDF List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(176) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myUdfCount
							myUdfName = "[" & myDbName & "].[" & myArrUdf(1, i) & "].[" & myArrUdf(0, i) & "]"
							myUdfShortName = "[" & myArrUdf(1, i) & "].[" & myArrUdf(0, i) & "]"
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/udf.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrUdf(0, i) & """ ALIGN=""MIDDLE""> "
							If mla_auth(9, 4) Then myObjStr.strAppend "<A HREF=""udfstruct.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&udf=" & Server.URLEncode(myUdfName) & """>"
							If Trim(myArrUdf(3, i)) <> 0 Then
								myObjStr.strAppend "<B><I>" & myArrUdf(0, i) & "</I></B>" & vbCrLf
							Else
								myObjStr.strAppend "<B>" & myArrUdf(0, i) & "</B>" & vbCrLf
							End If
							If mla_auth(9, 1) Then myObjStr.strAppend "</A>"
							myObjStr.strAppend "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrUdf(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrUdf(2, i) & "</TD>" & vbCrLf
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