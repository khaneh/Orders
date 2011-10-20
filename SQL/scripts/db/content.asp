<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myNodeId, myDbName, myObjName, myObjType, myStrTree, myArrCol, myColCount, i, j, myObjStr, myStrClass
	Dim myPgNb, myPgCount, myRecordCount, myStrRootLink, myStrBeginLnk, myStrPreviousLnk, myStrNextLnk, myStrEndLnk
	Dim mySortCol, mySortWay
	Dim myStrContent, myArrSQL, myObjRS
	Dim myStrImg
	Dim myStrCol, myStrWhere
	Dim myColSpan
	Dim myArrPK, ii, myPKFCount, myPKValue, myPKFlag
	Dim myStrLink

	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	myObjName = Request.QueryString("obj")
	myObjType = Request.QueryString("type")
	
	myStrCol = Request.QueryString("col")
	myStrWhere = Request.QueryString("cnd")

	mySortCol = Request.QueryString("sc")
	mySortWay = Request.QueryString("sw")
	If mySortWay <> "DESC" Then mySortWay = "ASC"
	
	myPgNb = Request("pg")
	If isNumeric(myPgNb) Then
		myPgNb = CLng(myPgNb)
		If myPgNb < 1 Then myPgNb = 1
	Else
		myPgNb = 1
	End If
	
	openConnection
	Select Case myObjType
		Case "tbl" :
			myStrTree = getTreeStr(Mid(myNodeId, 2) & "_1_1", Array(myDbName, myObjName))
			myArrPK = getTblPrimaryKey(myDbName, myObjName)
			If isArray(myArrPK) Then myPKFCount = UBound(myArrPK, 2)
			myStrImg = "table"
		Case "view" :
			myStrTree = getTreeStr(Mid(myNodeId, 2) & "_2_1", Array(myDbName, myObjName))
			myStrImg = "view"
	End Select
	myRecordCount = getSQLRecordCount(myDbName, myObjName, myStrWhere)
	myArrSQL = getSQLstr(myObjName, myStrCol, myStrWhere, mySortCol, mySortWay, mla_cfg_pagesize, myPgNb)

	Set myObjRS = Server.CreateObject("ADODB.Recordset")
	myObjRS.ActiveConnection = gObjDC
	myObjRS.CursorType = 3
	myObjRS.PageSize = mla_cfg_pagesize
	myObjRS.Open myArrSQL(0)

	myColCount = myObjRS.Fields.Count-1
	ReDim myArrCol(3, myColCount)

	myPKFlag = False
	If isArray(myArrPK) Then
		' checks whether each PKField was in the SQL query
		For i = 0 To myPKFCount
			myPKFlag = False
			For ii = 0 To myColCount
				If myObjRS(ii).Name = myArrPK(0, i) Then 
					myPKFlag = True
					Exit For
				End If
			Next
			If myPKFlag = False Then Exit For
		Next
	End If

	myPgCount = Int(myRecordCount / mla_cfg_pagesize)
	If (myRecordCount MOD mla_cfg_pagesize) > 0 Then myPgCount = myPgCount + 1
	If myPgNb > myPgCount Then myPgNb = myPgCount

	myStrRootLink = "content.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&obj=" & Server.URLEncode(myObjName) & "&type=" & Server.URLEncode(myObjType) & "&col=" & Server.URLEncode(myStrCol) & "&cnd=" & Server.URLEncode(myStrWhere)
	If myPgNb > 1 Then
		myStrBeginLnk = myStrRootLink & "&sc=" & mySortCol & "&sw=" & mySortWay & "&pg=1"
		myStrPreviousLnk = myStrRootLink & "&sc=" & mySortCol & "&sw=" & mySortWay & "&pg=" & (myPgNb - 1)
	Else
		myStrBeginLnk = Empty
		myStrPreviousLnk = Empty
	End If
	If myPgNb < myPgCount Then
		myStrNextLnk = myStrRootLink & "&sc=" & mySortCol & "&sw=" & mySortWay & "&pg=" & (myPgNb + 1)
		myStrEndLnk = myStrRootLink & "&sc=" & mySortCol & "&sw=" & mySortWay & "&pg=" & myPgCount
	Else
		myStrNextLnk = Empty
		myStrEndLnk = Empty
	End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>
	<FORM NAME="mla_content" METHOD=POST ACTION="<% = myStrRootLink & "&sc=" & mySortCol & "&sw=" & mySortWay %>">
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/<% = myStrImg %>.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Content"> <% = myTObj.getTerm(332) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=1 CLASS="content" SUMMARY="Query">
				<TR><TD CLASS="formlabel" NOWRAP><% = myTObj.getTerm(452) %> :</TD><TD NOWRAP><% Response.Write myArrSQL(1) %></TD><TD WIDTH="100%">&nbsp;</TD></TR>
				<TR><TD CLASS="formlabel" NOWRAP><% = myTObj.getTerm(453) %> :</TD><TD NOWRAP><% Response.Write myRecordCount %></TD><TD WIDTH="100%">&nbsp;</TD></TR>
				<TR><TD CLASS="formlabel" NOWRAP><% = myTObj.getTerm(454) %> :</TD>
					<TD NOWRAP>
						<A HREF="search.asp?nid=<% = myNodeId %>&db=<% = Server.URLEncode(myDbName) %>&obj=<% = Server.URLEncode(myObjName) %>&type=<% = Server.URLEncode(myObjType) %>"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/search.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="<% = myTObj.getTerm(154) %>"></A>
						&nbsp;
						<A HREF="<% If myObjType = "tbl" Then Response.Write "tblstruct" Else Response.Write "viewstruct" End If %>.asp?nid=<% = myNodeId %>&db=<% = Server.URLEncode(myDbName) %>&<% = myObjType%>=<% = Server.URLEncode(myObjName) %>"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/property.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="<% = myTObj.getTerm(153) %>"></A>
					</TD>
					<TD WIDTH="100%">&nbsp;</TD></TR>
			</TABLE>
		</TD></TR>
		<TR><TD>			
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Content">
				<THEAD>
					<%
						Set myObjStr = New mlt_string
						myObjStr.strAppend "<TR>" & vbCrLf
						If myPKFlag Then myObjStr.strAppend "<TD CLASS=""collabel"">&nbsp;</TD>" & vbCrLf
						For i = 0 To myColCount
							myArrCol(0, i) = myObjRS.Fields(i).Name
							myArrCol(1, i) = getColAlign(myObjRS.Fields(i).Type)
							myArrCol(2, i) = getColSort(myObjRS.Fields(i).Type)
							myArrCol(3, i) = getColDisplay(myObjRS.Fields(i).Type)
							myObjStr.strAppend "<TD CLASS=""collabel"">" & myArrCol(0, i) & "&nbsp;"
							If myArrCol(2, i) Then
								myObjStr.strAppend "<A HREF=""" & myStrRootLink & "&sc=" & Server.URLEncode(myArrCol(0, i)) & "&sw=ASC""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/asc.gif"" WIDTH=13 HEIGHT=9 BORDER=0 ALT=""ASC""></A>"
								myObjStr.strAppend "<A HREF=""" & myStrRootLink & "&sc=" & Server.URLEncode(myArrCol(0, i)) & "&sw=DESC""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/desc.gif"" WIDTH=13 HEIGHT=9 BORDER=0 ALT=""DESC""></A>"
							Else
								myObjStr.strAppend "<IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/asc_dis.gif"" WIDTH=13 HEIGHT=9 BORDER=0 ALT=""ASC"">"
								myObjStr.strAppend "<IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/desc_dis.gif"" WIDTH=13 HEIGHT=9 BORDER=0 ALT=""DESC"">"
							End If
							myObjStr.strAppend "</TD>"
						Next
						myObjStr.strAppend "</TR>" & vbCrLf
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>	
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						i = 0
						If myPgCount > 0 Then
							myObjRS.Move (mla_cfg_pagesize * (myPgNb-1))
							Do While Not myObjRS.EOF
								If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
								i = i + 1
								myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vbCrLf
								If (mla_auth(1, 10) Or mla_auth(1, 11)) And myPKFlag Then 
									myObjStr.strAppend "<TD NOWRAP ALIGN=CENTER WIDTH=""40"">"
									myObjStr.strAppend "<INPUT TYPE=""hidden"" NAME=""mla_pk_value_" & i & """>"
									myPKValue = ""
									For ii = 0 To myPKFCount
										If myPkValue <> "" Then myPKValue = myPKValue & " AND"
										myPKValue = myPKValue & " [" & rembracket(myArrPK(0, ii)) & "] = "
										Select Case myArrPK(1, ii)
											Case 1, 6 :	' STRING and UNIQUEIDENTIFIER
												myPKValue = myPKValue & "'" & remquote(myObjRS(myArrPK(0, ii))) & "'"
											Case 4 : ' DATE
												myPKValue = myPKValue & "CONVERT(DATETIME, '" & str2date(myObjRS(myArrPK(0, ii))) & "', 112)"
											Case Else : 
												myPKValue = myPKValue & myObjRS(myArrPK(0, ii))
										End Select
									Next
									myObjStr.strAppend "<SCRIPT LANGUAGE=""JavaScript"" TYPE=""text/javascript"">" & vbCrLf & "<!--" & vbCrlf
									myObjStr.strAppend "document.mla_content.mla_pk_value_" & i & ".value = """ & remdquote(myPKValue) & """;" & vbCrLf
									myObjStr.strAppend "//-->" & vbCrLf & "</SCRIPT>"
									myObjStr.strAppend "</TD>" & vbCrLf
								End If
								For j = 0 To myColCount
									myObjStr.strAppend "<TD ALIGN=""" & myArrCol(1, j) & """>"
									If isNull(myObjRS(j)) Then
										myObjStr.strAppend "<SPAN CLASS=""moreinfo"">(" & myTObj.getTerm(60) & ")</SPAN>"
									Else
										myStrLink = "txtviewer.asp?db=" & Server.URLEncode(myDbName) & "&tbl=" & Server.URLEncode(myObjName) & "&fld=" & Server.URLEncode(myObjRS(j).Name) & "&pk=" & Server.URLEncode(myPKValue)
										Select Case myArrCol(3, j) 
											Case 1 :	' CHAR, NCHAR, VARCHAR, NVARCHAR
												myStrContent = getStrBegin(myObjRS(j), mla_cfg_maxdisplayedchar)
												myObjStr.strAppend txt2html(myStrContent(0))
												If myStrContent(1) Then 
													If myObjType = "tbl" And myPKFlag Then 
														myObjStr.strAppend " <A HREF=# onclick=""openPopUp('" & myStrLink & "', 'moreinfo', 400, 400, 10, 10); return(false);"" CLASS=""moreinfo"">(...)</A>"
													Else
														myObjStr.strAppend " <SPAN CLASS=""moreinfo"">(...)</SPAN>"
													End If
												End If
											Case 2 :	' TEXT, NTEXT
												myStrContent = getStrBegin(myObjRS(j), mla_cfg_maxdisplayedchar)
												myObjStr.strAppend txt2html(myStrContent(0))
												If myStrContent(1) Then 
													If myObjType = "tbl" And myPKFlag Then 
														myObjStr.strAppend " <A HREF=# onclick=""openPopUp('" & myStrLink & "', 'moreinfo', 400, 400, 10, 10); return(false);"" CLASS=""moreinfo"">(...)</A>"
													Else
														myObjStr.strAppend " <SPAN CLASS=""moreinfo"">(...)</SPAN>"
													End If
												End If
											Case 3 :	' BINARY, VARBINARY, IMAGE
												myStrContent = bin2hex(myObjRS(j), mla_cfg_maxdisplayedbin)
												myObjStr.strAppend txt2html(myStrContent(0))
												If myStrContent(1) Then myObjStr.strAppend " <SPAN CLASS=""moreinfo"">(...)</SPAN>"
											Case Else :	' ALL OTHERS
												myObjStr.strAppend myObjRS(j)
										End Select
									End If
									myObjStr.strAppend "</TD>" & vbCrLf
								Next
								myObjStr.strAppend "</TR>" & vbCrLf
								myObjRS.MoveNext
							Loop
						End If
						myObjRS.Close
						Set myObjRS = Nothing
						closeConnection
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</TBODY>
				<THEAD>
					<%
						Set myObjStr = New mlt_string
						If myObjType = "tbl" And isArray(myArrPK) Then myColSpan = myColCount + 2 Else myColSpan = myColCount + 1 End If
						myObjStr.strAppend "<TR><TD CLASS=""collabel"" COLSPAN=" & myColSpan & ">"
						If isEmpty(myStrBeginLnk) Then
							myObjStr.strAppend "<IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/begin_dis.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""Begin"">"
							myObjStr.strAppend "<IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/previous_dis.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""Previous"">"
						Else
							myObjStr.strAppend "<A HREF=""" & myStrBeginLnk & """><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/begin.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""Begin""></A>"
							myObjStr.strAppend "<A HREF=""" & myStrPreviousLnk & """><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/previous.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""Previous""></A>"
						End If
						myObjStr.strAppend "<INPUT TYPE=""text"" NAME=""pg"" VALUE=""" & myPgNb & """ CLASS=""numeric"" TITLE=""Page"" >"
						myObjStr.strAppend " / " & myPgCount
						If isEmpty(myStrEndLnk) Then
							myObjStr.strAppend "<IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/next_dis.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""Next"">"
							myObjStr.strAppend "<IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/end_dis.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""End"">"
						Else
							myObjStr.strAppend "<A HREF=""" & myStrNextLnk & """><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/next.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""Next""></A>"
							myObjStr.strAppend "<A HREF=""" & myStrEndLnk & """><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/end.gif"" WIDTH=16 HEIGHT=16 BORDER=0 ALIGN=MIDDLE ALT=""End""></A>"
						End If
						Response.Write myObjStr.getStr()
						Set myObjStr = Nothing
					%>
				</THEAD>
			</TABLE>
		</TD></TR>
	</TABLE>
	</FORM>
</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->