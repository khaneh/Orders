<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myNodeId, myDbName, myObjName, myObjType, myStrTree, myStrClass
	Dim myObjStr, i
	Dim myArrCol, myColCount
	Dim myColStr, myWhereStr, myBoolOp

	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	myObjName = Request.QueryString("obj")
	myObjType = Request.QueryString("type")

	If Request.Form("mla_submit") <> "" Then
		openConnection
		Select Case myObjType
			Case "tbl" : myArrCol = getSimpleTblCol(myDbName, myObjName)
			Case "view" : myArrCol = getSimpleViewCol(myDbName, myObjName)
		End Select
		If isArray(myArrCol) Then myColCount = UBound(myArrCol, 2) Else myColCount = -1 End If
		myBoolOp = Request.Form("mla_match")
		myWhereStr = ""
		For i = 0 To myColCount
			If Request.Form("mla_display_" & i) <> "" Then myColStr = myColStr & "[" & rembracket(myArrCol(0, i)) & "], " 
			If Request.Form("mla_op_" & i) <> "-1" And Request.Form("mla_op_" & i) <> "" Then
				If myWhereStr <> "" Then myWhereStr = myWhereStr & " " & myBoolOp
				If Request.Form("mla_bool_" & i) <> "-1" Then myWhereStr = myWhereStr & " " & Request.Form("mla_bool_" & i) 
				myWhereStr = myWhereStr & " " & " ("
				myWhereStr = myWhereStr & "[" & rembracket(myArrCol(0, i)) & "]"
				If Request.Form("mla_op_" & i) <> "IS NULL" And Request.Form("mla_op_" & i) <> "IS NOT NULL"  Then
					myWhereStr = myWhereStr & " "
					If myArrCol(3, i) = 1 Or myArrCol(3, i) = 2 Then
						Select Case Request.Form("mla_op_" & i)
							Case "EXACT" : myWhereStr = myWhereStr & " LIKE '" & remquote(Request.Form("mla_value_" & i)) & "'"
							Case "STARTS" : myWhereStr = myWhereStr & " LIKE '" & remquote(Request.Form("mla_value_" & i)) & "%'"
							Case "ENDS" : myWhereStr = myWhereStr & " LIKE '%" & remquote(Request.Form("mla_value_" & i)) & "'"
							Case "CONTAINS" : myWhereStr = myWhereStr & " LIKE '%" & remquote(Request.Form("mla_value_" & i)) & "%'"
							Case Else : myWhereStr = myWhereStr & " " & Request.Form("mla_op_" & i) & " '" & remquote(Request.Form("mla_value_" & i)) & "'"
						End Select
					ElseIf myArrCol(3, i) = 4 Then
						myWhereStr = myWhereStr & " " & Request.Form("mla_op_" & i)
						myWhereStr = myWhereStr & " CONVERT(DATETIME, '" & str2date(Request.Form("mla_value_" & i)) & "', 112)"
						If Request.Form("mla_op_" & i) = "BETWEEN" Or Request.Form("mla_op_" & i) = "NOT BETWEEN" Then
							myWhereStr = myWhereStr & " AND CONVERT(DATETIME, '" & str2date(Request.Form("mla_value2_" & i)) & "', 112)"
						End If
					ElseIf myArrCol(3, i) = 5 Then
						Select Case Request.Form("mla_op_" & i)
							Case "IS TRUE" : myWhereStr = myWhereStr & " = 1"
							Case Else : myWhereStr = myWhereStr & " = 0"
						End Select
					ElseIf myArrCol(3, i) = 6 Then
						myWhereStr = myWhereStr & " " & Request.Form("mla_op_" & i) & " '" & remquote(Request.Form("mla_value_" & i)) & "'"
					Else
						myWhereStr = myWhereStr & " " & Request.Form("mla_op_" & i)
						myWhereStr = myWhereStr & " " & Request.Form("mla_value_" & i)
						If Request.Form("mla_op_" & i) = "BETWEEN" Or Request.Form("mla_op_" & i) = "NOT BETWEEN" Then
							myWhereStr = myWhereStr & " AND " & Request.Form("mla_value2_" & i)
						End If
					End If
				Else
					myWhereStr = myWhereStr & " " & Request.Form("mla_op_" & i)
				End If
				myWhereStr = myWhereStr & ")"
			End If
		Next
		If myColStr <> "" Then myColStr = Left(myColStr, Len(myColStr)-2)
		closeConnection
		Response.Redirect "content.asp?nid=" & myNodeId & "&db=" & Server.URLEncode(myDbName) & "&obj=" & Server.URLEncode(myObjName) & "&type=" & Server.URLEncode(myObjType) & "&col=" & Server.URLEncode(myColStr) & "&cnd=" & Server.URLEncode(myWhereStr)
		Response.End
	End If

	openConnection
	Select Case myObjType
		Case "tbl" : 
			myStrTree = getTreeStr(Mid(myNodeId, 2) & "_1_3", Array(myDbName, myObjName))
			myArrCol = getSimpleTblCol(myDbName, myObjName)
		Case "view" :
			myStrTree = getTreeStr(Mid(myNodeId, 2) & "_2_3", Array(myDbName, myObjName))
			myArrCol = getSimpleViewCol(myDbName, myObjName)
	End Select
	closeConnection
	If isArray(myArrCol) Then myColCount = UBound(myArrCol, 2) Else myColCount = -1 End If
	
	Dim myArrBool(0, 0)
	myArrBool(0, 0) = "NOT"

	' Numeric
	Dim myArrOpNumeric(0, 9)
	myArrOpNumeric(0, 0) = "="
	myArrOpNumeric(0, 1) = "&gt;"
	myArrOpNumeric(0, 2) = "&lt;"
	myArrOpNumeric(0, 3) = "&gt;="
	myArrOpNumeric(0, 4) = "&lt;="
	myArrOpNumeric(0, 5) = "&lt;&gt;"
	myArrOpNumeric(0, 6) = "IS NULL"
	myArrOpNumeric(0, 7) = "IS NOT NULL"
	myArrOpNumeric(0, 8) = "BETWEEN"
	myArrOpNumeric(0, 9) = "NOT BETWEEN"

	' Alpha
	Dim myArrOpAlpha(1, 6)
	myArrOpAlpha(0, 0) = myTObj.getTerm(430)
	myArrOpAlpha(1, 0) = "EXACT"
	myArrOpAlpha(0, 1) = myTObj.getTerm(431)
	myArrOpAlpha(1, 1) = "STARTS"
	myArrOpAlpha(0, 2) = myTObj.getTerm(438)
	myArrOpAlpha(1, 2) = "ENDS"
	myArrOpAlpha(0, 3) = myTObj.getTerm(432)
	myArrOpAlpha(1, 3) = "CONTAINS"
	myArrOpAlpha(0, 4) = "LIKE"
	myArrOpAlpha(1, 4) = "LIKE"
	myArrOpAlpha(0, 5) = "IS NULL"
	myArrOpAlpha(1, 5) = "IS NULL"
	myArrOpAlpha(0, 6) = "IS NOT NULL"
	myArrOpAlpha(1, 6) = "IS NOT NULL"

	' Text
	Dim myArrOpText(1, 5)
	myArrOpText(0, 0) = myTObj.getTerm(431)
	myArrOpText(1, 0) = "STARTS"
	myArrOpText(0, 1) = myTObj.getTerm(438)
	myArrOpText(1, 1) = "ENDS"
	myArrOpText(0, 2) = myTObj.getTerm(432)
	myArrOpText(1, 2) = "CONTAINS"
	myArrOpText(0, 3) = "LIKE"
	myArrOpText(1, 3) = "LIKE"
	myArrOpText(0, 4) = "IS NULL"
	myArrOpText(1, 4) = "IS NULL"
	myArrOpText(0, 5) = "IS NOT NULL"
	myArrOpText(1, 5) = "IS NOT NULL"

	' Binary
	Dim myArrOpBinary(0, 1)
	myArrOpBinary(0, 0) = "IS NULL"
	myArrOpBinary(0, 1) = "IS NOT NULL"

	' Boolean
	Dim myArrOpBool(0, 3)
	myArrOpBool(0, 0) = "IS NULL"
	myArrOpBool(0, 1) = "IS NOT NULL"
	myArrOpBool(0, 2) = "IS TRUE"
	myArrOpBool(0, 3) = "IS FALSE"

	' UNIQUE IDENTIFIER
	Dim myArrOpUId(0, 0)
	myArrOpUId(0, 0) = "="

%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
	<!--
	function openCalendar(pURL, pFormName, pControlName) {
		var myURL = pURL + "?form=" + pFormName + "&elt=" + pControlName + "&initialDate=" + escape(eval("document." + pFormName + "." + pControlName).value);
		openPopUp(myURL, "mlc_popup", 180, 180, 10, 10);
		return (true);
	}

	function checkall() {
		var myObj;
		var myCheck = (!(document.mla_search.mla_display_0.checked));
		for (var i = 0; i <= <% = myColCount %>; i++) {
			myObj = eval("document.mla_search.mla_display_" + i);
			myObj.checked = myCheck;
		}
		return (true);
	}
	//-->
	</SCRIPT>

	<P CLASS="treeinfo"><% = myStrTree %></P>
	
	<FORM NAME="mla_search" METHOD=POST ACTION="search.asp?nid=<% = myNodeId %>&db=<% = Server.URLEncode(myDbName) %>&obj=<% = Server.URLEncode(myObjName) %>&type=<% = myObjType %>">
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Search Form">
		<TR><TD CLASS="caption"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/search.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Search"> <% = myTObj.getTerm(154) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Search Form">
				<THEAD>
					<TR>
						<TD CLASS="collabel"><% = myTObj.getTerm(435) %></TD>
						<TD CLASS="collabel" COLSPAN=3><% = myTObj.getTerm(436) %></TD>
						<TD CLASS="collabel"><% = myTObj.getTerm(437) %> <A HREF=# onclick="return(checkall());"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/uncheck.gif" WIDTH="12" HEIGHT="12" BORDER=0 ALIGN="MIDDLE" ALT="Uncheck"></A></TD>
					</TR>
				</THEAD>
				<TBODY>
				<%
					Set myObjStr = New mlt_string
					For i = 0 To myColCount
						If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
						myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vbCrLf
						myObjStr.strAppend "<TD CLASS=""formlabel"">" & myArrCol(0, i) & "</TD>"
						Select Case myArrCol(3, i) 
							Case 1 : ' Alpha
								myObjStr.strAppend "<TD CLASS=""forminfo"">" & getListBox("mla_bool_" & i, myArrBool, 0, 0, " ", "", "", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">"  & getListBox("mla_op_" & i, myArrOpAlpha, 1, 0, " ", "", "alphanumeric", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo""><INPUT TYPE=""text"" NAME=""mla_value_" & i & """ CLASS=""alpha_edit""></TD>"
							Case 2 : ' Text
								myObjStr.strAppend "<TD CLASS=""forminfo"">" & getListBox("mla_bool_" & i, myArrBool, 0, 0, " ", "", "", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">"  & getListBox("mla_op_" & i, myArrOpText, 1, 0, " ", "", "alphanumeric", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo""><INPUT TYPE=""text"" NAME=""mla_value_" & i & """ CLASS=""alpha_edit""></TD>"
							Case 3 : ' Binary
								myObjStr.strAppend "<TD CLASS=""forminfo"">&nbsp;</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">"  & getListBox("mla_op_" & i, myArrOpBinary, 0, 0, " ", "", "alphanumeric", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">&nbsp;</TD>"
							Case 4 : ' Date
								myObjStr.strAppend "<TD CLASS=""forminfo"">&nbsp;</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">"  & getListBox("mla_op_" & i, myArrOpNumeric, 0, 0, " ", "", "alphanumeric", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo""><INPUT TYPE=""text"" NAME=""mla_value_" & i & """ CLASS=""num_edit"">"
								myObjStr.strAppend "<A HREF=# onclick=""openCalendar('mlc_popup.asp', 'mla_search',  'mla_value_" & i & "'); return false;""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/calendar.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALIGN=""MIDDLE"" ALT=""Calendar""></A>"
								myObjStr.strAppend " <B>AND</B> "
								myObjStr.strAppend "<INPUT TYPE=""text"" NAME=""mla_value2_" & i & """ CLASS=""num_edit"">"
								myObjStr.strAppend "<A HREF=# onclick=""openCalendar('mlc_popup.asp', 'mla_search', 'mla_value2_" & i & "'); return false;""><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/action/calendar.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALIGN=""MIDDLE"" ALT=""Calendar""></A>"
								myObjStr.strAppend "</TD>"
							Case 5 : ' Boolean
								myObjStr.strAppend "<TD CLASS=""forminfo"">&nbsp;</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">"  & getListBox("mla_op_" & i, myArrOpBool, 0, 0, " ", "", "alphanumeric", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">&nbsp;</TD>"
							Case 6 : ' Unique Identifier
								myObjStr.strAppend "<TD CLASS=""forminfo"">&nbsp;</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">"  & getListBox("mla_op_" & i, myArrOpUId, 0, 0, " ", "", "alphanumeric", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo""><INPUT TYPE=""text"" NAME=""mla_value_" & i & """ CLASS=""alpha_edit""></TD>"
							Case Else : ' numeric
								myObjStr.strAppend "<TD CLASS=""forminfo"">&nbsp;</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo"">"  & getListBox("mla_op_" & i, myArrOpNumeric, 0, 0, " ", "", "alphanumeric", "") & "</TD>"
								myObjStr.strAppend "<TD CLASS=""forminfo""><INPUT TYPE=""text"" NAME=""mla_value_" & i & """ CLASS=""num_edit"">"
								myObjStr.strAppend " <B>AND</B> "
								myObjStr.strAppend "<INPUT TYPE=""text"" NAME=""mla_value2_" & i & """ CLASS=""num_edit""></TD>"
						End Select
						myObjStr.strAppend "<TD CLASS=""forminfo""><INPUT TYPE=""checkbox"" NAME=""mla_display_" & i & """ CHECKED></TD>"
						myObjStr.strAppend "</TR>" & vbCrLf
					Next
					Response.Write myObjStr.getStr()
					Set myObjStr = Nothing
				%>	
				</TBODY>
			</TABLE>
		</TD></TR>
			<TR><TD>&nbsp;</TD></TR>
			<TR>
				<TD ALIGN=CENTER>
			<INPUT TYPE="radio" NAME="mla_match" VALUE="AND" ID="all" CHECKED><LABEL FOR="all"><% = myTObj.getTerm(433) %></LABEL>
			&nbsp;
			<INPUT TYPE="radio" NAME="mla_match" VALUE="OR" ID="any"><LABEL FOR="any"><% = myTObj.getTerm(434) %></LABEL>
				</TD>
			</TR>
			<TR>
				<TD ALIGN=CENTER>
					<INPUT TYPE="reset" VALUE="<% = myTObj.getTerm(50) %>" NAME="mla_reset"> &nbsp;
					<INPUT TYPE="submit" VALUE="<% = myTObj.getTerm(54) %>" NAME="mla_submit">
				</TD>
			</TR>
	</TABLE>
	</FORM>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->