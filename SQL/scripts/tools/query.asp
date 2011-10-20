<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree
	Dim myDbTbl
	Dim myObjRS, myStrRet, myStrQuery, Field, myArrBinary()
	Dim myObjStr, myStrClass, i, j, myStrValue, myArrTmp
	Dim myMaxCount

	openConnection

	If Request.Form("mla_query_submit") <> "" Then
		myStrQuery = Request.Form("mla_query_value")
		If myStrQuery <> "" Then
			Set myObjStr = New mlt_string
			Set myObjRS = Server.Createobject("ADODB.Recordset")
			myObjRS.ActiveConnection = gObjDC
			myObjRS.CursorLocation=3
			If Request.Form("mla_query_plan") <> "" Then 
				myObjRS.LockType = 1
			Else
				myObjRS.LockType = 3
			End If
			myMaxCount = CLng(Request.Form("mla_rec_count"))
			On Error Resume Next
			gObjDC.execute "USE [" & rembracket(Request.Form("mla_query_db")) & "]"
			If Request.Form("mla_query_plan") <> "" Then gObjDC.execute "SET SHOWPLAN_TEXT ON"
			myObjRS.Open myStrQuery
			If Err < 0 Then
				If Request.Form("mla_query_plan") <> "" Then gObjDC.execute "SET SHOWPLAN_TEXT OFF"
				Set myObjRS = Nothing
				Call mla_displayError(Err, myStrQuery)
			End If				
			Do Until myObjRS Is Nothing
				If myObjRS.Properties("Asynchronous Rowset Processing") = 16 Then
					myObjStr.strAppend "<P>" & vbCrLf
					myObjStr.strAppend "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS=""content"" WIDTH=""100%"" SUMMARY=""Result Content"">"
					myObjStr.strAppend "<THEAD><TR>" & vbCrLf
					i = 0
					For Each Field In myObjRS.Fields
						ReDim Preserve myArrBinary(i)
						myObjStr.strAppend "<TD CLASS=""collabel"">" & Field.Name & "</TD>" & vbCrLf
						myArrBinary(i) = (Field.Type = 128 Or Field.Type =  204 Or Field.Type = 205)
						i = i + 1
					Next
					myObjStr.strAppend "</TR></THEAD>" & vbCrLf
					myObjStr.strAppend "<TBODY>" & vbCrLf
					i = 0
					Do While Not myObjRS.EOF
						If myMaxCount > 0 And i > myMaxCount Then Exit Do
						If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
						myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vbCrLf
						j = 0
						For Each Field In myObjRS.Fields
							If isNull(Field.Value) Then
								myStrValue = "<SPAN CLASS=""moreinfo"">(" & myTObj.getTerm(60) & ")</SPAN>"
							ElseIf myArrBinary(j) Then
								myArrTmp= bin2hex(Field.Value, mla_cfg_maxdisplayedbin)
								If myArrTmp(1) Then
									myStrValue = txt2html(myArrTmp(0)) &" <SPAN CLASS=""moreinfo"">(...)</SPAN>"
								Else
									myStrValue = txt2html(myArrTmp(0))
								End If
							Else
								If Request.Form("mla_query_plan") = "" Then
									myArrTmp= getStrBegin(CStr(Field.Value), mla_cfg_maxdisplayedchar)
									If myArrTmp(1) Then
										myStrValue = txt2html(myArrTmp(0)) & " <SPAN CLASS=""moreinfo"">(...)</SPAN>"
									Else
										myStrValue = txt2html(myArrTmp(0))
									End If
								Else
									myStrValue = txt2html(CStr(Field.Value))
								End If
							End If
							myObjStr.strAppend "<TD>" & myStrValue & "</TD>" & vbCrLf
							j = j + 1
						Next
						myObjStr.strAppend "</TR>" & vbCrLf
						i = i + 1
						myObjRS.MoveNext
					Loop
					myObjStr.strAppend "</TBODY>" & vbCrLf
					myObjStr.strAppend "</TABLE>" & vbCrLf
					myObjStr.strAppend myObjRS.RecordCount &  " " & myTObj.getTerm(420) & "</P>" & vbCrLf
				Else
					myObjStr.strAppend myTObj.getTerm(421) &"<BR>" & vbCrLf
				End If
				Set myObjRS = myObjRS.NextRecordset
			Loop
			Set myObjRS = Nothing
			If Request.Form("mla_query_plan") <> "" Then gObjDC.execute "SET SHOWPLAN_TEXT OFF"
			myStrRet = myObjStr.getStr()
			Set myObjStr = Nothing
		End If
	End If
	myDbTbl = getDbList()
	myStrTree = getTreeStr("2_4", Array()) & " \ " & myTObj.getTerm(31) 
	closeConnection
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
	<!--
	function resetForm() {
		document.mla_query.mla_query_value.value=""; 
		document.mla_query.mla_query_value.focus(); 
		return (true);
	}

	function checkForm(pForm) {
		if (document.nocheck) return (true);
		if (isEmpty(pForm.mla_query_value)) {
			pForm.mla_query_value.focus();
			return (false);
		}
		return (true);
	}
	//-->
	</SCRIPT>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<FORM NAME="mla_query" METHOD=POST ACTION="query.asp" onSubmit = "return (checkForm(this));">
		<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Query Analyser">
			<TR><TD CLASS="caption"><A HREF=# onclick="openPopUp('query.asp', '<% = Replace(time(), ":", "") %>', 600, 400, 10, 10);"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/action/new.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="RIGHT" ALT="New"></A><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/query.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Query Analyser"> <% = myTObj.getTerm(31) %></TD></TR>
			<TR>
				<TD>
					<B><% = myTObj.getTerm(120) %></B> : <% = getListBox("mla_query_db", myDbTbl, 0, 0, "", Request.Form("mla_query_db"), "alphanumeric", "") %>
					&nbsp;
					<B><% = myTObj.getTerm(422) %></B> : <INPUT TYPE="checkbox" NAME="mla_query_plan">
					&nbsp;
					<B><% = myTObj.getTerm(423) %></B>   : 
						<SELECT NAME="mla_rec_count">
							<OPTION>50</OPTION>
							<OPTION>100</OPTION>
							<OPTION>500</OPTION>
							<OPTION>1000</OPTION>
							<OPTION VALUE="-1"><% = myTObj.getTerm(424) %></OPTION>
						</SELECT>
				</TD>
			</TR>
			<TR><TD><TEXTAREA NAME="mla_query_value" ROWS="10" COLS="30" CLASS="objtext_edit" TITLE="<% = myTObj.getTerm(31) %>"><% = Request.Form("mla_query_value") %></TEXTAREA></TD></TR>
			<TR><TD>&nbsp;</TD></TR>
			<TR>
				<TD ALIGN=CENTER>
					<INPUT TYPE="button" VALUE="<% = myTObj.getTerm(50) %>" NAME="mla_query_reset" onclick="resetForm();"> &nbsp;
					<INPUT TYPE="submit" VALUE="<% = myTObj.getTerm(54) %>" NAME="mla_query_submit">
				</TD>
			</TR>
		</TABLE>
	</FORM>

	<% = myStrRet %>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->