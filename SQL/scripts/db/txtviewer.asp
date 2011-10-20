<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myTblName, myFieldName, myStrCnd, myStrSQL, myArrRow

	myDbName = Request.QueryString("db")
	myTblName = Request.QueryString("tbl")
	myFieldName = Request.QueryString("fld")
	myStrCnd = Request.QueryString("pk")

	openConnection
	myStrSQL = "SELECT " & myFieldName & " FROM " & myTblName & " WHERE " & myStrCnd
	myArrRow = getArrFromSQL(myStrSQL)
	closeConnection

%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/table.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Content"> <% = myTObj.getTerm(459) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=1 CLASS="content" SUMMARY="Query">
				<THEAD>
				<TR><TD CLASS="collabel"><% = myFieldName %></TD></TR>
				</THEAD>
				<TBODY>
				<TR><TD>
				<%
					If isArray(myArrRow) Then
						Response.Write txt2html(myArrRow(0, 0)) 
					Else
						Response.Write "&nbsp;"
					End If
				%>
				</TD></TR>
				</TBODY>
			</TABLE>
		</TD></TR>
	</TABLE>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->