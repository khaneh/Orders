<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myNodeId, myDbName, myTblName, myTrName, myTrTxt, myStrTree
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	myTblName = Request.QueryString("tbl")
	myTrName = Request.QueryString("tr")
	openConnection
	'myStrTree = getTreeStr(Array(myTObj.getTerm(5), myDbName, myTObj.getTerm(7), myTblName, myTObj.getTerm(328), myTrName))
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_1_2_1", Array(myDbName, myTblName, myTrName))
	myTrTxt = txt2html(getObjTxt(myDbName, myTrName))
	closeConnection
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption"><% = myTObj.getTerm(329) & " " & myTrName %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" WIDTH="100%" SUMMARY="Triigger Definition">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(330) %></TD></TR>
				</THEAD>
				<TBODY>
					<TR><TD><% = myTrTxt %></TD></TR>
				</TBODY>
			</TABLE>
			
		</TD></TR>
	</TABLE>
</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->