<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myDbName, myNodeId, myArrRule, myRuleCount, i, myObjStr, myStrClass, myStrTree, myRuleName, myRuleTxt
	myNodeId = Request.QueryString("nid")
	myDbName = Request.QueryString("db")
	openConnection
	myStrTree = getTreeStr(Mid(myNodeId, 2) & "_6", Array(myDbName))
	myArrRule = getDbRule(myDbName)
	closeConnection
	If isArray(myArrRule) Then myRuleCount = UBound(myArrRule, 2) Else myRuleCount = -1 End If
%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="content" SUMMARY="Database Information">
		<TR><TD CLASS="caption">
		<IMG SRC="../../themes/<% = mla_cfg_theme %>/images/mylittletree/rule.gif" WIDTH="16" HEIGHT="16" BORDER=0 ALIGN="MIDDLE" ALT="Rules"> <% = myTObj.getTerm(12) %></TD></TR>
		<TR><TD>
			<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2 ALIGN=CENTER CLASS="content" SUMMARY="Rule List">
				<THEAD>
					<TR><TD CLASS="collabel"><% = myTObj.getTerm(180) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(121) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(122) %></TD><TD CLASS="collabel"><% = myTObj.getTerm(181) %></TD></TR>
				</THEAD>
				<TBODY>
					<%
						Set myObjStr = New mlt_string
						For i = 0 To myRuleCount
							myRuleName = "[" & rembracket(myArrRule(1, i)) & "].[" & rembracket(myArrRule(0, i)) & "]"
							myRuleTxt = txt2html(Trim(myArrRule(3, i)))
							If Left(myRuleTxt, 4) = "<BR>" Then myRuleTxt = Mid(myRuleTxt, 5)
							If i MOD 2 = 0 Then myStrClass = "odd" Else myStrClass = "even" End If
							myObjStr.strAppend "<TR CLASS=""" & myStrClass & """>" & vBCrLf
							myObjStr.strAppend "<TD><IMG SRC=""../../themes/" & mla_cfg_theme & "/images/mylittletree/rule.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=0 ALT=""" & myArrRule(0, i) & """ ALIGN=""MIDDLE""> "
							myObjStr.strAppend "<B>" & myArrRule(0, i) & "</B></TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrRule(1, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myArrRule(2, i) & "</TD>" & vbCrLf
							myObjStr.strAppend "<TD>" & myRuleTxt & "</TD>" & vbCrLf
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