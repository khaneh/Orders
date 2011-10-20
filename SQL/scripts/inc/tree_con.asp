<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	If Not Session("isConnected") Then
		Response.Redirect "tree.asp"
		Response.End
	End If

	Dim myArrDb, myDbCount, myStrSrvName
	openConnection
	myArrDB = getDbList()
	myStrSrvName = getSrvName()
	closeConnection
	If isArray(myArrDb) Then myDbCount = UBound(myArrDb, 2) Else myDbCount = -1 End If

%>
<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>

<SCRIPT LANGUAGE="JavaScript" SRC="../js/mylittletree.js" TYPE="text/javascript"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
<!--
	var oTree = new myLittleTree("oTree", "../../themes/<% = mla_cfg_theme %>/images/mylittletree/"); 

	oTree.addFolder("M0", null, "<% = myTObj.getTerm(1) %>", "root.gif", "../hlp/default.asp", "Content");

	oTree.addFolder("M1", "M0", "<% = myTObj.getTerm(2) %>", "folder.gif", "../conn/default.asp", "Content");
		oTree.addFolder("M1_1", "M1", "<% = myTObj.getTerm(3) %>", "server.gif", "../conn/dsnless.asp", "Content");
		oTree.addFolder("M1_2", "M1", "<% = myTObj.getTerm(4) %>", "server.gif", "../conn/dsn.asp", "Content");
		oTree.addFolder("M1_3", "M1", "<% = myTObj.getTerm(32) %>", "disconnect.gif", "../../restart.asp", "_top");

	oTree.addFolder("M2", "M0", "<% = myStrSrvName %>", "server.gif", "../srv/srvinfo.asp", "Content");
	oTree.addFolder("M2_1", "M2", "<% = myTObj.getTerm(5) %>", "folder.gif", "../db/default.asp", "Content");
	<%
		Dim i, myObjStr
		Set myObjStr = New mlt_string
		For i = 0 To myDbCount
			myObjStr.strAppend vbTab & "oTree.addFolder(""M2_1_" & i+1 & """, ""M2_1"", """ & myArrDb(0, i) & """, ""database.gif"", ""../db/dbinfo.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(1, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_1"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(7) & """, ""table.gif"", ""../db/dbtbl.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(2, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_2"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(8) & """, ""view.gif"", ""../db/dbview.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(3, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_3"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(9) & """, ""sp.gif"", ""../db/dbsp.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(4, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_4"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(10) & """, ""user.gif"", ""../db/dbuser.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(5, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_5"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(11) & """, ""role.gif"", ""../db/dbrole.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(6, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_6"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(12) & """, ""rule.gif"", ""../db/dbrule.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(7, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_7"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(13) & """, ""default.gif"", ""../db/dbdefault.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(8, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_8"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(14) & """, ""udt.gif"", ""../db/dbudt.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
			If mla_auth(9, 0) Then myObjStr.strAppend vbTab & vbTab & "oTree.addFolder(""M2_1_" & i+1 & "_9"", ""M2_1_" & i+1 & """, """ & myTObj.getTerm(16) & """, ""udf.gif"", ""../db/dbudf.asp?nid=M2_1_" & i+1 & "&db=" & Server.URLEncode(myArrDb(0, i)) & """, ""Content"");" & vbCrLf
		Next
		Response.Write myObjStr.getStr()
		Set myObjStr = Nothing
	%>
	
	<% If mla_auth(11, 0) Or mla_auth(12, 0) Then %>
		oTree.addFolder("M2_3", "M2", "<% = myTObj.getTerm(19) %>", "folder.gif", "../srv/default.asp", "Content");
		<% If mla_auth(11, 0) Then %>
			oTree.addFolder("M2_3_1", "M2_3", "<% = myTObj.getTerm(230) %>", "connection.gif", "../srv/srvcon.asp", "Content");
		<% End If %>
		<% If mla_auth(12, 0) Then %>
			oTree.addFolder("M2_3_2", "M2_3", "<% = myTObj.getTerm(231) %>", "srvrole.gif", "../srv/srvrole.asp", "Content");
		<% End If %>
	<% End If %>

	<% If mla_auth(13, 0) Then %>
		oTree.addFolder("M2_4", "M2", "<% = myTObj.getTerm(30) %>", "folder.gif", "../tools/default.asp", "Content");
			oTree.addFolder("M2_4_1", "M2_4", "<% = myTObj.getTerm(31) %>", "query.gif", "../tools/query.asp", "Content");
	<% End If %>

	<% If mla_auth(14, 0) Or mla_auth(14, 1) Or mla_auth(14, 2) Then %>
		oTree.addFolder("M10", "M0", "<% = myTObj.getTerm(22) %>", "folder.gif", "../pref/default.asp", "Content");
		<% If mla_auth(14, 0) Then %>
			oTree.addFolder("M10_1", "M10", "<% = myTObj.getTerm(23) %>", "", "../pref/theme.asp", "Content");
		<% End If %>
		<% If mla_auth(14, 1) Then %>
			oTree.addFolder("M10_2", "M10", "<% = myTObj.getTerm(24) %>", "", "../pref/language.asp", "Content");
		<% End If %>
		<% If mla_auth(14, 2) Then %>
			oTree.addFolder("M10_3", "M10", "<% = myTObj.getTerm(25) %>", "", "../pref/display.asp", "Content");
		<% End If %>
	<% End If %>

	document.write (oTree.createHTML());

	oTree.expandNode("M2"); 
	oTree.expandNode("M2_1"); 

	if (window.parent.frames['Content']) window.parent.frames['Content'].document.location = "../srv/srvinfo.asp";

//-->
</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->