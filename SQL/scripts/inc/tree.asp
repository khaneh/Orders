<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	If Session("isConnected") Then
		Response.Redirect "tree_con.asp"
		Response.End
	End If
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

	oTree.addFolder("M10", "M0", "<% = myTObj.getTerm(22) %>", "folder.gif", "../pref/default.asp", "Content");
		oTree.addFolder("M10_1", "M10", "<% = myTObj.getTerm(23) %>", "", "../pref/theme.asp", "Content");
		oTree.addFolder("M10_2", "M10", "<% = myTObj.getTerm(24) %>", "", "../pref/language.asp", "Content");
		oTree.addFolder("M10_3", "M10", "<% = myTObj.getTerm(25) %>", "", "../pref/display.asp", "Content");

	document.write (oTree.createHTML());

	if (window.parent.frames['Content']) window.parent.frames['Content'].document.location = "../hlp/default.asp";

//-->
</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->