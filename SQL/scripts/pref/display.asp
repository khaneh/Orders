<!-- #INCLUDE FILE="../inc/mla_sql_include.asp" -->
<%
	Dim myStrTree, myArrDay, i

	If Request.Form("mla_cfg_cancel") <> "" Then
		Response.Redirect "default.asp"
		Response.End
	End If

	If Request.Form("mla_cfg_submit") <> "" Then
		Response.Cookies("Option")("ShowSysDatabase") = (Request.Form("mla_cfg_showsysdatabases") <> "")
		Response.Cookies("Option")("ShowSysTable") = (Request.Form("mla_cfg_showsystables") <> "")
		Response.Cookies("Option")("ShowSysView") = (Request.Form("mla_cfg_showsysviews") <> "")
		Response.Cookies("Option")("ShowSysProcedure") = (Request.Form("mla_cfg_showsysprocedures") <> "")
		Response.Cookies("Option")("ShowSysFunction") = (Request.Form("mla_cfg_showsysfunctions") <> "")
		Response.Cookies("Option")("PageSize") = Request.Form("mla_cfg_pagesize") 
		Response.Cookies("Option")("MaxDisplayedChar") = Request.Form("mla_cfg_maxdisplayedchar") 
		Response.Cookies("Option")("MaxDisplayedBin") = Request.Form("mla_cfg_maxdisplayedbin") 
		Response.Cookies("Option")("RowDelimiter") = Request.Form("mla_cfg_rowdelimiter") 
		Response.Cookies("Option")("FirstDayOfWeek") = Request.Form("mla_cfg_firstdayofweek") 
		Response.Cookies("Option").Expires = #1/1/2020#
		Response.Redirect "default.asp?refresh=1"
		Response.End
	End If

	myStrTree = getTreeStr("3_3", Array())

	myArrDay = Split(myTObj.getTerm(443), ",", -1, 1)

%>

<!-- #INCLUDE FILE="../inc/metaheader.asp" -->
<BODY>
	<P CLASS="treeinfo"><% = myStrTree %></P>

	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
	<!--
	function checkForm(pForm) {
		if (document.nocheck) return (true);
		if (isEmpty(pForm.mla_cfg_pagesize)) {
			pForm.mla_cfg_pagesize.focus();
			return (false);
		}
		if (!isInteger(pForm.mla_cfg_pagesize)) {
			pForm.mla_cfg_pagesize.focus();
			return (false);
		}
		if (isEmpty(pForm.mla_cfg_maxdisplayedchar)) {
			pForm.mla_cfg_maxdisplayedchar.focus();
			return (false);
		}
		if (isEmpty(pForm.mla_cfg_maxdisplayedbin)) {
			pForm.mla_cfg_maxdisplayedbin.focus();
			return (false);
		}
		if (isEmpty(pForm.mla_cfg_rowdelimiter)) {
			pForm.mla_cfg_rowdelimiter.focus();
			return (false);
		}
		if (isEmpty(pForm.mla_cfg_firstdayofweek)) {
			pForm.mla_cfg_firstdayofweek.focus();
			return (false);
		}
		if (!isInteger(pForm.mla_cfg_maxdisplayedchar)) {
			pForm.mla_cfg_maxdisplayedchar.focus();
			return (false);
		}
		if (!isInteger(pForm.mla_cfg_maxdisplayedbin)) {
			pForm.mla_cfg_maxdisplayedbin.focus();
			return (false);
		}
		if (!isInteger(pForm.mla_cfg_firstdayofweek)) {
			pForm.mla_cfg_firstdayofweek.focus();
			return (false);
		}
		return (true);
	}
	//-->
	</SCRIPT>

	<FORM NAME="mla_cfg" METHOD=POST ACTION="display.asp" onSubmit = "return (checkForm(this));">
		<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 CLASS="hcontent" SUMMARY="Option Form">
			<TR><TD CLASS="caption" COLSPAN=2><% = myTObj.getTerm(22) & " \ " &  myTObj.getTerm(25) %></TD></TR>
			<TR>
				<TD CLASS="formlabel"><% = myTObj.getTerm(480) %> :</TD>
				<TD>&nbsp;</TD>
			</TR>
			<TR>
				<TD>&nbsp;</TD>			
				<TD><INPUT TYPE="checkbox" NAME="mla_cfg_showsysdatabases" ID="mla_cfg_showsysdatabases"> <LABEL FOR="mla_cfg_showsysdatabases"><% = myTObj.getTerm(481) %></LABEL></TD>
			</TR>
			<TR>
				<TD>&nbsp;</TD>			
				<TD><INPUT TYPE="checkbox" NAME="mla_cfg_showsystables" ID="mla_cfg_showsystables"> <LABEL FOR="mla_cfg_showsystables"><% = myTObj.getTerm(482) %></LABEL></TD>
			</TR>
			<TR>
				<TD>&nbsp;</TD>			
				<TD><INPUT TYPE="checkbox" NAME="mla_cfg_showsysviews" ID="mla_cfg_showsysviews"> <LABEL FOR="mla_cfg_showsysviews"><% = myTObj.getTerm(483) %></LABEL></TD>
			</TR>
			<TR>
				<TD>&nbsp;</TD>			
				<TD><INPUT TYPE="checkbox" NAME="mla_cfg_showsysprocedures" ID="mla_cfg_showsysprocedures"> <LABEL FOR="mla_cfg_showsysprocedures"><% = myTObj.getTerm(484) %></LABEL></TD>
			</TR>
			<TR>
				<TD>&nbsp;</TD>			
				<TD><INPUT TYPE="checkbox" NAME="mla_cfg_showsysfunctions" ID="mla_cfg_showsysfunctions"> <LABEL FOR="mla_cfg_showsysfunctions"><% = myTObj.getTerm(485) %></LABEL></TD>
			</TR>
			<TR><TD COLSPAN=2>&nbsp;</TD></TR>
			<TR>
				<TD CLASS="formlabel"><% = myTObj.getTerm(486) %> :</TD>
				<TD><INPUT TYPE="text" NAME="mla_cfg_pagesize" CLASS="numeric" TITLE=""></TD>
			</TR>
			<TR>
				<TD CLASS="formlabel"><% = myTObj.getTerm(487) %> :</TD>
				<TD><INPUT TYPE="text" NAME="mla_cfg_maxdisplayedchar" CLASS="numeric" TITLE=""></TD>
			</TR>
			<TR>
				<TD CLASS="formlabel"><% = myTObj.getTerm(488) %> :</TD>
				<TD><INPUT TYPE="text" NAME="mla_cfg_maxdisplayedbin" CLASS="numeric" TITLE=""></TD>
			</TR>
			<TR>
				<TD CLASS="formlabel"><% = myTObj.getTerm(489) %> :</TD>
				<TD><INPUT TYPE="text" NAME="mla_cfg_rowdelimiter" CLASS="numeric" TITLE=""></TD>
			</TR>
			<TR>
				<TD CLASS="formlabel"><% = myTObj.getTerm(490) %> :</TD>
				<TD>
					<SELECT NAME="mla_cfg_firstdayofweek" CLASS="alphanumeric">
					<%
						For i = 0 To 6
							Response.Write "<OPTION VALUE=" & (i+1) & ">" & myArrDay(i) & "</OPTION>" & vbCrLf
						Next
					%>
					</SELECT>
				</TD>
			</TR>
			<TR><TD COLSPAN=2>&nbsp;</TD></TR>
			<TR>
				<TD COLSPAN=2 ALIGN=CENTER>
					<INPUT TYPE="submit" VALUE="<% = myTObj.getTerm(51) %>" NAME="mla_cfg_cancel" onClick="document.nocheck = true;"> &nbsp;
					<INPUT TYPE="submit" VALUE="<% = myTObj.getTerm(54) %>" NAME="mla_cfg_submit">
				</TD>
			</TR>
		</TABLE>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
	<!--
		document.mla_cfg.mla_cfg_showsysdatabases.checked = <% = CInt(mla_cfg_showsysdatabases) %>;
		document.mla_cfg.mla_cfg_showsystables.checked = <% = CInt(mla_cfg_showsystables) %>;
		document.mla_cfg.mla_cfg_showsysviews.checked = <% = CInt(mla_cfg_showsysviews) %>;
		document.mla_cfg.mla_cfg_showsysprocedures.checked = <% = CInt(mla_cfg_showsysprocedures) %>;
		document.mla_cfg.mla_cfg_showsysfunctions.checked = <% = CInt(mla_cfg_showsysfunctions) %>;
		document.mla_cfg.mla_cfg_pagesize.value = "<% = mla_cfg_pagesize %>";
		document.mla_cfg.mla_cfg_maxdisplayedchar.value = "<% = mla_cfg_maxdisplayedchar %>";
		document.mla_cfg.mla_cfg_maxdisplayedbin.value = "<% = mla_cfg_maxdisplayedbin %>";
		document.mla_cfg.mla_cfg_rowdelimiter.value = "<% = mla_cfg_rowdelimiter %>";
		setSelectValue(document.mla_cfg.mla_cfg_firstdayofweek, "<% = mla_cfg_firstdayofweek %>", 1)
	//-->
	</SCRIPT>

</BODY>
</HTML>
<!-- #INCLUDE FILE="../inc/mla_sql_end.asp" -->