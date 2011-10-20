<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	body { font-family: tahoma; font-size: 8pt; background-color:#DDDDDD;}
	Input { font-family: tahoma; font-size: 9pt;}
	td { font-family: tahoma; font-size: 8pt;}
</style>
<TITLE> چاپ </TITLE>
<script language="JavaScript">
<!--
function documentKeyDown() {
	var theKey = event.keyCode;
	if (theKey == 27) { 
		window.close();
	}
}

document.onkeydown = documentKeyDown;

//-->
</script>
</HEAD>
<BODY>
<%
act=request.queryString("act")
if act="Fin" then
%>
	<div dir=RTL align=center>
		<br>
		<br>
		<br>
		<br> موفق شديم كه چاپ بكنيم! 
		<br>
		<br>
		<br>
		<br>
		<INPUT TYPE="button" onclick="window.close();" value=" Close ">
	</div>
<%
else
%>
<!--#Include file='config.asp' -->
<!--#include File="include_farsiDateHandling.asp"-->
<!--#include File="include_UtilFunctions.asp"-->
<%
	if act="print" OR act="view" then
		RFN=	request.queryString("RFN")
		RPNs=	request.queryString("RPNs")
		RPVs=	request.queryString("RPVs")
		RURL=	request.queryString("RURL")

		ReportFileName=			sqlSafe(RFN)
		ReportParameterNames=	sqlSafe(RPNs)
		ReportParameterValues=	sqlSafe(RPVs)
		ReturnURL=				sqlSafe(RURL)

		ReportLogRow = PrepareReport (ReportFileName, ReportParameterNames, ReportParameterValues, ReturnURL)

		if act="print" then
			response.redirect "/Reports/PrintReport.aspx?" & ReportLogRow
		else 
			response.redirect "/Reports/ViewReport.aspx?" & ReportLogRow
		end if

	else
	%>
		<div dir=RTL align=center>
			<br> چي؟؟؟ 
			<br> خطا ! خطا !<br> 
			<INPUT TYPE="button" onclick="window.close();" value=" Close ">
		</div>
	<%
	end if
end if
%>
</BODY>
</HTML>
