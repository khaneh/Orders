<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "äãÇíÔ ßåÇí ÏÑíÇÝÊí"
SubmenuItem=3
if not Auth("A" , 3) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_UtilFunctions.asp"-->


<BR><BR><BR>	
<CENTER>
		<% 	ReportLogRow = PrepareReport ("CheqRcpt.rpt", "cheqID", request("cheqID"), "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT TYPE="button" value=" Ç " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">
</CENTER>

<BR>
<iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:700; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no></iframe>


<!--#include file="tah.asp" -->