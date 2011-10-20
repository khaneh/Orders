<%
'IMPORTANT NOTE:	
'		This include file needs "include_farsiDateHandling.asp" to be included beforehand.
'		Because it uses shamsiToday() function.

'---------------------------------------------
'------------- Prepare a Crystal Report to be printed / viewed
'---------------------------------------------
function PrepareReport (ReportFileName, ReportParameterNames, ReportParameterValues, ReturnURL)
	dim NewInvoicePrintForm
	printDate=				shamsiToday()
	printTime=				currentTime10()
	clientIP=				Request.ServerVariables("REMOTE_ADDR")

	mySQL="INSERT INTO ReportLog (PrintedBy, PrintDate, PrintTime, clientIP, ReportFileName, ReportParameterNames, ReportParameterValues, ReturnURL) VALUES('"& session("ID") & "',N'"& printDate & "', N'"& printTime & "', N'"& clientIP & "', N'"& ReportFileName & "', N'"& ReportParameterNames & "', N'"& ReportParameterValues & "', N'"& ReturnURL & "' ); SELECT @@Identity AS NewInvoicePrintForm"

	Set RS_ReportLog = conn.Execute(mySQL).NextRecordSet
	NewInvoicePrintForm = RS_ReportLog("NewInvoicePrintForm")

	RS_ReportLog.close
	Set RS_ReportLog=Nothing
	PrepareReport=NewInvoicePrintForm
end function

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function printThisReport(src,id){
	if (document.getElementById("f1")){
		//alert("?");
		window.frames['f1'].document.getElementsByName("ReportViewer1$CrystalReportViewer1$ctl02$ctl01")[0].click();
		
	}
	else{
		tempIframe=document.createElement("iframe");
		tempIframe.style.visibility="hidden";
		tempIframe.id="f1";
		tempIframe.width="10";
		tempIframe.height="10";
		tempIframe.src="/CRReports/?Id="+id+"&P=1";
		document.getElementsByTagName("Body")[0].appendChild(tempIframe);
	}
}
//-->
</SCRIPT>
