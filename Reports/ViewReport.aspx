<%@ Register TagPrefix="cr" Namespace="CrystalDecisions.Web" Assembly="CrystalDecisions.Web, Version=9.2.3300.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" %>
<%@ Page language="c#" Codebehind="ViewReport.aspx.cs" AutoEventWireup="false" Inherits="Reports.GeneralReportWebForm" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title></title>
		<meta content="Microsoft Visual Studio 7.0" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="JavaScript">
	<!--
	function documentKeyDown() {
		var theKey = event.keyCode;
		if (theKey == 27) { 
			window.close();
		}
	}

	document.onkceydown = documentKeyDown;

	//-->
		</script>
	</HEAD>
	<BODY bottomMargin="0" topMargin="0" MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<CR:CRYSTALREPORTVIEWER id="viewer" style="Z-INDEX: 100;" DisplayGroupTree="False" HasToggleGroupTreeButton="False" HasSearchButton="False" HasExportButton="False" HasGotoPageButton="False" HasDrillUpButton="False" BestFitPage="False" ToolTip="Report Preview" Width="780px" runat="server" HasPrintButton="False"></CR:CRYSTALREPORTVIEWER>
		</form>
		<SCRIPT language="javascript">
	<!--
/*		
	document.getElementById("PageFooterSection1").parentNode.removeChild(document.getElementById("PageFooterSection1"));
	
	document.getElementById("viewer").align="center";

	document.getElementsByTagName("div")[1].removeAttribute("style");
	//document.getElementsByTagName("div")[1].style.removeAttribute("overflow");
*/		
	// -->
		</SCRIPT>
	</BODY>
</HTML>
