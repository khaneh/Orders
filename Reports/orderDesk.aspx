<%@ Page language="c#" Codebehind="orderDesk.aspx.cs" AutoEventWireup="false" Inherits="Reports.orderDeskWebForm" %>
<%@ Register TagPrefix="cr" Namespace="CrystalDecisions.Web" Assembly="CrystalDecisions.Web, Version=9.2.3300.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta content="Microsoft Visual Studio 7.0" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body bgColor="#c3dbeb" MS_POSITIONING="GridLayout">
		<center><font face="titr"></font>
			<form id="Form1" method="post" runat="server">
				<asp:button id="Button1" style="Z-INDEX: 102; LEFT: 600px; POSITION: absolute; TOP: 378px" runat="server" Text="Ç ÝÑã" Height="1px" Width="1px" Font-Names="Titr" Font-Size="Small"></asp:button><br>
				<CR:CRYSTALREPORTVIEWER id="CrystalReportViewer1" style="Z-INDEX: 100; LEFT: -17px; POSITION: absolute; TOP: 5px" runat="server" Height="362px" Width="749px" PageZoomFactor="90" ToolTip="Order Form Preview" BestFitPage="False" HasDrillUpButton="False" HasExportButton="False" HasSearchButton="False" HasToggleGroupTreeButton="False" DisplayGroupTree="False" HasPrintButton="False"></CR:CRYSTALREPORTVIEWER></form>
		</center>
		<SCRIPT language="javascript">
	<!--
	thediv=document.getElementsByTagName("div")[0];
	//thediv.style.overflow="hidden";
	//thediv.style.height=100;
	document.getElementById("PageFooterSection1").style.top=0;
	document.getElementById("PageFooterSection1").style.height=0;
	document.getElementById("PageFooterSection1").style.width=0;
	document.getElementById("PageFooterSection1").style.left=0;
	// -->
		</SCRIPT>
	</body>
</HTML>
