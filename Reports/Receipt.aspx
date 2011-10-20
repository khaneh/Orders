<%@ Page language="c#" Codebehind="Receipt.aspx.cs" AutoEventWireup="false" Inherits="Reports.ReceiptWebForm" %>
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
			<form id="Form1" method="post" runat="server" dir="rtl">
				<asp:button id="Button1" style="Z-INDEX: 102; LEFT: 625px; POSITION: absolute; TOP: 378px" runat="server" Font-Size="Small" Font-Names="Titr" Width="1px" Height="1px" Text="Ç ÝÑã"></asp:button><br> <!--center><select name="CrystalReportViewer1:_ctl1:_ctl0:_ctl13" onchange="__doPostBack('CrystalReportViewer1:_ctl1:_ctl0:_ctl13','')" language="javascript" style="height:23px;width:95px;">
					<option value="25">25%</option>
					<option value="50">50%</option>
					<option value="60" selected>60%</option>
					<option value="75">75%</option>
					<option value="100">100%</option>
					<option value="125">125%</option>
					<option value="150">150%</option>
					<option value="200">200%</option>
					<option value="300">300%</option>
					<option value="400">400%</option>
				</select></center><br--><CR:CRYSTALREPORTVIEWER id="CrystalReportViewer1" style="Z-INDEX: 100; LEFT: -24px; POSITION: absolute; TOP: 3px" runat="server" Width="1100px" Height="362px" DisplayToolbar="False" DisplayGroupTree="False" HasToggleGroupTreeButton="False" HasSearchButton="False" HasPageNavigationButtons="False" HasExportButton="False" HasGotoPageButton="False" HasDrillUpButton="False" BestFitPage="False" ToolTip="Order Form Preview"></CR:CRYSTALREPORTVIEWER></form>
		</center>
		<SCRIPT language="javascript">
	<!--
	thediv=document.getElementsByTagName("div")[0];
	thediv.style.overflow="hidden";
	//thediv.style.height=100;
	document.getElementById("PageFooterSection1").style.top=0;
	document.getElementById("PageFooterSection1").style.height=0;
	// -->
		</SCRIPT>
	</body>
</HTML>
