<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ReportViewer.ascx.cs" Inherits="UC_ReportViewer" %>
<%@ Register Assembly="CrystalDecisions.Web, Version=10.2.3600.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" Namespace="CrystalDecisions.Web" TagPrefix="CR" %>
<CR:CrystalReportViewer ID="CrystalReportViewer1" runat="server" AutoDataBind="true"
    DisplayGroupTree="False" EnableDatabaseLogonPrompt="False" PrintMode="ActiveX"
    ReportSourceID="CrystalReportSource1" EnableDrillDown="False" HasCrystalLogo="False" HasDrillUpButton="False" HasToggleGroupTreeButton="False" HasViewList="False" HasZoomFactorList="False" Height="50px" SeparatePages="True" Width="350px" DisplayToolbar="True" HasGotoPageButton="False" HasSearchButton="False" ReuseParameterValuesOnRefresh="True" />
<CR:CrystalReportSource ID="CrystalReportSource1" runat="server"  />
