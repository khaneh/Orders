<%@ page language="C#" autoeventwireup="true" inherits="_Default, App_Web_jtpuooaw" %>
<%@ Register Src="UC/ReportViewer.ascx" TagName="ReportViewer" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
</head>
<body onload="<%=OnLoad%>">
<script>
function doPrint()
{
    document.getElementById('ReportViewer1$CrystalReportViewer1$ctl02$ctl01').click();
}
</script>
    <form id="form1" runat="server">
    <div>
        <uc1:ReportViewer ID="ReportViewer1" runat="server" />
    </div>
    </form>
</body>
</html>
