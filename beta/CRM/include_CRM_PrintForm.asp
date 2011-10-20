<!--#include File="../include_UtilFunctions.asp"-->
	  <Tr><Td colspan="2" height="10px">
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable {font-family:tahoma; width:80%; border:1 solid black; direction: RTL; background-color:black;}
	.CustTable td {padding:5;}
	.CustTable a {text-decoration:none;color:#000088}
	.CustTable a:hover {text-decoration:underline;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	.CusTD3 {background-color: #DDDDDD; direction: LTR; text-align: center; font-size:9pt;}
	.CusTD4 {background-color: #CCCC66; direction: LTR; text-align: center; font-size:9pt;}
</STYLE>


	  &nbsp;</Td></Tr>
	  <Tr>
		<Td colspan="2" valign="top" align="center">
		<%
				'----   It's Print ( Crystal Reports )
		 %>
		<BR>
		<BR>
		<CENTER>
		<% 	ReportLogRow = PrepareReport ("CRM_PrintForm.rpt", "CustomerID", clng(request("selectedCustomer")), "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT TYPE="button" value=" █г│ бояс " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

		</CENTER>

		<BR><iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:650; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no ></iframe>
				<BR><BR>
		</Td>
	  </Tr>