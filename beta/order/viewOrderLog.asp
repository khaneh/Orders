<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<%
mySQL="SELECT OrderTraceLog.*, Users.RealName FROM OrderTraceLog INNER JOIN Users ON OrderTraceLog.InsertedBy = Users.ID WHERE (OrderTraceLog.ID='"& request("logid") & "')"
set RS1=conn.execute (mySQL)
if RS1.EOF then
	response.write "<BR><BR><BR><BR><CENTER>����� ����</CENTER>"
	response.end
end if
%>
<body topmargin=0 leftmargin=0 >
	<TABLE  border="0" cellspacing="0" cellpadding="2" align="center" style="background-color:#CCCCCC; color:black; direction:RTL; width:700;font-family:tahoma; font-size: 8pt">
	<TR bgcolor="black" height=30 style="color:yellow;">
		<TD align="left">����� �����:</TD>
		<TD align="right"><%=RS1("Order")%></TD>
		<TD align="left">����� ������ :</TD>
		<TD><span dir="LTR"><%=RS1("InsertedDate")%> (<%=RS1("InsertedTime")%>)</span></TD>
		<TD align="left">����:</TD>
		<TD align="right"><%=RS1("RealName")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">��� ����:</TD>
		<TD><%=RS1("company_name")%></TD>
		<TD align="left">���� �����:</TD>
		<TD><%=RS1("return_date")%></TD>
		<TD align="left">���� �����:</TD>
		<TD><%=RS1("return_time")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">��� �����:</TD>
		<TD><%=RS1("customer_name")%></TD>
		<TD align="left">��� �����:</TD>
		<TD><%=RS1("order_kind")%></TD>
		<TD align="left">����� ������:</TD>
		<TD><%=RS1("salesperson")%>	</TD>
	</TR>
	<TR height=30>
		<TD align="left">����:</TD>
		<TD><%=RS1("telephone")%></TD>
		<TD align="left">����� ��� ���� ����:</TD>
		<TD colspan="4"><%=RS1("order_title")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">�����:</TD>
		<TD><%=RS1("qtty")%></TD>
		<TD align="left">����:</TD>
		<TD><%=RS1("PaperSize")%></TD>
		<TD align="left">����/����:</TD>
		<TD><%=RS1("SimplexDuplex")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">�����:</TD>
		<TD><%=RS1("StepText")%></TD>
		<TD align="left">�����:</TD>
		<TD colspan="3"><%=RS1("StatusText")%></TD>
	</TR>
	<TR height=30>
		<TD align="left">���� ��:</TD>
		<TD><%=RS1("Price")%></TD>
		<TD colspan="4" height="30px">&nbsp;</TD>
	</TR>
	<!--TR height=30>
		<TD colspan="6" align="center"><input type="button" value="�����" onclick="window.location='OrderEdit.asp?e=y&radif=<%=request("order") %>';"></TD>
	</TR-->
	</TABLE>
</body>