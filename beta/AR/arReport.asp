<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="ê“«—‘ ›—Ê‘"
SubmenuItem=11
if not Auth("C" , 9) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<table>
	<tr>
		<td>‰«„</td>
		<td>„»·€ „«‰œÂ</td>
	</tr>
<%
	set rs=Conn.execute("select * from accounts where arBalance<0")
	while not rs.eof 
%>
	<tr>
		<td><a href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rs("id")%>"><%=rs("accountTitle")%></a></td>
		<td><%=Separate(abs(CDbl(rs("arBalance"))))%></td>
	</tr>
<%		
		rs.moveNext
	wend
%>
</table>
<!--#include file="tah.asp" -->