<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<?xml version="1.0" encoding="windows-1256"?> 
<%
InvItemID=request("id")
if isNumeric(InvItemID) then
	InvItemID=clng(InvItemID)
else
	InvItemID=0
end if
'sql="SELECT * FROM InvoiceItems WHERE (ID='" & InvItemID &"') AND (Enabled=1)"

mySQL="SELECT InvoiceItemCategories.* FROM InvoiceItemCategoryRelations INNER JOIN InvoiceItemCategories ON InvoiceItemCategoryRelations.InvoiceItemCategory = InvoiceItemCategories.ID WHERE (InvoiceItemCategoryRelations.InvoiceItem = '" & InvItemID &"')" 
set rs=Conn.Execute(mySQL)
if NOT rs.eof then
%><Categories>
<%
	Do while NOT rs.eof
		%><Category id="<%=rs("id")%>"><%=rs("name")%></Category> 
		<%
		rs.MoveNext
	Loop
%></Categories>
<%
end if
rs.close()
conn.close()
%>