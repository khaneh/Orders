<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
if request("id")="" then
	response.write escape ("���� ��� " )
	response.end
end if 

if not isnumeric(request("id")) then
	response.write escape ("���! " )
	response.end
end if 

sql="SELECT * From InventoryItems WHERE OldItemID="& request("id") & " ORDER BY id"
'response.write sql
set rs=Conn.Execute(sql)
if (rs.EOF) then
	response.write escape ("���� ������ ���� �����" )
else
    response.write escape( rs("name") )
end if

rs.close()
conn.close()
%>