<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
if request("id")="" then
	response.write escape ("ÎÇáí ÇÓÊ " )
	response.end
end if 

if not isnumeric(request("id")) then
	response.write escape ("ÎØÇ! " )
	response.end
end if 

sql="SELECT MAX(dbo.InventoryItems.OldItemID) + 1 AS maxID FROM dbo.InventoryItems INNER JOIN dbo.InventoryItemCategoryRelations ON dbo.InventoryItems.ID = dbo.InventoryItemCategoryRelations.Item_ID GROUP BY dbo.InventoryItemCategoryRelations.Cat_ID HAVING (dbo.InventoryItemCategoryRelations.Cat_ID = "& request("id") & ")"

'response.write sql
set rs=Conn.Execute(sql)
if (rs.EOF) then
	response.write escape ("")
else
    response.write escape( rs("maxID") )
end if

rs.close()
conn.close()
%>