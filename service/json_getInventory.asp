<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset = "utf-8"
	Response.CodePage = 65001
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
select case request("act")
	case "itemListFromInvoiceItem":
		set j = jsArray()
		invoiceItem = CInt(Request("invoiceItem"))
		set rs = Conn.Execute("select inventoryItem,InventoryItems.* from InventoryInvoiceRelations inner join InventoryItems on InventoryItems.id=InventoryInvoiceRelations.inventoryItem where invoiceItem=" & invoiceItem)
		while not rs.eof
			set j(null) = jsObject()
			j(null)("inventoryItem") = rs("inventoryItem")
			j(null)("name") = rs("Name")
			j(null)("unit") = rs("Unit")
			rs.MoveNext
		wend
		rs.close
		set rs = Nothing
	case "updateItemRequest":
		set j = jsObject()
		id= CDbl(Request("id"))
		Conn.Execute("update InventoryItemRequests set itemID= " & Request("itemID") & ",unit=N'" & Request("unit") & "' where id=" & id)
		j("status")="ok"
end select
Response.Write toJSON(j)
%>