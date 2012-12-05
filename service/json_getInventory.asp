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
		set rs = Conn.Execute("select inventoryItem,InventoryItems.* from InventoryInvoiceRelations inner join InventoryItems on InventoryItems.id=InventoryInvoiceRelations.inventoryItem where invoiceItem=" & invoiceItem & " order by InventoryItems.name")
		while not rs.eof
			set j(null) = jsObject()
			j(null)("inventoryItem") = rs("inventoryItem")
			j(null)("name") = rs("Name")
			j(null)("unit") = rs("Unit")
			rs.MoveNext
		wend
		rs.close
		set rs = Nothing
	case "findItem":
		set j = jsArray()
		mySearch = Request("search")
		mySQL =""
		if IsNumeric(mySearch) then 
			mySQL = "select * from inventoryItems where oldItemID=" & CDbl(mySearch)
		else
			mySQL = "select top 20 * from inventoryItems where enabled=1 and name like N'%" & Replace(mySearch," ","%") & "%'"
		end if
		if Len(mySearch)>3 and mySQL<>"" then 
			set rs = Conn.Execute(mySQL)
			while not rs.eof
				set j(null) = jsObject()
				j(null)("id") = rs("oldItemID") & "|" & rs("name")
				j(null)("text") = rs("oldItemID") & " " & rs("name") & " - موجودی " & rs("qtty") & " " & rs("unit") 
				rs.MoveNext
			wend		
			rs.close
			set rs = Nothing
		end if
	case "updateItemRequest":
		set j = jsObject()
		id= CDbl(Request("id"))
		Conn.Execute("update InventoryItemRequests set itemID= " & Request("itemID") & ",unit=N'" & Request("unit") & "' where id=" & id)
		j("status")="ok"
	case "delInvRequest":
		set j = jsObject()
		id= CDbl(Request("id"))
		Conn.Execute("update InventoryItemRequests set status= 'del' where id=" & id)
		j("status")="ok"
	case "updatePaperPrice":
		set j = jsObject()
		paperType = CInt(Request("paperType"))
		price = CDbl(Request("price"))
		cost = CDbl(Request("cost"))
		purchasePrice = CDbl(Request("purchasePrice"))
		set rs = Conn.Execute("select * from orderTypes where id=2")
		set prop=server.createobject("MSXML2.DomDocument")
		prop.loadXML(rs("property"))
		rs.close
		set rs = Nothing
		set paper = prop.SelectNodes("//key[@name='paper-type']/option[.=" & paperType & "]")
		paper(0).setAttribute "price",price
		paper(0).setAttribute "cost",cost
		paper(0).setAttribute "purchasePrice",purchasePrice
		conn.Execute("update orderTypes set property=N'" & prop.xml & "' where id=2")
		j("price") = paper(0).getAttribute("price")
	case "addWasteRequest":
		set j = jsObject()
		orderID = Request("orderID")
		rowName = Request("rowName")
		rowID = Request("rowID")
		maxRowID = CInt(Request("maxRowID")) + 1
		groupName = Request("groupName")
		qtty = Request("qtty")
		reqID = Request("reqID")
		comment = Request("comment")
		
		set rs = Conn.Execute("select * from orders where id=" & orderID)
		set orderProp = server.createobject("Msxml2.DOMDocument.6.0")
		orderProp.loadXML(rs("property"))
		set oldGroup = orderProp.selectNodes("/data/row[@id='" & rowID & "' and @name='" & rowName & "']/key[starts-with(@name,'" & groupName & "-')]") 
		set row  = orderProp.createElement("row")
		row.setAttribute "id", maxRowID
		row.setAttribute "name", rowName
		if rowName="printing" then 
			set tmp = orderProp.createElement("key")
			tmp.setAttribute "name", "description"
			tmp.text = " خرابي به دليل [" & comment & "]"
			row.AppendChild tmp
			set tmp = orderProp.createElement("key")
			tmp.setAttribute "name", "pages"
			tmp.text = "0"
			row.AppendChild tmp
		end if
		set tmp = orderProp.SelectSingleNode("/data/row[@id='" & rowID & "' and @name='" & rowName & "']/key[@name='" & groupName & "-qtty']")
 		orgQtty = CDbl(Replace(tmp.text,",",""))
 		if groupName = "plate" then 
 			set tmp = orderProp.SelectSingleNode("/data/row[@id='" & rowID & "' and @name='" & rowName & "']/key[@name='plate-color-count']") 
 			orgQtty = orgQtty * CDbl(Replace(tmp.text,",",""))
 		end if
 		set tmp = orderProp.SelectSingleNode("/data/row[@id='" & rowID & "' and @name='" & rowName & "']/key[@name='" & groupName & "-price']")
		orgPrice = CDbl(Replace(tmp.text,",",""))
		for each key in oldGroup 
			set tmp = orderProp.createElement("key")
			tmp.setAttribute "name",key.getAttribute("name")
			select case key.getAttribute("name")
				case "plate-color-count":
					tmp.text = 1
				case "paper-after-cut_qtty":
					tmp.text = qtty
				case groupName & "-price":
					tmp.text = 0
				case groupName & "-dis":
					tmp.text = 0
				case groupName & "-qtty":
					tmp.text = qtty
				case groupName & "-reverse":
					tmp.text = orgPrice / orgQtty * CDbl(qtty)
				case groupName & "-desc":
					tmp.text = " خرابي به دليل [" & comment & "]"
				case else
					tmp.text = key.text 
			end select
			row.AppendChild tmp
		next
		orderProp.SelectSingleNode("/data").AppendChild row
		conn.Execute("update orders set isApproved=0, property = N'" & orderProp.xml & "' where id=" & orderID)
		set rs = Conn.Execute("select * from InventoryItemRequests where ID=" & reqID)
		if IsNull(rs("itemID")) then 
			itemID = "null"
		else
			itemID = rs("itemID")
		end if
		if IsNull(rs("unit")) then 
			unit = "null"
		else
			unit = rs("itemID")
		end if
		Conn.Execute("insert into InventoryItemRequests (orderID, itemName, Comment, qtty, status, itemID, unit, createdBy, invoiceItem, rowID) values (" & orderID & ", N'" & rs("itemName") & "', N' خرابي به دليل [" & comment & "]', " & qtty & ", 'new', " & itemID & ", " & unit & ", " & Session("id") & ", " & rs("invoiceItem") & ", " & maxRowID & ")")
		j("status")="درخواست شما با موفقيت ثبت شد"
end select
Response.Write toJSON(j)
%>