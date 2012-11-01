<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
Response.Buffer = true
Response.ContentType = "text/xml;charset=windows-1256"
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
select case request("act")
	case "getProperty":
		orderType=CInt(request("type"))
		set rs = Conn.Execute("select * from orderTypes where id=" & orderType)
		set orderProp = server.createobject("MSXML2.DomDocument")
		
		if rs("property")<>"" then 
			orderProp.loadXML(rs("property"))
			response.write(orderProp.xml)
		end if
		rs.close
		set rs = nothing
	case "showLog":
		orderID = CDbl(request("id")) 
		set rs = Conn.Execute("select orderLogs.id, isnull(orderCustomerApprovedTypes.name,N' «ÌÌœ ‰‘œÂ') as customerApprovedTypeName, isnull(orderLogs.customerApprovedType,0) as customerApprovedType, isnull(orderLogs.isPrinted,0) as isPrinted, orderLogs.isOrder,orderLogs.isApproved, orderLogs.price, orderLogs.status, orderLogs.step, orderLogs.customer, orderLogs.returnDate, convert(varchar, orderLogs.insertedDate, 120) as insertedDate ,users.realName, orderSteps.name as stepName, orderStatus.name as statusName from orderLogs inner join users on users.id=orderLogs.insertedBy inner join orderStatus on orderLogs.status=orderStatus.id inner join orderSteps on orderSteps.id=orderLogs.step left outer join orderCustomerApprovedTypes on orderCustomerApprovedTypes.id=orderLogs.customerApprovedType where orderID = " & orderID & " order by insertedDate") 
		set logs=server.createobject("MSXML2.DomDocument")
		logs.loadXML("<logs/>")
		while not rs.eof
			set order = logs.createElement("order")
			For i = 0 to rs.Fields.Count - 1
				set tmp = logs.createElement(rs.Fields(i).Name)
				if IsNull(rs.Fields(i)) then
					tmp.text=""
				else
					tmp.text= rs.Fields(i)
				end if
				order.AppendChild tmp
			next				
			logs.documentElement.AppendChild order
			rs.moveNext
		wend
		response.write(logs.xml)
		rs.close
		set rs = nothing
	case "showOrder":
		if Request("logID")<>"" then
			logID=CDbl(Request("logID"))
			set rs = Conn.Execute("select * from orderLogs where id=" & logID)
		else
			orderID=Cdbl(request("id"))
			set rs = Conn.Execute("select * from orders where id=" & orderID)
		end if
		orderType=rs("type")
		set orderProp = server.createobject("MSXML2.DomDocument")
		orderProp.loadXML(rs("property"))
		set rs = Conn.Execute("select * from orderTypes where id=" & orderType)
		set typeProp = server.createobject("MSXML2.DomDocument")
		typeProp.loadXML(rs("property"))
		rs.close
		set rs = nothing
		set out = server.createobject("MSXML2.DomDocument")
		out.loadXML( "<keys/>" )
		for each row in orderProp.SelectNodes("//row")
			rowName=row.GetAttribute("name")
			rowID=row.GetAttribute("id")
			set rowType = typeProp.selectNodes("//rows[@name='" & rowName & "']")(0)
			set outRow=out.createElement("row")
			outRow.setAttribute "name",rowName
			outRow.setAttribute "id",rowID
			outRow.setAttribute "label",rowType.GetAttribute("label")
			for each key in row.SelectNodes("key")
				if typeProp.selectNodes("//key[@name='" & key.getAttribute("name") & "']").length>0 then 
					set typeKey = typeProp.selectNodes("//key[@name='" & key.getAttribute("name") & "']")(0)
					set typeGroup = typeKey.parentNode
					if outRow.selectNodes("//group[@name='" & typeGroup.getAttribute("name") & "']").length>0 then 
						set outGroup = outRow.selectNodes("//group[@name='" & typeGroup.getAttribute("name") & "']")(0)
						hasGroup=true
					else
						set outGroup = out.createElement("group")
						outGroup.setAttribute "label", ""&typeGroup.getAttribute("label")
						outGroup.setAttribute "name", ""&typeGroup.getAttribute("name")
						outGroup.setAttribute "isPrice", ""&typeGroup.getAttribute("isPrice")
						hasGroup=false
					end if
					set outKey = out.createElement("key")
					select case typeKey.getAttribute("type")
						case "check":
							outKey.text=""
						case "radio":
							set optionKey=typeKey.selectNodes("./option[.='" & key.text & "']")
							if optionKey.length>0 then 
								outKey.text = optionKey(0).getAttribute("label")
							else
								outKey.text = "Œÿ«ÌÌ ⁄ÃÌ» —Œ œ«œÂ!!!"
							end if
						case "option":
							set optionKey=typeKey.selectNodes("./option[.='" & key.text & "']")
							if optionKey.length>0 then 
								outKey.text = optionKey(0).getAttribute("label")
							else
								outKey.text = "Œÿ«ÌÌ ⁄ÃÌ» —Œ œ«œÂ!!!"
							end if
						case "option-other":
							set optionKey=typeKey.selectNodes("./option[.='" & key.text & "']")
							if optionKey.length>0 then 
								outKey.text = optionKey(0).getAttribute("label")
							else
								outKey.text = replace(key.text,"other:","")
							end if
						case "option-fromTable":
							tbl=typeKey.getAttribute("table")
							set rs= Conn.Execute("select * from "&tbl &" where id="&key.text)
							outKey.text=rs("name")
							rs.close
							set rs=nothing
						case else:
							outKey.text=key.text
					end select
					outKey.setAttribute "name",key.getAttribute("name")
					outKey.setAttribute "label",""&typeKey.getAttribute("label")
					outKey.setAttribute "type",typeKey.getAttribute("type")
					
					
					outGroup.AppendChild outKey
					if hasGroup=false then outRow.AppendChild outGroup
				else
					keyName=key.getAttribute("name")&""
					groupName=Split(keyName,"-")(0)
					keyDesc=Split(keyName,"-")(1)
					if keyDesc="price" or keyDesc="reverse" or keyDesc="dis" or keyDesc="over" or keyDesc="w" or keyDesc="l" or keyDesc="stockName" or keyDesc="stockDesc" or keyDesc="purchaseName" or keyDesc="purchaseDesc" or keyDesc="purchaseTypeID" then
						set typeGroup = typeProp.selectNodes("//group[@name='" & groupName & "']")(0)
					else
						set typeGroup = typeProp.selectNodes("//key[@name='" & keyName & "']/..")(0)
					end if
					if ((keyDesc="price") or keyDesc="reverse" or (keyDesc="dis") or (keyDesc="over")) then 
						if outRow.selectNodes("//group[@name='" & groupName & "']").length>0 then 
							set outGroup = outRow.selectNodes("//group[@name='" & groupName & "']")(0)
							hasGroup=true
						else
							set outGroup = out.createElement("group")
							outGroup.setAttribute "label", ""&typeGroup.getAttribute("label")
							outGroup.setAttribute "name", ""&typeGroup.getAttribute("name")
							outGroup.setAttribute "isPrice", ""&typeGroup.getAttribute("isPrice")
							hasGroup=false
						end if
						set outKey = out.createElement("key")
						outKey.setAttribute "name",key.getAttribute("name")
						outKey.setAttribute "type","price"
						outKey.text=key.text
						outGroup.AppendChild outKey

						if keyDesc="price" then 
							if outRow.selectNodes("//key[@name='" & groupName & "-dis']").length=0 and row.selectNodes("//key[@name='" & groupName & "-dis']").length=0 then 
								set outKey = out.createElement("key")
								outKey.setAttribute "name", groupName & "-dis"
								outKey.setAttribute "type","price"
								outKey.text="0%"
								outGroup.AppendChild outKey
							end if
							if outRow.selectNodes("//key[@name='" & groupName & "-over']").length=0 and row.selectNodes("//key[@name='" & groupName & "-over']").length=0 then 
								set outKey = out.createElement("key")
								outKey.setAttribute "name", groupName & "-over"
								outKey.setAttribute "type","price"
								outKey.text="0%"
								outGroup.AppendChild outKey
							end if
							if outRow.selectNodes("//key[@name='" & groupName & "-reverse']").length=0 and row.selectNodes("//key[@name='" & groupName & "-reverse']").length=0 then 
								set outKey = out.createElement("key")
								outKey.setAttribute "name", groupName & "-reverse"
								outKey.setAttribute "type","price"
								outKey.text="0%"
								outGroup.AppendChild outKey
							end if
						end if
						if hasGroup=false then outRow.AppendChild outGroup
					end if
				end if
			next
			if outRow.selectNodes("//key").length>0 then out.documentElement.AppendChild outRow
		next
		response.write(out.xml)
	case "editOrder":
		orderID=Cdbl(request("id"))
		set rs = Conn.Execute("select * from orders where id=" & orderID)
		orderType=rs("type")
		set orderProp = server.createobject("MSXML2.DomDocument")
		orderProp.loadXML(rs("property"))
		set rs = Conn.Execute("select * from orderTypes where id=" & orderType)
		set typeProp = server.createobject("MSXML2.DomDocument")
		typeProp.loadXML(rs("property"))
		rs.close
		set rs = nothing
		set out = server.createobject("MSXML2.DomDocument")
		out.loadXML( "<keys/>" )
		if orderProp.SelectSingleNode("/data") is Nothing then
			orderProp.loadXML("<data/>")
		end if
		for each group in typeProp.selectNodes("//group[@hasValue='yes']")
			group.removeAttribute("hasValue")
		next
		for each row in typeProp.SelectNodes("//rows")
			rowName = row.GetAttribute("name")
			if orderProp.SelectSingleNode("//row[@name='" & rowName & "']") is Nothing then
				Set tmp = orderProp.createElement("row")
				tmp.setAttribute "id","0"
				tmp.setAttribute "name",rowName
				orderProp.SelectSingleNode("/data").AppendChild tmp
			end if
		next
' 			Response.write orderProp.xml
' 			Response.end
		for each row in orderProp.SelectNodes("//row")
			rowName = row.GetAttribute("name")
			rowID = row.GetAttribute("id")
			set tmp=out.selectNodes("//rows[@name='" & rowName & "']")
			if tmp.length=0 then
				set rowType = typeProp.selectNodes("//rows[@name='" & rowName & "']")(0)
				set outRow = rowType.cloneNode(true)
				outRow.selectNodes("//row")(0).setAttribute "id",rowID
				hasRow=false
			else
				set rowType = typeProp.selectNodes("//rows[@name='" & rowName & "']/row")(0)
				set outRow = rowType.cloneNode(true)
				outRow.setAttribute "id",rowID
				hasRow=true
			end if
			
			for each key in row.selectNodes("./key")
				set thisRow = outRow.selectNodes("//key[@name='"&key.getAttribute("name")&"' and ../../@id='" & rowID & "']")
				if thisRow.length>0 then 
					thisRow(0).setAttribute "value",key.text
					set thisGroup = thisRow(0).selectNodes("..")
					thisGroup(0).setAttribute "hasValue" , "yes"
				end if
			next
			for each group in outRow.selectNodes("//group")
				thisPrice = "0"
				thisDis = "0%"
				thisOver = "0%"
				w="0"
				l="0"
				set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-price"&"' and ../@id='" & rowID & "']")
				if tmp.length>0 then thisPrice=tmp(0).text
				set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-dis"&"' and ../@id='" & rowID & "']")
				if tmp.length>0 then thisDis=tmp(0).text
				set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-over"&"' and ../@id='" & rowID & "']")
				if tmp.length>0 then thisOver=tmp(0).text
				set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-reverse"&"' and ../@id='" & rowID & "']")
				if tmp.length>0 then thisReverse=tmp(0).text
				set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-w"&"' and ../@id='" & rowID & "']")
				if tmp.length>0 then w=tmp(0).text
				set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-l"&"' and ../@id='" & rowID & "']")
				if tmp.length>0 then l=tmp(0).text
				group.setAttribute "price",thisPrice
				group.setAttribute "dis",thisDis
				group.setAttribute "over",thisOver
				group.setAttribute "reverse",thisReverse
				group.setAttribute "w",w
				group.setAttribute "l",l
				if group.getAttribute("hasStock")="yes" then 
					set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-stockName"&"' and ../@id='" & rowID & "']")
					if tmp.length>0 then group.setAttribute "stockName", tmp(0).text
					set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-stockDesc"&"' and ../@id='" & rowID & "']")
					if tmp.length>0 then group.setAttribute "stockDesc", tmp(0).text
				end if
				if group.getAttribute("hasPurchase")="yes" then 
					set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-purchaseName"&"' and ../@id='" & rowID & "']")
					if tmp.length>0 then group.setAttribute "purchaseName", tmp(0).text
					set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-purchaseDesc"&"' and ../@id='" & rowID & "']")
					if tmp.length>0 then group.setAttribute "purchaseDesc", tmp(0).text
					set tmp = row.selectNodes("//key[@name='"&group.getAttribute("name")&"-purchaseTypeID"&"' and ../@id='" & rowID & "']")
					if tmp.length>0 then group.setAttribute "purchaseTypeID", tmp(0).text
				end if
			next
			if hasRow then 
				set tmp=out.selectNodes("//rows[@name='" & rowName & "']")(0)
				tmp.AppendChild outRow
			else
				out.documentElement.AppendChild outRow
			end if
		next
		response.write(out.xml)
	case "getNew":
		customerID = Cdbl(request("id"))
		set head = server.createobject("MSXML2.DomDocument")
		head.loadXML( "<head/>" )
		set node = head.createElement("customer")
		set rs = Conn.Execute("select * from Accounts where id=" & customerID)
		head.documentElement.AppendChild node
		set customer = head.createElement("id")
		customer.text = customerID
		node.AppendChild customer
		set customer = head.createElement("accountTitle")
		customer.text = rs("accountTitle")
		node.AppendChild customer
		set customer = head.createElement("tel")
		if rs("tel1")<>"" then 
			customer.text = rs("tel1")
		else
			customer.text = rs("tel2")
		end if
		node.AppendChild customer
		set customer = head.createElement("companyName")
		customer.text = rs("companyName")
		node.AppendChild customer
		if rs("lastName1")<>"" or rs("firstName1")<>"" then 
		set customer = head.createElement("customerName")
			customer.text = rs("firstName1") & " " & rs("lastName1")
		else
			customer.text = rs("firstName2") & " " & rs("lastName2")
		end if
		node.AppendChild customer
		rs.close
		set rs=nothing
		Set node = head.createElement("orderTypeID")
		node.text=cint(request("typeID"))
		head.documentElement.AppendChild node
		Set node = head.createElement("orderTypeName")
		set rs = Conn.Execute("select * from orderTypes where id=" & cint(request("typeID")))
		node.text=rs("name")
		rs.close
		set rs=nothing
		head.documentElement.AppendChild node
		Set node = head.createElement("isOrder")
		node.text=request("isOrder")
		head.documentElement.AppendChild node
		Set node = head.createElement("orderID")
		node.text="######"
		head.documentElement.AppendChild node
		Set node = head.createElement("productionDuration")
		node.text=""
		head.documentElement.AppendChild node
		Set node = head.createElement("salesPerson")
		set rs = Conn.Execute("select * from users where id=" & session("ID"))
		node.text=rs("realName")
		rs.close
		set rs=nothing
		head.documentElement.AppendChild node
		Set node = head.createElement("orderTitle")
		node.text=""
		head.documentElement.AppendChild node
		Set node = head.createElement("paperSize")
		node.text=""
		head.documentElement.AppendChild node
		Set node = head.createElement("notes")
		node.text=""
		head.documentElement.AppendChild node
		Set node = head.createElement("qtty")
		node.text=""
		head.documentElement.AppendChild node
		Set node = head.createElement("totalPrice")
		node.text=""
		head.documentElement.AppendChild node
		set node = head.createElement("today")
		head.documentElement.AppendChild node
		set today = head.createElement("date")
		today.text=shamsiToday()
		node.AppendChild today
		set today = head.createElement("time")
		today.text=FormatDateTime(now(),vbshorttime)
		node.AppendChild today
		set today = head.createElement("retDate")
		if request("isOrder")="yes" then 
			today.text=""
		else
			today.text=shamsiDate(dateAdd("d",10,date()))
		end if
		node.AppendChild today
		set today = head.createElement("wName")
		today.text=weekdaynameFA(weekdayname(weekday(date)))
		node.AppendChild today
		response.write(head.xml)
	case "showHead":
		if Request("logID")<>"" then 
			LogID = CDbl(Request("logID"))
			orderID=0
		else
			orderID=Cdbl(request("id"))
			LogID = 0
		end if
		set head = server.createobject("MSXML2.DomDocument")
		head.loadXML( "<head/>" )
		if orderID>0 and logID=0 then 
			mySQL = "select orders.*, Accounts.tel1, Accounts.tel2, Accounts.dear1, Accounts.dear2, Accounts.firstName1, Accounts.firstName2, Accounts.lastName1,Accounts.lastName2, Accounts.companyName, Accounts.accountTitle,orderTypes.name as orderTypeName,users.realName, users.extention, isnull(invoices.Approved,-1) as invoiceApproved, isnull(invoices.Issued,-1) as invoiceIssued, orderSteps.name as stepName, orderStatus.name as statusName from Orders inner join orderSteps on orders.step=orderSteps.id inner join orderStatus on orderStatus.id=orders.status inner join Accounts on orders.customer=Accounts.ID inner join orderTypes on orders.type=orderTypes.id inner join users on orders.createdBy=users.id left outer join InvoiceOrderRelations on InvoiceOrderRelations.[order]=orders.id left outer join invoices on InvoiceOrderRelations.invoice = invoices.id and invoices.voided=0 where orders.id=" & orderID
		elseif orderID=0 and logID>0 then
			mySQL = "select orderlogs.*, orders.createdDate, Accounts.tel1, Accounts.tel2, Accounts.dear1, Accounts.dear2, Accounts.firstName1, Accounts.firstName2, Accounts.lastName1,Accounts.lastName2, Accounts.companyName, Accounts.accountTitle,orderTypes.name as orderTypeName,users.realName, users.extention, isnull(invoices.Approved,-1) as invoiceApproved, isnull(invoices.Issued,-1) as invoiceIssued, orderSteps.name as stepName, orderStatus.name as statusName from orderlogs inner join orderSteps on orderlogs.step=orderSteps.id inner join orderStatus on orderStatus.id=orderlogs.status inner join orders on orderLogs.orderID=orders.id inner join Accounts on orderlogs.customer=Accounts.ID inner join orderTypes on orderlogs.type=orderTypes.id inner join users on orders.createdBy=users.id left outer join InvoiceOrderRelations on InvoiceOrderRelations.[order]=orderlogs.orderID left outer join invoices on InvoiceOrderRelations.invoice = invoices.id and invoices.voided=0 where orderlogs.id=" & logID
		end if
		set rs = Conn.Execute(mySQL)
		set node = head.createElement("status")	
		head.documentElement.AppendChild node 
		set stat = head.createElement("isOrder")
		stat.text = CBool(rs("isOrder"))
		node.AppendChild stat
		set stat = head.createElement("isApproved")
		stat.text = CBool(rs("isApproved"))
		node.AppendChild stat
		set stat = head.createElement("isClosed")
		stat.text = CBool(rs("isClosed"))
		node.AppendChild stat
		set stat = head.createElement("step")
		stat.text = rs("step") 
		node.AppendChild stat
		set stat = head.createElement("statusName")
		stat.text = rs("statusName") 
		node.AppendChild stat
		set stat = head.createElement("stepName")
		stat.text = rs("stepName") 
		node.AppendChild stat
		set stat = head.createElement("invoiceApproved")
		stat.text = rs("invoiceApproved") 
		node.AppendChild stat
		set stat = head.createElement("invoiceIssued")
		stat.text = rs("invoiceIssued") 
		node.AppendChild stat
		
		set node = head.createElement("customer")
		head.documentElement.AppendChild node
		set customer = head.createElement("id")
		customer.text = rs("customer")
		node.AppendChild customer
		set customer = head.createElement("accountTitle")
		customer.text = rs("accountTitle")
		node.AppendChild customer
		set customer = head.createElement("tel")
		if rs("tel1")<>"" then 
			customer.text = rs("tel1")
		else
			customer.text = rs("tel2")
		end if
		node.AppendChild customer
		set customer = head.createElement("companyName")
		customer.text = rs("companyName")
		node.AppendChild customer
		set dear = head.createElement("dear")
		if rs("lastName1")<>"" or rs("firstName1")<>"" then 
		set customer = head.createElement("customerName")
			customer.text = rs("firstName1") & " " & rs("lastName1")
			dear.text = rs("dear1")
		else
			customer.text = rs("firstName2") & " " & rs("lastName2")
			dear.text = rs("dear2")
		end if
		node.AppendChild customer
		node.AppendChild dear
		
		Set node = head.createElement("orderTypeID")
		node.text=rs("type")
		head.documentElement.AppendChild node
		Set node = head.createElement("orderTypeName")
		node.text=rs("orderTypeName")
		head.documentElement.AppendChild node
		Set node = head.createElement("isOrder")
		if rs("isOrder") then 
			node.text="yes"
		else
			node.text=""
		end if
		head.documentElement.AppendChild node
		Set node = head.createElement("orderID")
		if orderID=0 then 
			node.text = rs("orderID")
		else
			node.text=orderID
		end if
		head.documentElement.AppendChild node
		Set node = head.createElement("productionDuration")
		node.text=rs("productionDuration")
		head.documentElement.AppendChild node
		Set node = head.createElement("salesPerson")
		node.text=rs("realName")
		head.documentElement.AppendChild node
		Set node = head.createElement("extention")
		if IsNull(rs("extention")) then 
			node.text = ""
		else
			node.text = rs("extention")
		end if 
		head.documentElement.AppendChild node			
		Set node = head.createElement("orderTitle")
		node.text=rs("orderTitle")
		head.documentElement.AppendChild node
		Set node = head.createElement("paperSize")
		if IsNull(rs("paperSize")) then
			node.text=""
		else
			node.text=rs("paperSize")
		end if
		head.documentElement.AppendChild node
		Set node = head.createElement("notes")
		if IsNull(rs("notes")) then 
			node.text = ""
		else
			node.text=rs("notes")
		end if
		head.documentElement.AppendChild node
		Set node = head.createElement("qtty")
		if IsNull(rs("qtty")) then
			node.text=""
		else
			node.text=rs("qtty")
		end if
		head.documentElement.AppendChild node
		Set node = head.createElement("totalPrice")
		if IsNull(rs("price")) then
			node.text=""
		else
			node.text=rs("Price")
		end if
		head.documentElement.AppendChild node
		set node = head.createElement("today")
		head.documentElement.AppendChild node
		set today = head.createElement("date")
		today.text=shamsiDate(rs("createdDate"))
		node.AppendChild today
		set today = head.createElement("time")
		today.text=FormatDateTime(rs("createdDate"),vbshorttime)
		node.AppendChild today
		set today = head.createElement("retDate")
		if IsNull(rs("returnDate")) then 
			today.text=""
			node.AppendChild today
			set today = head.createElement("retTime")
			today.text=""
			node.AppendChild today
			set today = head.createElement("retIsNull")
			today.text="yes"
			node.AppendChild today
		else
			today.text=shamsiDate(rs("returnDate"))
			node.AppendChild today
			set today = head.createElement("retTime")
			today.text=FormatDateTime(rs("returnDate"),vbshorttime)
			node.AppendChild today
		end if
		
		set today = head.createElement("wName")
		today.text=weekdaynameFA(weekdayname(weekday(rs("createdDate"))))
		node.AppendChild today
		set today = head.createElement("shamsiToday")
		today.text = shamsiToday()
		node.AppendChild today
		rs.close
		set rs=nothing
		response.write(head.xml)
	case "stock":
		orderID=Cdbl(request("id"))
		set rs = Conn.Execute("select InventoryItemRequests.*, InventoryPickuplists.status as pickStatus, InventoryPickuplistItems.pickupListID, InventoryPickuplistItems.qtty as pickQtty, InventoryPickuplistItems.unit as pickUnit from InventoryItemRequests left outer join InventoryPickuplistItems on InventoryPickuplistItems.requestID=InventoryItemRequests.id left outer join InventoryPickuplists on InventoryPickuplists.id=InventoryPickuplistItems.pickupListID where isnull(InventoryPickuplists.status,'')<>'del' and  InventoryItemRequests.orderID=" & orderID)
		set stock=server.createobject("MSXML2.DomDocument")
		stock.loadXML("<stock/>")
		while not rs.eof
			set req = stock.createElement("req")
			set tmp = stock.createElement("ID")
			tmp.text=rs("ID")
			req.AppendChild tmp
			set tmp = stock.createElement("Comment")
			tmp.text=rs("Comment") 
			req.AppendChild tmp
			set tmp = stock.createElement("ItemName")
			tmp.text=rs("ItemName") 
			req.AppendChild tmp
			set tmp = stock.createElement("customer")
			tmp.text=rs("customerHaveInvItem") 
			req.AppendChild tmp
			set tmp = stock.createElement("rowID")
			if IsNull(rs("rowID")) then 
				tmp.text = ""
			else
				tmp.text=rs("rowID") 
			end if
			req.AppendChild tmp
			set tmp = stock.createElement("invoiceItem")
			if IsNull(rs("invoiceItem")) then
				tmp.text=""
			else
				tmp.text=rs("invoiceItem") 
			end if
			req.AppendChild tmp
			set tmp = stock.createElement("reqStatus")
			tmp.text=rs("Status") 
			req.AppendChild tmp
			set st = stock.createElement("Status")
			set stClass = stock.createElement("StatusClass")
			set unit = stock.createElement("unit")
			set qtty = stock.createElement("Qtty")
			set link = stock.createElement("link")
			if rs("pickStatus")="done" then 
				st.text="Œ—ÊÃ «‰Ã«„ ‘œÂ"
				qtty.text = rs("pickQtty")
				unit.text = rs("pickUnit")
				stClass.text="label label-success"
				link.text = rs("pickupListID")
			elseif rs("Status")="pick" then
				st.text="ÕÊ«·Â"
				qtty.text = rs("pickQtty")
				unit.text = rs("pickUnit")
				stClass.text="label label-inverse"
				link.text = rs("pickupListID")
			elseif rs("Status")="new" then
				st.text="ÃœÌœ"
				qtty.text = rs("Qtty")
				if IsNull(rs("unit")) then
					unit.text=""
				else
					unit.text = rs("unit")
				end if
				stClass.text="label label-info"
			elseif rs("Status")="del" then 
				st.text="œ—ŒÊ«”  Å«ﬂ ‘œÂ"
				qtty.text = rs("Qtty")
				if IsNull(rs("unit")) then 
					unit.text=""
				else
					unit.text = rs("unit")
				end if
				stClass.text="label label-important"
			elseif rs("pickStatus")="del" then
				st.text="ÕÊ«·Â Å«ﬂ ‘œÂ"
				qtty.text = rs("pickQtty")
				unit.text = rs("pickUnit")
				stClass.text="label label-important"
				link.text = rs("pickupListID")
			else
				st.text="⁄ÃÌ»!"
				qtty.text = rs("Qtty")
				unit.text = rs("unit")
				stClass.text="label label-warning"
			end if
			req.AppendChild st
			req.AppendChild stClass
			req.AppendChild qtty
			req.AppendChild unit
			req.AppendChild link
			stock.documentElement.AppendChild req
			rs.moveNext
		wend
		response.write(stock.xml)
		rs.close
		set rs = nothing
		set tmp=Nothing
		set st=Nothing
		set stClass=Nothing
		set unit=Nothing
		set qtty=Nothing
		set req=Nothing
		set stock=Nothing
	case "purchase":
		orderID=Cdbl(request("id"))
		set rs = Conn.Execute("select PurchaseRequests.*,isnull(PurchaseRequestOrderRelations.Ord_ID,0) as Ord_ID from PurchaseRequests left outer join PurchaseRequestOrderRelations on PurchaseRequests.ID=PurchaseRequestOrderRelations.Req_ID where PurchaseRequests.orderID=" & orderID)
		set purchase=server.createobject("MSXML2.DomDocument")
		purchase.loadXML("<purchase/>")
		while not rs.eof
			set req = purchase.createElement("req")
			set tmp = purchase.createElement("Comment")
			tmp.text = rs("Comment")
			req.AppendChild tmp
			set tmp = purchase.createElement("Status")
			tmp.text = rs("Status")
			req.AppendChild tmp
			set tmp = purchase.createElement("TypeName")
			tmp.text = rs("TypeName")
			req.AppendChild tmp
			set tmp = purchase.createElement("Qtty")
			tmp.text = rs("Qtty")
			req.AppendChild tmp
			set tmp = purchase.createElement("Price")
			if IsNull(rs("Price")) then 
				tmp.text = ""
			else
				tmp.text = rs("Price")
			end if
			req.AppendChild tmp
			set tmp = purchase.createElement("Ord_ID")
			tmp.text = rs("Ord_ID")
			req.AppendChild tmp
			purchase.documentElement.AppendChild req
			rs.moveNext
		wend
		response.write(purchase.xml)
		rs.close
		set rs = nothing
end select
%>