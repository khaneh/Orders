<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
Response.Buffer = true
Response.ContentType = "text/xml;charset=windows-1256"
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
select case request("act")
	case "search":
		searchBox= request("searchBox")
		if searchBox="" then
			'By Default show Open Quotes of Current User
			myCriteria= "orders.isClosed=0 and Orders.CreatedBy = " & session("ID") 
		elseif isNumeric(searchBox) then
			search=clng(searchBox)
			myCriteria= "Orders.ID = '"& searchBox & "'"
		else
			searchBox=sqlSafe(searchBox)
			myCriteria= "orders.isClosed=0 and (REPLACE(accounts.companyName, ' ', '') LIKE REPLACE(N'%"& searchBox & "%', ' ', '') OR REPLACE(accounts.firstName1+accounts.firstName2+accounts.lastName1+accounts.lastName2, ' ', '') LIKE REPLACE(N'%"& searchBox & "%', ' ', ''))"
		End If
		mySQL="select Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName,orderStatus.Icon as statusIcon, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, 'test' as invStatus from Orders inner join orderStatus on orders.status=orderStatus.ID inner join orderSteps on orderSteps.ID=orders.step inner join Accounts on accounts.ID=orders.Customer inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type where  (" & myCriteria & ") order by Orders.createdDate desc"
		'response.write mysql
		'response.end
		Set rs = conn.execute(mySQL)
		set rows=server.createobject("MSXML2.DomDocument")
		rows.loadXML("<rows/>")
		while not rs.eof
			set order = rows.createElement("row")
			For i = 0 to rs.Fields.Count - 1
				set tmp = rows.createElement(rs.Fields(i).Name)
				if not isnull(rs.Fields(i))  then tmp.text= rs.Fields(i)
				order.AppendChild tmp
			next				
			rows.documentElement.AppendChild order
			rs.moveNext
		wend
		response.write(rows.xml)
		rs.close
		set rs = nothing
	case "advanceSearch":
		myCriteria="1=1"
		if request("checkID")="on" then 
			if request("fromID")<>"" then myCriteria=myCriteria & " and orders.id >= " & request("fromID")
			if request("toID")<>"" then myCriteria=myCriteria & " and orders.id <= " & request("toID")
		end if
		if request("checkIsOrder")="on" then 
			if request("isOrder")<>"" then myCriteria=myCriteria & " and orders.isOrder = " & request("isOrder")
		end if
		if request("checkCreatedDate")="on" then 
			if request("fromCreatedDate")<>"" then 
				d=splitDate(request("fromCreatedDate"))
				myCriteria=myCriteria & " and orders.createdDate >= dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&")" 
			end if
			if request("toCreatedDate")<>"" then 
				d=splitDate(request("toCreatedDate"))
				myCriteria=myCriteria & " and orders.createdDate < dateAdd(dd,1,dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&"))" 
			end if
		end if
		if request("checkRetDate")="on" then 
			if request("fromRetDate")<>"" then 
				d=splitDate(request("fromRetDate"))
				myCriteria=myCriteria & " and orders.returnDate >= dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&")" 
			end if
			if request("toRetDate")<>"" then 
				d=splitDate(request("toRetDate"))
				myCriteria=myCriteria & " and orders.returnDate < dateAdd(dd,1,dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&"))" 
			end if
		end if
		if request("checkContractDate")="on" then 
			if request("fromContractDate")<>"" then 
				d=splitDate(request("fromContractDate"))
				myCriteria=myCriteria & " and orders.id in (select orderID from OrderLogs where returnDate >= dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&") and ID in (select min(id) from OrderLogs where returnDate is not null group by OrderID))" 
			end if
			if request("toContractDate")<>"" then 
				d=splitDate(request("toContractDate"))
				myCriteria=myCriteria & " and orders.id in (select orderID from OrderLogs where returnDate < dateAdd(dd,1,dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&")) and ID in (select min(id) from OrderLogs where returnDate is not null group by OrderID))" 
			end if
		end if 
		if request("checkRetIsNull")="on" then 
			myCriteria=myCriteria & " and orders.returnDate is null"
		end if
		if request("checkCompany")="on" then 
			myCriteria=myCriteria & " and REPLACE(accounts.companyName, ' ', '') LIKE REPLACE(N'%"& request("companyName") & "%', ' ', '')"
		end if
		if request("checkCustomer")="on" then 
			myCriteria=myCriteria & " and REPLACE(accounts.firstName1+accounts.firstName2+accounts.lastName1+accounts.lastName2, ' ', '') LIKE REPLACE(N'%"& request("customerName") & "%', ' ', '')"
		end if
		if request("checkOrderType")="on" then 
			if request("checkNotOrderType")="on" then 
				myCriteria=myCriteria & " and orders.type<>" & request("orderType")
			else
				myCriteria=myCriteria & " and orders.type=" & request("orderType")
			end if
		end if
		if request("checkStep")="on" then
			if request("checkNotStep")="on" then 
				myCriteria=myCriteria & " and orders.step<>" & request("orderStep")
			else
				myCriteria=myCriteria & " and orders.step=" & request("orderStep")
			end if
		end if
		if request("checkStatus")="on" then
			if request("checkNotStatus")="on" then 
				myCriteria=myCriteria & " and orders.status<>" & request("orderStatus")
			else
				myCriteria=myCriteria & " and orders.status=" & request("orderStatus")
			end if
		end if
		if request("checkCreatedBy")="on" then
			if request("checkNotCreatedBy")="on" then 
				myCriteria=myCriteria & " and orders.CreatedBy<>" & request("orderCreatedBy")
			else
				myCriteria=myCriteria & " and orders.CreatedBy=" & request("orderCreatedBy")
			end if
		end if
		if request("checkTel")="on" then 
			myCriteria=myCriteria & " and (accounts.tel1 like '%"& request("tel") & "%' or accounts.tel2 like '%"& request("tel") & "%' or accounts.mobile1 like '%"& request("tel") & "%' or accounts.mobile2 like '%"& request("tel") & "%')"
		end if
		if request("checkClosed")="on" then 
			myCriteria=myCriteria & " and orders.isClosed=0"
		end if
		if request("checkTitle")="on" then 
			myCriteria=myCriteria & " and REPLACE(Orders.orderTitle, ' ', '') LIKE REPLACE(N'%"& request("orderTitle") & "%', ' ', '')"
		end if
		if cint(request("resultCount"))>1000 then 
			myCount = 1000
		else
			myCount = cint(request("resultCount"))
		end if
		mySQL="select top " & myCount & " Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName,orderStatus.Icon as statusIcon, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, 'test' as invStatus from Orders inner join orderStatus on orders.status=orderStatus.ID inner join orderSteps on orderSteps.ID=orders.step inner join Accounts on accounts.ID=orders.Customer inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type where  (" & myCriteria & ")"
		'response.write mysql
		'response.end
		Set rs = conn.execute(mySQL)
		set rows=server.createobject("MSXML2.DomDocument")
		rows.loadXML("<rows/>")
		while not rs.eof
			set order = rows.createElement("row")
			For i = 0 to rs.Fields.Count - 1
				set tmp = rows.createElement(rs.Fields(i).Name)
				if not isnull(rs.Fields(i))  then tmp.text= rs.Fields(i)
				order.AppendChild tmp
			next				
			rows.documentElement.AppendChild order
			rs.moveNext
		wend
		response.write(rows.xml)
		rs.close
		set rs = nothing
	case "getFolow":
		myCriteria=""
		if request("fromDate")<>"" then myCriteria = " and ((orders.returnDate between '" & request("fromDate") & "' AND '" & request("toDate") & "') or orders.returnDate is null) "
		if request("orderTypes")<>"" then myCriteria = myCriteria & " and orders.type in (" & request("orderTypes") & ")"
		myCriteria = myCriteria & " and orders.step=" & request("step")
		mySQL="SELECT Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName,orderStatus.Icon as statusIcon, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, 'test' as invStatus, DRV_Invoice.price FROM Orders inner join Accounts on accounts.ID=orders.Customer inner join orderSteps on orderSteps.ID=orders.step inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type INNER JOIN OrderStatus ON orders.status = OrderStatus.ID left outer join (select InvoiceOrderRelations.[Order],SUM(InvoiceLines.Price + InvoiceLines.Vat - InvoiceLines.Discount -InvoiceLines.Reverse) as price from InvoiceOrderRelations inner join Invoices on InvoiceOrderRelations.Invoice=Invoices.ID inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice where Invoices.Voided=0 group by InvoiceOrderRelations.[Order]) DRV_Invoice on Orders.ID=DRV_Invoice.[Order] WHERE (orders.isClosed=0 and isOrder=" & Request("isOrder") & myCriteria & ")  ORDER BY orders.createdDate DESC, orders.id DESC"	
		'response.write mysql
		'response.end
		Set rs = conn.execute(mySQL)
		set rows=server.createobject("MSXML2.DomDocument")
		rows.loadXML("<rows/>")
		while not rs.eof
			set order = rows.createElement("row")
			For i = 0 to rs.Fields.Count - 1
				set tmp = rows.createElement(rs.Fields(i).Name)
				if not isnull(rs.Fields(i))  then tmp.text= rs.Fields(i)
				order.AppendChild tmp
			next				
			rows.documentElement.AppendChild order
			rs.moveNext
		wend
		response.write(rows.xml)
		rs.close
		set rs = nothing
	case "getQuoteFolow":
		myCriteria=""
		if request("fromDate")<>"" then myCriteria = " and ((orders.returnDate between '" & request("fromDate") & "' AND '" & request("toDate") & "') or orders.returnDate is null) "
		if request("orderTypes")<>"" then myCriteria = myCriteria & " and orders.type in (" & request("orderTypes") & ")"
		myCriteria = myCriteria & " and orders.isClosed=" & Request("isClose")
		mySQL="SELECT Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName,orderStatus.Icon as statusIcon, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, 'test' as invStatus, DRV_Invoice.price FROM Orders inner join Accounts on accounts.ID=orders.Customer inner join orderSteps on orderSteps.ID=orders.step inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type INNER JOIN OrderStatus ON orders.status = OrderStatus.ID left outer join (select InvoiceOrderRelations.[Order],SUM(InvoiceLines.Price + InvoiceLines.Vat - InvoiceLines.Discount -InvoiceLines.Reverse) as price from InvoiceOrderRelations inner join Invoices on InvoiceOrderRelations.Invoice=Invoices.ID inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice where Invoices.Voided=0 group by InvoiceOrderRelations.[Order]) DRV_Invoice on Orders.ID=DRV_Invoice.[Order] WHERE (orders.isOrder=0  "& myCriteria & ")  ORDER BY orders.createdDate DESC, orders.id DESC"	
		'response.write mysql
		'response.end
		Set rs = conn.execute(mySQL)
		set rows=server.createobject("MSXML2.DomDocument")
		rows.loadXML("<rows/>")
		while not rs.eof
			set order = rows.createElement("row")
			For i = 0 to rs.Fields.Count - 1
				set tmp = rows.createElement(rs.Fields(i).Name)
				if not isnull(rs.Fields(i))  then tmp.text= rs.Fields(i)
				order.AppendChild tmp
			next				
			rows.documentElement.AppendChild order
			rs.moveNext
		wend
		response.write(rows.xml)
		rs.close
		set rs = nothing
	
end select
%>