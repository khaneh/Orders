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
		mySQL="select Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, isnull(convert(varchar, Orders.approvedDate, 120),0) as approvedDate, isnull(convert(varchar, Orders.orderDate, 120),0) as orderDate, isnull(convert(varchar, contract.contractDate, 120),0) as contractDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, case when invoices.voided=1 then N'»«ÿ· ‘œÂ' else case when invoices.issued=1 then N'’«œ— ‘œÂ' else case when invoices.approved=1 then N' «ÌÌœ ‘œÂ' else case when invoices.id is null then N'‰œ«—œ' else N'«ÌÃ«œ ‘œÂ' end end end end as invStatus, isnull(invoices.id,0) as invoiceID, isnull(orderCustomerApprovedTypes.name,N' «ÌÌœ ‰‘œÂ') as customerApprovedTypeName, isnull(orders.customerApprovedType,0) as customerApprovedType, isnull(orders.isPrinted,0) as isPrinted, orders.isOrder, orders.isApproved from Orders inner join orderStatus on orders.status=orderStatus.ID inner join orderSteps on orderSteps.ID=orders.step inner join Accounts on accounts.ID=orders.Customer inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id left outer join orderCustomerApprovedTypes on orderCustomerApprovedTypes.id=orders.customerApprovedType left outer join (select orderID, returnDate as contractDate from orderLogs where id in (select min(id) from orderLogs where returnDate is not null and isOrder=1 group by orderID)) as contract on orders.id = contract.orderID where  (" & myCriteria & ") order by Orders.orderDate DESC, Orders.createdDate desc"
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
				if Request("checkIsOrder")="on" and Request("isOrder") = "1" then 
					myCriteria=myCriteria & " and orders.orderDate >= dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&")" 
				elseif Request("checkIsOrder")="on" and Request("isOrder") = "0" then 
					myCriteria=myCriteria & " and orders.createdDate >= dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&")" 
				else
					myCriteria=myCriteria & " and (orders.createdDate >= dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&") or orders.orderDate >= dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&"))" 
				end if
			end if
			if request("toCreatedDate")<>"" then 
				d=splitDate(request("toCreatedDate"))
				if Request("checkIsOrder")="on" and Request("isOrder") = "1" then 
					myCriteria=myCriteria & " and orders.orderDate < dateAdd(dd,1,dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&"))" 
				elseif Request("checkIsOrder")="on" and Request("isOrder") = "0" then 
					myCriteria=myCriteria & " and orders.createdDate < dateAdd(dd,1,dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&"))" 
				else
					myCriteria=myCriteria & " and (orders.createdDate < dateAdd(dd,1,dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&")) or orders.orderDate < dateAdd(dd,1,dbo.udf_date_solarToDate("&d(0)&","&d(1)&","&d(2)&"))" 
				end if
				
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
		mySQL="select top " & myCount & " Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, isnull(convert(varchar, Orders.approvedDate, 120),0) as approvedDate, isnull(convert(varchar, Orders.orderDate, 120),0) as orderDate, isnull(convert(varchar, contract.contractDate, 120),0) as contractDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName,orderStatus.Icon as statusIcon, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, case when invoices.voided=1 then N'»«ÿ· ‘œÂ' else case when invoices.issued=1 then N'’«œ— ‘œÂ' else case when invoices.approved=1 then N' «ÌÌœ ‘œÂ' else case when invoices.id is null then N'‰œ«—œ' else N'«ÌÃ«œ ‘œÂ' end end end end as invStatus, isnull(invoices.id,0) as invoiceID, isnull(orderCustomerApprovedTypes.name,N' «ÌÌœ ‰‘œÂ') as customerApprovedTypeName, isnull(orders.customerApprovedType,0) as customerApprovedType, isnull(orders.isPrinted,0) as isPrinted, orders.isOrder, orders.isApproved from Orders inner join orderStatus on orders.status=orderStatus.ID inner join orderSteps on orderSteps.ID=orders.step inner join Accounts on accounts.ID=orders.Customer inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id left outer join orderCustomerApprovedTypes on orderCustomerApprovedTypes.id=orders.customerApprovedType left outer join (select orderID, returnDate as contractDate from orderLogs where id in (select min(id) from orderLogs where returnDate is not null and isOrder=1 group by orderID)) as contract on orders.id = contract.orderID where  (" & myCriteria & ") order by Orders.orderDate DESC, Orders.createdDate desc"
		
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
		mySQL="SELECT Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, isnull(convert(varchar, Orders.approvedDate, 120),0) as approvedDate, isnull(convert(varchar, Orders.orderDate, 120),0) as orderDate, isnull(convert(varchar, contract.contractDate, 120),0) as contractDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName,orderStatus.Icon as statusIcon, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, case when invoices.voided=1 then N'»«ÿ· ‘œÂ' else case when invoices.issued=1 then N'’«œ— ‘œÂ' else case when invoices.approved=1 then N' «ÌÌœ ‘œÂ' else case when invoices.id is null then N'‰œ«—œ' else N'«ÌÃ«œ ‘œÂ' end end end end as invStatus, isnull(invoices.id,0) as invoiceID, isnull(orderCustomerApprovedTypes.name,N' «ÌÌœ ‰‘œÂ') as customerApprovedTypeName, isnull(orders.customerApprovedType,0) as customerApprovedType, isnull(orders.isPrinted,0) as isPrinted, orders.isOrder, orders.isApproved FROM Orders inner join Accounts on accounts.ID=orders.Customer inner join orderSteps on orderSteps.ID=orders.step inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type INNER JOIN OrderStatus ON orders.status = OrderStatus.ID left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id left outer join orderCustomerApprovedTypes on orderCustomerApprovedTypes.id=orders.customerApprovedType left outer join (select orderID, returnDate as contractDate from orderLogs where id in (select min(id) from orderLogs where returnDate is not null and isOrder=1 group by orderID)) as contract on orders.id = contract.orderID WHERE (orders.isClosed=0 and isOrder=" & Request("isOrder") & myCriteria & ")  ORDER BY  Orders.orderDate DESC, orders.createdDate DESC, orders.id DESC"	
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
		mySQL="SELECT Orders.id, isnull(convert(varchar, Orders.createdDate, 120),0) as createdDate, Orders.Customer, isnull(convert(varchar, Orders.returnDate, 120),0) as returnDate, isnull(convert(varchar, Orders.approvedDate, 120),0) as approvedDate, isnull(convert(varchar, Orders.orderDate, 120),0) as orderDate, isnull(convert(varchar, contract.contractDate, 120),0) as contractDate, Orders.orderTitle, accounts.CompanyName, accounts.AccountTitle, accounts.Tel1,accounts.Tel2, accounts.Mobile1, accounts.Mobile2, orderStatus.Name as statusName,orderStatus.Icon as statusIcon, orderSteps.Name as stepName, users.RealName, orderTypes.name as typeName, orders.Price, case when invoices.voided=1 then N'»«ÿ· ‘œÂ' else case when invoices.issued=1 then N'’«œ— ‘œÂ' else case when invoices.approved=1 then N' «ÌÌœ ‘œÂ' else case when invoices.id is null then N'‰œ«—œ' else N'«ÌÃ«œ ‘œÂ' end end end end as invStatus, isnull(invoices.id,0) as invoiceID, isnull(orderCustomerApprovedTypes.name,N' «ÌÌœ ‰‘œÂ') as customerApprovedTypeName, isnull(orders.customerApprovedType,0) as customerApprovedType, isnull(orders.isPrinted,0) as isPrinted, orders.isOrder, orders.isApproved FROM Orders inner join Accounts on accounts.ID=orders.Customer inner join orderSteps on orderSteps.ID=orders.step inner join Users on users.ID=orders.CreatedBy inner join orderTypes on orderTypes.id=orders.type INNER JOIN OrderStatus ON orders.status = OrderStatus.ID left outer join InvoiceOrderRelations on orders.id=InvoiceOrderRelations.[order] left outer join invoices on InvoiceOrderRelations.invoice=invoices.id left outer join orderCustomerApprovedTypes on orderCustomerApprovedTypes.id=orders.customerApprovedType left outer join (select orderID, returnDate as contractDate from orderLogs where id in (select min(id) from orderLogs where returnDate is not null and isOrder=1 group by orderID)) as contract on orders.id = contract.orderID WHERE (orders.isOrder=0  "& myCriteria & ")  ORDER BY Orders.orderDate DESC, orders.createdDate DESC, orders.id DESC"	
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