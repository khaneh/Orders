<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
'dim result
set result = jsObject()
select case request("act")
	case "prop":
		if IsNumeric(request("orderID")) then 
			set rs=Conn.Execute("select orders.id,orderTypes.name as typeName,orders.orderTitle, accounts.AccountTitle,orderStatus.Name as statusName, orderSteps.Name as stepName,orders.isClosed,orders.isApproved,orders.isOrder, orders.step from orders inner join orderTypes on orders.type=orderTypes.id inner join Accounts on orders.Customer = accounts.ID inner join orderStatus on orders.status=orderStatus.ID inner join orderSteps on orders.step=orderSteps.ID where orders.id=" & request("orderID"))
			if not rs.eof then 
				result("order") = rs("id")
				result("orderTitle") = rs("orderTitle")
				result("typeName") = rs("typeName")
				result("customerName") = rs("AccountTitle")
				result("stepName") = rs("stepName")
				result("step") = rs("step")
				result("statusName") = rs("statusName")
				result("isOrder") = rs("isOrder")
				result("isApproved") = rs("isApproved")
				result("isClosed") = rs("isClosed")
			end if
			rs.close
			set rs = Nothing
		end if
	case "set":
		if IsNumeric(request("orderID")) then 
			Conn.Execute ("UPDATE orders set step=" & Request("step") & ", isPrinted=0,lastUpdatedBy=" & Session("ID") & ", lastUpdatedDate=getdate() where id=" & Request("orderID"))
			set rs = Conn.Execute("select orderSteps.name from orders inner join orderSteps on orders.step=orderSteps.id where orders.id=" & Request("orderID"))
			if not rs.eof then
				result("status")="done"
				result("stepName")=rs("name")
			end if
			rs.Close
			set rs = Nothing
		end if
	case "print":
		if IsNumeric(request("orderID")) then 
			orderID= CDbl(Request("orderID"))
			conn.Execute("update orders set isPrinted=1, LastUpdatedDate=getDate(),LastUpdatedBy=" & Session("id") & " where id=" & orderID)
			result("status")="ok"
		end if
end select
Response.Write toJSON(result)
%>