<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
'dim result
set result = jsObject()
select case request("act")
	case "order":
		result("order") = -1
		if IsNumeric(request("orderID")) then 
			set rs=Conn.Execute("select orders.*, orderTypes.name as typeName, Accounts.AccountTitle, orderSteps.name as stepName from orders inner join orderTypes on orders.type=orderTypes.id inner join Accounts on orders.customer=Accounts.id inner join orderSteps on orders.step = orderSteps.id where orders.id=" & request("orderID"))
			if not rs.eof then 
				result("order") = rs("id")
				result("orderTitle") = rs("orderTitle")
				result("orderKind") = rs("typeName")
				result("customerName") = rs("AccountTitle")
				result("orderStep") = rs("stepName")
			end if
			rs.close
		end if
	case "counter":
		set rs=Conn.Execute("select max(end_counter) as lastCounter from costs where operation_type in (select id from cost_operation_type where driver_id=" & request("driver_id") & ")")
		result("lastCounter")=0
		if not IsNull(rs("lastCounter")) then result("lastCounter")=CDbl(rs("lastCounter"))
		rs.close
	case "time":
		set rs=Conn.Execute("select convert(varchar, max(end_time), 120) as lastTime from costs where operation_type in (select id from cost_operation_type where driver_id=" & request("driver_id") & ") and user_id=" & session("id"))
		result("foundLastTime")=0
		current=now()
		if not IsNull(rs("lastTime")) then 
			result("foundLastTime")=1
			if dateDiff("d", rs("lastTime"), current)>0 then 
				result("lastTime") = rs("lastTime")
				result("isNewDay") = 1
				result("day") = dateDiff("d", rs("lastTime"), current)
			else
				result("lastTime") = rs("lastTime")
				result("isNewDay") = 0
			end if
		end if
end select
Response.Write toJSON(result)
%>