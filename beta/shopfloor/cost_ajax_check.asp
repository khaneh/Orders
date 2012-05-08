<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../JSON_2.0.4.asp"-->
<%
'dim result
set result = jsObject()
select case request("act")
	case "order":
		result("order") = -1
		if IsNumeric(request("orderID")) then 
			set rs=Conn.Execute("select * from orders_trace where radif_sefareshat=" & request("orderID"))
			if not rs.eof then 
				result("order") = rs("radif_sefareshat")
				result("orderTitle") = rs("order_title")
				result("orderKind") = rs("order_kind")
				result("customerName") = rs("customer_name")
				result("orderStep") = rs("marhale")
			end if
			rs.close
		end if
	case "counter":
		set rs=Conn.Execute("select max(end_counter) as lastCounter from costs where operation_type=" & request("operation_type"))
		result("lastCounter")=0
		if not IsNull(rs("lastCounter")) then result("lastCounter")=CDbl(rs("lastCounter"))
		rs.close
	case "time":
		set rs=Conn.Execute("select convert(varchar, max(end_time), 120) as lastTime from costs where operation_type in (select id from cost_operation_type where driver_id in ( select driver_id from cost_operation_type where id=" & request("operation_type") & ")) and user_id=" & session("id"))
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