<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
select case request("act")
	case "user":
		set j = jsObject()
		id = CInt(Request("id"))
		set rs = Conn.Execute("select realName from users where id=" & id)
		j("name") = rs("realName")
		rs.close
		set rs = Nothing
	case "orderType":
		set j = jsObject()
		id = CInt(Request("id"))
		set rs = Conn.Execute("select name from orderTypes where id = " & id)
		j("name") = rs("name")
		rs.close
		set rs = Nothing
	case "orderStep":
		set j = jsObject()
		id = CInt(Request("id"))
		set rs = Conn.Execute("select name from orderSteps where id = " & id)
		j("name") = rs("name")
		rs.close
		set rs = Nothing
	case "orderStatus":
		set j = jsObject()
		id = CInt(Request("id"))
		set rs = Conn.Execute("select name from orderStatus where id = " & id)
		j("name") = rs("name")
		rs.close
		set rs = Nothing
end select
Response.Write toJSON(j)
%>