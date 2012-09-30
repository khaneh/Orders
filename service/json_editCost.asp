<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
set result = jsObject()
select case request("act")
	case "del":
		costID = CDbl(Request("id"))
		conn.Execute("delete costs where id=" & costID)
		result("status") = "ok"
end select
Response.Write toJSON(result)
%>