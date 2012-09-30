<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
select case request("act")
	case "getList":
		set j = jsArray()
		set rs=Conn.Execute("select * from OutServices order by name")
		while not rs.eof  
			set j(null) = jsObject()
			j(null)("id")=rs("id")
			j(null)("name")=rs("name")
			rs.moveNext
		wend
		rs.close
		set rs=nothing
end select
Response.Write toJSON(j)
%>