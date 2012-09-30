<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../JSON_2.0.4.asp"-->
<%
select case request("act")
	case "getAddr":
		id=0
		set j = jsObject()
		if isNumeric(request("id")) then id=request("id")
		set rs=Conn.Execute("select * from Accounts where id=" & id)
		if not rs.eof then 
			j("addr1")=rs("address1")
			j("addr2")=rs("address2")
		end if
		rs.close
		set rs=nothing
end select
Response.Write toJSON(j)
%>