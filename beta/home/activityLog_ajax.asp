<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../JSON_2.0.4.asp"-->
<%
Set result = jsArray()

select case request("act")
	case "message":
		id=0
		if isNumeric(request("id")) then id=request("id")
		set rs=Conn.Execute("select messages.*,users.realName as toName, message_Types.name as typeName from messages inner join users on messages.msgTo=users.id inner join message_Types on messages.type=message_Types.id where messages.type>0 and messages.id>" & id & " and msgDate='" & request("date") & "' order by messages.id desc")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("id")=rs("id") 
			result(null)("body")=rs("msgBody")
			result(null)("time")=rs("msgTime")
			result(null)("date")=rs("msgDate")
			result(null)("to")=rs("msgTo")
			result(null)("from")=rs("msgFrom")
			result(null)("toName")=rs("toName")
			result(null)("typeName")=rs("typeName")
			rs.moveNext
		wend
		rs.close
		set rs=nothing
	case "quote":
		id=0
		if isNumeric(request("id")) then id=request("id")
		set rs=Conn.Execute("select * from Quotes where id>" & id &" and createdDate='" & request("date") & "'")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("id")=rs("id") 
			result(null)("date")=rs("createdDate")
			result(null)("time")=rs("order_time") 
			result(null)("company")=rs("company_name")  
			result(null)("customer")=rs("customer_name") 
			result(null)("kind")=rs("order_kind") 
			result(null)("title")=rs("order_title") 
			result(null)("qtty")=rs("qtty") 
			result(null)("size")=rs("paperSize") 
			result(null)("price")=rs("price") 
			result(null)("note")=rs("notes") 
			result(null)("createdBy")=rs("createdBy") 
			rs.moveNext
		wend
		rs.close
		set rs=nothing
	case "order":
		id=0
		if isNumeric(request("id")) then id=request("id")
		set rs=Conn.Execute("select * from orders_trace inner join orders on orders_trace.radif_sefareshat=orders.id where id>" & id &" and createdDate='" & request("date") & "'")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("id")=rs("id") 
			result(null)("date")=rs("createdDate")
			result(null)("time")=rs("order_time") 
			result(null)("company")=rs("company_name")  
			result(null)("customer")=rs("customer_name") 
			result(null)("kind")=rs("order_kind") 
			result(null)("title")=rs("order_title") 
			result(null)("qtty")=rs("qtty") 
			result(null)("size")=rs("paperSize") 
			result(null)("price")=rs("price") 
			result(null)("createdBy")=rs("createdBy") 
			rs.moveNext
		wend
		rs.close
		set rs=nothing
end select
Response.Write toJSON(result)
%>