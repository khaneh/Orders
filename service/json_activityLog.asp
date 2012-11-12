<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
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
		'Response.write "select orders.*, orderType.name as order_kind, accounts.accountTitle from orders inner join Accounts on orders.customer=accounts.id inner join orderTypes on orders.type = orderTypes.id where orders.isOrder=0 and orders.id>" & id &" and createdDate between DATEADD(DAY, 0, DATEDIFF(DAY, 0, " & myDate & ")) and DATEADD(DAY, 1, DATEDIFF(DAY, 0, " & myDate & "))"
		'Response.end
		set rs=Conn.Execute("declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate(" & mid(request("date"),1,4) & "," & mid(request("date"),6,2) & "," & mid(request("date"),9,2) & ");select orders.*, orderTypes.name as order_kind, accounts.accountTitle from orders inner join Accounts on orders.customer=accounts.id inner join orderTypes on orders.type = orderTypes.id where orders.isOrder=0 and orders.id>" & id &" and orders.createdDate between DATEADD(DAY, 0, DATEDIFF(DAY, 0, @myDate)) and DATEADD(DAY, 1, DATEDIFF(DAY, 0, @myDate))")
		while not rs.eof
			set result(null) = jsObject()
			result(null)("id")=rs("id") 
			result(null)("date")=shamsiDate(rs("createdDate"))
			result(null)("time")=FormatDateTime(rs("createdDate"), 4)
			result(null)("accountTitle")=rs("accountTitle")  
			result(null)("kind")=rs("order_kind") 
			result(null)("title")=rs("orderTitle") 
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
		'Response.end
		id=0
		if isNumeric(request("id")) then id=request("id")
		myDate = "dbo.udf_date_solarToDate(" & mid(request("date"),1,4) & "," & mid(request("date"),6,2) & "," & mid(request("date"),9,2) & ")"
		set rs=Conn.Execute("declare @myDate datetime;set @myDate = dbo.udf_date_solarToDate(" & mid(request("date"),1,4) & "," & mid(request("date"),6,2) & "," & mid(request("date"),9,2) & ");select orders.*, orderTypes.name as order_kind, accounts.accountTitle from orders inner join Accounts on orders.customer=accounts.id inner join orderTypes on orders.type = orderTypes.id where orders.isOrder=1 and orders.id>" & id &" and orders.createdDate between DATEADD(DAY, 0, DATEDIFF(DAY, 0, @myDate)) and DATEADD(DAY, 1, DATEDIFF(DAY, 0, @myDate))")
		'DATEADD(DAY, 0, DATEDIFF(DAY, 0, createdDate)),DATEADD(DAY, 1, DATEDIFF(DAY, 0, createdDate))
		while not rs.eof
			set result(null) = jsObject()
			result(null)("id")=rs("id") 
			result(null)("date")=shamsiDate(rs("createdDate"))
			result(null)("time")=FormatDateTime(rs("createdDate"), 4)
			result(null)("accountTitle")=rs("accountTitle")
			result(null)("kind")=rs("order_kind") 
			result(null)("title")=rs("orderTitle") 
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