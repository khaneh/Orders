<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
Response.Buffer = true
Response.ContentType = "text/xml;charset=windows-1256"
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
select case request("act")
	case "related":
		table = request("table")
		id= request("id")
		mySQL="SELECT Messages.*, fu.realName as fromName, tu.realName as toName,Messages.Urgent, message_types.name typeName FROM Messages INNER JOIN Users as fu ON Messages.MsgFrom = fu.ID inner join users as tu on tu.id=messages.msgTo inner join message_types on Messages.type=message_types.id WHERE (Messages.RelatedTable = '" & table & "') AND (Messages.RelatedID = " & id & ") ORDER BY Messages.Urgent desc,Messages.type desc, Messages.ID DESC"
		Set rs = conn.execute(mySQL)
		set rows=server.createobject("MSXML2.DomDocument")
		rows.loadXML("<rows/>")
		oldBody = "<!----First----!>"
		while not rs.eof
			if oldBody = rs("MsgBody") then 
				set tmp = order.selectNodes("./toName")(0)
				if rs("toName")<>rs("fromName") then 
					if tmp.text="" then
						tmp.text = "Èå " & rs("toName")
					else
						tmp.text = tmp.text & "¡ " & rs("toName")
					end if
				end if
			else
				set order = rows.createElement("row")
				set tmp = rows.createElement("Urgent")
				tmp.text = rs("Urgent")
				order.AppendChild tmp
				set tmp = rows.createElement("fromName")
				tmp.text = rs("fromName")
				order.AppendChild tmp
				set tmp = rows.createElement("toName")
				if rs("toName")=rs("fromName") then 
					tmp.text = ""
				else
					tmp.text = "Èå " & rs("toName")
				end if
				order.AppendChild tmp
				set tmp = rows.createElement("MsgBody")
				tmp.text = rs("MsgBody")
				oldBody = rs("MsgBody")
				order.AppendChild tmp
				set tmp = rows.createElement("MsgDate")
				tmp.text = rs("MsgDate")
				order.AppendChild tmp
				set tmp = rows.createElement("MsgTime")
				tmp.text = rs("MsgTime")
				order.AppendChild tmp
				set tmp = rows.createElement("typeName")
				if rs("type")="0" then 
					tmp.text = ""
				else
					tmp.text = rs("typeName")
				end if
				order.AppendChild tmp
				rows.documentElement.AppendChild order
			end if
			rs.moveNext
		wend
		response.write(rows.xml)
		rs.close
		set rs = nothing
end select
%>