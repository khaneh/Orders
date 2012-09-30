<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
Response.Buffer = true
Response.ContentType = "text/xml;charset=windows-1256"
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
select case request("act")
	case "showType" 
		typ = CInt(request("type"))
		gl = CInt(request("gl"))
		mySQL="select arItems.EffectiveDate,arItems.Link,arMemo.Description,arItems.IsCredit,arItems.AmountOriginal,arItems.Voided, accounts.AccountTitle,arItems.Account from ARItems inner join ARMemo on arItems.Link=arMemo.ID inner join Accounts on arItems.Account=accounts.ID where ARItems.Type=3 and arMemo.Type=" & typ & " and ARItems.gl=" & gl
		'response.write mysql
		'response.end
		Set rs = conn.execute(mySQL)
		set rows=server.createobject("MSXML2.DomDocument")
		rows.loadXML("<rows/>")
		while not rs.eof
			set order = rows.createElement("row")
			For i = 0 to rs.Fields.Count - 1
				set tmp = rows.createElement(rs.Fields(i).Name)
				if not isnull(rs.Fields(i))  then tmp.text= rs.Fields(i)
				order.AppendChild tmp
			next				
			rows.documentElement.AppendChild order
			rs.moveNext
		wend
		response.write(rows.xml)
		rs.close
		set rs = nothing
end select
%>