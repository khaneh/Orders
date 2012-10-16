<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset = "utf-8"
	Response.CodePage = 65001
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
select case request("act")
	case "account":
		set j = jsArray()
		if IsNumeric(Request("search")) then
			id = CDbl(Request("search"))
			mySQL = "select id,accountTitle + '(' + cast(id as varchar(10)) + ')' as accountTitle from accounts where id= " & id
		else
			text = Replace(Request("search")," ","")
			mySQL = "select top 10 id,accountTitle from accounts where replace(accountTitle, ' ', '') like N'%" & text & "%' order by accountTitle"
		end if
		if Len(Request("search"))>3 then 
			set rs = Conn.Execute(mySQL)
			while not rs.eof
				set j(null) = jsObject()
				j(null)("id") = rs("id")
				j(null)("text") = rs("accountTitle")
				rs.MoveNext
			wend		
			rs.close
			set rs = Nothing
		end if
end select
Response.Write toJSON(j)
%>