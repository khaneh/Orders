<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset = "utf-8"
	Response.CodePage = 65001
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
select case request("act")
	case "find":
		set j = jsArray()
		mySearch = Request("search")
		mySQL =""
		if (IsNumeric(mySearch) or LCase(mySearch) <> UCase(mySearch)) then 
			mySQL = "select top 15 * from catalogItems inner join catalogGroup on catalogItems.groupID=catalogGroup.id where lower(itemCode) like N'%" & lcase(mySearch) & "%'"
		elseif LCase(mySearch) = UCase(mySearch) then 
			mySQL = "select top 15 * from catalogItems inner join catalogGroup on catalogItems.groupID=catalogGroup.id where description like N'%" & mySearch & "%' or name like N'%" & mySearch & "%'"
		end if
		if Len(mySearch)>3 and mySQL<>"" then 
			
			set rs = Conn.Execute(mySQL)
			while not rs.eof
				set j(null) = jsObject()
				j(null)("id") = rs("price") & "|" & rs("qtty") & "|" & rs("itemCode") & "|" & rs("name") & "|" & rs("description")
				j(null)("text") = rs("itemCode") & " " & rs("name") & " - " & rs("description") & " (" & rs("qtty") & " عدد)"
				rs.MoveNext
			wend		
			rs.close
			set rs = Nothing
		end if
end select
Response.Write toJSON(j)
%>