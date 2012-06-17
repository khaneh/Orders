<%@ Language=VBScript %>
<%
Function GenHtmlTable()
	Dim sRet 
	sRet = ""
	sRet = "<table border=1 class=mysmalltext><tr bgcolor=""#d8ecea"">"
	sRet = sRet & "<td>Name</td>"	
	sRet = sRet & "<td>Country</td>"	
	sRet = sRet & "</tr>"	
	
	'Beginning rows...Might read from Recordset - here it's hardcoded
	sRet = sRet & "<tr>"	
	sRet = sRet & "<td>John Doe</td>"	
	sRet = sRet & "<td>US</td>"	
	sRet = sRet & "</tr>"	

	sRet = sRet & "<tr>"	
	sRet = sRet & "<td>Stefan Holmberg</td>"	
	sRet = sRet & "<td>Sweden</td>"	
	sRet = sRet & "</tr>"	

	sRet = sRet & "</table>"	
                GenHtmlTable = sRet
End Function





jhkjh

Response.Clear()
Response.Buffer = True
Response.AddHeader "Content-Disposition", "attachment;filename=export.xls"  
Response.ContentType = "application/vnd.ms-excel"
Response.Write GenHtmlTable()
Response.End()
%>
