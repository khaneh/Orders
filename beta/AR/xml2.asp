<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
InvItemID=request("id")
if isNumeric(InvItemID) then
	InvItemID=clng(InvItemID)
else
	InvItemID=0
end if
sql="SELECT replace(Name,'&#1740;',N'í') as name,type,isnull(fee,0) as fee,hasVat FROM InvoiceItems WHERE (ID='" & InvItemID &"') AND (Enabled=1)"
set rs=Conn.Execute(sql)
if (request("id")="") then
	response.write escape ("ÓØÑ ÎÇáí" )
elseif (rs.EOF) then
	response.write escape ("ßÏ ßÇáÇ ÛáØ ÇÓÊ" )
else
    response.write escape( rs("name") & "#" & rs("type") & "#" & rs("fee") & "#" & rs("hasVat"))
end if

rs.close()
conn.close()
%>