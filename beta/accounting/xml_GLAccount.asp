<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
if request("id")="" then
'	response.write escape ("ÎÇáí ÇÓÊ " )
	response.write escape ("äíä ÍÓÇÈí æÌæÏ äÏÇÑÏ" )
	response.end
end if 

if not isnumeric(request("id")) then
'	response.write escape ("ÎØÇ! " )
	response.write escape ("äíä ÍÓÇÈí æÌæÏ äÏÇÑÏ" )
	response.end
end if 

sql="SELECT * From GLAccounts WHERE id="& request("id") & " AND (GL="& OpenGL & ") ORDER BY Name"
'response.write sql
set rs=Conn.Execute(sql)
if (rs.EOF) then
	response.write escape ("äíä ÍÓÇÈí æÌæÏ äÏÇÑÏ" )
else
    response.write escape( rs("name") )
end if

rs.close()
conn.close()
%>