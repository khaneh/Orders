<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
ErrorFound=False
ON ERROR Resume Next
	ID=clng(request("id"))

	if Err.Number<>0 then
		Err.clear
		ErrorFound=True
	end if
ON ERROR GOTO 0

if NOT ErrorFound then
	'Changed by Kid ! 831125
	'Considering Status <> 3 (3=Account is Blocked)
	sql="SELECT AccountTitle From Accounts WHERE ID="& ID & " AND Status <> 3"
	set rs=Conn.Execute(sql)
	if (rs.EOF) then
		response.write escape ("äíä ÍÓÇÈí æÌæÏ äÏÇÑÏ" )
	else
		response.write escape( rs("AccountTitle") )
	end if
	rs.close
else
		response.write escape ("äíä ÍÓÇÈí æÌæÏ äÏÇÑÏ" )
end if
conn.close
%>