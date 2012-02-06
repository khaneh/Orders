<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<% if Auth(4 , 6) then %>
<%
	value=request.querystring("value")
	id=request.querystring("id")
	Conn.Execute ("update purchaseRequests SET order_ID = "& value & " where id = "& id )	
	response.write ("ok")	
%>
<% end if %>