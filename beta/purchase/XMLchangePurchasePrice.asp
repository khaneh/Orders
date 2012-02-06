<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<% if Auth(4 , 7) then %>
<%
	value=request.querystring("value")
	id=request.querystring("id")
	'response.write "update purchaseOrders SET price = "& value & " where id = "& id  
	Conn.Execute ("update purchaseOrders SET price = "& value & " where id = "& id )	
	response.write ("ok")	
%>
<% end if %>