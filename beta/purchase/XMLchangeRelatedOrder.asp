<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<% if Auth(4 , 6) then %>
<%
	ordID=request.querystring("ordID")
	reqID=request.querystring("reqID")
	Conn.Execute ("update purchaseRequests SET order_ID = "& ordID & " where id = "& reqID )	
	response.write ("ok")	
%>
<% end if %>