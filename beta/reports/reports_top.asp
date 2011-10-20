<%
if (session("ID")="") then
	session.abandon
	response.redirect "../login.asp"
end if
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=utf-8">
<title><%=reportTitle%></title>
</head>
<body>
<%
if Request.QueryString("act")="show" then

'	conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open conStr

'	------------------ 
'	-----  
'	-----  Create the Query from posted parameters
'	-----  
'	------------------ 

%>
