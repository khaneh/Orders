<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle=server.HTMLEncode("ãÏíÑíÊ ÎØÇ")
SubmenuItem=10
%>
<!--#include file="top.asp" -->
<style>
.myTable
TD { 
	border-bottom: 1pt solid black;
	border-left: 1pt solid gray;
	border-right: 1pt solid gray;
	}
</style>
<%
	session.codepage="1252"

	errID = request("id")

	mySQL = "SELECT errLog.[file] as myFile, errLog.[line] as myLine, errLog.*, Users.RealName, Users_1.RealName AS checkerName FROM errLog INNER JOIN Users ON errLog.errBy = Users.ID LEFT OUTER JOIN Users Users_1 ON errLog.CheckedBy = Users_1.ID where errLog.id = "& errID & " order by errLog.ID DESC"
	
	set myRS = Conn.Execute (mySQL)

	aspCode		= myRS("ASPCode")
	Number		= myRS("Number")
	Source		= myRS("Source")
	Category	= myRS("Category")
	File		= myRS("File")
	Line		= myRS("Line")
	Column		= myRS("Column")
	Description = myRS("Description")
	Extended	= myRS("Extended")
	rform		= replace(replace(unescape(myRS("requestForm")),"&", "<br>"),"+"," ")
	rquerystring = replace(replace(unescape(myRS("requestQuerystring")),"&", "<br>"),"+"," ")
	rDate		= myRS("errDate")
	rTime		= myRS("errTime")
	httpReferer = myRS("httpReferer")
	clientIP	= myRS("clientIP")
	rURL		= myRS("url")
	RealName	= myRS("RealName")
	checkerName	= myRS("checkerName")
	CheckerComment	= myRS("CheckerComment")

%>
<div dir="ltr"><BR><BR><BR>
<TABLE align="center" dir="ltr" style="border:0pt solid white; width=50%">
<TR>
	<TD align="right" bgcolor="bbbbbb" bgcolor="bbbbbb">aspCode:</TD>
	<TD align="left"> <%=aspCode%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Number:</TD>
	<TD align="left"> <%=Number%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Source:</TD>
	<TD align="left"> <%=Source%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Category:</TD>
	<TD align="left"> <%=Category%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">File:</TD>
	<TD align="left"> <%=File%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Line:</TD>
	<TD align="left"> <%=Line%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Column:</TD>
	<TD align="left"> <%=Column%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Description:</TD>
	<TD align="left"> <%=Description%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Extended:</TD>
	<TD align="left"> <%=Extended%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">request.form:</TD>
	<TD align="left"> <%=replace(rform,"&","<br>")%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">request.querystring:</TD>
	<TD align="left"> <%=replace(replace(rquerystring,"&","<br>"),"_", " ")%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Date:</TD>
	<TD align="left"> <%=rdate%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Time:</TD>
	<TD align="left"> <%=rtime%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">HTTP Referer:</TD>
	<TD align="left"> <%=httpReferer%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Client IP:</TD>
	<TD align="left"> <%=clientIP%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Real Name:</TD>
	<TD align="left"> <%=server.HTMLEncode(RealName)%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">checker Name:</TD>
	<TD align="left"> <%=server.HTMLEncode(" " & checkerName)%></TD>
</TR>
<TR>
	<TD align="right" bgcolor="bbbbbb">Checker Comment:</TD>
	<TD align="left"> <%=CheckerComment%></TD>
</TR>
</TABLE><BR>
<%
session.codepage="1256"
%>
<!--#include file="tah.asp" -->