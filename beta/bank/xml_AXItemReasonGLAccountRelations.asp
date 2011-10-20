<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<%
if request("id")="" then
	response.write escape ("äíä ÍÓÇÈí æÌæÏ äÏÇÑÏ" )
	response.end
end if 

if not isnumeric(request("id")) then
	response.write escape ("äíä ÍÓÇÈí æÌæÏ äÏÇÑÏ" )
	response.end
end if 


mySQL = "SELECT ID,name FROM GLAccounts WHERE (HasAppendix = 1) AND (GL = "& openGL & ") AND (ID = "& request("id") & ")"
Set RS1=Conn.execute(mySQL)
if RS1.eof then
	response.write escape ("Çíä ãÚíä ÊÝÕíáí äãí ÐíÑÏ" )
	response.end
	accountName = ""
else
	accountName = RS1("name")
end if
RS1.close

mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& request("id") & ")"
Set RS1=Conn.execute(mySQL)
if RS1.eof then
	'Using default reason (sys: AO, reason: Misc.)
	Reason=6
else
	Reason=	cint(RS1("ItemReason"))
end if
RS1.close

mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
Set RS1=Conn.execute(mySQL)
sys=			RS1("Acron")
ReasonName =	RS1("Name")

response.write escape( Reason & "-" & ReasonName & " (" & accountName & ")")
RS1.close
conn.close
%>