<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Home (0)
PageTitle= " ãÞÇáÇÊ "
SubmenuItem=7
if not Auth(0 , 8) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<%
remotID = request.serverVariables("REMOTE_ADDR")
if mid(remotID,1,11)="192.168.10." then 
	wikiADDR = "http://192.168.10.12/mediawiki"
else
	wikiADDR = "http://jame.pdhco.com:1375/mediawiki"
end if
'response.write wikiADDR
%>
<iframe src=<%=wikiADDR%> name="wiki" width="900px" height="900px">
<!--#include file="tah.asp" -->
