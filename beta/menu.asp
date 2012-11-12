<!--#include file="config.asp" -->
<%
AppBgColor = "#555"	'Other:"#99AACC" change 2 bbb
AppFgColor =  "#C3DBEB" '"#DEEBD9"
SelectedMenuColor = "#0E5499"
UnSelectedMenuColor = "#309261"
SelectedSubMenuColor = "#DEEBD9" '"#C3DBEB" 
UnSelectedSubMenuColor = "#0E5499" ' "#609250"
TabWidth=65
ImgTabSelected="/images/tab-1.gif"
ImgTabNotSelected="/images/tab-2.gif"

Function WriteMessagesStatus()
	result=""
	MySQL="SELECT (SELECT COUNT(*) AS CntTotal FROM Messages WHERE (MsgTo = '"& session("ID") & "') AND (IsRead = 0) AND (IsSmall = 0)) AS CntTotal, (SELECT COUNT(*) AS CntTotal FROM Messages WHERE (MsgTo = '"& session("ID") & "') AND (IsRead = 0) AND (IsSmall = 0) AND (Urgent = 2)) AS CntUrgent2, (SELECT COUNT(*) AS CntTotal FROM Messages WHERE (MsgTo = '"& session("ID") & "') AND (IsRead = 0) AND (IsSmall = 0) AND (Urgent = 3)) AS CntUrgent3"
	set tmpRS1=Conn.Execute (MySQL)
	if not tmpRS1.EOF then
		result="<A href='../home/default.asp?sub=1'><span style='font-size:10pt;color:#CCCCCC;'><FONT Face='wingdings'>*</FONT> ÅÌ«„ ÃœÌœ (" & tmpRS1("CntTotal") & ") </span>"
		if tmpRS1("CntUrgent2")>0 then
			result = result & "<span style='font-size:10pt;color:yellow;'><FONT Face='wingdings'>*</FONT> ŒÌ·Ì ›Ê—Ì (" & tmpRS1("CntUrgent2") & ") </span>"
		end if
		if tmpRS1("CntUrgent3")>0 then
			result = result & "<span style='font-size:10pt;color:#33FF99;'><FONT Face='wingdings'>*</FONT> ¬„«œÂ  ÕÊÌ· (" & tmpRS1("CntUrgent3") & ") </span>"
		end if
			result = result & "</A>"
	else
		result = "&nbsp;"
	end if
	tmpRS1.close
	set tmpRS1 = Nothing
	'-------------------------------------- this added by sam--------------------------------------------------
	'MySQL = "SELECT Extention FROM Users WHERE ID = " & session("ID")
	'set tmpRS1 = Conn.Execute (MySQL)
	'if not tmpRS1.EOF then
	'	dim xml
	'	set xml = server.CreateObject("MSXML2.ServerXMLHTTP")
	'	xml.open "GET", "https://192.168.0.9/test/getvmi.php?exten=" & tmpRS1("Extention"), false
	'	xml.SetOption(2) = 13056
	'	xml.send
	'	vmCount = xml.ResponseText
	'	set xml = Nothing
	'	if CInt(vmCount) > 0 then
	'		result = result & "<span style='font-size:10pt;color:red;'><FONT Face='wingdings'>*</FONT> ÅÌ«„ ’Ê Ì (" & vmCount & ") </span>"
	'	end if
	'end if
	'tmpRS1.close
	'----------------------------------------------------------------------------------------------------------
	set tmpRS1 = nothing
	WriteMessagesStatus = result
End Function

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<TITLE> <%=PageTitle%> </TITLE>
<style>
	body { font-family: tahoma; font-size: 8pt;}
	body A { Text-Decoration : none ;}
	Input { font-family: tahoma; font-size: 9pt;}
	td { font-family: tahoma; font-size: 8pt;}
	.tt { font-family: tahoma; font-size: 10pt; color:yellow;}
	.tt2 { font-family: tahoma; font-size: 8pt; color:yellow;}
	.inputBut { font-family: tahoma; font-size: 8pt; richness: 10}
	.t7pt { font-size: 8pt;}
	.t8pt { font-size: 10pt;}
	.alak a { color: #cccccc; text-decoration: none;  font-size: 10pt;}
	.alak2 a { color: black; text-decoration: none;  font-size: 10pt; }
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; height:20px; }
</style>
<link type="text/css" href="/css/jquery-ui-1.8.16.custom.css" rel="stylesheet" />
<link rel="StyleSheet" href="/css/jame.css" type="text/css">
<link rel="StyleSheet" href="/css/font.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
<script type="text/javascript" src="/js/jquery.dateFormat-1.0.js"></script>
<script type="text/javascript" src="/js/jquery.jsonSuggest-2.min.js"></script>
<script type="text/javascript" src="/js/jalaliCalendar.js"></script>
<script type="text/javascript" src="/js/xslTransform.js"></script>
<script type="text/javascript" src="/js/jame.1.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.8.16.custom.min.js"></script>
</HEAD>
<BODY bgcolor=<%=AppBgColor%> topmargin=0 leftmargin=0><!<% if onunload<>"" then %> onunload="<%=onunload%>"<% end if %> >
<TABLE topmargin=0 leftmargin=0 align=center width=760 border=0>
<TR>
	<TD align=LEFT>
	<%if Auth(8,4) then  
		response.write "<a href='../accounting/manageGL.asp?act=select'>"&session("OpenGLName")&"</a>"
	else
		response.write session("OpenGLName")
	end if%> -  
	<% if session("differentGL") then%> <FONT SIZE="" COLOR="yellow"><B>¬ê«Â »«‘ ﬂÂ œ› — ﬂ·  Ê »« »ﬁÌÂ ›—ﬁ œ«—œ </B></FONT><% end if %></TD>
	<TD align=right>”·«„ <%=session("CSRName")%> Ã«‰</TD>
</TR>
</TABLE><%
if application("syslock") <> "" and application("syslock") <> session("id")  then 
response.write "<BR><BR><BR><BR><BR><CENTER><H1>”Ì” „  Ê”ÿ  "& application("syslockerName") & " ﬁ›· ‘œÂ «” </H1></CENTER>"
response.end
end if 

if menuItem="0" then 
	rootLink="../"
else
	rootLink="../"
end if

CSRName = session("CSRName")

%>
<BR>
<TABLE cellspacing=0 cellpadding=0 width=760 border=0 dir=rtl align=center>
<TR >
	<TD>
	<TABLE cellspacing=0 cellpadding=0>
	<TR height=30 class="alak">

	<%if Auth(0 , 0) then %>
	<%if menuItem="0" then %> 
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>’›ÕÂ «Ê·</TD>
	<%else %>  
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='<%=rootLink%>home'>’›ÕÂ «Ê·</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(1 , 0) then %>
	<%if menuItem="1" then %> 
		<TD class="tt2" width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>›Â—”  <br>‰«„ Â«</TD>
	<%else %>  
		<TD class="tt2" width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='<%=rootLink%>CRM/AccountInfo.asp?act=search'>›Â—”  <br>‰«„ Â«</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(2 , 0) then %>
	<%if menuItem="2" then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>”›«—‘« </TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='<%=rootLink%>order/order.asp?act=makeNewQoute'>”›«—‘« </A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(3 , 0) then %>
	<%if menuItem="3" then %> 
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'> Ê·Ìœ</TD>
	<%else %>  
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='<%=rootLink%>shopfloor'> Ê·Ìœ</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(4 , 0) then %>
	<%if menuItem="4" then %> 
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>Œ—Ìœ</TD>
	<%else %>  
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='<%=rootLink%>purchase'>Œ—Ìœ</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(5 , 0) then %>
	<%if menuItem="5" then %> 
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>«‰»«—</TD>
	<%else %>  
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='<%=rootLink%>inventory'>«‰»«—</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(6 , 0) then %>
	<%if menuItem="6" then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>Õ”«»œ«—Ì ›—Ê‘</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='<%=rootLink%>AR/Invoice.asp'>Õ”«»œ«—Ì ›—Ê‘</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(7 , 0) then %>
	<%if menuItem="7" then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>Õ”«»œ«—Ì Œ—Ìœ</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='<%=rootLink%>AP/voucherInput.asp'>Õ”«»œ«—Ì Œ—Ìœ</A></TD>
	<%end if %>
	<%end if %>

	<%if Auth("B" , 0) then %>
	<%if menuItem="B" then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>Õ”«»œ«—Ì ”«Ì—</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='<%=rootLink%>AO/AccountReport.asp?act=search'>Õ”«»œ«—Ì ”«Ì—</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(8 , 0) then %>
	<%if menuItem="8" then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>Õ”«»œ«—Ì</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='<%=rootLink%>accounting/GLMemoInput.asp'>Õ”«»œ«—Ì</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth(9 , 0) then %>
	<%if menuItem="9" then %> 
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>’‰œÊﬁ</TD>
	<%else %>  
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='<%=rootLink%>cashReg/ReceiptInput.asp'>’‰œÊﬁ</A></TD>
	<%end if %>
	<%end if %>


	<%if Auth("A" , 0) then %>
	<%if menuItem="A" then %> 
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>»«‰ﬂ</TD>
	<%else %>  
		<TD class=tt width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='<%=rootLink%>bank/cheq.asp'>»«‰ﬂ</A></TD>
	<%end if %>
	<%end if %>
	
	</TR>
	</TABLE>
	</TD>
</TR>
<TR BGCOLOR="<%=SelectedMenuColor%>">
	<TD height=20>
	<div id="MessagesStatusPanel">
	<%=WriteMessagesStatus()%>
	</div>
<%

' ---	LOGGING

'	if not DONT_LOG_THIS then
'		function pad (inpStr,padSize)
'			result = inpStr
'			while len(result) < padSize 
'				result = "0" & result
'			wend
'			pad = result
'		end function
'
'
'		rform		= sqlSafe(request.form)
'		rquerystring = sqlSafe(request.querystring)
'
'		tmpDate		= date()
'		tmpTime		= time()
'		rdate		= year(tmpDate) & "/" & pad(Month(tmpDate),2) & "/" & pad(Day(tmpDate),2)
'		rtime		= hour(tmpTime)  & ":" & minute(tmpTime)  & ":" & second(tmpTime)
'		httpReferer = sqlSafe(Request.ServerVariables("HTTP_REFERER"))
'		clientIP	= Request.ServerVariables("REMOTE_ADDR")
'		rURL		= Request.ServerVariables("URL")
'		errBy = session("id") 
'
'		errSQL = "INSERT INTO useageLog (errDate, errTime, errBy, url, httpReferer, clientIP, requestForm, requestQuerystring) VALUES ('" & rdate & "','" & rtime & "'," & errBy & ",'" & rURL & "' ,N'" & httpReferer & "','" & clientIP & "','" & left(rform,500) & "','" & rquerystring & "')"
'
'		Conn.Execute errSQL 
'
'	end if
' ---	LOGGING end
%>
	</TD>
</TR>
<TR >
	<TD >
