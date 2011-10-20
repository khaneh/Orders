<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><BR><BR>
<CENTER><H3>Ÿ«Â—« «‘ﬂ«·Ì ÅÌ‘ ¬„œÂ «” </H3>
<!--„„ﬂ‰ «”  «Ì‰ «‘ﬂ«· ‰«‘Ì «“   Ê—ÊœÌ €·ÿ  Ê”ÿ ‘„« »«‘œ.<BR><BR>
·ÿ›« <A HREF="javascript:history.go(-1)">œÊ»«—Â ”⁄Ì ﬂ‰Ìœ° </A>Ì« œ— ’Ê—  »«ﬁÌ „«‰œ‰ «‘ﬂ«· „œÌ— ”Ì” „ —« „ÿ·⁄ ‰„«ÌÌœ.
-->
</CENTER>
<BR><BR>
<%
'Response.CharSet = "arabic-1256"

	function sqlSafe (inpStr)
	  tmpStr=inpStr
	  tmpStr=replace(tmpStr,"'","`")
	  tmpStr=replace(tmpStr,chr(34),"`")
	  sqlsafe=tmpStr
	end function
	
	function pad (inpStr,padSize)
		result = inpStr
		while len(result) < padSize 
			result = "0" & result
		wend
		pad = result
	end function

	
	dim obje
	set obje = server.GetLastError()

	aspCode		= obje.ASPCode
	Number		= obje.Number
	Source		= sqlSafe(obje.Source)
	Category	= sqlSafe(obje.Category)
	File		= sqlSafe(obje.File)
	Line		= obje.Line
	Column		= obje.Column
	Description = sqlSafe (obje.Description)
	Extended	= sqlSafe (obje.ASPDescription)

	rquerystring = sqlSafe(request.querystring)
	tmpDate		= date()
	tmpTime		= time()
	rdate		= year(tmpDate) & "/" & pad(Month(tmpDate),2) & "/" & pad(Day(tmpDate),2)
	rtime		= hour(tmpTime)  & ":" & minute(tmpTime)  & ":" & second(tmpTime)
	httpReferer = sqlSafe(Request.ServerVariables("HTTP_REFERER"))
	clientIP	= Request.ServerVariables("REMOTE_ADDR")
	rURL		= Request.ServerVariables("URL")

	rform	= "Binary Form"
	ON ERROR Resume Next
		rform	= sqlSafe(request.form)
	ON ERROR GOTO 0

%>
<div dir="ltr">
<TABLE align="center" dir="ltr" style="border:2pt solid white; width=50%">
<TR>
	<TD align="center" colspan=2 bgcolor="eeeeee">Ã“ÌÌ«  ›‰Ì «‘ﬂ«·</TD>
</TR>
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
	<TD align="left"> <%=replace(rquerystring,"&","<br>")%></TD>
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
</TABLE><BR>
<%
'errConStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
errConStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"
'***** NOTE: If conStr changed, you should change it in /config.asp too

Set errConn = Server.CreateObject("ADODB.Connection")
errConn.open errConStr
if session("id")<>"" then errBy = session("id") else errBy = "-100"

errSQL = "INSERT INTO errLog (errDate, errTime, errBy, url, httpReferer, clientIP, aspCode, Number, Source, Category, [File], Line, [Column], Description, Extended, requestForm, requestQuerystring) VALUES ('" & rdate & "','" & rtime & "'," & errBy & ",'" & rURL & "' ,N'" & httpReferer & "','" & clientIP & "','" & aspCode & "','" & Number & "','" & Source & "','" & Category & "','" & File & "','" & Line & "','" & Column & "','" & Description & "','" & Extended & "','" & rform & "','" & rquerystring & "')"

errConn.Execute errSQL 
'response.write errSQL 
'response.end

errSQL = "SELECT id FROM errLog where errDate='" & rdate & "' and errTime='" & rtime & "'"

set errRS = errConn.Execute (errSQL)
	if not errRS.EOF then
		response.write "<CENTER>»—«Ì ÅÌêÌ—Ì »⁄œÌ «Ì‰ ‘„«—Â Ê ‰ÕÊÂ —”Ìœ‰ »Â «Ì‰ «‘ﬂ«· —« Ì«œœ«‘  ﬂ‰Ìœ:  " & 		"<BR><div dir=ltr style='font:17pt'> " & errRS("id") & "</div></CENTER><BR>"
		
	end if
'Show All Server Variables
'for i = 1 to 49
'response.write Request.ServerVariables.Key(i) & " = " & Request.ServerVariables(i) & "<br>" 
'next 
 %>
</div>
