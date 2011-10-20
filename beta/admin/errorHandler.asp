<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle=server.HTMLEncode("„œÌ—Ì  Œÿ«")
SubmenuItem=10
%>
<!--#include file="top.asp" -->
<!--#include file="../include_farsiDateHandling.asp" -->
<style>
.myTable
TD { 
	border-bottom: 1pt solid black;
	border-left: 1pt solid gray;
	border-right: 1pt solid gray;
	}
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function selectAll(src){
		totalItems=document.getElementsByName("errID").length
		checked=src.checked
		for (i=0;i<totalItems;i++)
			document.getElementsByName("errID")[i].checked=checked;
	}
	
//-->
</SCRIPT>
<%

function pad (inpStr,padSize)
	result = inpStr
	while len(result) < padSize 
		result = "0" & result
	wend
	pad = result
end function

function sqlSafe (s)
  st=s
  st=replace(St,"'","`")
  st=replace(St,chr(34),"`")
  sqlSafe=st
end function

session.codepage="1252"
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit new AO MEMO
'-----------------------------------------------------------------------------------------------------
if request("act")="update" then
	
	CheckerComment = sqlSafe(request.form("CheckerComment"))

	for i=1 to request("errID").count
		errID=request("errID")(i)
		Conn.Execute("UPDATE errLog SET Checked = 1, CheckedBy ="& session("id") & ", CheckerComment =N'" & CheckerComment & "' where id = " & errID)
	next
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit new AO MEMO
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showDetails" then
	

	'======================================================= Khotoote Zir Kaare Khaasi nemikonan 
	'set RSV=Conn.Execute ("SELECT id,errdate FROM errLog") 
	'Do while not RSV.eof
	'	errDate = RSV("errDate")
	'	a = right(errDate,len(errDate)-5)
	'	newErrDate = left(errDate,4) & "/" & pad(left(a,instr(a,"/")-1),2) & "/"& pad(mid(a,instr(a,"/")+1,len(a)-instr(a,"/")),2) 
	'	Conn.Execute ("update errLog set errDate=N'"& newErrDate & "' WHERE id=" & RSV("id")) 
	'RSV.moveNext
	'Loop
	'RSV.close
	'response.write "<br>Radife" 
	'response.end
	'===========================================================================================  

end if
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit new AO MEMO
'-----------------------------------------------------------------------------------------------------
DateFrom		= sqlSafe(request.form("DateFrom"))
DateTo			= sqlSafe(request.form("DateTo"))
status			= sqlSafe(request.form("status"))
errBy			= sqlSafe(request.form("errBy"))
perPage			= sqlSafe(request.form("perPage"))
File			= sqlSafe(request.form("File"))
LineNo			= sqlSafe(request.form("LineNo"))

if request("s") = "y" then
	DateFrom		= sqlSafe(session("DateFrom"))
	DateTo			= sqlSafe(session("DateTo"))
	status			= sqlSafe(session("status"))
	errBy			= sqlSafe(session("errBy"))
	perPage			= sqlSafe(session("perPage"))
	File			= sqlSafe(session("File"))
	LineNo			= sqlSafe(session("LineNo"))

end if

session("DateFrom")	= DateFrom
session("DateTo")	= DateTo
session("status")	= status
session("errBy")	= errBy
session("perPage")	= perPage
session("File")		= File
session("LineNo")	= LineNo

if not ( perPage <> "" and isNumeric(perPage) ) then
	perPage = 20
end if

%>
<BR><BR>

<FORM METHOD=POST ACTION="errorHandler.asp">
	&nbsp;&nbsp;
	«“  «—ÌŒ <INPUT  dir="LTR"  TYPE="text" NAME="DateFrom" maxlength="10" size="10"  value="<%=DateFrom%>"> &nbsp;&nbsp;
	 «  «—ÌŒ <INPUT  dir="LTR"  TYPE="text" NAME="DateTo" maxlength="10" size="10"  value="<%=DateTo%>"> &nbsp;&nbsp;
	 Ê”ÿ <select name="errBy" class=inputBut >
		<option value="1000">Â„Â</option>
		<option value="1000">-------------------</option>
	<% set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
	Do while not RSV.eof
	%>
		<option value="<%=RSV("ID")%>" <%
			if trim(RSV("ID"))=trim(errBy) then
				response.write " selected "
			end if
			%>><%=server.HTMLEncode(RSV("RealName"))%></option>
	<%
	RSV.moveNext
	Loop
	RSV.close
	%>
	</select> 
	&nbsp;&nbsp;Ê÷⁄Ì  <select name="status" class=inputBut >
		<option value="1000">Â„Â</option>
		<option value="1" <% if trim(status)="1" then %> selected <% end if %>>»——”Ì ‘œÂ</option>
		<option value="0" <% if trim(status)="0" then %> selected <% end if %>>»——”Ì ‰‘œÂ</option>
	</select> &nbsp;&nbsp;
	&nbsp;&nbsp; ⁄œ«œ<select name="perPage" class=inputBut >
		<option value="10" <% if trim(perPage)="10" then %> selected <% end if %>>10</option>
		<option value="20" <% if trim(perPage)="20" then %> selected <% end if %>>20</option>
		<option value="50" <% if trim(perPage)="50" then %> selected <% end if %>>50</option>
		<option value="100" <% if trim(perPage)="100" then %> selected <% end if %>>100</option>
		<option value="200" <% if trim(perPage)="200" then %> selected <% end if %>>200</option>
		<option value="500" <% if trim(perPage)="500" then %> selected <% end if %>>500</option>
		<option value="1000" <% if trim(perPage)="1000" then %> selected <% end if %>>1000</option>
	</select> &nbsp;&nbsp;
	<br><br>
	&nbsp;&nbsp;
	›«Ì· <INPUT  dir="LTR"  TYPE="text" NAME="File" size="20"  value="<%=File%>"> &nbsp;&nbsp;
	‘„«—Â Œÿ <INPUT  dir="LTR"  TYPE="text" NAME="LineNo" maxlength="5" size="5"  value="<%=LineNo%>"> &nbsp;&nbsp;

	<INPUT TYPE="submit" NAME="submit" value="„‘«ÂœÂ">
</FORM>

<%
	conditions = "(1=1)"

	if DateFrom <> "" then
		conditions = conditions & " AND errLog.errDate >= N'"& DateFrom & "' "
	end if

	if DateTo <> "" then
		conditions = conditions & " AND errLog.errDate <= N'"& DateTo & "' "
	end if

	if errBy <> "1000" and errBy <> "" then
		conditions = conditions & " AND errLog.errBy = " & errBy
	end if

	if status <> "1000" and status <> "" then
		conditions = conditions & " AND errLog.Checked = "& status 
	end if

	if File <> "" then
		conditions = conditions & " AND errLog.[File] Like '%"& File & "%'"
	end if

	if LineNo <> "" then
		conditions = conditions & " AND errLog.Line = "& cint(LineNo) 
	end if

SA_RecordLimit = perPage

Set SA_RS1 = Server.CreateObject("ADODB.Recordset")
SA_RS1.CursorLocation=3 '---------- Alix: Because in ADOVBS_INC adUseClient=3
SA_RS1.PageSize = SA_RecordLimit

SA_mySQL="SELECT errLog.[file] as myFile, errLog.[line] as myLine, errLog.*, Users.RealName, Users_1.RealName AS checkerName FROM errLog INNER JOIN Users ON errLog.errBy = Users.ID LEFT OUTER JOIN Users Users_1 ON errLog.CheckedBy = Users_1.ID where "& conditions & " order by errLog.ID DESC"

'		SA_RS1.Open SA_mySQL ,Conn,adOpenStatic,,adcmdcommand 
SA_RS1.Open SA_mySQL ,Conn,3,,adcmdcommand 
FilePages = SA_RS1.PageCount
Records = SA_RS1.RecordCount

Page=1
if isnumeric(Request.QueryString("p")) then
  if (clng(Request.QueryString("p")) <= FilePages) and clng(Request.QueryString("p") > 0) then
	Page=clng(Request.QueryString("p"))
  end if
end if

if not SA_RS1.eof then
	SA_RS1.AbsolutePage=Page
end if

%>
&nbsp;&nbsp;&nbsp;&nbsp; ⁄œ«œ ﬂ· ‰ «ÌÃ Ã” ÃÊ : <%=Records%><BR><BR>
<TABLE class="myTable" border=0 align=center width=95% cellspacing=0 cellpadding=2>
	<TR bgcolor=white>
		<TD><INPUT TYPE="checkbox" NAME="aaa" onclick="selectAll(this)"></TD>
		<TD> «—ÌŒ</TD>
		<TD> Ê”ÿ</TD>
		<TD>›«Ì·</TD>
		<TD>Œÿ</TD>
		<TD>ÅÌ€«„</TD>
		<TD>Form</TD>
		<TD>QueryString</TD>
	</TR>
	<FORM METHOD=POST ACTION="?act=update&p=<%=page%>&s=y">
<%
if (SA_RS1.eof) then ' Not Found
	response.write "<br><TR align=center><TD colspan=8>"

	response.write server.HTMLEncode("çÌ“Ì Ì«›  ‰‘œ" )

	response.write "</TD></TR>"
	session.codepage="1256"
	response.end

end if 
Do while not SA_RS1.eof and (SA_RS1.AbsolutePage = Page)
	%>
	<A HREF="errorDetails.asp?id=<%=SA_RS1("id")%>" target="_blank">
	<TR dir=LTR style="cursor:hand">
		
		<TD dir=rtl title="<%=SA_RS1("id")%>&nbsp;<% if SA_RS1("checkerName")<>"" then %> <%=chr(13)%><%=server.HTMLEncode(" Ê”ÿ " & SA_RS1("checkerName"))%> [<%=SA_RS1("CheckerComment")%>]<% end if %>" style="border-bottom: 1pt solid black">
			<INPUT TYPE="checkbox" name="errID" <% if SA_RS1("Checked") then %> checked disabled<% end if %> value="<%=SA_RS1("id")%>">
		</TD>
		
		<TD title="<%=shamsiDate(cdate(SA_RS1("errDate"))) & vbCrLf & SA_RS1("errTime")%>">
			<%=SA_RS1("errDate")%>
		</TD>
		
		<TD dir=rtl title="clientIP:<%=SA_RS1("clientIP")%>">
			<%=server.HTMLEncode(SA_RS1("RealName"))%>
		</TD>
		

		<TD title="URL:<%=replace(SA_RS1("URL"),"/beta/","")%><%=chr(13)%>-------------------------------------<%=chr(13)%><%=SA_RS1("httpReferer")%>">
			<INPUT TYPE="text" style="background:transparent;border:none; width:140pt" readonly value="<%=replace(unescape(SA_RS1("myFile")),"/beta/", "")%>">
		</TD>
		<TD title="Column:<%=SA_RS1("Column")%>">
			<%=SA_RS1("myLine")%>
		</TD>

		<TD title="<%=SA_RS1("Description")%><% if SA_RS1("Source")<>"" then%><%=chr(13)%>-----Source: ---------------------------------<%=chr(13)%><% end if %><%=server.HTMLEncode(SA_RS1("Source"))%>">
			<INPUT TYPE="text" style="background:transparent;border:none; width:140pt" readonly value="<%=SA_RS1("Description")%>">
		</TD>

		<TD title="<%=replace(replace(unescape(SA_RS1("requestForm")),"&", chr(13)),"+"," ")%>">
			<INPUT TYPE="text" style="background:transparent;border:none; width:70pt" readonly value="<%=replace(unescape(SA_RS1("requestForm")),"&", chr(13))%>">
		</TD>

		<TD title="<%=replace(replace(unescape(SA_RS1("requestQuerystring")),"&", chr(13)),"+"," ")%>">
			<INPUT TYPE="text" style="background:transparent;border:none; width:70pt" readonly value="<%=replace(unescape(SA_RS1("requestQuerystring")),"&", chr(13))%>"  title="<%=replace(replace(unescape(SA_RS1("requestQuerystring")),"&", chr(13)),"+"," ")%>">
		</TD>
	</TR>
	</A>
	<%
	SA_RS1.moveNext
Loop
SA_RS1.close
%>
	<TR bgcolor=white>
		<TD colspan=8>
	<%
			if FilePages>1 then
				response.write server.HTMLEncode("’›ÕÂ ") & "&nbsp;"
				for i=1 to FilePages 
					if i = page then %>
						[<%=i%>]&nbsp;
<%					else%>
						<A HREF="?p=<%=i%>&s=y&act=HX6600"><%=i%></A>&nbsp;
<%					end if 
				next 
			end if
%>
		</TD>
	</TR>
</TABLE>
<TABLE align=center>
<TR>
		<TD colspan=7 align=center>
			<BR><BR>
			 Ê÷ÌÕ: <INPUT TYPE="text" NAME="CheckerComment" size=126>
			<INPUT TYPE="submit" value="À» ">
		</TD>
	</TR>
	</FORM>
</TABLE><BR><BR>
<%
session.codepage="1256"
%>
<!--#include file="tah.asp" -->