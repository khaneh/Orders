<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle= "„œÌ—Ì  œ› — ﬂ·"
SubmenuItem=9


%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------- Submit Change Default GL
'-----------------------------------------------------------------------------------------------------
if request("act")="delRow" then
	rowNo =request("rowNo")
	mySQL="UPDATE UserDefaults SET WorkingGL=Null WHERE ([User]="& rowNo & ")"
	conn.Execute(mySQL)

	Set RS2 = conn.Execute("SELECT GLs.EndDate, GLs.StartDate, GLs.Name, GLs.ID, GLs.FiscalYear, UserDefaults.[User] FROM GLs INNER JOIN UserDefaults ON GLs.ID = UserDefaults.WorkingGL WHERE (UserDefaults.[User] = '"& session("id") & "') OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")
	session("OpenGL")=RS2("id")
	session("FiscalYear")=RS2("FiscalYear")
	session("OpenGLName")=RS2("name")
	session("OpenGLStartDate")=RS2("StartDate")
	session("OpenGLEndDate")=RS2("EndDate")
	RS2.movenext
	session("differentGL") = False
	RS2.close
	
	response.redirect "?act=select"

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------- Submit Change Default GL
'-----------------------------------------------------------------------------------------------------
elseif request("act")="change" then

	newOpenGL=request("OpenGL")
	forWhom=request("forWhom")

	if forWhom = 0 then

		conn.Execute("UPDATE GLs SET IsDefault = 0")
		conn.Execute("UPDATE GLs SET IsDefault = 1 , LastChangedBy="& session("ID")& " , LastChangedDate=N'"& shamsiToday() & "', LastChange=N'set to default' WHERE (ID = "& newOpenGL & ")")

		'Changed 13282/01/09 for using UserDefaults table instead
		'Set RS2 = conn.Execute("SELECT FiscalYear, name, ID FROM GLs WHERE (IsDefault = 1) ORDER BY ID DESC")

		mySQL="UPDATE UserDefaults SET WorkingGL='"& newOpenGL & "' WHERE ([User]=0)"
		conn.Execute(mySQL)
		responseRedirectAddress = "AccountInfo.asp?msg=" & Server.URLEncode("”«· „«·Ì ”Ì” „ »—«Ì Â„Â ﬂ«—»—«‰  €ÌÌ— ﬂ—œ.")
	
	else
		Set RS2 = conn.Execute("SELECT [User] as ID, WorkingGL FROM UserDefaults WHERE ([User] = "& forWhom & ") OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")

		if RS2("id")="0" then
			mySQL="INSERT INTO UserDefaults ([User], WorkingGL) VALUES ("& forWhom & ","& newOpenGL & ")"
			conn.Execute(mySQL)
		else
			mySQL="UPDATE UserDefaults SET WorkingGL='"& newOpenGL & "' WHERE ([User]="& forWhom & ")"
			conn.Execute(mySQL)
		end if

		responseRedirectAddress = "AccountInfo.asp?msg=" & Server.URLEncode("”«· „«·Ì ﬂ«—»— „Ê—œ ‰Ÿ— (" & forWhom & ")  €ÌÌ— ﬂ—œ")
	end if

	Set RS2 = conn.Execute("SELECT GLs.EndDate, GLs.StartDate, GLs.Name, GLs.ID, GLs.FiscalYear, UserDefaults.[User] FROM GLs INNER JOIN UserDefaults ON GLs.ID = UserDefaults.WorkingGL WHERE (UserDefaults.[User] = '"& session("id") & "') OR (UserDefaults.[User] = 0) ORDER BY ABS(UserDefaults.[User]) DESC")
	session("OpenGL")=RS2("id")
	session("FiscalYear")=RS2("FiscalYear")
	session("OpenGLName")=RS2("name")
	session("OpenGLStartDate")=RS2("StartDate")
	session("OpenGLEndDate")=RS2("EndDate")
	RS2.movenext
	session("differentGL") = False
	if not RS2.EOF then
		temp=RS2("id")
		if temp <> session("OpenGL") then
			session("differentGL") = True
		end if
	end if
	RS2.close

	response.redirect responseRedirectAddress

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- Submit Create New GL
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitNew" then

	name=request.form("name")
	FiscalYear=request.form("FiscalYear")
	copy = request.form("copy")
	OpenGL = request.form("OpenGL")


	CurrentDate = shamsiToday()
	tempStartDate=	FiscalYear & "/01/01" 'left(CurrentDate,4)&"/01/01"
	tempEndDate= FiscalYear & "/12/30" 'left(CurrentDate,4)&"/12/30"

	Set RS2 = conn.Execute("SELECT * FROM GLs WHERE ID=" & OpenGL)
	SourceGL = RS2("ID")
	SourceGL_Name = RS2("Name")
	RS2.close

	conn.Execute("INSERT INTO GLs (Name, FiscalYear, StartDate, EndDate, IsDefault, IsClosed, LastChangedBy, LastChangedDate, LastChange) VALUES ( N'"& name & "', N'"& FiscalYear & "', N'"& tempStartDate & "', N'"& tempEndDate & "', 0, 0, "& session("id") & ", N'"& CurrentDate & "', N'Created')")

	Call showAlert ("œ› — ﬂ· ÃœÌœ «ÌÃ«œ ‘œ",CONST_MSG_INFORM)

	if copy="on" then
		Set RS2 = conn.Execute("SELECT * from GLs where FiscalYear=" & FiscalYear & " order by id DESC")
		GLID=RS2("id")
		conn.Execute("INSERT INTO GLAccountSuperGroups SELECT ID, Name, " & GLID & " AS GL, Type FROM GLAccountSuperGroups WHERE (GL = "& SourceGL & ")")
		conn.Execute("INSERT INTO GLAccountGroups SELECT ID, GLSuperGroup, Name, " & GLID & " AS GL FROM GLAccountGroups WHERE (GL = "& SourceGL & ")")
		conn.Execute("INSERT INTO GLAccounts SELECT ID, GLGroup, Name, " & GLID & " AS GL,HasAppendix, IsChequeAccount FROM GLAccounts WHERE (GL = "& SourceGL & ")")
		response.write "<BR><BR>"
		Call showAlert ("Õ”«» Â« Ê ê—ÊÂ Â« «“ " & SourceGL_Name & " ﬂÅÌ ‘œ",CONST_MSG_INFORM)


		conn.Execute("INSERT INTO AXItemReasonGLAccountRelations (ItemReason, GLAccount, GL) SELECT ItemReason, GLAccount, " & GLID & " AS GL FROM AXItemReasonGLAccountRelations WHERE (GL = " & SourceGL & ")")
		response.write "<BR><BR>"
		Call showAlert ("AXItemReasonGLAccountRelations «“ " & SourceGL_Name & " ﬂÅÌ ‘œ",CONST_MSG_INFORM)

		conn.Execute("INSERT INTO BankerCheqStatusGLAccountRelation (Banker, CheqStatus, GLAccount, GL) SELECT Banker, CheqStatus, GLAccount, " & GLID & " AS GL FROM BankerCheqStatusGLAccountRelation WHERE (GL = " & SourceGL & ") ORDER BY Banker, CheqStatus, GLAccount")
		response.write "<BR><BR>"
		Call showAlert ("BankerCheqStatusGLAccountRelation «“ " & SourceGL_Name & " ﬂÅÌ ‘œ",CONST_MSG_INFORM)

	end if



'"INSERT INTO GLAccounts SELECT ID, GLGroup, Name, 3 AS GL FROM GLAccounts WHERE (GL = 1)"

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------- Close an open GL Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="new" then
%>
<CENTER>
<BR><BR><BR><BR>
<FORM METHOD=POST ACTION="manageGL.asp?act=submitNew">
<TABLE>
<TR>
	<TD colspan=2 align=center><H3>«ÌÃ«œ œ› — ﬂ· ÃœÌœ</H3></TD>
</TR>
<TR>
	<TD align=left>⁄‰Ê«‰ œ› —: </TD>
	<TD><INPUT TYPE="text" NAME="name"></TD>
</TR>
<TR>
	<TD align=left>”«· „«·Ì: </TD>
	<TD><INPUT TYPE="text" NAME="FiscalYear" onkeypress="return maskNumber(this)" maxlength=4></TD>
</TR>
<TR>
	<TD colspan=2><BR></TD>
</TR>
<TR>
	<TD align=left><INPUT TYPE="checkbox" NAME="copy" checked> ﬂÅÌ Õ”«» Â« Ê ê—ÊÂ Â« «“ </TD>
	<TD>
		<%
		set RSSX=Conn.Execute ("SELECT * from GLs where IsClosed=0 order by id desc")
		%>
		<SELECT NAME="OpenGL" style="font-family:tahoma" class="GenButton">
		<%
		Do while not RSSX.eof
		%>
		<option <%	if cint(OpenGL)=cint(RSSX("ID")) then 
						response.write "selected"
						GLname = RSSX("name")
					end if %> value="<%=RSSX("ID")%>" ><%=RSSX("name")%> - ”«· „«·Ì <%=RSSX("FiscalYear")%></option>
		<% 

		RSSX.moveNext
		Loop
		RSSX.close
		%>
		</SELECT>
	</TD>
</TR>
<TR>
	<TD colspan=2 align=center> 
		<BR><BR>
		<INPUT TYPE="submit" value="«ÌÃ«œ"class="GenButton">
		<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp'">
	</TD>
</TR>
</TABLE>
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------- Close an open GL Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="close" then
%>
<CENTER>
<BR><BR>
<% call showAlert ("<B> ÊÃÂ: </B><BR>»” ‰ Ìﬂ ﬂ· œ› — »—ê‘  ‰«Å–Ì— «” <BR>Ê »«⁄À „Ì ‘Êœ «“ «Ì‰ Å” À»  ÂÌç ”‰œÌ<BR>œ— ¬‰ «„ﬂ«‰ Å–Ì— ‰»«‘œ.<BR>",CONST_MSG_INFORM)%>

<BR>
<% call showAlert ("œ› — ﬂ· Ã«—Ì —« ﬁ»· «“  ⁄ÌÌ‰ œ› — ÅÌ‘ ›—÷ ÃœÌœ ‰„Ì  Ê«‰ »” ",CONST_MSG_INFORM)%>

<BR><BR>
<%
set RSSX=Conn.Execute ("SELECT * from GLs where IsClosed=0 and IsDefault=0 order by id desc")
%>
<FORM METHOD=POST ACTION="manageGL.asp?act=submitClose">

<B>»” ‰ œ› — ﬂ· </B><SELECT NAME="OpenGL" style="font-family:tahoma" class="GenButton">
<%
Do while not RSSX.eof
%>
<option <%	if cint(OpenGL)=cint(RSSX("ID")) then 
				response.write "selected"
				GLname = RSSX("name")
			end if %> value="<%=RSSX("ID")%>" ><%=RSSX("name")%> - ”«· „«·Ì <%=RSSX("FiscalYear")%></option>
<% 

RSSX.moveNext
Loop
RSSX.close
%>
</SELECT><BR><BR>
<INPUT TYPE="submit" value="»” ‰"class="GenButton">
<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp'">
</FORM>
</CENTER>

<%

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------- submit Closing a GL
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitClose" then
	conn.Execute("UPDATE GLs SET IsClosed = 1 , LastChangedBy="& session("ID")& " , LastChangedDate=N'"& shamsiToday() & "', LastChange=N'Closed' WHERE (ID = "& request("OpenGL") & ")")

	response.redirect "AccountInfo.asp"


'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- Show GLs Changes Log
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showLog" then
	set RSSX=Conn.Execute ("SELECT Users.RealName, GLLog.ChangedDate, GLLog.Change, GLs.Name ,GLs.id , GLs.FiscalYear  FROM GLLog INNER JOIN GLs ON GLLog.GLID = GLs.ID INNER JOIN Users ON GLLog.ChangedBy = Users.ID ORDER BY GLLog.id DESC")
	%>
	<TABLE dir=rtl align=center width=600>
	<TR >
		<TD colspan=6 align=center>
			<BR><H3>ê“«—‘  €ÌÌ—«  œ› — ﬂ·</H3><BR>
		</TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD><SMALL> # </SMALL></TD>
		<TD><SMALL> œ› — </SMALL></TD>
		<TD><SMALL> ”«· „«·Ì</SMALL></TD>
		<TD><SMALL>  «—ÌŒ  €ÌÌ— </SMALL></TD>
		<TD><SMALL>  €ÌÌ— œÂ‰œÂ </SMALL></TD>
		<TD><SMALL>  Ê÷ÌÕ </SMALL></TD>
	</TR>
	<%
	tmpCounter=0
	Do while not RSSX.eof
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 
		%>
		<TR bgcolor="<%=tmpColor%>" >
			<TD><%=RSSX("id")%> </TD>
			<TD><%=RSSX("Name")%> </TD>
			<TD><%=RSSX("FiscalYear")%></TD>
			<TD><span dir=ltr><%=RSSX("ChangedDate")%></span></TD>
			<TD><%=RSSX("RealName")%></TD>
			<TD><%=RSSX("Change")%></TD>
			</TD>
		</TR>
		<%
		RSSX.moveNext
	Loop
	%>
	</TABLE><br>
	<CENTER><INPUT TYPE="button" value="»—ê‘ "class="GenButton" onclick="window.location='AccountInfo.asp'"></CENTER><BR><BR>

	<%
	set RSSX = nothing

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- Select Another Default GL Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="select" or request("act")="" then
%>
<CENTER>
<BR><BR><BR>
<% call showAlert ("<B> ÊÃÂ: </B><BR> €ÌÌ— œ› — ﬂ· ÅÌ‘ ›—÷ œ— «Ì‰ ﬁ”„ <BR> »«⁄À „Ì ‘Êœ «“ «Ì‰ Å”  „«„  —«ﬂ‰‘ Â« <BR> œ— «Ì‰œ› — À»  ‘Ê‰œ",CONST_MSG_INFORM)%>
<BR><BR>
<% call showAlert (" €ÌÌ— œ› — ﬂ·  »— Ê÷⁄Ì  ﬂ«—»—«‰Ì ﬂÂ «ﬂ‰Ê‰<BR> œ— ”Ì” „ „‘€Ê· ﬂ«— Â” ‰œ „ÊÀ— ‰Ì” .",CONST_MSG_INFORM)%>
<BR><BR>
<% call showAlert ("ﬂ«—»—«‰Ì ﬂÂ œ— ›Â—”  “Ì— »—«Ì‘«‰ œ› — ﬂ· Œ«’Ì <BR>„‘Œ’ ‘œÂ «“ œ› — ﬂ· ÅÌ‘ ›—÷ ”Ì” „ <BR>«” ›«œÂ ‰„Ì ﬂ‰‰œ.",CONST_MSG_INFORM)%>
<BR><BR>
<%
set RSSX=Conn.Execute ("SELECT * from GLs where IsClosed=0 order by id desc")
set RS_UsersList=Conn.Execute ("SELECT Users.ID, Users.RealName, GLs.Name FROM Users LEFT OUTER JOIN UserDefaults ON Users.ID = UserDefaults.[User] LEFT OUTER JOIN GLs ON UserDefaults.WorkingGL = GLs.ID ORDER BY Users.RealName")
%>
<TABLE>
<TR>
	<TD colspan=3>
		<FORM METHOD=POST ACTION="manageGL.asp?act=change">
		<B> €ÌÌ— œ› — ﬂ· »Â </B><SELECT NAME="OpenGL" style="font-family:tahoma" class="GenButton">
		<%
		Do while not RSSX.eof
		%>
		<option <%	if cint(OpenGL)=cint(RSSX("ID")) then 
						response.write "selected"
						GLname = RSSX("name")
					end if %> value="<%=RSSX("ID")%>" ><%=RSSX("name")%> - ”«· „«·Ì <%=RSSX("FiscalYear")%></option>
		<% 
		RSSX.moveNext
		Loop
		RSSX.close
		%>
		</SELECT><BR><BR>
	</TD>
</TR>
<%
Do while not RS_UsersList.eof
%>
<TR>
	<TD>
		<% if RS_UsersList("id")="0" then%>
			<INPUT TYPE="radio" NAME="forWhom" value="0"><FONT COLOR="red"><B>œ› — ﬂ· ÅÌ‘ ›—÷ ”Ì” „</B></FONT><br>
		<% else 
			UsersRealName = RS_UsersList("RealName")
			UsersID = RS_UsersList("id")
			%>
			<INPUT TYPE="radio" NAME="forWhom" <% if UsersID=session("id") then %>checked <% end if %> value="<%=UsersID%>">›ﬁÿ »—«Ì ﬂ«—»— <%=UsersRealName%>
		<% end if %>
	</TD>
	<TD align=left width=120>
		<%=RS_UsersList("Name")%> &nbsp;&nbsp;
	</TD>
	<TD align=right>
		<% if not RS_UsersList("id")="0" and not isNull(RS_UsersList("Name")) then%>
			<A HREF="?act=delRow&rowNo=<%=RS_UsersList("id")%>" title="»« Õ–› «Ì‰ ”ÿ— «Ì‰ ﬂ«—»— ‰Ì“ «“ «Ì‰ Å” «“ œ› — ﬂ· ÅÌ‘ ›—÷ ”Ì” „ «” ›«œÂ ŒÊ«Âœ ﬂ—œ.">[Õ–›]</A>
		<% else %>
		 &nbsp;
		<% end if %>
	</TD>
</TR>
<% 
RS_UsersList.moveNext
Loop
RS_UsersList.close
%>
</TABLE><BR><BR>
<INPUT TYPE="submit" value=" €ÌÌ—"class="GenButton">
<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp'">
</FORM>
</CENTER>

<%
end if
%>
<!--#include file="tah.asp" -->

<%
	'------------- Related Triggers ----------------
	'
	'	CREATE TRIGGER [setDefaultGL] ON [dbo].[GLs] 
	'	FOR INSERT, UPDATE
	'	AS 
	'	DECLARE 
	'		@IsDefault binary,
	'		@id int
	'		SELECT 
	'			@IsDefault = IsDefault, 
	'			@id = [id]
	'		FROM inserted 
	'	if @IsDefault = 1
	'	begin
	'	UPDATE GLs SET 
	'		IsDefault = 0
	'	UPDATE GLs SET 
	'		IsDefault = 1
	'	WHERE ID = @id
	'	end
	'
	'----------------------------------------------
	'
	'	CREATE TRIGGER Trig_GLLOGGER ON [dbo].[GLs] 
	'	FOR INSERT, UPDATE
	'	AS 
	'	DECLARE 
	'		@GLID int,
	'		@LastChange nvarchar(100),
	'		@LastChangedDate nvarchar(10),
	'		@LastChangedBy int
	'		SELECT 
	'			@GLID=[ID], 
	'			@LastChange = LastChange, 
	'			@LastChangedDate = LastChangedDate, 
	'			@LastChangedBy = LastChangedBy
	'		FROM inserted
	'	IF UPDATE (LastChange)
	'	INSERT INTO GLLog (GLID, Change, ChangedDate, ChangedBy)
	'			VALUES (@GLID, @LastChange, @LastChangedDate, @LastChangedBy)
	'
	'
	'----------------------------------------------

%>