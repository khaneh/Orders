<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle= "„œÌ—Ì  œ› — ﬂ·"
SubmenuItem=9

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%

' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * *   Account 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Edit GL Accounts submit
'-----------------------------------------------------------------------------------------------------
if request("act")="EditAccountSubmit" then
	ACCID = request("ACCID")
	name = request("name")
	GroupID = request("GroupID")
	accountType = request("accountType")
	If request("tafsil")="on" Then 
		tafsil = 1
	Else 
		tafsil = 0
	End If 

	conn.Execute("UPDATE GLAccounts SET Name = N'"& name & "', HasAppendix = "&tafsil&", accountType=" & accountType & " WHERE (ID = "& ACCID & ") AND (GL = "& OpenGL & ")")
	response.redirect "AccountInfo.asp?act=account&GroupID=" & GroupID

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Add GL Accounts submit
'-----------------------------------------------------------------------------------------------------
elseif request("act")="NewAccountSubmit" then
	ACCID = request("ACCID")
	name = request("name")
	GroupID = request("GroupID")

	Set RS2 = conn.Execute("SELECT * FROM GLAccounts WHERE ID="& ACCID & " AND GL="& OpenGL )
	if not RS2.EOF then
		response.write "<BR><BR><BR><BR><CENTER>Œÿ«! «Ì‰ ‘„«—Â Õ”«» ÊÃÊœ œ«—œ"
		response.write "<BR><BR><A HREF='AccountInfo.asp?act=editAccountForms&GroupID="& GroupID & "&submit=«›“Êœ‰ Õ”«»'>»—ê‘ </A></CENTER>"
		response.end
	else
		conn.Execute("INSERT INTO GLAccounts (GLGroup, Name, GL, ID) VALUES ("& GroupID & ", N'"& name & "',"& OpenGL & ", "& ACCID & ")")
		response.redirect "AccountInfo.asp?act=account&GroupID=" & GroupID
	end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------- Add / Delete / Edit GL Accounts Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="editAccountForms" then
	ACCID = request("ACCID")
	submit = request("submit")
	GroupID = request("GroupID")

	Set RS2 = conn.Execute("SELECT * from GLAccountGroups where id="& GroupID & " and GL="& OpenGL )
	GroupName=RS2("name")
	
	'======================================== «›“Êœ‰ Õ”«»
	'====================================================
	if submit="«›“Êœ‰ Õ”«»" then
		Set RS2 = conn.Execute("SELECT MAX(id) AS MaxID FROM GLAccounts WHERE GLGroup="& GroupID & " AND GL="& OpenGL )
		MaxID=RS2("MaxID")
		if isnull(MaxID) then MaxID= GroupID
		%>
		<BR><BR>
		<TABLE align=center>
		<FORM METHOD=POST ACTION="editGL.asp?act=NewAccountSubmit">
		<TR>
			<TD colspan=2 align=center><H3>«›“Êœ‰ Õ”«»</H3><BR>
			œ— ê—ÊÂ <%=GroupName%> (ﬂœ <%=GroupID%>)<BR><BR>
			œ— <%=OpenGLName%> (”«· „«·Ì <%=FiscalYear%>)
			<INPUT TYPE="hidden" name="GroupID" value="<%=GroupID%>">
			<BR><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD>‘„«—Â Õ”«»</TD>
			<TD><INPUT TYPE="text" NAME="ACCID" onkeypress="return maskNumber(this)" maxlength=5 value="<%=MaxID+1%>"></TD>
		</TR>
		<TR>
			<TD>‰«„ Õ”«»</TD>
			<TD><INPUT TYPE="text" NAME="name"></TD>
		</TR>
		<TR>
			<TD colspan=2 align=center>
				<BR><BR>
				<INPUT TYPE="submit" value="«ÌÃ«œ"class="GenButton">
				<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp?act=account&GroupID=<%=GroupID%>'">
			</TD>
		</TR>
		</FORM>
		</TABLE>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.all.name.focus();
		//-->
		</SCRIPT>
		<%

	'======================================== ÊÌ—«Ì‘ Õ”«»
	'====================================================
	elseif submit="ÊÌ—«Ì‘" then
		Set RS2 = conn.Execute("SELECT * FROM GLAccounts WHERE ID="& ACCID & " AND GL="& OpenGL )
		AccountName=RS2("name")
		%>
		<BR><BR>
		<TABLE align=center>
		<FORM METHOD=POST ACTION="editGL.asp?act=EditAccountSubmit">
		<TR>
			<TD colspan=2 align=center><H3>ÊÌ—«Ì‘ Õ”«»</H3><BR>
			œ— ê—ÊÂ <%=GroupName%> (ﬂœ <%=GroupID%>)<BR><BR>
			œ— <%=OpenGLName%> (”«· „«·Ì <%=FiscalYear%>)
			<INPUT TYPE="hidden" name="GroupID" value="<%=GroupID%>">
			<BR><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD>‘„«—Â Õ”«»</TD>
			<TD><INPUT TYPE="text" NAME="ACCID" onkeypress="return maskNumber(this)" maxlength=5 value="<%=ACCID%>" readonly></TD>
		</TR>
		<TR>
			<TD>‰«„ Õ”«»</TD>
			<TD><INPUT TYPE="text" NAME="name" value="<%=AccountName%>" ></TD>
		</TR>
		<tr>
			<td> ›’Ì·</td>
			<td><input type='checkbox' name='tafsil' <% If CBool(rs2("HasAppendix")) Then response.write " checked='checked' " %>></td>
		</tr>
		<tr>
			<td>‰Ê⁄ Õ”«»</td>
			<td>
				<select name="accountType">
<%
set rs=Conn.Execute("select * from glAccountTypes")
while not rs.eof
%>
					<option <%if rs2("accountType")=rs("id") then response.write " selected='selected' "%> value="<%=rs("id")%>">
						<%=rs("name")& " (" &rs("name_en")& ")"%>
					</option>
<%
	rs.moveNext
wend
%>
				</select>
			</td>
		</tr>
		<TR>
			<TD colspan=2 align=center>
				<BR><BR>
				<INPUT TYPE="submit" value="À» "class="GenButton">
				<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp?act=account&GroupID=<%=GroupID%>'">
			</TD>
		</TR>
		</FORM>
		</TABLE>
		<%

	
	'======================================== Õ–› Õ”«» 
	'====================================================
	elseif submit="Õ–›" then

	Set RS2 = conn.Execute("SELECT * FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc WHERE (GLDocs.GL = "& OpenGL & ") AND (GLRows.GLAccount = "& ACCID & ")")
		if not RS2.EOF then
			response.write "<BR><BR><BR><BR><CENTER>Œÿ«! œ— «Ì‰ Õ”«»  ⁄œ«œÌ ”‰œ ÊÃÊœ œ«—œ. ·–«  Õ–› ¬‰ „ﬁœÊ— ‰Ì” ."
			response.write "<BR><BR><A HREF='AccountInfo.asp?act=account&GroupID=" & GroupID & "'>»—ê‘ </A></CENTER>"
			response.end
		else
			conn.Execute("DELETE FROM GLAccounts WHERE (ID = "& ACCID & ") AND (GL = "& OpenGL & ")")
			response.redirect "AccountInfo.asp?act=account&GroupID=" & GroupID
		end if

	end if

' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * *   Account Groups
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 


'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------- Edit GL Account Groups submit
'-----------------------------------------------------------------------------------------------------
elseif request("act")="EditAccountGroupSubmit" then
	GroupID = request("GroupID")
	name = request("name")
	SuperGroupID = request("SuperGroupID")

	conn.Execute("UPDATE GLAccountGroups SET Name = N'"& name & "' WHERE (ID = "& GroupID & ") AND (GL = "& OpenGL & ")")
	
	if cint(request("accountType"))<>-1 then 
		conn.Execute("update glAccounts set accountType=" & request("accountType") &" where glGroup=" & groupID & " and gl=" & openGL)
	end if
	response.redirect "AccountInfo.asp?act=groups&SuperGroupID=" & SuperGroupID
'response.end
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------ Add GL Account Groups submit
'-----------------------------------------------------------------------------------------------------
elseif request("act")="NewAccountGroupSubmit" then
	GroupID = request("GroupID")
	name = request("name")
	SuperGroupID = request("SuperGroupID")

	Set RS2 = conn.Execute("SELECT * FROM GLAccountGroups WHERE ID="& GroupID & " AND GL="& OpenGL )
	if not RS2.EOF then
		response.write "<BR><BR><BR><BR><CENTER>Œÿ«! «Ì‰  ﬂœ ê—ÊÂ ÊÃÊœ œ«—œ"
		response.write "<BR><BR><A HREF='AccountInfo.asp?act=editAccountGroupForms&SuperGroupID="& SuperGroupID & "&submit=«›“Êœ‰ ê—ÊÂ'>»—ê‘ </A></CENTER>"
		response.end
	else
		conn.Execute("INSERT INTO GLAccountGroups (GLSuperGroup, Name, GL, ID) VALUES ("& SuperGroupID & ", N'"& name & "',"& OpenGL & ", "& GroupID & ")")
		response.redirect "AccountInfo.asp?act=groups&SuperGroupID=" & SuperGroupID
	end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------- Add / Delete / Edit GL Account Groups Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="editAccountGroupForms" then
	GroupID = request("GroupID")
	submit = request("submit")
	SuperGroupID = request("SuperGroupID")

	Set RS2 = conn.Execute("SELECT * FROM GLAccountSuperGroups WHERE ID="& SuperGroupID & " AND GL="& OpenGL )
	SuperGroupName=RS2("name")
	
	'======================================== «›“Êœ‰ ê—ÊÂ
	'====================================================
	if submit="«›“Êœ‰ ê—ÊÂ" then
		Set RS2 = conn.Execute("SELECT MAX(id) AS MaxID FROM GLAccountGroups WHERE GLSuperGroup="& SuperGroupID & " AND GL="& OpenGL )
		MaxID=RS2("MaxID")
		if isnull(MaxID) then MaxID= SuperGroupID
		%>
		<BR><BR>
		<TABLE align=center>
		<FORM METHOD=POST ACTION="editGL.asp?act=NewAccountGroupSubmit">
		<TR>
			<TD colspan=2 align=center><H3>«›“Êœ‰ ê—ÊÂ</H3><BR>
			œ—  ”—ê—ÊÂ <%=SuperGroupName%> (ﬂœ <%=SuperGroupID%>)<BR><BR>
			œ— <%=OpenGLName%> (”«· „«·Ì <%=FiscalYear%>)
			<INPUT TYPE="hidden" name="SuperGroupID" value="<%=SuperGroupID%>">
			<BR><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD>‘„«—Â ê—ÊÂ</TD>
			<TD><INPUT TYPE="text" NAME="GroupID" onkeypress="return maskNumber(this)" maxlength=5 value="<%=MaxID+100%>"></TD>
		</TR>
		<TR>
			<TD>‰«„ ê—ÊÂ </TD>
			<TD><INPUT TYPE="text" NAME="name"></TD>
		</TR>
		<TR>
			<TD colspan=2 align=center>
				<BR><BR>
				<INPUT TYPE="submit" value="«ÌÃ«œ"class="GenButton">
				<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp?act=groups&SuperGroupID=<%=SuperGroupID%>'">
			</TD>
		</TR>
		</FORM>
		</TABLE>
		<%

	'======================================== ÊÌ—«Ì‘ ê—ÊÂ
	'====================================================
	elseif submit="ÊÌ—«Ì‘" then
		Set RS2 = conn.Execute("SELECT * FROM GLAccountGroups WHERE ID="& GroupID & " AND GL="& OpenGL )
		GroupName=RS2("name")
		%>
		<BR><BR>
		<TABLE align=center>
		<FORM METHOD=POST ACTION="editGL.asp?act=EditAccountGroupSubmit">
		<TR>
			<TD colspan=2 align=center><H3>ÊÌ—«Ì‘ ê—ÊÂ</H3><BR>
			œ— ”—ê—ÊÂ <%=SuperGroupName%> (ﬂœ <%=SuperGroupID%>)<BR><BR>
			œ— <%=OpenGLName%> (”«· „«·Ì <%=FiscalYear%>)
			<INPUT TYPE="hidden" name="SuperGroupID" value="<%=SuperGroupID%>">
			<BR><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD>‘„«—Â ê—ÊÂ</TD>
			<TD><INPUT TYPE="text" NAME="GroupID" onkeypress="return maskNumber(this)" maxlength=5 value="<%=GroupID%>" readonly></TD>
		</TR>
		<TR>
			<TD>‰«„ ê—ÊÂ</TD>
			<TD><INPUT TYPE="text" NAME="name" value="<%=GroupName%>" ></TD>
		</TR>
		<tr>
			<td>ù‰Ê⁄ Õ”«»ùÂ«</td>
			<td>
				<select name="accountType">
					<option value="-1" selected="selected">ÂÌç ﬂœ«„</option>
<%
set rs=Conn.Execute("select * from glAccountTypes")
while not rs.eof
%>
					<option value="<%=rs("id")%>">
						<%=rs("name")& " (" &rs("name_en")& ")"%>
					</option>
<%
	rs.moveNext
wend
%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2">* œ— ’Ê— Ì ﬂÂ ‰Ê⁄ Õ”«» —«  €ÌÌ— œÂÌœ  „«„Ì Õ”«»ùÂ« «“ «Ì‰ ‰Ê⁄ ŒÊ«Â‰œ ‘œ</td>
		</tr>
		<TR>
			<TD colspan=2 align=center>
				<BR><BR>
				<INPUT TYPE="submit" value="À» "class="GenButton">
				<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp?act=groups&SuperGroupID=<%=SuperGroupID%>'">
			</TD>
		</TR>
		</FORM>
		</TABLE>
		<%

	
	'======================================== Õ–› ê—ÊÂ 
	'====================================================
	elseif submit="Õ–›" then

	Set RS2 = conn.Execute("SELECT * FROM GLAccounts WHERE (GL = "& OpenGL & ") AND (GLGroup = "& GroupID & ")")
		if not RS2.EOF then
			response.write "<BR><BR><BR><BR><CENTER>Œÿ«! œ— «Ì‰ ê—ÊÂ  ⁄œ«œÌ Õ”«» ÊÃÊœ œ«—œ. ·–«  Õ–› ¬‰ „ﬁœÊ— ‰Ì” ."
			response.write "<BR><BR><A HREF='AccountInfo.asp?act=groups&SuperGroupID=" & SuperGroupID & "'>»—ê‘ </A></CENTER>"
			response.end
		else
			conn.Execute("DELETE FROM GLAccountGroups WHERE (ID = "& GroupID & ") AND (GL = "& OpenGL & ")")
			response.redirect "AccountInfo.asp?act=groups&SuperGroupID=" & SuperGroupID
		end if

	end if


' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * *   Account Super Groups
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 
' * * * * * * * * 


'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------- Edit GL Account Super Groups submit
'-----------------------------------------------------------------------------------------------------
elseif request("act")="EditAccountSuperGroupSubmit" then
	name = request("name")
	SuperGroupID = request("SuperGroupID")
	SuperGroupType = request("type")

	conn.Execute("UPDATE GLAccountSuperGroups SET type="& SuperGroupType & ", Name = N'"& name & "' WHERE (ID = "& SuperGroupID & ") AND (GL = "& OpenGL & ")")
	if cint(request("accountType"))<>-1 then 
		conn.Execute("update glAccounts set accountType=" & request("accountType") & "where gl=" & openGL & " and glGroup in (select id from GLAccountGroups where gl=" &openGL& " and glSuperGroup=" &SuperGroupID& ")")
	end if
	response.redirect "AccountInfo.asp"

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------ Add GL Account Super Groups submit
'-----------------------------------------------------------------------------------------------------
elseif request("act")="NewAccountSuperGroupSubmit" then
	name = request("name")
	SuperGroupID = request("SuperGroupID")
	SuperGroupType= request("type")

	Set RS2 = conn.Execute("SELECT * FROM GLAccountSuperGroups WHERE ID="& SuperGroupID & " AND GL="& OpenGL )
	if not RS2.EOF then
		response.write "<BR><BR><BR><BR><CENTER>Œÿ«! «Ì‰  ﬂœ ”—ê—ÊÂ ÊÃÊœ œ«—œ"
		response.write "<BR><BR><A HREF='AccountInfo.asp?act=editAccountSuperGroupForms&submit=«›“Êœ‰ ”—ê—ÊÂ'>»—ê‘ </A></CENTER>"
		response.end
	else
		conn.Execute("INSERT INTO GLAccountSuperGroups (type, Name, GL, ID) VALUES ("& SuperGroupType & ", N'"& name & "',"& OpenGL & ", "& SuperGroupID & ")")
		response.redirect "AccountInfo.asp"
	end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------- Add / Delete / Edit GL Account Super Groups Form
'-----------------------------------------------------------------------------------------------------
elseif request("act")="editAccountSuperGroupForms" then
	submit = request("submit")
	SuperGroupID = request("SuperGroupID")

	'====================================== «›“Êœ‰ ”—ê—ÊÂ
	'====================================================
	if submit="«›“Êœ‰ ”—ê—ÊÂ" then
		Set RS2 = conn.Execute("SELECT MAX(id) AS MaxID FROM GLAccountSuperGroups WHERE GL="& OpenGL )
		MaxID=RS2("MaxID")
		if isnull(MaxID) then MaxID= 0
		%>
		<BR><BR>
		<TABLE align=center>
		<FORM METHOD=POST ACTION="editGL.asp?act=NewAccountSuperGroupSubmit">
		<TR>
			<TD colspan=2 align=center><H3>«›“Êœ‰ ”—ê—ÊÂ</H3><BR>
			œ— <%=OpenGLName%> (”«· „«·Ì <%=FiscalYear%>)
			<BR><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD>‘„«—Â ”—ê—ÊÂ</TD>
			<TD><INPUT TYPE="text" NAME="SuperGroupID" onkeypress="return maskNumber(this)" maxlength=5 value="<%=MaxID+1000%>"></TD>
		</TR>
		<TR>
			<TD>‰«„ ”—ê—ÊÂ </TD>
			<TD><INPUT TYPE="text" NAME="name"></TD>
		</TR>
		<TR>
			<TD>‰Ê⁄ ”—ê—ÊÂ</TD>
			<TD>
			<SELECT NAME="type">
				<option value=1>œ«—«ÌÌ (”„  —«”   —«“)</option>
				<option value=2>»œÂÌ (”„  çÅ  —«“)</option>
				<option value=3>”—„«ÌÂ (”„  çÅ  —«“)</option>
				<option value=4>Â“Ì‰Â / Œ—Ìœ</option>
				<option value=5>œ—¬„œ / ›—Ê‘ </option>
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
		</FORM>
		</TABLE>
		<%

	'====================================== ÊÌ—«Ì‘ ”—ê—ÊÂ
	'====================================================
	elseif submit="ÊÌ—«Ì‘" then
		Set RS2 = conn.Execute("SELECT * FROM GLAccountSuperGroups WHERE ID="& SuperGroupID & " AND GL="& OpenGL )
		SuperGroupName=RS2("name")
		SuperGroupType=RS2("type")
		%>
		<BR><BR>
		<TABLE align=center>
		<FORM METHOD=POST ACTION="editGL.asp?act=EditAccountSuperGroupSubmit">
		<TR>
			<TD colspan=2 align=center><H3>ÊÌ—«Ì‘ ”—ê—ÊÂ</H3><BR>
			œ— <%=OpenGLName%> (”«· „«·Ì <%=FiscalYear%>)
			<BR><BR><BR>
			</TD>
		</TR>
		<TR>
			<TD>‘„«—Â ”—ê—ÊÂ</TD>
			<TD><INPUT TYPE="text" NAME="SuperGroupID" onkeypress="return maskNumber(this)" maxlength=5 value="<%=SuperGroupID%>" readonly></TD>
		</TR>
		<TR>
			<TD>‰«„ ”—ê—ÊÂ</TD>
			<TD><INPUT TYPE="text" NAME="name" value="<%=SuperGroupName%>" ></TD>
		</TR>
		<TR>
			<TD>‰Ê⁄ ”—ê—ÊÂ</TD>
			<TD>
			<SELECT NAME="type">
				<option <% if SuperGroupType=1 then %>selected<% end if %> value=1>œ«—«ÌÌ (”„  —«”   —«“)</option>
				<option <% if SuperGroupType=2 then %>selected<% end if %> value=2>»œÂÌ (”„  çÅ  —«“)</option>
				<option <% if SuperGroupType=3 then %>selected<% end if %> value=3>”—„«ÌÂ (”„  çÅ  —«“)</option>
				<option <% if SuperGroupType=4 then %>selected<% end if %> value=4>Â“Ì‰Â / Œ—Ìœ</option>
				<option <% if SuperGroupType=5 then %>selected<% end if %> value=5>œ—¬„œ / ›—Ê‘ </option>
			</SELECT>
			</TD>
		</TR>
		<tr>
			<td>ù‰Ê⁄ Õ”«»ùÂ«</td>
			<td>
				<select name="accountType">
					<option value="-1" selected="selected">ÂÌç ﬂœ«„</option>
<%
set rs=Conn.Execute("select * from glAccountTypes")
while not rs.eof
%>
					<option value="<%=rs("id")%>">
						<%=rs("name")& " (" &rs("name_en")& ")"%>
					</option>
<%
	rs.moveNext
wend
%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2">* œ— ’Ê— Ì ﬂÂ ‰Ê⁄ Õ”«» —«  €ÌÌ— œÂÌœ  „«„Ì Õ”«»ùÂ« «“ «Ì‰ ‰Ê⁄ ŒÊ«Â‰œ ‘œ</td>
		</tr>
		<TR>
			<TD colspan=2 align=center>
				<BR><BR>
				<INPUT TYPE="submit" value="À» "class="GenButton">
				<INPUT TYPE="button" value="«‰’—«›"class="GenButton" onclick="window.location='AccountInfo.asp'">
			</TD>
		</TR>
		</FORM>
		</TABLE>
		<%

	
	'======================================== Õ–› ”—ê—ÊÂ 
	'====================================================
	elseif submit="Õ–›" then

	Set RS2 = conn.Execute("SELECT * FROM GLAccountGroups WHERE (GL = "& OpenGL & ") AND (GLSuperGroup = "& SuperGroupID & ")")
		if not RS2.EOF then
			response.write "<BR><BR><BR><BR><CENTER>Œÿ«! œ— «Ì‰ ”—ê—ÊÂ  ⁄œ«œÌ ê—ÊÂ ÊÃÊœ œ«—œ. ·–«  Õ–› ¬‰ „ﬁœÊ— ‰Ì” ."
			response.write "<BR><BR><A HREF='AccountInfo.asp'>»—ê‘ </A></CENTER>"
			response.end
		else
			conn.Execute("DELETE FROM GLAccountSuperGroups WHERE (ID = "& SuperGroupID & ") AND (GL = "& OpenGL & ")")
			response.redirect "AccountInfo.asp"
		end if

	end if


end if


%>
