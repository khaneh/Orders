<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle= "гѕн—н  ѕЁ — яб"
SubmenuItem=9

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
dateStr = "«“  «—нќ <span dir=rtl>" & fiscalYear & "/1/1</span>  « <span dir=rtl>" & shamsiToday() & "</span>"

function Separate2(inputTxt)
myMinus=""
input=inputTxt
if left(input,1)="-" then
	myMinus="-"
	input=right(input,len(input)-1)
end if
if len(input) > 3 then
	tmpr=right(input ,3)
	tmpl=left(input , len(input) - 3 )
	result = tmpr
	while len(tmpl) > 3
		tmpr=right(tmpl,3)
		result = tmpr & "," & result 
		tmpl=left(tmpl , len(tmpl) - 3 )
	wend
	if len(tmpl) > 0 then
		result = tmpl & "," & result
	end if 
else
	result = input
end if 
	Separate2=myMinus & result
end function


'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- List GL Super Groups
'-----------------------------------------------------------------------------------------------------
if request("act")="" then
%><BR><BR>
<CENTER>
<FORM METHOD=POST ACTION="AccountInfo.asp">
<INPUT TYPE="button" onclick="window.location='manageGL.asp?act=showLog'" class="GenButton"  value="<%=OpenGLName%>" readonly style="width:276px; background-color: white ">
<!--INPUT TYPE="button" onclick="window.location='balance.asp'" class="GenButton" value=" —«“" style="width:90px; ">
<INPUT TYPE="button" onclick="window.location='sood.asp'" class="GenButton" value="”жѕ ж “н«д" style="width:90px; "--><br><br>
<INPUT TYPE="button" onclick="window.location='manageGL.asp?act=select'" class="GenButton" value=" џнн—" style="width:90px; background-color: #FFFFBB ">
<INPUT TYPE="button" onclick="window.location='manageGL.asp?act=new'" class="GenButton" value="ћѕнѕ" style="width:90px; background-color: #FFFFBB ">
<INPUT TYPE="button" onclick="window.location='manageGL.asp?act=close'" class="GenButton" value="»” д" style="width:90px; background-color: #FFFFBB ">
</FORM>
</CENTER>
<%
Conn.CommandTimeout = 300
set RSS=Conn.Execute ("SELECT GLAccountSuperGroups.Name, ISNULL(DERIVEDTBL.totalCredit, 0) AS totalCredit, ISNULL(DERIVEDTBL.totalDebit, 0) AS totalDebit, totalDebit-totalCredit as totalRemained, GLAccountSuperGroups.ID, GLAccountSuperGroups.Type, ISNULL(DERIVEDTBL.GroupsCount,0) as GroupsCount, ISNULL(DERIVEDTBL.accountsCount,0) as accountsCount FROM (SELECT SUM(DERIVEDTBL.totalCredit) AS totalCredit, SUM(DERIVEDTBL.totalDebit) AS totalDebit, GLAccountGroups.GLSuperGroup, SUM(accountsCount) AS accountsCount, COUNT(GLAccountGroups.ID) AS GroupsCount FROM (SELECT SUM(DERIVEDTBL.totalCredit) AS totalCredit, SUM(DERIVEDTBL.totalDebit) AS totalDebit, GLAccounts.GLGroup, ISNULL(COUNT(GLAccounts.ID),0) AS accountsCount FROM (SELECT SUM(IsCredit * Amount) AS totalCredit, SUM((convert(tinyint,IsCredit)-1) * -1 * Amount) AS totalDebit, GLAccount FROM  EffectiveGlRows WHERE (GL = "& OpenGL&") GROUP BY GLAccount) DERIVEDTBL RIGHT OUTER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL&") GROUP BY GLAccounts.GLGroup) DERIVEDTBL  RIGHT OUTER JOIN GLAccountGroups ON DERIVEDTBL.GLGroup = GLAccountGroups.ID WHERE (GLAccountGroups.GL = "& OpenGL&") GROUP BY GLAccountGroups.GLSuperGroup) DERIVEDTBL RIGHT OUTER JOIN GLAccountSuperGroups ON DERIVEDTBL.GLSuperGroup = GLAccountSuperGroups.ID WHERE (GLAccountSuperGroups.GL = "& OpenGL&") ")	
%>

<TABLE dir=rtl align=center width=600>
<TR >
	<TD colspan=5>
		<%=OpenGLName%>  &nbsp;<%=dateStr%><BR><hR>
	</TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD><!A HREF="default.asp?s=1"><SMALL>яѕ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>д«г ”—Р—же</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=3"><SMALL> Џѕ«ѕ Р—же е«</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL> Џѕ«ѕ Ќ”«»е«</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL>ћгЏ »” «дя«—  (—н«б)</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL>ћгЏ »ѕея«—  (—н«б)</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL>г«дѕе (—н«б)</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=5"><SMALL>Џгбн« </SMALL></A></TD>
</TR>
<%
tmpCounter=0
Do while not RSS.eof
	tmpCounter = tmpCounter + 1
	if tmpCounter mod 2 = 1 then
		tmpColor="#FFFFFF"
		tmpColor2="#FFFFBB"
	Else
		tmpColor="#DDDDDD"
		tmpColor2="#EEEEBB"
	End if 

%>
<TR bgcolor="<%=tmpColor%>" height=30>
	<FORM METHOD=POST ACTION="editGL.asp?act=editAccountSuperGroupForms&SuperGroupID=<%=RSS("ID")%>">
	<TD><A HREF="AccountInfo.asp?act=groups&SuperGroupID=<%=RSS("ID")%>"><%=RSS("ID")%></A></TD>
	<TD><A HREF="AccountInfo.asp?act=groups&SuperGroupID=<%=RSS("ID")%>"><%=RSS("Name")%></A></TD>
	<TD><%=RSS("GroupsCount")%></TD>
	<TD><%=RSS("accountsCount")%></TD>
	<TD  dir=ltr align=right><span dir=ltr><%=Separate2(RSS("totalCredit"))%></span></TD>
	<TD dir=ltr align=right><span dir=ltr><%=Separate2(RSS("totalDebit"))%></span></TD>
	<TD dir=ltr align=right><span dir=ltr><%=Separate2(RSS("totalRemained"))%></span></TD>
	<TD  dir=ltr align=center><INPUT style="font-size:7pt" TYPE="submit" value="Ќ–Ё" NAME="submit" onclick="return confirm('¬н« Ќёнё « гн ќж«енѕ «нд ”—Р—же —« Б«я яднѕњ')">
	<INPUT TYPE="submit" value="жн—«н‘" NAME="submit" style="font-size:7pt">
	</TD>
	</form>
</TR>
	  
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
<FORM METHOD=POST ACTION="editGL.asp?act=editAccountSuperGroupForms">
<CENTER><INPUT TYPE="submit" value="«Ё“жѕд ”—Р—же" NAME="submit"></CENTER><br>
</FORM>
<%

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------ List GL Groups under a super group
'-----------------------------------------------------------------------------------------------------
elseif request("act")="groups" then
SuperGroupID = request("SuperGroupID")

set RSS=Conn.Execute ("SELECT SUM(ISNULL(DERIVEDTBL.totalCredit, 0)) AS totalCredit, SUM(ISNULL(DERIVEDTBL.totalDebit, 0)) AS totalDebit, GLAccountGroups.GLSuperGroup, SUM(ISNULL(DERIVEDTBL.accountsCount, 0)) AS accountsCount, GLAccountGroups.Name, GLAccountGroups.ID FROM (SELECT SUM(DERIVEDTBL.totalCredit) AS totalCredit, SUM(DERIVEDTBL.totalDebit) AS totalDebit, GLAccounts.GLGroup, COUNT(GLAccounts.ID) AS accountsCount FROM (SELECT SUM(IsCredit * Amount) AS totalCredit, SUM((convert(tinyint,IsCredit)-1) * -1 * Amount) AS totalDebit, GLAccount FROM EffectiveGLRows Where (GL = "& OpenGL & ")  GROUP BY GLAccount) DERIVEDTBL RIGHT OUTER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") GROUP BY GLAccounts.GLGroup) DERIVEDTBL RIGHT OUTER JOIN GLAccountGroups ON DERIVEDTBL.GLGroup = GLAccountGroups.ID WHERE (GLAccountGroups.GL = "& OpenGL & ") AND (GLAccountGroups.GLSuperGroup = "& SuperGroupID & ") GROUP BY GLAccountGroups.GLSuperGroup, GLAccountGroups.Name, GLAccountGroups.ID ORDER BY GLAccountGroups.ID")	

%><BR><BR>
<TABLE dir=rtl align=center width=600>
<TR >
	<TD colspan=4>
		<%
		set RSS2=Conn.Execute ("SELECT *, GLs.ID AS GLID ,GLs.Name AS GLname FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL WHERE (GLAccountSuperGroups.GL = "& OpenGL & ") AND (GLAccountSuperGroups.ID = "& SuperGroupID & ")")
		%><A HREF="AccountInfo.asp?OpenGL=<%=RSS2("GLID")%>"><%=RSS2("GLname")%></A> > <%=RSS2("name")%>
		<BR><hR>
	</TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD><!A HREF="default.asp?s=1"><SMALL>яѕ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>д«г  Р—же</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL> Џѕ«ѕ Ќ”«»е«</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL>ћгЏ »” «дя«—  (—н«б)</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL>ћгЏ »ѕея«—  (—н«б)</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=5"><SMALL>Џгбн« </SMALL></A></TD>
</TR>
<%
tmpCounter=0
Do while not RSS.eof
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
	<FORM METHOD=POST ACTION="editGL.asp?act=editAccountGroupForms&SuperGroupID=<%=SuperGroupID%>&GroupID=<%=RSS("id")%>">
	<TD><A HREF="AccountInfo.asp?act=account&GroupID=<%=RSS("id")%>"><%=RSS("id")%></A></TD>
	<TD><A HREF="AccountInfo.asp?act=account&GroupID=<%=RSS("id")%>"><%=RSS("Name")%></A></TD>
	<TD><%=RSS("accountsCount")%></TD>
	<TD><span dir=ltr><%=RSS("totalCredit")%></span></TD>
	<TD><span dir=ltr><%=RSS("totalDebit")%></span></TD>
	<TD align=center><INPUT TYPE="submit" value="Ќ–Ё" NAME="submit" onclick="return confirm('¬н« Ќёнё « гн ќж«енѕ «нд Р—же —« Б«я яднѕњ')">
	<INPUT TYPE="submit" value="жн—«н‘" NAME="submit">
	</TD>
	</form>
</TR>
	  
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
<FORM METHOD=POST ACTION="editGL.asp?act=editAccountGroupForms&SuperGroupID=<%=SuperGroupID%>">
<CENTER><INPUT TYPE="submit" value="«Ё“жѕд Р—же" NAME="submit"></CENTER><br>
</FORM>
<%

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- List GL Accounts under a group
'-----------------------------------------------------------------------------------------------------
elseif request("act")="account" then
GroupID = request("GroupID")

set RSS=Conn.Execute ("SELECT GLs.Name AS GLName, GLs.ID AS GLID, GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName, GLAccountGroups.ID AS GroupID, GLAccountGroups.Name AS GroupName FROM GLAccountGroups INNER JOIN GLAccountSuperGroups ON GLAccountGroups.GLSuperGroup = GLAccountSuperGroups.ID INNER JOIN GLs ON GLAccountSuperGroups.GL = GLs.ID WHERE (GLAccountSuperGroups.GL = "& OpenGL & ") AND (GLAccountGroups.GL = "& OpenGL & ") AND (GLAccountGroups.ID = "& GroupID & ")")
%><BR><BR>
<TABLE dir=rtl align=center width=600>
<TR >
	<TD colspan=5>
		<%
		set RSS2=Conn.Execute ("SELECT GLAccounts.Name, DERIVEDTBL.totalDebit AS totalDebit, DERIVEDTBL.totalCredit AS totalCredit, GLAccounts.ID, GLAccounts.HasAppendix FROM (SELECT SUM(CONVERT(tinyint, IsCredit) * Amount) AS totalCredit, SUM((CONVERT(tinyint, IsCredit) - 1) * (- 1) * Amount) AS totalDebit, GLAccount FROM EffectiveGLRows where (GL = "& OpenGL & ")   GROUP BY GLAccount) DERIVEDTBL RIGHT OUTER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.GLGroup = "& GroupID & ") order by GLAccounts.ID")
		
		%><A HREF="AccountInfo.asp?OpenGL=<%=RSS("GLID")%>"><%=RSS("GLname")%></A> > <A HREF="AccountInfo.asp?act=groups&SuperGroupID=<%=RSS("SuperGroupID")%>"><%=RSS("SuperGroupName")%></a>  > <%=RSS("GroupName")%>
		<BR><hR>
	</TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD><!A HREF="default.asp?s=1"><SMALL>яѕ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>Џдж«д Ќ”«»</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ћгЏ »” «дя«—</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ћгЏ »ѕея«—</SMALL></A></TD>
	<td><small> Ё’нб</small></td>
	<TD><!A HREF="default.asp?s=5"><SMALL>Џгбн« </SMALL></A></TD>
</TR>
<%
tmpCounter=0
Do while not RSS2.eof
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
	<FORM METHOD=POST ACTION="editGL.asp?act=editAccountForms&GroupID=<%=GroupID%>&ACCID=<%=RSS2("id")%>">
	<TD><A HREF="AccountInfo.asp?act=accountRows&accountID=<%=RSS2("id")%>"><%=RSS2("id")%></A></TD>
	<TD><A HREF="AccountInfo.asp?act=accountRows&accountID=<%=RSS2("id")%>"><%=RSS2("Name")%></A></TD>
	<TD><span dir=ltr><%=RSS2("totalCredit")%></span></TD>
	<TD><span dir=ltr><%=RSS2("totalDebit")%></span></TD>
	<td><span dir=rtl><% If Cbool(RSS2("HasAppendix")) then response.write "ѕ«—ѕ" Else response.write "дѕ«—ѕ" End If %></span></td>
	<TD align=center><INPUT TYPE="submit"  value="Ќ–Ё" NAME="submit" onclick="return confirm('¬н« Ќёнё « гн ќж«енѕ «нд Ќ”«» —« Б«я яднѕњ')">
	<INPUT TYPE="submit"  value="жн—«н‘" NAME="submit">
	</TD>
	</FORM>
</TR>
	  
<% 
RSS2.moveNext
Loop
%>
</TABLE><br>
<FORM METHOD=POST ACTION="editGL.asp?act=editAccountForms&GroupID=<%=GroupID%>">
<CENTER><INPUT TYPE="submit" value="«Ё“жѕд Ќ”«»" NAME="submit"></CENTER><br>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var submits=document.getElementsByName("submit");
	submits[submits.length -1].focus();
//-->
</SCRIPT>
<%

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------- List Rows of a GL Accounts 
'-----------------------------------------------------------------------------------------------------
elseif request("act")="accountRows" then
account = request("accountID")

DateFrom = request.form("DateFrom")
if DateFrom = "" then DateFrom = shamsiToday()

DateTo = request.form("DateTo")
if DateTo = "" then DateTo = shamsiToday()

set RSS=Conn.Execute ("SELECT *, GLDocs.id as GLDoc, GLRows.id as GLRowsID FROM GLRows inner join GLDocs on GLDocs.id=GLRows.GLDoc where ((GLDocs.IsTemporary=1 or GLDocs.IsChecked=1 or  GLDocs.IsFinalized=1) and GLDocs.deleted=0 and GLDocs.IsRemoved=0) and  ((GLDocs.GL = "& OpenGL & ") AND (GLRows.GLaccount = "& account & ") and (GLDocs.GLDocDate <= N'"& DateTo & "') and (GLDocs.GLDocDate >= N'"& DateFrom & "'))")	
%><BR><BR>
<TABLE dir=rtl align=center width=640 cellspacing=1 cellpadding=3>
<TR >
	<TD colspan=7>
		<%
		set RSS2=Conn.Execute ("SELECT *, GLAccountGroups.ID AS GroupID, GLAccountGroups.Name AS GroupName, GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName, GLs.ID AS GLID, GLs.Name AS GLname, GLAccounts.Name AS name FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL INNER JOIN GLAccountGroups ON GLs.ID = GLAccountGroups.GL AND GLAccountSuperGroups.ID = GLAccountGroups.GLSuperGroup INNER JOIN GLAccounts ON GLs.ID = GLAccounts.GL AND GLAccountGroups.ID = GLAccounts.GLGroup WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.ID = "& account & ")")
		
		%><A HREF="AccountInfo.asp?OpenGL=<%=RSS2("GLID")%>"><%=RSS2("GLname")%></A> > <A HREF="AccountInfo.asp?act=groups&SuperGroupID=<%=RSS2("SuperGroupID")%>"><%=RSS2("SuperGroupName")%></a>  > <A HREF="AccountInfo.asp?act=account&GroupID=<%=RSS2("GroupID")%>"><%=RSS2("GroupName")%></a>  >  <%=RSS2("Name")%>
		<BR><hR>
	</TD>
</TR>
<TR >
	<TD colspan=7 align=center>
	<FORM METHOD=POST ACTION="AccountInfo.asp?act=accountRows&accountID=<%=account%>">
	«“  «—нќ <INPUT  dir="LTR"  TYPE="text" NAME="DateFrom" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=DateFrom%>">
	 «  «—нќ <INPUT  dir="LTR"  TYPE="text" NAME="DateTo" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=DateTo%>">
	<INPUT TYPE="submit" NAME="submit" value="ћ” ћж">
	</FORM>
	</TD>
</TR>
<% if request("submit")="ћ” ћж" then %>
<TR bgcolor="eeeeee" >
	<TD align=center><!A HREF="default.asp?s=2"><SMALL>”дѕ</SMALL></A></TD>
	<TD align=center><!A HREF="default.asp?s=2"><SMALL> «—нќ ”дѕ</SMALL></A></TD>
	<TD align=center width=100><!A HREF="default.asp?s=2"><SMALL>‘—Ќ</SMALL></A></TD>
	<TD align=center><!A HREF="default.asp?s=5"><SMALL>»ѕея«—</SMALL></A></TD>
	<TD align=center><!A HREF="default.asp?s=5"><SMALL>»” «дя«—</SMALL></A></TD>
	<TD align=center><!A HREF="default.asp?s=5"><SMALL>-</SMALL></A></TD>
	<TD align=center><!A HREF="default.asp?s=5"><SMALL>г«дѕе</SMALL></A></TD>
</TR>
<%
tmpCounter=0
remainedAmount = 0
Do while not RSS.eof
	tmpCounter = tmpCounter + 1
	if tmpCounter mod 2 = 1 then
		tmpColor="#FFFFFF"
		tmpColor2="#FFFFBB"
	Else
		tmpColor="#DDDDDD"
		tmpColor2="#EEEEBB"
	End if 
	
	debit = 0
	credit = 0
	if RSS("IsCredit") then
		credit = RSS("Amount")
	else
		debit = RSS("Amount")
	end if
	remainedAmount = remainedAmount - cdbl(credit) + cdbl(debit)

%>
<TR bgcolor="<%=tmpColor%>" >
	<TD><A HREF="GLMemoDocShow.asp?id=<%=RSS("GLDoc")%>"><%=RSS("GLDocID")%></A></TD>
	<TD  align=center dir=ltr><%=RSS("GLDocDate")%></TD>
	<TD  width=100><!A HREF="AccountInfo.asp?act=accountRows&accountID=<%=RSS("id")%>"><INPUT TYPE="text" value="<%=RSS("Description")%>" style="width=200pt; border:solid 0pt; font-size:8pt; background-color:transparent"></A></TD>
	
	
	<TD dir=ltr align=right><% if debit<>"0" then %> <%=Separate2(debit)%><% end if %></TD>
	<TD dir=ltr align=right><% if credit<>"0" then %> <%=Separate2(credit)%><% end if %></TD>
	<TD dir=ltr align=center><% if remainedAmount > 0 then %>ѕ<% else %>”<% end if %></TD>
	<TD dir=ltr align=right><%=Separate2(remainedAmount)%></TD>
</TR>
	  
<% 
RSS.moveNext
Loop
%>
<% end if %>
</TABLE><br>
<%
end if


%>

<!--#include file="tah.asp" -->
