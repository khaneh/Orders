<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= "สัวา"
SubmenuItem=4
if not Auth(8 , 4) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->

<%


'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- List GL Super Groups
'-----------------------------------------------------------------------------------------------------
if request("act")="" then
%><BR><BR>
<TABLE dir=rtl align=center width=600 border=0>
<TR>
	<TD   align=right><A HREF="AccountInfo.asp">ฯÝสั ฿แ </A>> สัวา ยาใวํิํ
	</TD>
</TR>
<TR  bgcolor="eeeeee" >
	<TD colspan=2 align=center>สัวา ไวใๅ ยาใวํิํ
	</TD>
</TR>
<TR>
	<TD bgcolor="#FFFFFF" width=50%><!----Right---->
	<TABLE width=70% align=center>
	<%
	mySQL="SELECT GLAccountSuperGroups.Name, ISNULL(DERIVEDTBL.total, 0) AS total, GLAccountSuperGroups.ID, GLAccountSuperGroups.Type FROM (SELECT SUM(DERIVEDTBL.Total) AS total, GLAccountGroups.GLSuperGroup FROM (SELECT SUM(DERIVEDTBL.total) AS Total, GLAccounts.GLGroup FROM (SELECT SUM(-1*(CONVERT(tinyint, IsCredit) - .5) * 2 * Amount) AS total, GLAccount FROM GLDocs inner join GLRows on GLDocs.id=GLRows.GLDoc where ((GLDocs.IsTemporary=1 or GLDocs.IsChecked=1 or  GLDocs.IsFinalized=1) and GLDocs.deleted=0)  GROUP BY GLAccount, GL HAVING (GL = "& OpenGL&")) DERIVEDTBL INNER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL&") GROUP BY GLAccounts.GLGroup) DERIVEDTBL INNER JOIN GLAccountGroups ON DERIVEDTBL.GLGroup = GLAccountGroups.ID WHERE (GLAccountGroups.GL = "& OpenGL&") GROUP BY GLAccountGroups.GLSuperGroup) DERIVEDTBL RIGHT OUTER JOIN GLAccountSuperGroups ON DERIVEDTBL.GLSuperGroup = GLAccountSuperGroups.ID WHERE (GLAccountSuperGroups.GL = "& OpenGL&" and GLAccountSuperGroups.Type=1) "
'	mySQL="SELECT GLAccountSuperGroups.Name,SUM((CONVERT(tinyint, GLRows.IsCredit) - .5) * 2 * GLRows.Amount) AS total, GLAccountSuperGroups.ID FROM GLAccountGroups INNER JOIN GLAccounts ON GLAccountGroups.ID = GLAccounts.GLGroup INNER JOIN GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc ON GLAccounts.ID = GLRows.GLAccount RIGHT OUTER JOIN GLAccountSuperGroups ON GLAccountGroups.GLSuperGroup = GLAccountSuperGroups.ID WHERE (GLDocs.IsTemporary = 1) AND (GLDocs.deleted = 0) AND (GLDocs.IsRemoved = 0) OR (GLDocs.deleted = 0) AND (GLDocs.IsChecked = 1) AND (GLDocs.IsRemoved = 0) OR (GLDocs.deleted = 0) AND (GLDocs.IsFinalized = 1) AND (GLDocs.IsRemoved = 0) GROUP BY GLDocs.GL, GLAccountSuperGroups.ID,GLAccountSuperGroups.Name HAVING (GLDocs.GL = 82)"

'	response.write "<br>" & mySQL
'	response.end
	set RSS=Conn.Execute (mySQL)

	Do while not RSS.eof
		%>
		<TR>
		<TD><A HREF="balance.asp?act=groups&SuperGroupID=<%=RSS("ID")%>"><%=RSS("name")%></a></TD>
		<TD><span dir=ltr><%=RSS("total")%></span><br></TD>
		</TR>
		<%
		totalRight = totalRight + cdbl(RSS("total"))
		RSS.moveNext
	Loop
	%>	
	</TABLE>
	</TD>

	<TD bgcolor="#DDDDDD" width=50%><!----left---->
	<TABLE width=70% align=center>
	<%
	mySQL="SELECT GLAccountSuperGroups.Name, ISNULL(DERIVEDTBL.total, 0) AS total, GLAccountSuperGroups.ID, GLAccountSuperGroups.Type FROM (SELECT SUM(DERIVEDTBL.Total) AS total, GLAccountGroups.GLSuperGroup FROM (SELECT SUM(DERIVEDTBL.total) AS Total, GLAccounts.GLGroup FROM (SELECT SUM((CONVERT(tinyint, IsCredit) - .5) * 2 * Amount) AS total, GLAccount FROM GLDocs inner join GLRows on GLDocs.id=GLRows.GLDoc where ((GLDocs.IsTemporary=1 or GLDocs.IsChecked=1 or  GLDocs.IsFinalized=1) and GLDocs.deleted=0)  GROUP BY GLAccount, GL HAVING (GL = "& OpenGL&")) DERIVEDTBL INNER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL&") GROUP BY GLAccounts.GLGroup) DERIVEDTBL INNER JOIN GLAccountGroups ON DERIVEDTBL.GLGroup = GLAccountGroups.ID WHERE (GLAccountGroups.GL = "& OpenGL&") GROUP BY GLAccountGroups.GLSuperGroup) DERIVEDTBL RIGHT OUTER JOIN GLAccountSuperGroups ON DERIVEDTBL.GLSuperGroup = GLAccountSuperGroups.ID WHERE (GLAccountSuperGroups.GL = "& OpenGL&" and GLAccountSuperGroups.Type=2) "
'	response.write "<br>" & mySQL
'	response.end

	set RSS=Conn.Execute (mySQL)

	Do while not RSS.eof
		%>
		<TR>
		<TD><A HREF="balance.asp?act=groups&SuperGroupID=<%=RSS("ID")%>"><%=RSS("name")%></a></TD>
		<TD><span dir=ltr><%=RSS("total")%></span><br></TD>
		</TR>
		<%
		totalLeft = totalLeft + cdbl(RSS("total"))
		RSS.moveNext
	Loop
	%>	
		<TR>
		<TD><!A HREF="balance.asp?act=groups&SuperGroupID=">ำัใวํๅ</a></TD>
		<TD><span dir=ltr>???</span><br></TD>
		</TR>
		<%
		totalLeft = totalLeft + 0
		%>
	</TABLE>
	</TD>
</TR>
<TR  bgcolor="eeeeee" >
	<TD align=center><span dir=ltr><%=totalRight%></span>
	</TD>
	<TD align=center><span dir=ltr><%=totalLeft%></span>
	</TD>
</TR>
</TABLE>

<BR><BR><BR>
<% if not totalRight=totalLeft then %>
<FONT COLOR="RED"><H3><CENTER>อำวศ ๅว สัวา ไํำส</CENTER></H3></FONT>
<% else %>
<H3><CENTER>อำวศ ๅว สัวา วำส </CENTER></H3>
<% end if  %>
<% 
response.end


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------ List GL Groups under a super group
'-----------------------------------------------------------------------------------------------------
elseif request("act")="groups" then
SuperGroupID = request("SuperGroupID")

set RSS=Conn.Execute ("SELECT SUM(ISNULL(DERIVEDTBL.Total, 0)) AS total, GLAccountGroups.GLSuperGroup, SUM(DERIVEDTBL.accountsCount) AS accountsCount, GLAccountGroups.Name, GLAccountGroups.ID FROM (SELECT SUM(DERIVEDTBL.total) AS Total, GLAccounts.GLGroup, COUNT(GLAccounts.ID) AS accountsCount FROM (SELECT SUM((CONVERT(tinyint, IsCredit) - .5) * 2 * Amount) AS total, GLAccount FROM GLDocs inner join GLRows on GLDocs.id=GLRows.GLDoc where ((GLDocs.IsTemporary=1 or GLDocs.IsChecked=1 or  GLDocs.IsFinalized=1) and GLDocs.deleted=0)  GROUP BY GLAccount, GL HAVING (GL = "& OpenGL & ")) DERIVEDTBL RIGHT OUTER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") GROUP BY GLAccounts.GLGroup) DERIVEDTBL INNER JOIN GLAccountGroups ON DERIVEDTBL.GLGroup = GLAccountGroups.ID WHERE (GLAccountGroups.GL = "& OpenGL & ") AND (GLAccountGroups.GLSuperGroup = "& SuperGroupID & ") GROUP BY GLAccountGroups.GLSuperGroup, GLAccountGroups.Name, GLAccountGroups.ID ")	

%><BR><BR>
<TABLE dir=rtl align=center width=600>
<TR >
	<TD colspan=4>
		<%
		set RSS2=Conn.Execute ("SELECT *, GLs.ID AS GLID ,GLs.Name AS GLname FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL WHERE (GLAccountSuperGroups.GL = "& OpenGL & ") AND (GLAccountSuperGroups.ID = "& SuperGroupID & ")")
		%><A HREF="AccountInfo.asp">ฯÝสั ฿แ </A>> <A HREF="balance.asp?OpenGL=<%=RSS2("GLID")%>">สัวา ยาใวํิํ</A> > <%=RSS2("name")%>
		<BR><hR>
	</TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD><!A HREF="default.asp?s=1"><SMALL>฿ฯ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ไวใ  ัๆๅ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL>สฺฯวฯ อำวศๅวํ าํัํไ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=4"><SMALL>ฬใฺ อำวศๅว (ัํวแ)</SMALL></A></TD>
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
	<TD><A HREF="balance.asp?act=account&GroupID=<%=RSS("id")%>"><%=RSS("id")%></A></TD>
	<TD><A HREF="balance.asp?act=account&GroupID=<%=RSS("id")%>"><%=RSS("Name")%></A></TD>
	<TD><%=RSS("accountsCount")%></TD>
	<TD><span dir=ltr><%=RSS("total")%></span></TD>
</TR>
	  
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
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
		set RSS2=Conn.Execute ("SELECT GLAccounts.Name, DERIVEDTBL.totalDebit AS totalDebit, DERIVEDTBL.totalCredit AS totalCredit, GLAccounts.ID FROM (SELECT SUM(CONVERT(tinyint, IsCredit) * Amount) AS totalCredit, SUM((CONVERT(tinyint, IsCredit) - 1) * (- 1) * Amount) AS totalDebit, GLAccount FROM GLDocs inner join GLRows on GLDocs.id=GLRows.GLDoc where ((GLDocs.IsTemporary=1 or GLDocs.IsChecked=1 or  GLDocs.IsFinalized=1) and GLDocs.deleted=0)  GROUP BY GLAccount, GL HAVING (GL = "& OpenGL & ")) DERIVEDTBL RIGHT OUTER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.GLGroup = "& GroupID & ")")
		
		%><A HREF="AccountInfo.asp">ฯÝสั ฿แ </A>> <A HREF="balance.asp?OpenGL=<%=RSS("GLID")%>">สัวา ยาใวํิํ</A> > <A HREF="balance.asp?act=groups&SuperGroupID=<%=RSS("SuperGroupID")%>"><%=RSS("SuperGroupName")%></a>  > <%=RSS("GroupName")%>
		<BR><hR>
	</TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD><!A HREF="default.asp?s=1"><SMALL>฿ฯ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ฺไๆวไ อำวศ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ฬใฺ ศฯๅ฿วั</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ฬใฺ ศำสวไ฿วั</SMALL></A></TD>
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
	<TD><A HREF="balance.asp?act=accountRows&accountID=<%=RSS2("id")%>"><%=RSS2("id")%></A></TD>
	<TD><A HREF="balance.asp?act=accountRows&accountID=<%=RSS2("id")%>"><%=RSS2("Name")%></A></TD>
	<TD><span dir=ltr><%=RSS2("totalDebit")%></span></TD>
	<TD><span dir=ltr><%=RSS2("totalCredit")%></span></TD>
</TR>
	  
<% 
RSS2.moveNext
Loop
%>
</TABLE><br>
<%

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------- List Rows of a GL Accounts 
'-----------------------------------------------------------------------------------------------------
elseif request("act")="accountRows" then
account = request("accountID")

set RSS=Conn.Execute ("SELECT * FROM GLDocs inner join GLRows on GLDocs.id=GLRows.GLDoc where ((GLDocs.IsTemporary=1 or GLDocs.IsChecked=1 or  GLDocs.IsFinalized=1) and GLDocs.deleted=0) AND (GL = "& OpenGL & ") AND (GLaccount = "& account & ")")	
%><BR><BR>
<TABLE dir=rtl align=center width=600>
<TR >
	<TD colspan=6>
		<%
		set RSS2=Conn.Execute ("SELECT *, GLAccountGroups.ID AS GroupID, GLAccountGroups.Name AS GroupName, GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName, GLs.ID AS GLID, GLs.Name AS GLname, GLAccounts.Name AS name FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL INNER JOIN GLAccountGroups ON GLs.ID = GLAccountGroups.GL AND GLAccountSuperGroups.ID = GLAccountGroups.GLSuperGroup INNER JOIN GLAccounts ON GLs.ID = GLAccounts.GL AND GLAccountGroups.ID = GLAccounts.GLGroup WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.ID = "& account & ")")
		
		%><A HREF="AccountInfo.asp">ฯÝสั ฿แ </A>> <A HREF="balance.asp?OpenGL=<%=RSS2("GLID")%>">สัวา ยาใวํิํ</A> > <A HREF="balance.asp?act=groups&SuperGroupID=<%=RSS2("SuperGroupID")%>"><%=RSS2("SuperGroupName")%></a>  > <A HREF="balance.asp?act=account&GroupID=<%=RSS2("GroupID")%>"><%=RSS2("GroupName")%></a>  >  <%=RSS2("Name")%>
		<BR><hR>
	</TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD><!A HREF="default.asp?s=1"><SMALL>฿ฯ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ิัอ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>ิใวัๅ ำไฯ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>สวัํฮ ำไฯ</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=5"><SMALL>ศฯๅ฿วั</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=5"><SMALL>ศำสวไ฿วั</SMALL></A></TD>
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
	
	debit = ""
	credit = ""
	if RSS("IsCredit") then
		credit = RSS("Amount")
	else
		debit = RSS("Amount")
	end if
%>
<TR bgcolor="<%=tmpColor%>" >
	<TD><!A HREF="balance.asp?act=accountRows&accountID=<%=RSS("id")%>"><%=RSS("id")%></A></TD>
	<TD><!A HREF="balance.asp?act=accountRows&accountID=<%=RSS("id")%>"><%=RSS("Description")%></A></TD>
	<TD><A HREF="javascript:void(0);" onclick="window.open('GLMemoDocShow.asp?id=<%=RSS("GLDoc")%>',null,'menubars=no')"><%=RSS("GLDocID")%></A></TD>
	<TD><%=RSS("GLDocDate")%></TD>
	<TD><%=debit%></TD>
	<TD><%=credit%></TD>
</TR>
	  
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
<%
end if


%>

<!--#include file="tah.asp" -->