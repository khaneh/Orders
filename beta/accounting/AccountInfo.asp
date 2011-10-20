<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
Response.Buffer=false
'Accounting (8)
PageTitle= "œ› — ﬂ·"
SubmenuItem=4
if not Auth(8 , 4) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'------------------------------------------------------------------------------------------		+
'--------------------------------------------------------------------------------------------- Scripts
'-----------------------------------------------------------------------------------------------------
if request.form("act")="" then
	DateFrom	=	request("DateFrom")
	DateTo		=	request("DateTo")
	act			=	request("act")
	id			=	request("id")
	ShowRemained=	request("ShowRemained")
else
	DateFrom	=	request.form("DateFrom")
	DateTo		=	request.form("DateTo")
	act			=	request.form("act")
	id			=	request.form("id")
	ShowRemained=	request.form("ShowRemained")
end if

if ShowRemained="on" then
	ShowRemained=true
else
	ShowRemained=false
end if

if DateFrom		= "" then	DateFrom = fiscalYear & "/01/01"
if DateTo		= "" then	DateTo = shamsiToday()

%>
<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>
<BR>
<FORM METHOD=POST ACTION="">
<input type="hidden" Name='act' value='<%=act%>'>
<input type="hidden" Name='id' value='<%=id%>'>
<TABLE style="border:1 solid #330066;" align=center>
<TR>
	<TD align=left>«“  «—ÌŒ</TD>
	<TD><INPUT dir="LTR"  TYPE="text" NAME="DateFrom" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=DateFrom%>"></TD>
	<TD align=left> «  «—ÌŒ </TD>
	<TD><INPUT dir="LTR"  TYPE="text" NAME="DateTo" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=DateTo%>"></TD>
	<TD align=center>
		<INPUT TYPE="submit" NAME="submitt" value="„‘«ÂœÂ">
		<INPUT TYPE="hidden" Name="Order" Value="">
	</TD>
	<TD>
		<INPUT TYPE="checkbox" NAME="ShowRemained" <%if ShowRemained then response.write "checked"%>> „«‰œÂ ﬁ»· ‰„«Ì‘ œ«œÂ ‘Êœ.
	</TD>
</TR>
</FORM>
</TABLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
function drill(act,id){
	document.all.act.value=act;
	document.all.id.value=id;
	document.forms[0].submit();
}
//-->
</SCRIPT>
<%
dateStr = "«“  «—ÌŒ <span dir=rtl>" & fiscalYear & "/1/1</span>  « <span dir=rtl>" & shamsiToday() & "</span>"

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- List GL Super Groups
'-----------------------------------------------------------------------------------------------------
if request("act")="" then

%>
	<br>
	<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
	<TR bgcolor="#CCCCEE" height="30">
		<TD colspan=6>
			<A HREF="?"><%=OpenGLName%></A>
		</TD>
	</TR>
	<TR bgcolor="black" height="2">
		<TD colspan="6" style="padding:0;"></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD>ﬂœ</TD>
		<TD>‰«„ ”—ê—ÊÂ</TD>
		<TD>ê—œ‘ »œÂﬂ«—</TD>
		<TD>ê—œ‘ »” «‰ﬂ«—</TD>
		<TD>„«‰œÂ »œÂﬂ«—</TD>
		<TD>„«‰œÂ »” «‰ﬂ«—</TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
	</TR>
<%
	SumCredit=0
	SumDebit=0
	SumCreditRemained=0
	SumDebitRemained=0
	tmpCounter=0

	'Added By kid 830805
	if ShowRemained then
		mySQL="SELECT ISNULL(SUM(IsCredit * Amount),0) AS totalCredit, ISNULL(SUM(- ((IsCredit - 1) * Amount)),0) AS totalDebit FROM EffectiveGLRows WHERE (GLDocDate < N'"& DateFrom & "') AND (GL = "& OpenGL & ")"
		set RSS2=Conn.Execute (mySQL)

		if RSS2.eof then
			totalDebit =	0
			totalCredit =	0
		else
			totalDebit =	cdbl(RSS2("totalDebit"))
			totalCredit =	cdbl(RSS2("totalCredit"))
		end if

		RSS2.close
		
		tmpCounter=1

		if totalCredit > totalDebit then
			creditRemained =	totalCredit - totalDebit
			debitRemained  =	0
		else
			creditRemained =	0
			debitRemained  =	totalDebit - totalCredit
		end if

		SumCredit = SumCredit + totalCredit 
		SumDebit =	SumDebit + totalDebit

		SumCreditRemained = SumCreditRemained + creditRemained
		SumDebitRemained =	SumDebitRemained + debitRemained
%>
		<TR bgcolor="#FFFFFF">
			<TD style="border-bottom:1 solid black;" colspan=2 align=left>„«‰œÂ ﬁ»· «“  «—ÌŒ <%=replace(DateFrom,"/",".")%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(totalDebit)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(totalCredit)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(debitRemained)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(creditRemained)%></TD>
		</TR>
<% 
	end if


	'changed by kid 830805
	mySQL="SELECT ISNULL(DRV.totalCredit, 0) AS totalCredit, ISNULL(DRV.totalDebit, 0) AS totalDebit, GLAccountSuperGroups.ID, GLAccountSuperGroups.Name FROM (SELECT SUM(EffectiveGLRows.IsCredit * EffectiveGLRows.Amount) AS totalCredit, SUM(- ((EffectiveGLRows.IsCredit - 1)  * EffectiveGLRows.Amount)) AS totalDebit, GLAccountGroups.GLSuperGroup FROM EffectiveGLRows INNER JOIN GLAccounts INNER JOIN GLAccountGroups ON GLAccounts.GLGroup = GLAccountGroups.ID AND GLAccounts.GL = GLAccountGroups.GL ON  EffectiveGLRows.GL = GLAccounts.GL AND EffectiveGLRows.GLAccount = GLAccounts.ID WHERE (GLDocDate >= N'"& DateFrom & "') AND (GLDocDate <= N'"& DateTo & "') AND (EffectiveGLRows.GL = "& OpenGL&") GROUP BY GLAccountGroups.GLSuperGroup) DRV RIGHT OUTER JOIN GLAccountSuperGroups ON DRV.GLSuperGroup = GLAccountSuperGroups.ID WHERE (GLAccountSuperGroups.GL = "& OpenGL&") ORDER BY GLAccountSuperGroups.ID"

	set RSS=Conn.Execute (mySQL)

	Do while not RSS.eof
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 

		totalDebit =	cdbl(RSS("totalDebit"))
		totalCredit =	cdbl(RSS("totalCredit"))

		if totalCredit > totalDebit then
			creditRemained =	totalCredit - totalDebit
			debitRemained  =	0
		else
			creditRemained =	0
			debitRemained  =	totalDebit - totalCredit
		end if

		SumCredit = SumCredit + totalCredit 
		SumDebit =	SumDebit + totalDebit

		SumCreditRemained = SumCreditRemained + creditRemained
		SumDebitRemained =	SumDebitRemained + debitRemained

%>
		<TR bgcolor="<%=tmpColor%>" height=30>
			<TD><A HREF="javascript:drill('groups','<%=RSS("ID")%>');"><%=RSS("ID")%></A></TD>
			<TD><A HREF="javascript:drill('groups','<%=RSS("ID")%>');"><%=RSS("Name")%></A></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(totalDebit)%></span></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(totalCredit)%></span></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(debitRemained)%></span></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(creditRemained)%></span></TD>
		</TR>
		  
<% 
	RSS.moveNext
	Loop

	if SumDebitRemained > SumCreditRemained then
		TotalSumDebitRemained = SumDebitRemained - SumCreditRemained 
		TotalSumCreditRemained=0
	else
		TotalSumDebitRemained=0
		TotalSumCreditRemained = SumCreditRemained - SumDebitRemained 
	end if

%>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD style="border-bottom:1 solid black;" colspan=2 align=left>Ã„⁄:</TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumDebit)%></TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumCredit)%></TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumDebitRemained)%></TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumCreditRemained)%></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=4 align=left>„«‰œÂ œ— ﬂ·:</TD>
		<TD dir=ltr align=right><%=Separate(TotalSumDebitRemained)%></TD>
		<TD dir=ltr align=right><%=Separate(TotalSumCreditRemained)%></TD>
	</TR>
	</TABLE><br>

	<CENTER>
	<FORM METHOD=POST ACTION="AccountInfo.asp">
	<!--INPUT TYPE="button" onclick="window.location='balance.asp'" class="GenButton" value=" —«“" style="width:90px; ">
	<INPUT TYPE="button" onclick="window.location='sood.asp'" class="GenButton" value="”Êœ Ê “Ì«‰" style="width:90px; "--><br><br>
	<INPUT TYPE="button" onclick="window.location='manageGL.asp?act=select'" class="GenButton" value=" €ÌÌ—" style="width:90px; background-color: #FFFFBB ">
	</FORM>
	</CENTER>
<%

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------ List GL Groups under a super group
'-----------------------------------------------------------------------------------------------------
elseif request("act")="groups" then
	SuperGroupID = clng(id)
	'changed by kid 830804
	mySQL="SELECT GLs.ID AS GLID, GLs.Name AS GLname, GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL WHERE (GLAccountSuperGroups.GL = "& OpenGL & ") AND (GLAccountSuperGroups.ID = "& SuperGroupID & ")"
	set rsGLNames=Conn.Execute (mySQL)

	if not rsGLNames.EOF then
%>
		<br>
		<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
		<TR bgcolor="#CCCCEE" height="30">
			<TD colspan=7>
				<A HREF="javascript:drill('','');"><%=rsGLNames("GLname")%></A>
				> <A HREF="javascript:drill('groups','<%=rsGLNames("SuperGroupID")%>');"><%=rsGLNames("SuperGroupName")%></A>
				[<%=rsGLNames("SuperGroupID")%>]
			</TD>
		</TR>
		<TR bgcolor="black" height="2">
			<TD colspan="7" style="padding:0;"></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD>ﬂœ</TD>
			<TD>‰«„ ê—ÊÂ</TD>
			<TD>ê—œ‘ »œÂﬂ«—</TD>
			<TD>ê—œ‘ »” «‰ﬂ«—</TD>
			<TD>„«‰œÂ »œÂﬂ«—</TD>
			<TD>„«‰œÂ »” «‰ﬂ«—</TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD colspan=7 height=2 bgcolor=0></TD>
		</TR>
<%	end if 
	rsGLNames.close

	SumCredit=0
	SumDebit=0
	SumCreditRemained=0
	SumDebitRemained=0
	tmpCounter=0
	'Added By kid 830805
	if ShowRemained then
		mySQL="SELECT ISNULL(SUM(IsCredit * Amount),0) AS totalCredit, ISNULL(SUM(- ((IsCredit - 1) * Amount)),0) AS totalDebit FROM GLAccounts INNER JOIN EffectiveGLRows ON GLAccounts.GL = EffectiveGLRows.GL AND GLAccounts.ID = EffectiveGLRows.GLAccount INNER JOIN GLAccountGroups ON GLAccounts.GLGroup = GLAccountGroups.ID AND GLAccounts.GL = GLAccountGroups.GL WHERE (GLDocDate < N'"& DateFrom & "') AND (EffectiveGLRows.GL = "& OpenGL & ") GROUP BY GLAccountGroups.GLSuperGroup HAVING (GLAccountGroups.GLSuperGroup = "& SuperGroupID & ")"
		set RSS2=Conn.Execute (mySQL)

		if RSS2.eof then
			totalDebit =	0
			totalCredit =	0
		else
			totalDebit =	cdbl(RSS2("totalDebit"))
			totalCredit =	cdbl(RSS2("totalCredit"))
		end if

		RSS2.close
		
		tmpCounter=1

		if totalCredit > totalDebit then
			creditRemained =	totalCredit - totalDebit
			debitRemained  =	0
		else
			creditRemained =	0
			debitRemained  =	totalDebit - totalCredit
		end if

		SumCredit = SumCredit + totalCredit 
		SumDebit =	SumDebit + totalDebit

		SumCreditRemained = SumCreditRemained + creditRemained
		SumDebitRemained =	SumDebitRemained + debitRemained
%>
		<TR bgcolor="#FFFFFF">
			<TD style="border-bottom:1 solid black;" colspan=2 align=left>„«‰œÂ ﬁ»· «“  «—ÌŒ <%=replace(DateFrom,"/",".")%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(totalDebit)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(totalCredit)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(debitRemained)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(creditRemained)%></TD>
		</TR>
<% 
	end if


'	mySQL="SELECT SUM(IsCredit * Amount) AS totalCredit, SUM(- ((IsCredit - 1) * Amount)) AS totalDebit, GLAccounts.GLGroup, GLAccountGroups.Name FROM GLAccounts INNER JOIN EffectiveGLRows ON GLAccounts.GL = EffectiveGLRows.GL AND GLAccounts.ID = EffectiveGLRows.GLAccount INNER JOIN GLAccountGroups ON GLAccounts.GLGroup = GLAccountGroups.ID AND GLAccounts.GL = GLAccountGroups.GL WHERE (GLDocDate >= N'"& DateFrom & "') AND (GLDocDate <= N'"& DateTo & "') AND (EffectiveGLRows.GL = "& OpenGL & ") GROUP BY GLAccountGroups.GLSuperGroup, GLAccounts.GLGroup, GLAccountGroups.Name HAVING (GLAccountGroups.GLSuperGroup = "& SuperGroupID & ") ORDER BY GLAccounts.GLGroup"
	'Changed By Kid 830805
	mySQL="SELECT ISNULL(DRV.totalCredit, 0) AS totalCredit, ISNULL(DRV.totalDebit, 0) AS totalDebit, GLAccountGroups.ID, GLAccountGroups.Name FROM (SELECT SUM(EffectiveGLRows.IsCredit * EffectiveGLRows.Amount) AS totalCredit, SUM(- ((EffectiveGLRows.IsCredit - 1)  * EffectiveGLRows.Amount)) AS totalDebit, GLAccounts.GLGroup FROM GLAccounts INNER JOIN EffectiveGLRows ON GLAccounts.GL = EffectiveGLRows.GL AND GLAccounts.ID = EffectiveGLRows.GLAccount WHERE (EffectiveGLRows.GLDocDate >= N'"& DateFrom & "') AND (EffectiveGLRows.GLDocDate <= N'"& DateTo & "') AND (EffectiveGLRows.GL = "& OpenGL & ") GROUP BY GLAccounts.GLGroup) DRV RIGHT OUTER JOIN GLAccountGroups ON DRV.GLGroup = GLAccountGroups.ID WHERE (GLAccountGroups.GLSuperGroup = "& SuperGroupID & ") AND (GLAccountGroups.GL = "& OpenGL & ") ORDER BY ID"


	set RSS=Conn.Execute (mySQL)

	Do while not RSS.eof
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 

		totalDebit =	cdbl(RSS("totalDebit"))
		totalCredit =	cdbl(RSS("totalCredit"))

		if totalCredit > totalDebit then
			creditRemained =	totalCredit - totalDebit
			debitRemained  =	0
		else
			creditRemained =	0
			debitRemained  =	totalDebit - totalCredit
		end if

		SumCredit = SumCredit + totalCredit 
		SumDebit =	SumDebit + totalDebit

		SumCreditRemained = SumCreditRemained + creditRemained
		SumDebitRemained =	SumDebitRemained + debitRemained

%>
		<TR bgcolor="<%=tmpColor%>" >
			<TD><A HREF="javascript:drill('account','<%=RSS("ID")%>');"><%=RSS("ID")%></A></TD>
			<TD><A HREF="javascript:drill('account','<%=RSS("ID")%>');"><%=RSS("Name")%></A></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(totalDebit)%></span></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(totalCredit)%></span></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(debitRemained)%></span></TD>
			<TD dir=ltr align=right><span dir=ltr><%=Separate(creditRemained)%></span></TD>
		</TR>
		  
<% 
	RSS.moveNext
	Loop

	if SumDebitRemained > SumCreditRemained then
		TotalSumDebitRemained = SumDebitRemained - SumCreditRemained 
		TotalSumCreditRemained=0
	else
		TotalSumDebitRemained=0
		TotalSumCreditRemained = SumCreditRemained - SumDebitRemained 
	end if

%>
	<TR bgcolor="eeeeee" >
		<TD colspan=7 height=2 bgcolor=0></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD style="border-bottom:1 solid black;" colspan=2 align=left>Ã„⁄:</TD>
		<TD style="border-bottom:1 solid black;" align=right><%=Separate(SumDebit)%></TD>
		<TD style="border-bottom:1 solid black;" align=right><%=Separate(SumCredit)%></TD>
		<TD style="border-bottom:1 solid black;" align=right><%=Separate(SumDebitRemained)%></TD>
		<TD style="border-bottom:1 solid black;" align=right><%=Separate(SumCreditRemained)%></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=4 align=left>„«‰œÂ œ— ﬂ·:</TD>
		<TD dir=ltr align=right><%=Separate(TotalSumDebitRemained)%></TD>
		<TD dir=ltr align=right><%=Separate(TotalSumCreditRemained)%></TD>
	</TR>
	</TABLE><br>
<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------- List GL Accounts under a group
'-----------------------------------------------------------------------------------------------------
elseif request("act")="account" then
	GroupID = clng(id)

	mySQL = "SELECT GLs.Name AS GLName, GLs.ID AS GLID, GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName, GLAccountGroups.ID AS GroupID, GLAccountGroups.Name AS GroupName FROM GLAccountGroups INNER JOIN GLAccountSuperGroups ON GLAccountGroups.GLSuperGroup = GLAccountSuperGroups.ID INNER JOIN GLs ON GLAccountSuperGroups.GL = GLs.ID WHERE (GLAccountSuperGroups.GL = "& OpenGL & ") AND (GLAccountGroups.GL = "& OpenGL & ") AND (GLAccountGroups.ID = "& GroupID & ")"

	set rsGLNames=Conn.Execute (mySQL)

	if not rsGLNames.EOF then
%>
		<br>
		<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
		<TR bgcolor="#CCCCEE" height="30">
			<TD colspan=6>
				<A HREF="javascript:drill('','');"><%=rsGLNames("GLname")%></A>
				> <A HREF="javascript:drill('groups','<%=rsGLNames("SuperGroupID")%>');"><%=rsGLNames("SuperGroupName")%></A>
				> <A HREF="javascript:drill('account','<%=rsGLNames("GroupID")%>');"><%=rsGLNames("GroupName")%></A>
				[<%=rsGLNames("GroupID")%>]
			</TD>
		</TR>
		<TR bgcolor="black" height="2">
			<TD colspan="6" style="padding:0;"></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD>ﬂœ</TD>
			<TD>⁄‰Ê«‰ Õ”«»</TD>
			<TD>ê—œ‘ »œÂﬂ«—</TD>
			<TD>ê—œ‘ »” «‰ﬂ«—</TD>
			<TD>„«‰œÂ »œÂﬂ«—</TD>
			<TD>„«‰œÂ »” «‰ﬂ«—</TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD colspan=6 height=2 bgcolor=0></TD>
		</TR>
<%	end if 
	rsGLNames.close

	SumCredit=0
	SumDebit=0
	SumCreditRemained=0
	SumDebitRemained=0
	tmpCounter=0
	'Added By kid 830805
	if ShowRemained then

		mySQL="SELECT SUM(IsCredit * Amount) AS totalCredit, SUM(- ((IsCredit - 1) * Amount)) AS totalDebit FROM GLAccounts INNER JOIN EffectiveGLRows ON GLAccounts.GL = EffectiveGLRows.GL AND GLAccounts.ID = GLAccount WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLDocDate < N'"& DateFrom & "') GROUP BY GLAccounts.GLGroup HAVING (GLAccounts.GLGroup = "& GroupID & ")"

		set RSS2=Conn.Execute (mySQL)

		if RSS2.eof then
			totalDebit =	0
			totalCredit =	0
		else
			totalDebit =	cdbl(RSS2("totalDebit"))
			totalCredit =	cdbl(RSS2("totalCredit"))
		end if

		RSS2.close
		
		tmpCounter=1

		if totalCredit > totalDebit then
			creditRemained =	totalCredit - totalDebit
			debitRemained  =	0
		else
			creditRemained =	0
			debitRemained  =	totalDebit - totalCredit
		end if

		SumCredit = SumCredit + totalCredit 
		SumDebit =	SumDebit + totalDebit

		SumCreditRemained = SumCreditRemained + creditRemained
		SumDebitRemained =	SumDebitRemained + debitRemained
%>
		<TR bgcolor="#FFFFFF">
			<TD style="border-bottom:1 solid black;" colspan=2 align=left>„«‰œÂ ﬁ»· «“  «—ÌŒ <%=replace(DateFrom,"/",".")%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(totalDebit)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(totalCredit)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(debitRemained)%></TD>
			<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(creditRemained)%></TD>
		</TR>
<% 
	end if

	'mySQL="SELECT GLAccounts.Name, ISNULL(DERIVEDTBL.totalDebit,0) AS totalDebit, ISNULL(DERIVEDTBL.totalCredit,0) AS totalCredit, GLAccounts.ID FROM (SELECT SUM(GLRows.IsCredit * GLRows.Amount) AS totalCredit, SUM(- ((GLRows.IsCredit - 1) * GLRows.Amount)) AS totalDebit,  GLRows.GLAccount FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc WHERE (GLDocs.IsTemporary = 1 OR GLDocs.IsChecked = 1 OR GLDocs.IsFinalized = 1) AND (GLDocs.deleted = 0) AND (GLDocs.IsRemoved = 0) AND (GLRows.deleted = 0) GROUP BY GLRows.GLAccount, GLDocs.GL HAVING (GLDocs.GL = "& OpenGL & ")) DERIVEDTBL RIGHT OUTER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.GLGroup = "& GroupID & ") ORDER BY GLAccounts.ID"
	'Changed By Kid 830805
	mySQL="SELECT GLAccounts.Name, GLAccounts.HasAppendix, ISNULL(DERIVEDTBL.totalDebit, 0) AS totalDebit, ISNULL(DERIVEDTBL.totalCredit, 0) AS totalCredit, GLAccounts.ID FROM (SELECT SUM(IsCredit * Amount) AS totalCredit, SUM(- ((IsCredit - 1) * Amount)) AS totalDebit, GLAccount FROM EffectiveGLRows WHERE (GLDocDate >= N'"& DateFrom & "') AND (GLDocDate <= N'"& DateTo & "') GROUP BY GLAccount, GL HAVING (GL = "& OpenGL & ")) DERIVEDTBL RIGHT OUTER JOIN GLAccounts ON DERIVEDTBL.GLAccount = GLAccounts.ID WHERE (GLAccounts.GL = "& OpenGL & ") AND (GLAccounts.GLGroup = "& GroupID & ") ORDER BY GLAccounts.ID"

	set RSS2=Conn.Execute (mySQL)

	Do while not RSS2.eof
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 

		HasAppendix = RSS2("HasAppendix")

		totalDebit =	cdbl(RSS2("totalDebit"))
		totalCredit =	cdbl(RSS2("totalCredit"))

		if totalCredit > totalDebit then
			creditRemained =	totalCredit - totalDebit
			debitRemained  =	0
		else
			creditRemained =	0
			debitRemained  =	totalDebit - totalCredit
		end if

		SumCredit = SumCredit + totalCredit 
		SumDebit =	SumDebit + totalDebit

		SumCreditRemained = SumCreditRemained + creditRemained
		SumDebitRemained =	SumDebitRemained + debitRemained

		if ShowRemained then
			MoeenShowRemained = "on"
		else
			MoeenShowRemained = "off"
		end if
%>
		<TR bgcolor="<%=tmpColor%>" >
			<TD>
			<%if HasAppendix then%>
				* <A Target="_blank" HREF="OtherReports.asp?act=MoeenRep&GLAccount=<%=RSS2("id")%>&FromDate=<%=DateFrom%>&ToDate=<%=DateTo%>&FromTafsil=0&ToTafsil=999999"><%=RSS2("id")%></A>
			<%else%>
				<A Target="_blank" HREF="moeen.asp?act=Show&accountID=<%=RSS2("id")%>&accountName=<%=Server.URLEncode(RSS2("Name"))%>&FromDate=<%=DateFrom%>&ToDate=<%=DateTo%>&ShowRemained=<%=MoeenShowRemained%>"><%=RSS2("id")%></A>
			<%end if %>
			</TD>
			<TD><A Target="_blank" HREF="moeen.asp?act=Show&accountID=<%=RSS2("id")%>&accountName=<%=Server.URLEncode(RSS2("Name"))%>&FromDate=<%=DateFrom%>&ToDate=<%=DateTo%>&ShowRemained=<%=MoeenShowRemained%>"><%=RSS2("Name")%></A></TD>
			<TD dir=ltr align=right><%=Separate(totalDebit)%></TD>
			<TD dir=ltr align=right><%=Separate(totalCredit)%></TD>
			<TD dir=ltr align=right><%=Separate(debitRemained)%></TD>
			<TD dir=ltr align=right><%=Separate(creditRemained)%></TD>
		</TR>
		  
<% 
	RSS2.moveNext
	Loop

	if SumDebitRemained > SumCreditRemained then
		TotalSumDebitRemained = SumDebitRemained - SumCreditRemained 
		TotalSumCreditRemained=0
	else
		TotalSumDebitRemained=0
		TotalSumCreditRemained = SumCreditRemained - SumDebitRemained 
	end if

%>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD style="border-bottom:1 solid black;" colspan=2 align=left>Ã„⁄  «  «—ÌŒ <%=replace(DateTo,"/",".")%></TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumDebit)%></TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumCredit)%></TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumDebitRemained)%></TD>
		<TD style="border-bottom:1 solid black;" dir=ltr align=right><%=Separate(SumCreditRemained)%></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=4 align=left>„«‰œÂ œ— ﬂ·:</TD>
		<TD dir=ltr align=right><%=Separate(TotalSumDebitRemained)%></TD>
		<TD dir=ltr align=right><%=Separate(TotalSumCreditRemained)%></TD>
	</TR>
	</TABLE><br>
<%
end if
%>
<!--#include file="tah.asp" -->