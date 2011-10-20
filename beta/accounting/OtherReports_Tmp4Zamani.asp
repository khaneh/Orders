<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle="”«Ì— ê“«—‘ Â«"
SubmenuItem=10
if not Auth(8 , "E") then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {border:1pt solid white;vertical-align:top;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTable2 th {font-size:9pt; background-color:#666699;height:25px;}
	.RepTable2 input {font-family:tahoma; font-size:9pt; border:1 solid black;}
</STYLE>
<BR>
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'-----------------------------------------------------------------------------------------------------
if request("act")="MoeenRep" OR request("act")="" then

	Ord=request("Ord")

	select case Ord
	case "1":
		order="Tafsil"
	case "-1":
		order="Tafsil DESC"
	case "2":
		order="AccountTitle"
	case "-2":
		order="AccountTitle DESC"
	case "3":
		order="totalDebit"
	case "-3":
		order="totalDebit DESC"
	case "4":
		order="totalCredit"
	case "-4":
		order="totalCredit DESC"
	case "5","-6":
		order="ISNULL(SUM(ARItems.AmountOriginal * ARItems.IsCredit), 0) - ISNULL(SUM(ARItems.AmountOriginal * (1 - ARItems.IsCredit)), 0) DESC"
	case "6","-5":
		order="ISNULL(SUM(ARItems.AmountOriginal * ARItems.IsCredit), 0) - ISNULL(SUM(ARItems.AmountOriginal * (1 - ARItems.IsCredit)), 0)"
	case else:
		order="Tafsil"
		Ord=1
	end select

%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function showAcc(acc){
		window.open('tafsili.asp?accountID='+acc+'&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>&moeenFrom=0&moeenTo=99999&act=Show');
	}
	//-->
	</SCRIPT>
	<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
	<TR bgcolor="#EEEECC" height="30">
		<TD colspan=7>
		„«‰œÂ Õ”«» Â«Ì „‘ —Ì«‰ œ— ”Ì” „ Ã«„⁄ œ— Å«Ì«‰ ”«· 82	
		</TD>
	</TR>
<%

	'mySQL = "SELECT SUM(SumCred) AS SumCred, SUM(SumDeb) AS SumDeb, SUM(Flow * (Sgn + 1) / 2) AS sumFlowCred, SUM(Flow * (Sgn - 1) / 2) AS sumFlowdeb FROM (SELECT SUM(IsCredit * Amount) AS SumCred, SUM(- ((IsCredit - 1) * Amount)) AS SumDeb, SUM(IsCredit * Amount) - SUM(- ((IsCredit - 1) * Amount)) AS Flow, SIGN(SUM(IsCredit * Amount) - SUM(- ((IsCredit - 1) * Amount))) AS Sgn FROM EffectiveGLRows WHERE (GLAccount = "& GLAccount & ") AND (GL = "& OpenGL & ") AND (ISNULL(Tafsil, 0) >= "& FromTafsil & ") AND (ISNULL(Tafsil, 0) <= "& ToTafsil & ") AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "') GROUP BY Tafsil) FlowTbl"
	mySQL="SELECT SUM(totalCredit) AS sumCred, SUM(totalDebit) AS sumDeb, SUM(Flow * (Sgn + 1) / 2) AS sumFlowCred, SUM(Flow * (Sgn - 1) / 2) AS sumFlowdeb FROM (SELECT Account AS Tafsil, ISNULL(SUM(AmountOriginal * IsCredit), 0) AS totalCredit, ISNULL(SUM(AmountOriginal * (1 - IsCredit)), 0) AS totalDebit, ISNULL(SUM(AmountOriginal * (CONVERT(tinyint, IsCredit) - .5) * 2), 0) AS Flow, ISNULL(SIGN(SUM(AmountOriginal * (CONVERT(tinyint, IsCredit) - .5) * 2)), 0) AS Sgn FROM ARItems WHERE (Voided = 0) AND (EffectiveDate < '1383/01/01') GROUP BY Account) FlowTbl"

	set rs=Conn.Execute (mySQL)

%>
	<TR bgcolor="#EEEECC" >
		<TD colspan=2 rowspan=2>
		Ã„⁄
		</TD>
		<TD width=70 >ê—œ‘ »œÂﬂ«—		</TD>
		<TD width=70 >ê—œ‘ »” «‰ﬂ«—		</TD>
		<TD width=70 >„«‰œÂ »œÂﬂ«—		</TD>
		<TD width=70 >„«‰œÂ »” «‰ﬂ«—	</TD>
	</TR>
	<TR bgcolor="#EEEECC" >
		<TD width=70 ><%=Separate(rs("SumDeb"))%></TD>
		<TD width=70 ><%=Separate(rs("SumCred"))%></TD>
		<TD width=70 ><%=Separate(rs("SumFlowDeb"))%></TD>
		<TD width=70 ><%=Separate(rs("SumFlowCred"))%></TD>
	</TR>
	<TR bgcolor="black" height="2">
		<TD colspan="6" style="padding:0;"></TD>
	</TR>
<%
	rs.close

	if ord<0 then
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>6 6 6</span>"
	else
		style="background-color: #33CC99;"
		arrow="<br><span style='font-family:webdings'>5 5 5</span>"
	end if
%>
	<TR bgcolor="eeeeee" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
		<TD width=50 onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>"> ›’Ì·Ì			<%if abs(ord)=1 then response.write arrow%></TD>
		<TD width='*' onclick='go2Page(1,2);' style="<%if abs(ord)=2 then response.write style%>">⁄‰Ê«‰ Õ”«»		<%if abs(ord)=2 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-3);' style="<%if abs(ord)=3 then response.write style%>">ê—œ‘ »œÂﬂ«—		<%if abs(ord)=3 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-4);' style="<%if abs(ord)=4 then response.write style%>">ê—œ‘ »” «‰ﬂ«—		<%if abs(ord)=4 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-5);' style="<%if abs(ord)=5 then response.write style%>">„«‰œÂ »œÂﬂ«—		<%if abs(ord)=5 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-6);' style="<%if abs(ord)=6 then response.write style%>">„«‰œÂ »” «‰ﬂ«—	<%if abs(ord)=6 then response.write arrow%></TD>
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

	'mySQL="SELECT GLRows.Tafsil, SUM(GLRows.IsCredit * GLRows.Amount) AS totalCredit, SUM(- ((GLRows.IsCredit - 1) * GLRows.Amount)) AS totalDebit, Accounts.AccountTitle AS AccountTitle FROM (SELECT ID AS GLDoc, GLDocDate FROM GLDocs WHERE (GLDocs.IsTemporary = 1 OR GLDocs.IsChecked = 1 OR GLDocs.IsFinalized = 1) AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "') AND (GL = "& openGL & " ) AND (IsRemoved = 0) AND (deleted = 0)) EffectiveGLDocs INNER JOIN GLRows ON EffectiveGLDocs.GLDoc = GLRows.GLDoc INNER JOIN Accounts ON GLRows.Tafsil = Accounts.ID WHERE (GLRows.GLAccount = "& GLAccount & ") AND (GLRows.deleted = 0) GROUP BY GLRows.Tafsil, Accounts.AccountTitle HAVING (ISNULL(GLRows.Tafsil, 0) >= "& FromTafsil & ") AND (ISNULL(GLRows.Tafsil, 0) <= "& ToTafsil & ") ORDER BY " & order
	mySQL="SELECT ARItems.Account AS Tafsil, ISNULL(SUM(ARItems.AmountOriginal * ARItems.IsCredit), 0) AS totalCredit, ISNULL(SUM(ARItems.AmountOriginal * (1 - ARItems.IsCredit)), 0) AS totalDebit, Accounts.AccountTitle FROM ARItems INNER JOIN Accounts ON ARItems.Account = Accounts.ID WHERE (ARItems.Voided = 0) AND (ARItems.EffectiveDate < '1383/01/01') GROUP BY ARItems.Account, Accounts.AccountTitle ORDER BY "& order

'
' The Differences:
' mySQL="SELECT FromHan.Tafsil, FromHan.totalCredit AS HanCredit, FromHan.totalDebit AS HanDebit, DRV.totalCredit AS sysCred, DRV.totalDebit AS sysDeb FROM (SELECT GLRows.Tafsil, SUM(GLRows.IsCredit * GLRows.Amount) AS totalCredit, SUM(- ((GLRows.IsCredit - 1) * GLRows.Amount)) AS totalDebit FROM (SELECT ID AS GLDoc, GLDocDate FROM GLDocs WHERE (GLDocs.IsTemporary = 1 OR GLDocs.IsChecked = 1 OR GLDocs.IsFinalized = 1) AND (GLDocDate >= N'1383/01/01') AND (GLDocDate <= N'1383/01/01') AND (GL = 83) AND (IsRemoved = 0) AND (deleted = 0)) EffectiveGLDocs INNER JOIN GLRows ON EffectiveGLDocs.GLDoc = GLRows.GLDoc WHERE (GLRows.deleted = 0) AND (LEFT(GLRows.GLAccount, 2) = 13) AND (GLRows.Tafsil IS NOT NULL) GROUP BY GLRows.Tafsil) FromHan INNER JOIN (SELECT Account AS Tafsil, ISNULL(SUM(AmountOriginal * IsCredit), 0) AS totalCredit, ISNULL(SUM(AmountOriginal * (1 - IsCredit)), 0) AS totalDebit FROM ARItems WHERE (Voided = 0) AND (EffectiveDate < '1383/01/01') GROUP BY Account) DRV ON FromHan.Tafsil = DRV.Tafsil AND FromHan.totalCredit - FromHan.totalDebit <> DRV.totalCredit - DRV.totalDebit"
'
'
	Set rs=Server.CreateObject("ADODB.Recordset")'Conn.Execute(mySQL)

	PageSize = 25
	rs.PageSize = PageSize

	rs.CursorLocation=3 'in ADOVBS_INC adUseClient=3
	rs.Open mySQL ,Conn,3
	TotalPages = rs.PageCount

	CurrentPage=1

	if isnumeric(Request.QueryString("p")) then
		pp=clng(Request.QueryString("p"))
		if pp <= TotalPages AND pp > 0 then
			CurrentPage = pp
		end if
	end if

	if not rs.eof then
		rs.AbsolutePage=CurrentPage
	end if

	if rs.eof then
%>		<tr>
			<td bgcolor="#BBBBBB" height="30" colspan="7" align=center><b>ÂÌç .</b></td>
		</tr>
<%	else
		Do While NOT rs.eof AND (rs.AbsolutePage = CurrentPage)
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

			totalDebit =	cdbl(rs("totalDebit"))
			totalCredit =	cdbl(rs("totalCredit"))

			if totalCredit > totalDebit then
				creditRemained =	totalCredit - totalDebit
				debitRemained =	0
			else
				creditRemained =	0
				debitRemained =	totalDebit - totalCredit
			end if

			SumCredit = SumCredit + totalCredit 
			SumDebit =	SumDebit + totalDebit

			SumCreditRemained = SumCreditRemained + creditRemained
			SumDebitRemained =	SumDebitRemained + debitRemained

	%>
			<TR bgcolor="<%=tmpColor%>" >
				<TD dir=ltr align=right><A HREF="javascript:showAcc(<%=rs("Tafsil")%>);"><%=rs("Tafsil")%></A></TD>
				<TD><%=rs("AccountTitle")%></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(totalDebit)%></span></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(totalCredit)%></span></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(debitRemained)%></span></TD>
				<TD dir=ltr align=right><span dir=ltr><%=Separate(creditRemained)%></span></TD>
			</TR>
			 
	<% 
		rs.moveNext
		Loop

		if TotalPages > 1 then
			pageCols=20
%>			
			<TR bgcolor="eeeeee" >
				<TD colspan=6 height=2 bgcolor=0></TD>
			</TR>
			<TR class="RepTableTitle">
				<TD bgcolor="#CCCCEE" height="30" colspan="6">
				<table width=100% cellspacing=0 style="cursor:hand;color:gray;">
				<tr>
					<td style="height:25;border-bottom:1 solid black;" colspan=<%=pagecols%>>
						<b>’›ÕÂ <%=CurrentPage%> «“ <%=TotalPages%></b>
						&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>,0);">’›ÕÂ »⁄œ &gt;</a>
					</td>
				</tr>
				<tr>
<%				for i=1 to TotalPages 
					if i = CurrentPage then 
%>						<td style="color:black;"><b>[<%=i%>]</b></td>
<%					else
%>						<td onclick="go2Page(<%=i%>,0);"><%=i%></td>
<%					end if 
					if i mod pageCols = 0 then response.write "</tr><tr>" 
				next 

%>				</tr>
				</table>
				</TD>
			</TR>
<%		end if
%>
		</TABLE><br>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function go2Page(p,ord) {
			if(ord==0){
				ord=<%=Ord%>;
			}
			else if(ord==<%=Ord%>){
				ord= 0-ord;
			}
			str='?act=MoeenRep&GLAccount='+escape('<%=GLAccount%>')+'&FromDate='+escape('<%=FromDate%>')+'&ToDate='+escape('<%=ToDate%>')+'&FromTafsil='+escape('<%=FromTafsil%>')+'&ToTafsil='+escape('<%=ToTafsil%>')+'&Ord='+escape(ord)+'&p='+escape(p) //+'& ='+escape(' ')+'& ='+escape(' ')+'& ='+escape(' ')
			window.location=str;
		}
		//-->
		</SCRIPT>
<%
	end if
end if


%>

<!--#include file="tah.asp" -->