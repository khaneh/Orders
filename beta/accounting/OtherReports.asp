<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle="���� ����� ��"
SubmenuItem=10
if not Auth(8 , "E") then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {border:1pt solid white;vertical-align:top;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTable2 th {font-size:9pt; background-color:#666699;height:25px;}
	.RepTable2 input {font-family:tahoma; font-size:9pt; border:1 solid black;}
	.RepTR0 {background-color: #DDDDDD;}
	.RepTR1 {background-color: #FFFFFF;}
	.RepTableHeader {background-color: #BBBBFF; text-align: center; font-weight:bold;}
</STYLE>
<BR>

<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'-----------------------------------------------------------------------------------------------------
if request("act")="" then
%>

<TABLE class="RepTable" align=center cellspacing=10>
<TR>
	<TD width="130">
		<FORM METHOD=POST ACTION="?act=MoeenRep">
		<table class="RepTable2" id="MoeenRep01">
			<tr>
			<th colspan="2">����� ����</td>
		</tr>
		<tr>
			<td align=left>����</td>
			<td><INPUT TYPE="text" NAME="GLAccount" style="width:50px;" maxlength=5></td>
		</tr>
		<tr>
			<td align=left>�� �����</td>
			<td><INPUT TYPE="text" NAME="FromDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
		</tr>
		<tr>
			<td align=left>�� �����</td>
			<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
		</tr>
		<tr>
			<td align=left>�� �����</td>
			<td><INPUT TYPE="text" NAME="FromTafsil" style="width:50px;" maxlength=6 value="0"></td>
		</tr>
		<tr>
			<td align=left>�� �����</td>
			<td><INPUT TYPE="text" NAME="ToTafsil" style="width:50px;" maxlength=6 value="999999">
				<INPUT TYPE="hidden" NAME="GL" value="<%=OpenGL%>">
			</td>
		</tr>
		</table>
		<div align=center>
		<% 	ReportLogRow = PrepareReport ("Receipt.rpt", "Recept_ID", Receipt, "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT Class="GenButton" TYPE="submit" name="action" value=" ����� ">
		<INPUT Class="GenButton" TYPE="submit" name="action" style="border:1 solid green;" value=" �ǁ ">
		</div>
		</FORM>
	</TD>
	<TD width="130" align=center>
	 &nbsp;
	 <A HREF="OtherReports_Tmp4Zamani.asp" style='font-weight:bold;'>����� ���� ��� ������� �� ����� ���� �� ����� ��� 82</A>
	</TD>
</TR>
<TR>
	<TD width="130">
	<%
	if Auth(8 , "G") then
	%>
		<FORM METHOD=POST ACTION="?act=cash">
			<table class="RepTable2" id="MoeenRep01">
				<tr>
					<th colspan="2">������ ���������</td>
				</tr>
				
			</table>		
			 <div align="center">
			 	
				 	<INPUT Class="GenButton" TYPE="submit" name="action" value=" ����� ">
				
			 </div>
	 	</form>
	 <%end if%>
	</TD>
	<TD>
	 &nbsp;
	</TD>
</TR>
</TABLE>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="cash" then 
	dim fmonth(12)
	fmonth(0)="�������"
	fmonth(1)="��������"
	fmonth(2)="�����"
	fmonth(3)="���"
	fmonth(4)="�����"
	fmonth(5)="������"
	fmonth(6)="���"
	fmonth(7)="����"
	fmonth(8)="���"
	fmonth(9)="��"
	fmonth(10)="����"
	fmonth(11)="�����"
	function echoTD(num,sys)
		result="<td style=""direction: ltr;text-align: right;color:"
		if num<0 then 
			result = result & "red"
		else
			reslt = result & "black"
		end if
		result = result & ";"">" 
		select case sys
		case "ar" ' ----------- AR Items
			result = result & "<a target=�_blank� href=""../AR/rep_dailySale.asp?input_date_start="
			if EffectiveDate<>"old" then 
				result = result & Server.URLEncode(EffectiveDate&"01")&"&input_date_end=" & Server.URLEncode(EffectiveDate&"31")&"&fullyApplied=on"" >" 
			else
				result = result & Server.URLEncode("0000/00/00")&"&input_date_end=" & Server.URLEncode("1388/01/01") &"&fullyApplied=on"">" 
			end if
		case "ap" '------------ AP Items
			result = result & "<a target=�_blank� href=""../AP/report.asp?fromDate="
			if EffectiveDate<>"old" then 
				result = result & Server.URLEncode(EffectiveDate&"01")&"&toDate=" & Server.URLEncode(EffectiveDate&"31")&"&vouchers=on&payments=on"">" 
			else
				result = result & Server.URLEncode(shamsiDate(DateAdd("m",-10,Date)))&"&toDate=" & Server.URLEncode(mid(shamsiDate(DateAdd("m",-6,Date)),1,8)) & "31" &"&vouchers=on&payments=on"">" 
			end if
		case else
			if mid(sys,1,4)="bank" then 
				result = result & "<a target=�_blank� href=""../bank/CheqBook.asp?act=showBook&GLAccount=" & mid(sys,5,5) & "&FromDate="
				result = result & Server.URLEncode(Ref2&"01")&"&ToDate=" & Server.URLEncode(Ref2&"31")&"&ShowRemained=&displayMode=0"">" 
			end if
		end select
		result = result & Separate(num) 
		if sys="ar" then result = result & "</a>"
		result = result & "</td>"
		echoTD = result
	end function
	mySQL="select * from (select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17003 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17003 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17004 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17004 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42001 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42001 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42004 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42004 group by SUBSTRING(Ref2,1,8),GLAccount) drv order by Ref2,GLAccount"
	'response.write mySQL
	set rs=Conn.Execute(mySQL)
	If rs.EOF then
		Conn.Close
		response.redirect "?errMsg=" & Server.URLEncode("���� ����! ���� ���� ���")
	end if
	sub setValue()
		select case rs("GLAccount")
			case "17003":
				amount17003 = CDbl(rs("Amount"))
			case "17004":
				amount17004 = CDbl(rs("Amount"))
			case "42001":
				amount42001 = CDbl(rs("Amount"))
			case "42004":
				amount42004 = CDbl(rs("Amount"))			
		end select
		rs.MoveNext
		if not rs.eof then 
			if ref2=rs("Ref2") then
				ref2=rs("Ref2")
				call setValue()
			end if
		end if
	end sub
%>
<table>
	<tr class="RepTableHeader">
		<td rowspan="2" colspan="2">&nbsp;</td>
		<td colspan="2">�� ����� ����</td>
		<td rowspan="2">��� �� �����</td>
		<td colspan="2">����� ��������</td>
		<td rowspan="2">��� ����� ��������</td>
		<td rowspan="2">����</td>
	</tr>
	<tr class="RepTableHeader">
		<td>17003</td>
		<td>17004</td>
		<td>42001</td>
		<td>42004</td>
	</tr>
	<%
	rowColor="RepTR0"
	do while not rs.eof
	ref2=rs("Ref2")
	if rowColor="RepTR0" then 
		rowColor="RepTR1"
	else
		rowColor="RepTR0"
	end if
	%>
	<tr class="<%=rowColor%>">
		<%
		yyyy=mid(ref2,1,4)
		mm=fmonth(cint(mid(ref2,6,2))-1)
		amount17003=0
		amount17004=0
		amount42001=0
		amount42004=0
		
		call setValue()
		%>
		<td title="<%=ref2%>"><%=yyyy%></td>
		<td><%=mm%></td>
		<%
		response.write echoTD(amount17003,"bank17003")
		response.write echoTD(amount17004,"bank17003")
		response.write echoTD(amount17003+amount17004,"")
		response.write echoTD(amount42001,"bank42001")
		response.write echoTD(amount42004,"bank42004")
		response.write echoTD(amount42001+amount42004,"")
		response.write echoTD(amount17003+amount17004-(amount42001+amount42004),"")
		%>
		
	</tr>
	<%
	loop

	mySQL="select * from (select sum(RemainedAmount) as RemainedAmount, SUBSTRING(EffectiveDate,1,8) as EffectiveDate, 'ar' as sys from ARItems where FullyApplied=0 and Voided=0 and ARItems.Type=1 and EffectiveDate>='1388/01/01' group by SUBSTRING(EffectiveDate,1,8) UNION select sum(RemainedAmount) as RemainedAmount, SUBSTRING(EffectiveDate,1,8) as EffectiveDate, 'ap' as sys from APItems where FullyApplied=0 and Voided=0 and APItems.Type=6 and EffectiveDate>='1388/01/01' group by SUBSTRING(EffectiveDate,1,8) ) drv order by EffectiveDate,sys"
	
	mySQLold="select sum(RemainedAmount) as RemainedAmount,'old' as EffectiveDate,'ar' as sys from ARItems where FullyApplied=0 and Voided=0 and ARItems.Type=1 and EffectiveDate<'1388/01/01' UNION select sum(RemainedAmount) as RemainedAmount, 'old' as EffectiveDate,'ap' as sys from APItems where FullyApplied=0 and Voided=0 and APItems.Type=6 and EffectiveDate<'1388/01/01'"

	set rs=Conn.Execute(mySQL)
	set rsOLD=conn.Execute(mySQLold)
	If rs.EOF or rsOLD.EOF then
		Conn.Close
		response.redirect "?errMsg=" & Server.URLEncode("���� ����! ���� ���� ���")
	end if
	sub setAXValue(RSs)
		select case RSs("sys")
			case "ap"
				apRemainedAmount = CDbl(RSs("RemainedAmount"))
			case "ar"
				arRemainedAmount = CDbl(RSs("RemainedAmount"))	
		end select
		if not RSs.EOF then RSs.MoveNext
		if not RSs.eof then 
			if EffectiveDate = RSs("EffectiveDate") then
				ref2 = RSs("EffectiveDate")
				call setAXValue(RSs)
			end if
		end if
	end sub

	%>
</table>
<table>
	<tr class="RepTableHeader">
		<td colspan="2">&nbsp;</td>
		<td>����</td>
		<td>����</td>
	
	</tr>
	
	<%
	rowColor="RepTR0"
	do while not rsOLD.eof
		EffectiveDate=rsOLD("EffectiveDate")
		%>
		<tr class="<%=rowColor%>">
			<%
			arRemainedAmount=0
			apRemainedAmount=0			
			call setAXValue(rsOLD)
			%>
			<td colspan="2">����� �����</td>
			<%
			response.write echoTD(arRemainedAmount, "ar")
			response.write echoTD(apRemainedAmount, "ap")
			%>
		</tr>
		<%
			
	loop
	arRemainedAmountSum=0
	apRemainedAmountSum=0
	do while not rs.eof
		EffectiveDate=rs("EffectiveDate")
		if rowColor="RepTR0" then 
			rowColor="RepTR1"
		else
			rowColor="RepTR0"
		end if
		%>
		<tr class="<%=rowColor%>">
			<%
			yyyy=mid(EffectiveDate,1,4)
			mm=fmonth(cint(mid(EffectiveDate,6,2))-1)
			arRemainedAmount=0
			apRemainedAmount=0
			call setAXValue(rs)
			arRemainedAmountSum=arRemainedAmountSum+arRemainedAmount
			apRemainedAmountSum=apRemainedAmountSum+apRemainedAmount
			%>
			<td title="<%=EffectiveDate%>"><%=yyyy%></td>
			<td><%=mm%></td>
			<%
			response.write echoTD(arRemainedAmount, "ar")
			response.write echoTD(apRemainedAmount, "ap")
			%>
		</tr>
		<%
	loop
	if rowColor="RepTR0" then 
		rowColor="RepTR1"
	else
		rowColor="RepTR0"
	end if
	
%>
		<tr class="<%=rowColor%>">
			<td colspan="2">��� ���� �������</td>
			<td><%=Separate(arRemainedAmountSum)%></td>
			<td><%=Separate(apRemainedAmountSum)%></td>
		</tr>
</table>
<%
elseif request("act")="MoeenRep" then

	ON ERROR RESUME NEXT
		GLAccount =		clng(request("GLAccount"))
		FromDate =		sqlSafe(request("FromDate"))
		ToDate = 		sqlSafe(request("ToDate"))

		if FromDate	= "" then FromDate = fiscalYear & "/01/01"
		if ToDate = "" then ToDate = shamsiToday()

		FromTafsil =	clng(request("FromTafsil"))
		ToTafsil =		clng(request("ToTafsil"))


		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "?errMsg=" & Server.URLEncode("���!")
		end if
	ON ERROR GOTO 0

  '-----------------------------------------------------------------------------------------------------
  if request("action")=" �ǁ " then
  '----   It's Print ( Crystal Reports )
  %>
	<BR>
	<BR>
	<CENTER>
	<% 	ReportLogRow = PrepareReport ("MoeenRep01.rpt", "GLAccountGLFromDateToDateFromTafsilToTafsil", GLAccount & "" & OpenGL & "" & FromDate & "" & ToDate & "" & FromTafsil& "" & ToTafsil, "/beta/dialog_printManager.asp?act=Fin") %>
	<INPUT TYPE="button" value=" �ǁ " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

	</CENTER>

	<BR><iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no ></iframe>

  <%
  else
  '----   It's Not Print

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
		order="(SUM(GLRows.IsCredit * GLRows.Amount) - SUM(- ((GLRows.IsCredit - 1) * GLRows.Amount))) DESC"
	case "6","-5":
		order="SUM(GLRows.IsCredit * GLRows.Amount) - SUM(- ((GLRows.IsCredit - 1) * GLRows.Amount))"
	case else:
		order="Tafsil"
		Ord=1
	end select

	mySQL="SELECT GLAccounts.ID, GLAccounts.Name AS AccountName, GLAccountGroups.ID AS GroupID, GLAccountGroups.Name AS GroupName,  GLAccountSuperGroups.ID AS SuperGroupID, GLAccountSuperGroups.Name AS SuperGroupName, GLs.ID AS GLID, GLs.Name AS GLname FROM GLs INNER JOIN GLAccountSuperGroups ON GLs.ID = GLAccountSuperGroups.GL INNER JOIN GLAccountGroups ON GLs.ID = GLAccountGroups.GL AND GLAccountSuperGroups.ID = GLAccountGroups.GLSuperGroup INNER JOIN GLAccounts ON GLs.ID = GLAccounts.GL AND GLAccountGroups.ID = GLAccounts.GLGroup WHERE (GLs.ID = "& OpenGL & ") AND (GLAccounts.ID = "& GLAccount & ")"

	set rsGLNames=Conn.Execute (mySQL)

	If rsGLNames.EOF then
		Conn.Close
		response.redirect "?errMsg=" & Server.URLEncode("���� ���� ��� [" & GLAccount & "] �� ��� ���� ���� (" & OpenGL & ") ���� ����� ")
	End If
	AccountInfoParams = "&DateFrom=" & FromDate & "&DateTo=" & ToDate
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function showAcc(acc){
		window.open('tafsili.asp?accountID='+acc+'&FromDate=<%=FromDate%>&ToDate=<%=ToDate%>&moeenFrom=0&moeenTo=99999&act=Show');
	}
	//-->
	</SCRIPT>
	<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">
	<TR bgcolor="#CCCCEE" height="30">
		<TD colspan=7>
			<A HREF="AccountInfo.asp?OpenGL=<%=rsGLNames("GLID")&AccountInfoParams%>" Target="_blank"><%=rsGLNames("GLname")%></A>
			> <A HREF="AccountInfo.asp?act=groups&id=<%=rsGLNames("SuperGroupID")&AccountInfoParams%>" Target="_blank"><%=rsGLNames("SuperGroupName")%></A>
			> <A HREF="AccountInfo.asp?act=account&id=<%=rsGLNames("GroupID")&AccountInfoParams%>" Target="_blank"><%=rsGLNames("GroupName")%></A>
			> <%=rsGLNames("AccountName")%>
			[<%=GLAccount%>]
		</TD>
	</TR>
<%
	rsGLNames.close
	Set rsGLNames = Nothing

	mySQL = "SELECT SUM(SumCred) AS SumCred, SUM(SumDeb) AS SumDeb, SUM(Flow * (Sgn + 1) / 2) AS sumFlowCred, SUM(Flow * (Sgn - 1) / 2) AS sumFlowdeb FROM (SELECT SUM(IsCredit * Amount) AS SumCred, SUM(- ((IsCredit - 1) * Amount)) AS SumDeb, SUM(IsCredit * Amount) - SUM(- ((IsCredit - 1)  * Amount)) AS Flow, SIGN(SUM(IsCredit * Amount) - SUM(- ((IsCredit - 1) * Amount))) AS Sgn FROM EffectiveGLRows WHERE (GLAccount = "& GLAccount & ") AND (GL = "& OpenGL & ") AND (ISNULL(Tafsil, 0) >= "& FromTafsil & ") AND (ISNULL(Tafsil, 0) <= "& ToTafsil & ") AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "') GROUP BY Tafsil) FlowTbl"

	set rs=Conn.Execute (mySQL)

%>
	<TR bgcolor="#CCCCEE" >
		<TD colspan=2 rowspan=2>
			�� ����� <B><%=replace(FromDate,"/",".")%></B> �� ����� <B><%=replace(ToDate,"/",".")%></B><br>
			�� ������ <B><%=FromTafsil%></B> �� ������ <B><%=ToTafsil%></B>
		</TD>
		<TD width=70 >���� ������		</TD>
		<TD width=70 >���� ��������		</TD>
		<TD width=70 >����� ������		</TD>
		<TD width=70 >����� ��������	</TD>
	</TR>
	<TR bgcolor="#CCCCEE" >
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
	<TR bgcolor="eeeeee" style="cursor:hand;" title="����� �����">
		<TD width=50  onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>">������			<%if abs(ord)=1 then response.write arrow%></TD>
		<TD width='*' onclick='go2Page(1,2);' style="<%if abs(ord)=2 then response.write style%>">����� ����		<%if abs(ord)=2 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-3);' style="<%if abs(ord)=3 then response.write style%>">���� ������		<%if abs(ord)=3 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-4);' style="<%if abs(ord)=4 then response.write style%>">���� ��������		<%if abs(ord)=4 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-5);' style="<%if abs(ord)=5 then response.write style%>">����� ������		<%if abs(ord)=5 then response.write arrow%></TD>
		<TD width=70 onclick='go2Page(1,-6);' style="<%if abs(ord)=6 then response.write style%>">����� ��������	<%if abs(ord)=6 then response.write arrow%></TD>
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

	mySQL="SELECT GLRows.Tafsil, SUM(GLRows.IsCredit * GLRows.Amount) AS totalCredit, SUM(- ((GLRows.IsCredit - 1) * GLRows.Amount)) AS totalDebit, Accounts.AccountTitle AS AccountTitle FROM (SELECT ID AS GLDoc, GLDocDate FROM GLDocs WHERE (GLDocs.IsTemporary = 1 OR GLDocs.IsChecked = 1 OR GLDocs.IsFinalized = 1) AND (GLDocDate >= N'"& FromDate & "') AND (GLDocDate <= N'"& ToDate & "') AND (GL = "& openGL & " ) AND (IsRemoved = 0) AND (deleted = 0)) EffectiveGLDocs INNER JOIN GLRows ON EffectiveGLDocs.GLDoc = GLRows.GLDoc INNER JOIN Accounts ON GLRows.Tafsil = Accounts.ID WHERE (GLRows.GLAccount = "& GLAccount & ") AND (GLRows.deleted = 0) GROUP BY GLRows.Tafsil, Accounts.AccountTitle HAVING (ISNULL(GLRows.Tafsil, 0) >= "& FromTafsil & ") AND (ISNULL(GLRows.Tafsil, 0) <= "& ToTafsil & ") ORDER BY " & order

	Set rs=Server.CreateObject("ADODB.Recordset")'Conn.Execute(mySQL)

	PageSize = 50
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
			<td bgcolor="#BBBBBB" height="30" colspan="7" align=center><b>�� .</b></td>
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
						<b>���� <%=CurrentPage%> �� <%=TotalPages%></b>
						&nbsp;&nbsp;<a href="javascript:go2Page(<%=CurrentPage+1%>,0);">���� ��� &gt;</a>
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
			str='?act=MoeenRep&GLAccount='+escape('<%=GLAccount%>')+'&FromDate='+escape('<%=FromDate%>')+'&ToDate='+escape('<%=ToDate%>')+'&FromTafsil='+escape('<%=FromTafsil%>')+'&ToTafsil='+escape('<%=ToTafsil%>')+'&Ord='+escape(ord)+'&p='+escape(p)  //+'& ='+escape(' ')+'& ='+escape(' ')+'& ='+escape(' ')
			window.location=str;
		}
		//-->
		</SCRIPT>
<%
	end if
  end if
  '-----------------------------------------------------------------------------------------------------
end if
%>

<!--#include file="tah.asp" -->