<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle="”«Ì— ê“«—‘ Â«"
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
	.RepTDb {font-weight: bold;}
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
			<th colspan="2">ê“«—‘ „⁄Ì‰</td>
		</tr>
		<tr>
			<td align=left>„⁄Ì‰</td>
			<td><INPUT TYPE="text" NAME="GLAccount" style="width:50px;" maxlength=5></td>
		</tr>
		<tr>
			<td align=left>«“  «—ÌŒ</td>
			<td><INPUT TYPE="text" NAME="FromDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
		</tr>
		<tr>
			<td align=left> «  «—ÌŒ</td>
			<td><INPUT TYPE="text" NAME="ToDate" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
		</tr>
		<tr>
			<td align=left>«“  ›’Ì·</td>
			<td><INPUT TYPE="text" NAME="FromTafsil" style="width:50px;" maxlength=6 value="0"></td>
		</tr>
		<tr>
			<td align=left> «  ›’Ì·</td>
			<td><INPUT TYPE="text" NAME="ToTafsil" style="width:50px;" maxlength=6 value="999999">
				<INPUT TYPE="hidden" NAME="GL" value="<%=OpenGL%>">
			</td>
		</tr>
		</table>
		<div align=center>
		<% 	ReportLogRow = PrepareReport ("Receipt.rpt", "Recept_ID", Receipt, "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT Class="GenButton" TYPE="submit" name="action" value=" ‰„«Ì‘ ">
		<INPUT Class="GenButton" TYPE="submit" name="action" style="border:1 solid green;" value=" ç«Å ">
		</div>
		</FORM>
	</TD>
	<TD width="130" align=center>
	 &nbsp;
	 <A HREF="OtherReports_Tmp4Zamani.asp" style='font-weight:bold;'>„«‰œÂ Õ”«» Â«Ì „‘ —Ì«‰ œ— ”Ì” „ Ã«„⁄ œ— Å«Ì«‰ ”«· 82</A>
	</TD>
</TR>
<TR>
	<TD width="130">
	<%
	if Auth(8 , "G") then
	%>
		<FORM METHOD=POST ACTION="?act=cash">
			<table class="RepTable2">
				<tr>
					<th colspan="2">œ‘»Ê—œ Œ“«‰Âùœ«—Ì</td>
				</tr>
				
			</table>		
			 <div align="center">
			 	
				 	<INPUT Class="GenButton" TYPE="submit" name="action" value=" ‰„«Ì‘ ">
				
			 </div>
	 	</form>
	 <%end if%>
	</TD>
	<TD>
	 <%
	 if Auth(8,"G") then 
	 	set rs=Conn.Execute ("select * from gls where id="&openGL)
	 	if not rs.eof then 
	 		startDate=rs("startDate")
	 		if shamsiToday()>rs("endDate") then 
	 			endDate=rs("endDate")
	 		else
	 			endDate=shamsiToday()
	 		end if
	 	end if
	 %>
	 	<form method="post" action="?act=finState">
	 		<table class="RepTable2">
	 			<tr>
	 				<th colspan="2">’Ê— ùÂ«Ì „«·Ì</th>
	 			<tr>
					<td align=left>«“  «—ÌŒ</td>
					<td><INPUT TYPE="text" NAME="FromDate" value="<%=startDate%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
				</tr>
				<tr>
					<td align=left> «  «—ÌŒ</td>
					<td><INPUT TYPE="text" NAME="ToDate" value="<%=endDate%>" style="width:75px;direction:LTR;" maxlength=10 OnBlur="return acceptDate(this);"></td>
				</tr>
	 				<td colspan="2">
	 					<input name="rep" type="submit" value="”Êœ Ê “Ì«‰">
	 					<input name="rep" type="submit" value=" —«“ ‰«„Â">
	 					<input name="rep" type="submit" value="ê—œ‘ ÊÃÊÂ ‰ﬁœ">
	 				</td>
	 			</tr>
	 		</table>
	 	</form>
	 <%
	 end if
	 %>
	</TD>
</TR>
<TR>
	<TD width="130">
	<%
	if Auth(8,"G") then 
	%>
		<table class="RepTable2">
			<tr>
				<td colspan="2">ê“«—‘  Ã“ÌÂ «”‰«œ</td>
			</tr>
			<tr>
				<td><a href="otherReports.asp?act=sepPay">«”‰«œ Å—œ«Œ ‰Ì</a></td>
				<td><a href="otherReports.asp?act=sepRec">«”‰«œ œ—Ì«› ‰Ì</a></td>
			</tr>
		</table>
	<%	
	end if
	%>
	</td>
	<td></td>
</TABLE>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="sepPay" or request("act")="sepRec" then 
	if request("act")="sepPay" then 
		mySQL="select g42.*, case when g41.glAccount=41001 then N' Ã«—Ì' else N'€Ì—  Ã«—Ì' end as state from EffectiveGLRows as g42 inner join EffectiveGLRows as g41 on g41.glDoc=g42.glDoc and g41.sys=g42.sys and g41.link=g42.link where g42.GL=" & openGL & " and g42.glAccount between 42000 and 42999 and g41.glAccount between 41000 and 41999"
	else
		mySQL="select g17.*, case when g13.glAccount=13003 then N' Ã«—Ì' else N'€Ì—  Ã«—Ì' end as state from EffectiveGLRows as g17 inner join EffectiveGLRows as g13 on g17.glDoc=g13.glDoc and g17.sys=g13.sys and g17.link=g13.link where g17.GL=" & openGL & " and g17.glAccount between 17000 and 17999 and g17.ref1<>'' and not g13.glAccount between 17000 and 17999 "
	end if
	set rs=Conn.Execute(mySQL)
%>
<table class="RepTable">
	<tr class="RepTableHeader">
		<th>‘„«—Â ”‰œ</th>
		<th> «—ÌŒ ”‰œ</th>
		<th>„⁄Ì‰</th>
		<th>„»·€</th>
		<th>‘—Õ</th>
		<th>‘„«—Â çﬂ</th>
		<th> «—ÌŒ çﬂ</th>
		<th>Ê÷⁄Ì </th>
	</tr>
<%
	rowColor="RepTR0"
	while not rs.eof
		theLink = "ShowItem.asp?sys=" & rs("sys") & "&Item=" & rs("Link")
		if rowColor="RepTR0" then 
			rowColor="RepTR1"
		else
			rowColor="RepTR0"
		end if
	%>
	<tr class="<%=rowColor%>">
		<td style="cursor:hand;" title="»—«Ì ‰„«Ì‘ ”‰œ Õ”«»œ«—Ì „—»ÊÿÂ ﬂ·Ìﬂ ﬂ‰Ìœ." onclick="window.open('GLMemoDocShow.asp?id=<%=rs("glDoc")%>');"><%=rs("glDocID")%></td>
		<td><%=rs("glDocDate")%></td>
		<td><%=rs("glAccount")%></td>
		<td><%=Separate(rs("amount"))%></td>
		<td style="cursor:hand;" title="»—«Ì ‰„«Ì‘ ”‰œ “Ì— ”Ì” „ „—»ÊÿÂ ﬂ·Ìﬂ ﬂ‰Ìœ." onclick="window.open('<%=theLink%>');"><%=rs("description")%></td>
		<td style="cursor:hand;" title="»—«Ì ‰„«Ì‘ çﬂ „—»ÊÿÂ ﬂ·Ìﬂ ﬂ‰Ìœ." onclick="window.open('../bank/CheqBook.asp?act=findCheq&cheque=<%=rs("ref1")%>&amount=<%=rs("amount")%>');"><%=rs("ref1")%></td>
		<td><%=rs("ref2")%></td>
		<td><%=rs("state")%></td>

	</tr>
	<%
		rs.moveNext
	wend
%>
</table>
<%
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="finState" then
	fromDate=request("fromDate")
	toDate=request("toDate")
	function getRem(v,t)
		if t="group" then 
			mySQL="select isnull(sum(effectiveGlRows.Amount*2*(cast(effectiveGlRows.IsCredit as int)-.5)),0) as amount from effectiveGlRows inner join glAccounts on effectiveGlRows.glAccount=glAccounts.id and effectiveGlRows.gl=glAccounts.gl where effectiveGlRows.gl=" & OpenGL & " and glAccounts.GLGroup=" & v & " and effectiveGLRows.glDocDate between '" & fromDate & "' and '" & toDate & "'"
		elseif t="super" then 
			mySQL="select isnull(sum(effectiveGlRows.Amount*2*(cast(effectiveGlRows.IsCredit as int)-.5)),0) as amount from effectiveGlRows inner join glAccounts on effectiveGlRows.glAccount=glAccounts.id and effectiveGlRows.gl=glAccounts.gl inner join GLAccountGroups on glAccounts.GLGroup=GLAccountGroups.ID and GLAccountGroups.gl=glAccounts.gl where effectiveGlRows.gl=" & OpenGL & " and GLAccountGroups.GLSuperGroup=" & v & " and effectiveGLRows.glDocDate between '" & fromDate & "' and '" & toDate & "'"
		end if
		set rs=Conn.Execute(mySQL)
		if not rs.eof then 
			getRem = CDbl(rs("amount"))
		else
			getRem = "Œÿ«"
		end if
	end function
	if request("rep")="”Êœ Ê “Ì«‰" then
		sales = getRem(91000,"group")
		cogs = abs(getRem(30000,"super"))
		grossProfit = sales - cogs
		cgA = abs(getRem(62000,"group")+getRem(63000,"group")+getRem(68000,"group")+getRem(69000,"group"))
		rANDd = 0
		other = 0
		Interest = abs(getRem(61000,"group"))
		netProfit = grossProfit - cgA - rANDd - other - interest
		tax = 0
		Depreciation = 0
		EBITDA = netProfit + interest + tax + depreciation
%>
<table class="RepTable">
	<tr>
		<th> —«“‰«„Â</th>
		<th>«“ <%=fromDate%> «·Ì <%=toDate%></th>
	</tr>
	<tr>
		<td> ›—Ê‘ù Œ«·’</td>
		<td><%=Separate(sales)%></td>
	</tr>
	<tr>
		<td> »Â«Ìù  „«„ù‘œÂù ﬂ«·«Ìù ›—Ê‘ù—› Âù</td>
		<td><%=Separate(cogs)%></td>
	</tr>
	<tr>
		<td>”Êœ (“Ì«‰ù) ‰«Œ«·’</td>
		<td><%=Separate(grossProfit)%></td>
	</tr>
	<tr>
		<td> ›—Ê‘ù° «œ«—Ìù Ê ⁄„Ê„Ìù</td>
		<td><%=Separate(cgA)%></td>
	</tr>
	<tr>
		<td> ÕﬁÌﬁ Ê  Ê”⁄Â</td>
		<td><%=Separate(rANDd)%></td>
	</tr>
	<tr>
		<td>Â“Ì‰Â (œ—¬„œ) »Â—Â</td>
		<td><%=Separate(interest)%></td>
	</tr>
	<tr>
		<td>”«Ì— (⁄‰«ÊÌ‰ «’·Ì –ﬂ— ‘Êœ)</td>
		<td><%=Separate(other)%></td>
	</tr>
	<tr>
		<td>”Êœ (“Ì«‰ù) Œ«·’ù</td>
		<td><%=Separate(netProfit)%></td>
	</tr>
	<tr>
		<td>Â“Ì‰ÂùÂ«Ì »Â—Â</td>
		<td><%=Separate(interest)%></td>
	</tr>
	<tr>
		<td>„«·Ì« </td>
		<td><%=Separate(tax)%></td>
	</tr>
	<tr>
		<td>«” Â·«ﬂ</td>
		<td><%=Separate(depreciation)%></td>
	</tr>
	<tr>
		<td>«»Ì œ«</td>
		<td><%=Separate(EBITDA)%></td>
	</tr>
</table>
<%	
	elseif request("rep")=" —«“ ‰«„Â" then
		'mySQL="select sum(g42.amount) as amount from EffectiveGLRows as g42 inner join EffectiveGLRows as g41 on g41.glDoc=g42.glDoc and g41.sys=g42.sys and g41.link=g42.link where g42.GL=" & openGL & " and g42.glAccount between 42000 and 42999 and g41.glAccount between 41000 and 41999 and g42.isCredit=1 and g41.glAccount=41001"
		mySQL="select isnull(sum(amount),0) as amount from (select sum(amount)/count(id) as amount from EffectiveGLRows inner join (select g42.ref1,g42.ref2 from EffectiveGLRows as g42 inner join EffectiveGLRows as g41 on g41.glDoc=g42.glDoc and g41.sys=g42.sys and g41.link=g42.link where g42.GL=" & openGL & " and g42.glAccount between 42000 and 42999 and g41.glAccount between 41000 and 41999 and g42.isCredit=1 and g41.glAccount=41001 and g42.ref1<>'' and g42.glDocDate between '" & fromDate & "' and '" & toDate & "') as drv on EffectiveGLRows.Ref1=drv.ref1 and EffectiveGLRows.ref2=drv.ref2 where EffectiveGLRows.gl=" & openGL & " and glAccount between 42000 and 42999 and glDocDate between '" & fromDate & "' and '" & toDate & "' group by GLAccount, Tafsil, Amount,EffectiveGLRows.ref1,EffectiveGLRows.ref2 having count(EffectiveGLRows.ref1) % 2 =1) as d"
		rs=conn.Execute(mySQL)
		accountsPayableTej = cdbl(rs("amount"))
		'mySQL="select sum(g42.amount) as amount from EffectiveGLRows as g42 inner join EffectiveGLRows as g41 on g41.glDoc=g42.glDoc and g41.sys=g42.sys and g41.link=g42.link where g42.GL=" & openGL & " and g42.glAccount between 42000 and 42999 and g41.glAccount between 41000 and 41999 and g42.isCredit=1 and g41.glAccount<>41001"
		mySQL="select isnull(sum(amount),0) as amount from (select sum(amount)/count(id) as amount from EffectiveGLRows inner join (select g42.ref1,g42.ref2 from EffectiveGLRows as g42 inner join EffectiveGLRows as g41 on g41.glDoc=g42.glDoc and g41.sys=g42.sys and g41.link=g42.link where g42.GL=" & openGL & " and g42.glAccount between 42000 and 42999 and g41.glAccount between 41000 and 41999 and g42.isCredit=1 and g41.glAccount<>41001 and g42.ref1<>'' and g42.glDocDate between '" & fromDate & "' and '" & toDate & "') as drv on EffectiveGLRows.Ref1=drv.ref1 and EffectiveGLRows.ref2=drv.ref2 where EffectiveGLRows.gl=" & openGL & " and glAccount between 42000 and 42999 and glDocDate between '" & fromDate & "' and '" & toDate & "' group by GLAccount, Tafsil, Amount,EffectiveGLRows.ref1,EffectiveGLRows.ref2 having count(EffectiveGLRows.ref1) % 2 =1) as d"
		rs=conn.Execute(mySQL)
		accountsPayableNonTej = cdbl(rs("amount"))
		'mySQL="select sum(g17.amount) as amount from EffectiveGLRows as g17 inner join EffectiveGLRows as g13 on g17.glDoc=g13.glDoc and g17.sys=g13.sys and g17.link=g13.link where g17.GL=" & openGL & " and g17.glAccount between 17000 and 17999 and g17.ref1<>'' and not g13.glAccount between 17000 and 17999 and g17.isCredit=0 and g13.glAccount <> 13003"
		mySQL="select isnull(sum(amount),0) as amount from (select sum(amount)/count(id) as amount from EffectiveGLRows inner join (select g17.ref1,g17.ref2 from EffectiveGLRows as g17 inner join EffectiveGLRows as g13 on g17.glDoc=g13.glDoc and g17.sys=g13.sys and g17.link=g13.link where g17.GL=" & openGL & " and g17.glAccount between 17000 and 17999 and g17.ref1<>'' and not g13.glAccount between 17000 and 17999 and g17.isCredit=0 and g13.glAccount <> 13003 and g17.glDocDate between '" & fromDate & "' and '" & toDate & "') as drv on EffectiveGLRows.Ref1=drv.ref1 and EffectiveGLRows.ref2=drv.ref2 where EffectiveGLRows.gl=" & openGL & " and glAccount between 17000 and 17999 and glDocDate between '" & fromDate & "' and '" & toDate & "' group by GLAccount, Tafsil, Amount,EffectiveGLRows.ref1,EffectiveGLRows.ref2 having count(EffectiveGLRows.ref1) % 2 =1) as d"
		rs=conn.Execute(mySQL) 
		accountsReceviableNonTej = cdbl(rs("amount"))
		'mySQL="select sum(g17.amount) as amount from EffectiveGLRows as g17 inner join EffectiveGLRows as g13 on g17.glDoc=g13.glDoc and g17.sys=g13.sys and g17.link=g13.link where g17.GL=" & openGL & " and g17.glAccount between 17000 and 17999 and g17.ref1<>'' and not g13.glAccount=13003 and g17.isCredit=0"
		mySQL="select isnull(sum(amount),0) as amount from (select sum(amount)/count(id) as amount from EffectiveGLRows inner join (select g17.ref1,g17.ref2 from EffectiveGLRows as g17 inner join EffectiveGLRows as g13 on g17.glDoc=g13.glDoc and g17.sys=g13.sys and g17.link=g13.link where g17.GL=" & openGL & " and g17.glAccount between 17000 and 17999 and g17.ref1<>'' and not g13.glAccount=13003 and g17.isCredit=0 and g17.glDocDate between '" & fromDate & "' and '" & toDate & "') as drv on EffectiveGLRows.Ref1=drv.ref1 and EffectiveGLRows.ref2=drv.ref2 where EffectiveGLRows.gl=" & openGL & " and glAccount between 17000 and 17999 and glDocDate between '" & fromDate & "' and '" & toDate & "' group by GLAccount, Tafsil, Amount,EffectiveGLRows.ref1,EffectiveGLRows.ref2 having count(EffectiveGLRows.ref1) % 2 =1) as d"
		rs=conn.Execute(mySQL)
		accountsReceviableTej = cdbl(rs("amount"))
		
		cash=abs(getRem(11000,"group")) + abs(getRem(12000,"group"))
		'accountsReceviable = abs(getRem(13000,"group")) + abs(getRem(15000,"group")) + abs(getRem(17000,"group")) + abs(getRem(18000,"group"))
		inventory = abs(getRem(14000,"group"))
		orders = abs(getRem(16000,"group"))
		plant = abs(getRem(26000,"group"))-abs(getRem(26100,"group"))
		properties = abs(getRem(27000,"group"))-abs(getRem(27100,"group"))
		otherAsset = abs(getRem(29000,"group"))
		totalAsset = cash + accountsReceviableTej + accountsReceviableNonTej + inventory + orders + plant + properties + otherAsset
		'accountsPayable = abs(getRem(41000,"group"))+abs(getRem(42000,"group"))
		bankDebit = abs(getRem(46000,"group"))
		personsDebit = abs(getRem(44000,"group"))
		otherDebit =  abs(getRem(43000,"group")) + abs(getRem(45000,"group")) + abs(getRem(47000,"group")) + abs(getRem(49000,"group")) + abs(getRem(71000,"group"))
		totalDebit = accountsPayableTej + accountsPayableNonTej + bankDebit + personsDebit + otherDebit
		equity = abs(getRem(50100,"group"))
		retainedEarnings = getRem(51000,"group") + getRem(52000,"group") + getRem(53000,"group") + getRem(54000,"group") + getRem(55000,"group") + getRem(59000,"group")
		periodEarnings = 0
		totalLiabilities = totalDebit + equity + retainedEarnings
%>
<table class="RepTable">
	<tr>
		<th> —«“‰«„Â</th>
		<th><%=toDate%></th>
	</tr>
	<tr>
		<td>„ÊÃÊœÌù ‰ﬁœ</td>
		<td><%=Separate(cash)%></td>
	</tr>
	<tr>
		<td>Õ”«»Â« Ê «”‰«œ œ—Ì«› ‰Ìù  Ã«—Ìù</td>
		<td><%=Separate(accountsReceviableTej)%></td>
	</tr>
	<tr>
		<td>Õ”«»Â« Ê «”‰«œ œ—Ì«› ‰Ìù €Ì—  Ã«—Ì</td>
		<td><%=Separate(accountsReceviableNonTej)%></td>
	</tr>
	<tr>
		<td>„ÊÃÊœÌ «‰»«— „Ê«œ «Ê·ÌÂ Ê ﬂ«·«Ì ”«Œ Â ‘œÂ</td>
		<td><%=Separate(inventory)%></td>
	</tr>
	<tr>
		<td>”›«—‘«  Ê ÅÌ‘ Å—œ«Œ Â«</td>
		<td><%=Separate(orders)%></td>
	</tr>
	<tr>
		<td>œ«—«ÌÌÂ«Ìù À«» ù „‘ÂÊœ</td>
		<td><%=Separate(plant)%></td>
	</tr>
	<tr>
		<td>œ«—«ÌÌÂ«Ìù ‰« „‘ÂÊœ</td>
		<td><%=Separate(properties)%></td>
	</tr>
	<tr>
		<td>Õ’Â ”‰Ê«  ¬ Ì Ê«„ »«‰ﬂÌ</td>
		<td><%=Separate(otherAsset)%></td>
	</tr>
	<tr>
		<td class="RepTDb">Ã„⁄ù œ«—«ÌÌÂ«</td>
		<td><%=Separate(totalAsset)%></td>
	</tr>
	<tr>
		<td>Õ”«»Â« Ê «”‰«œ Å—œ«Œ ‰Ìù  Ã«—Ìù</td>
		<td><%=Separate(accountsPayableTej)%></td>
	</tr>
	<tr>
		<td>Õ”«»Â« Ê «”‰«œ Å—œ«Œ ‰Ìù €Ì—  Ã«—Ìù</td>
		<td><%=Separate(accountsPayableNonTej)%></td>
	</tr>
	<tr>
		<td> ”ÂÌ·« ù „«·Ì »«‰ﬂÌ</td>
		<td><%=Separate(bankDebit)%></td>
	</tr>
	<tr>
		<td> ”ÂÌ·«  „«·Ì «‘Œ«’</td>
		<td><%=Separate(personsDebit)%></td>
	</tr>
	<tr>
		<td>”«Ì— »œÂÌùÂ« (⁄‰«ÊÌ‰ «’·Ì –ﬂ— ‘Êœ(</td>
		<td><%=Separate(otherDebit)%></td>
	</tr>
	<tr>
		<td class="RepTDb">Ã„⁄ù »œÂÌÂ«</td>
		<td><%=Separate(totalDebit)%></td>
	</tr>
	<tr>
		<td>”—„«ÌÂù</td>
		<td><%=Separate(equity)%></td>
	</tr>
	<tr>
		<td>«‰»«‘ Âù ( ”Êœ (“Ì«‰ù</td>
		<td><%=Separate(retainedEarnings)%></td>
	</tr>
	<tr>
		<td>ù”Êœ Ê “Ì«‰ œÊ—Â Ã«—Ì</td>
		<td></td>
	</tr>
	<tr>
		<td class="RepTDb">ùÃ„⁄ù »œÂÌÂ« Ê ÕﬁÊﬁù ’«Õ»«‰ù ”Â«„ù</td>
		<td><%=Separate(totalLiabilities)%></td>
	</tr>
</table>
<%
	elseif request("rep")="ê—œ‘ ÊÃÊÂ ‰ﬁœ" then 
		response.write "Â‰Ê“ ”«Œ Â ‰‘œÂ"
	end if
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
elseif request("act")="cash" then 
	myMonth = replace(request("month"),"-","/")
	dim fmonth(12)
	fmonth(0)="›—Ê—œÌ‰"
	fmonth(1)="«—œÌ»Â‘ "
	fmonth(2)="Œ—œ«œ"
	fmonth(3)=" Ì—"
	fmonth(4)="„—œ«œ"
	fmonth(5)="‘Â—ÌÊ—"
	fmonth(6)="„Â—"
	fmonth(7)="¬»«‰"
	fmonth(8)="¬–—"
	fmonth(9)="œÌ"
	fmonth(10)="»Â„‰"
	fmonth(11)="«”›‰œ"
	
	'----------------------------------------------------------------------------------------
	
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
			result = result & "<a target=î_blankî href=""../AR/rep_dailySale.asp?input_date_start="
			if EffectiveDate<>"old" then 
				result = result & Server.URLEncode(EffectiveDate&"01")&"&input_date_end=" & Server.URLEncode(EffectiveDate&"31")&"&fullyApplied=on"" >" 
			else
				result = result & Server.URLEncode("0000/00/00")&"&input_date_end=" & Server.URLEncode("1388/01/01") &"&fullyApplied=on"">" 
			end if
		case "ap" '------------ AP Items
			result = result & "<a target=î_blankî href=""../AP/report.asp?fromDate="
			if EffectiveDate<>"old" then 
				if myMonth<>"" then 
					result = result & Server.URLEncode(EffectiveDate)&"&toDate=" & Server.URLEncode(EffectiveDate)&"&vouchers=on&payments=on&Effective=on"">"
				else
					result = result & Server.URLEncode(EffectiveDate&"01")&"&toDate=" & Server.URLEncode(EffectiveDate&"31")&"&vouchers=on&payments=on&Effective=on"">" 
				end if
			else
				result = result & Server.URLEncode(shamsiDate(DateAdd("m",-10,Date)))&"&toDate=" & Server.URLEncode(mid(shamsiDate(DateAdd("m",-6,Date)),1,8)) & "31" &"&vouchers=on&payments=on"">" 
			end if
		case "ar-full"
			result = result & "<a target=î_blankî href=""../AR/rep_dailySale.asp?input_date_start="
			if EffectiveDate<>"old" then 
				result = result & Server.URLEncode(EffectiveDate&"01")&"&input_date_end=" & Server.URLEncode(EffectiveDate&"31")&"&fullyApplied="" >" 
			else
				result = result & Server.URLEncode("0000/00/00")&"&input_date_end=" & Server.URLEncode("1388/01/01") &"&fullyApplied="">" 
			end if
		case "ap-full"
			result = result & "<a target=î_blankî href=""../AP/report.asp?fromDate="
			if EffectiveDate<>"old" then
				if myMonth<>"" then 
					result = result & Server.URLEncode(EffectiveDate)&"&toDate=" & Server.URLEncode(EffectiveDate)&"&vouchers=on&payments=on&Effective=on"">" 
				else 
					result = result & Server.URLEncode(EffectiveDate&"01")&"&toDate=" & Server.URLEncode(EffectiveDate&"31")&"&vouchers=on&payments=on&Effective=on"">" 
				end if
			else
				result = result & Server.URLEncode(shamsiDate(DateAdd("m",-10,Date)))&"&toDate=" & Server.URLEncode(mid(shamsiDate(DateAdd("m",-6,Date)),1,8)) & "31" &"&vouchers=on&payments=on"">" 
			end if
		case "cash"
			if myMonth<>"" then 
				result = result & "<a target=î_blankî href=""AccountInfo.asp?act=account&id=11000&fromDate=" & Server.URLEncode(ref2) & "&DateTo=" & Server.URLEncode(ref2) & """>"
			else
				result = result & "<a target=î_blankî href=""AccountInfo.asp?act=account&id=11000&fromDate=" & Server.URLEncode(ref2&"01") & "&DateTo=" & Server.URLEncode(ref2&"31") & """>"
			end if
		case "bnk"
			if myMonth<>"" then 
				result = result & "<a target=î_blankî href=""AccountInfo.asp?act=account&id=12000&fromDate=" & Server.URLEncode(ref2) & "&DateTo=" & Server.URLEncode(ref2) & """>"
			else
				result = result & "<a target=î_blankî href=""AccountInfo.asp?act=account&id=12000&fromDate=" & Server.URLEncode(ref2&"01") & "&DateTo=" & Server.URLEncode(ref2&"31") & """>"
			end if
		case else
			if mid(sys,1,4)="bank" then 
				if myMonth<>"" then 
					result = result & "<a  target=î_blankî href=""../bank/CheqBook.asp?act=showBook&GLAccount=" & mid(sys,5,5) & "&FromDate="
					result = result & Server.URLEncode(Ref2)&"&ToDate=" & Server.URLEncode(Ref2)&"&ShowRemained=&displayMode=0"">" 		
				else
					result = result & "<a target=î_blankî href=""../bank/CheqBook.asp?act=showBook&GLAccount=" & mid(sys,5,5) & "&FromDate="
					result = result & Server.URLEncode(Ref2&"01")&"&ToDate=" & Server.URLEncode(Ref2&"31")&"&ShowRemained=&displayMode=0"">" 
				end if
			end if
		end select
		result = result & Separate(num) 
		if sys="ar" or sys="ar-full" or sys="ap" or sys="ap-full" or sys="cash" or sys="bnk" then result = result & "</a>"
		result = result & "</td>"
		echoTD = result
	end function
	if myMonth<>"" then 
		mySQL="select * from (select sum(amount) as amount,ref2,glAccount from (select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17003 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17003 group by Ref2,glAccount union select sum(amount) as amount,ref2,glAccount from (select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17004 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17004 group by Ref2,glAccount union select sum(amount) as amount,ref2,glAccount from (select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17016 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17016 group by Ref2,glAccount union select sum(amount) as amount,ref2,glAccount from (select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42001 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42001 group by Ref2,glAccount union select sum(amount) as amount,ref2,glAccount from (select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42004 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42004 group by Ref2,glAccount union select sum(amount) as amount,ref2,glAccount from (select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42011 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42011 group by Ref2,glAccount union select sum(Amount) as Amount, GLDocDate collate SQL_Latin1_General_CP1_CI_AS as ref2,11000 as GLAccount from EffectiveGLRows where gl=" & OpenGL & " and SUBSTRING(GLDocDate,1,8)='" & myMonth & "' and isCredit=0 and glAccount in (11004,11005,11006,11007) group by GLDocDate union select sum(2*(.5-cast(isCredit as int))*amount) as amount, d.ref2 collate SQL_Latin1_General_CP1_CI_AS as ref2,12000 as GLAccount from EffectiveGLRows inner join (select GLDocDate as ref2 from EffectiveGLRows where GL=" & OpenGL & " and glAccount between 12000 and 12999 and  SUBSTRING(GLDocDate,1,8)='" & myMonth & "' group by GLDocDate) d on GLDocDate<=d.ref2 where GL=" & OpenGL & " and glAccount between 12000 and 12999 group by d.ref2) dev order by Ref2,GLAccount"
		
		'"select * from (select Amount,Ref2,ref1,GLAccount,max(description) as description from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17003 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) union select Amount,Ref2,ref1,GLAccount,max(description) as description from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17004 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) union select Amount,Ref2,ref1,GLAccount,max(description) as description from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42001 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) union select Amount,Ref2,ref1,GLAccount,max(description) as description from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42004 and Ref1<>'' and SUBSTRING(Ref2,1,8)='" & myMonth & "' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) dev order by Ref2,GLAccount"
	else
		mySQL="select * from (select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17003 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17003 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17004 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17004 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=17016 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv17016 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42001 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42001 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42004 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42004 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(Ref2,1,8) as Ref2,GLAccount from ( select Amount,Ref2,GLAccount from EffectiveGLRows where GL=" & OpenGL & " and GLAccount=42011 and Ref1<>'' group by GLAccount, Tafsil, Amount, Ref1, Ref2, GL HAVING (COUNT(Ref1) % 2 = 1) ) drv42011 group by SUBSTRING(Ref2,1,8),GLAccount union select sum(Amount) as Amount, SUBSTRING(GLDocDate,1,8) collate SQL_Latin1_General_CP1_CI_AS as ref2,11000 as GLAccount from EffectiveGLRows where gl=" & OpenGL & " and isCredit=0 and glAccount in (11004,11005,11006,11007) group by SUBSTRING(GLDocDate,1,8) union select sum(2*(.5-cast(isCredit as int))*amount) as amount, d.ref2,12000 as GLAccount from EffectiveGLRows inner join (select SUBSTRING(GLDocDate,1,8) collate SQL_Latin1_General_CP1_CI_AS as ref2 from EffectiveGLRows where GL=" & OpenGL & " and glAccount between 12000 and 12999 group by SUBSTRING(GLDocDate,1,8)) d on GLDocDate<d.ref2+'32' where GL=" & OpenGL & " and glAccount between 12000 and 12999 group by d.ref2) drv order by Ref2,GLAccount"
	end if
	'response.write mySQL
	set rs=Conn.Execute(mySQL)
	amount12000=0
		'Conn.Close
		'response.redirect "?errMsg=" & Server.URLEncode("Œÿ«Ì ⁄ÃÌ»! çÌ“Ì ÅÌœ« ‰‘œ")
	
	sub setValue()
		select case rs("GLAccount")
			case "17003":
				amount17003 = CDbl(rs("Amount"))
			case "17004":
				amount17004 = CDbl(rs("Amount"))
			case "17016":
				amount17016 = CDbl(rs("Amount"))
			case "42001":
				amount42001 = CDbl(rs("Amount"))
			case "42004":
				amount42004 = CDbl(rs("Amount"))
			case "42011":
				amount42011 = CDbl(rs("Amount"))	
			case "11000":
				amount11000 = CDbl(rs("Amount"))
			case "12000":
				amount12000 = CDbl(rs("Amount"))		
		end select
		
		rs.MoveNext
		if not rs.eof then 
			if ref2=rs("Ref2") then
				ref2=rs("Ref2")
				call setValue()
			end if
		end if
	end sub
	If not rs.EOF then
	%>
	<table>
		<tr class="RepTableHeader">
			<td rowspan="2" colspan="2">&nbsp;</td>
			<td colspan="3">œ— Ã—Ì«‰ Ê’Ê·</td>
			<td rowspan="2">Ã„⁄ œ— Ã—Ì«‰</td>
			<td colspan="3">«”‰«œ Å—œ«Œ ‰Ì</td>
			<td rowspan="2">Ã„⁄ «”‰«œ Å—œ«Œ ‰Ì</td>
			<td rowspan="2"> —«“ çﬂ</td>
			<td rowspan="2">’‰œÊﬁ</td>
			<td rowspan="2">»«‰ﬂ —Ê“</td>
			<td rowspan="2">»«‰ﬂ ÅÌ‘ù»Ì‰Ì</td>
		</tr>
		<tr class="RepTableHeader">
			<td>17003</td>
			<td>17004</td>
			<td>17016</td>
			<td>42001</td>
			<td>42004</td>
			<td>42011</td>
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
				if myMonth<>"" then 
					mm=ref2
				else
					mm=fmonth(cint(mid(ref2,6,2))-1)
				end if
				amount17003=0
				amount17004=0
				amount17016=0
				amount42001=0
				amount42004=0
				amount42011=0
				amount11000=0
				
				call setValue()
			%>
			<td title="<%=ref2%>"><%=yyyy%></td>
				<%if myMonth<>"" then%>
			<td><%=mm%></td>
				<%else%>
			<td title='»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ'><a href="otherReports.asp?act=cash&month=<%=replace(mid(ref2,1,8),"/","-")%>"><%=mm%></a></td>
			<%
				end if
				response.write echoTD(amount17003,"bank17003")
				response.write echoTD(amount17004,"bank17004")
				response.write echoTD(amount17016,"bank17016")
				response.write echoTD(amount17003+amount17004+amount17016,"")
				response.write echoTD(amount42001,"bank42001")
				response.write echoTD(amount42004,"bank42004")
				response.write echoTD(amount42011,"bank42011")
				response.write echoTD(amount42001+amount42004+amount17016,"")
				response.write echoTD(amount17003+amount17004-(amount42001+amount42004),"")
				response.write echoTD(amount11000,"cash")
				response.write echoTD(amount12000,"bnk")
				response.write echoTD(amount12000+amount17003+amount17004-(amount42001+amount42004),"")
			%>
			
		</tr>
		<%
		loop
	%>
</table>
<%
	end if
	if myMonth<>"" then 
		mySQL="select * from (select RemainedAmount, EffectiveDate, 'ar' as sys from ARItems where FullyApplied=0 and Voided=0 and ARItems.Type=1 and SUBSTRING(EffectiveDate,1,8)='" & myMonth & "' UNION select RemainedAmount, Vouchers.EffectiveDate, 'ap' as sys from APItems inner join Vouchers on apItems.link=Vouchers.id where FullyApplied=0 and apItems.Voided=0 and APItems.Type=6 and SUBSTRING(Vouchers.EffectiveDate,1,8)='" & myMonth & "' UNION select amountOriginal as RemainedAmount, EffectiveDate, 'ar-full' as sys from ARItems where  Voided=0 and ARItems.Type=1 and SUBSTRING(EffectiveDate,1,8)='" & myMonth & "' UNION select amountOriginal as RemainedAmount, EffectiveDate, 'ap-full' as sys from APItems where Voided=0 and APItems.Type=6 and SUBSTRING(EffectiveDate,1,8)='" & myMonth & "' ) dev order by EffectiveDate,sys"
	else
		mySQL="select * from (select sum(RemainedAmount) as RemainedAmount, SUBSTRING(EffectiveDate,1,8) as EffectiveDate, 'ar' as sys from ARItems where FullyApplied=0 and Voided=0 and ARItems.Type=1 and EffectiveDate>='1388/01/01' group by SUBSTRING(EffectiveDate,1,8) UNION select sum(RemainedAmount) as RemainedAmount, SUBSTRING(Vouchers.EffectiveDate,1,8) as EffectiveDate, 'ap' as sys from APItems inner join Vouchers on apItems.link=Vouchers.id where FullyApplied=0 and apItems.Voided=0 and APItems.Type=6 and Vouchers.EffectiveDate>='1388/01/01' group by SUBSTRING(Vouchers.EffectiveDate,1,8) UNION select sum(amountOriginal) as RemainedAmount, SUBSTRING(EffectiveDate,1,8) as EffectiveDate, 'ar-full' as sys from ARItems where  Voided=0 and ARItems.Type=1 and EffectiveDate>='1388/01/01' group by SUBSTRING(EffectiveDate,1,8) UNION select sum(amountOriginal) as RemainedAmount, SUBSTRING(EffectiveDate,1,8) as EffectiveDate, 'ap-full' as sys from APItems where Voided=0 and APItems.Type=6 and EffectiveDate>='1388/01/01' group by SUBSTRING(EffectiveDate,1,8)) dev order by EffectiveDate,sys"
	end if
	mySQLold="select sum(RemainedAmount) as RemainedAmount,'old' as EffectiveDate,'ar' as sys from ARItems where FullyApplied=0 and Voided=0 and ARItems.Type=1 and EffectiveDate<'1388/01/01' UNION select sum(RemainedAmount) as RemainedAmount, 'old' as EffectiveDate,'ap' as sys from APItems inner join Vouchers on apItems.link=Vouchers.id where FullyApplied=0 and apItems.Voided=0 and APItems.Type=6 and Vouchers.EffectiveDate<'1388/01/01' UNION select sum(amountOriginal) as RemainedAmount,'old' as EffectiveDate,'ar-full' as sys from ARItems where Voided=0 and ARItems.Type=1 and EffectiveDate<'1388/01/01' UNION select sum(amountOriginal) as RemainedAmount, 'old' as EffectiveDate,'ap-full' as sys from APItems where Voided=0 and APItems.Type=6 and EffectiveDate<'1388/01/01'"

	set rs=Conn.Execute(mySQL)
	set rsOLD=conn.Execute(mySQLold)
	
		'Conn.Close
		'response.redirect "?errMsg=" & Server.URLEncode("Œÿ«Ì ⁄ÃÌ»! çÌ“Ì ÅÌœ« ‰‘œ")
	
	sub setAXValue(RSs)
		select case RSs("sys")
			case "ap"
				apRemainedAmount = CDbl(RSs("RemainedAmount"))
			case "ar"
				arRemainedAmount = CDbl(RSs("RemainedAmount"))	
			case "ap-full"
				apTotal = CDbl(RSs("RemainedAmount"))
				case "ar-full"
				arTotal = CDbl(RSs("RemainedAmount"))	
		end select
		if not RSs.EOF then RSs.MoveNext
		if not RSs.eof then 
			if EffectiveDate = RSs("EffectiveDate") then
				ref2 = RSs("EffectiveDate")
				call setAXValue(RSs)
			end if
		end if
	end sub
	If not rs.EOF and not rsOLD.EOF then
	%>
	<table>
		<tr class="RepTableHeader">
			<td colspan="2">&nbsp;</td>
			<td>„«‰œÂ ›—Ê‘</td>
			<td>ﬂ· ›—Ê‘</td>
			<td>œ—’œ „«‰œÂ</td>
			<td>„«‰œÂ Œ—Ìœ</td>
			<td>ﬂ· Œ—Ìœ</td>
			<td>œ—’œ „«‰œÂ</td>
		</tr>
		
		<%
		rowColor="RepTR0"
		if myMonth="" then 
			do while not rsOLD.eof
				EffectiveDate=rsOLD("EffectiveDate")
				%>
				<tr class="<%=rowColor%>">
					<%
					arRemainedAmount=0
					apRemainedAmount=0	
					arTotal=0
					apTotal=0		
					call setAXValue(rsOLD)
					%>
					<td colspan="2">„«‰œÂ ﬁœÌ„Ì</td>
					<%
					response.write echoTD(arRemainedAmount, "ar")
					response.write echoTD(arTotal, "ar-full")
					response.write echoTD(round(arRemainedAmount/arTotal*100), "")
					response.write echoTD(apRemainedAmount, "ap")
					response.write echoTD(apTotal, "ap-full")
					response.write echoTD(round(apRemainedAmount/apTotal*100), "")
					%>
				</tr>
				<%
					
			loop
		end if
		arRemainedAmountSum=0
		apRemainedAmountSum=0
		arTotalAmountSum=0
		apTotalAmountSum=0
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
				if myMonth<>"" then 
					mm= mid(EffectiveDate,6,5)
				else
					mm=fmonth(cint(mid(EffectiveDate,6,2))-1)
				end if
				arRemainedAmount=0
				apRemainedAmount=0
				call setAXValue(rs)
				arRemainedAmountSum=arRemainedAmountSum+arRemainedAmount
				apRemainedAmountSum=apRemainedAmountSum+apRemainedAmount
				arTotalAmountSum=arTotalAmountSum+arTotal
				apTotalAmountSum=apTotalAmountSum+apTotal
				%>
				<td title="<%=EffectiveDate%>"><%=yyyy%></td>
				<%if myMonth<>"" then%>
				<td><%=mm%></td>
				<%else%>
				<td title='»—«Ì „‘«ÂœÂ Ã“∆Ì«  ﬂ·Ìﬂ ﬂ‰Ìœ'><a href="otherReports.asp?act=cash&month=<%=replace(mid(effectiveDate,1,8),"/","-")%>"><%=mm%></a></td>
				<%
				end if
				response.write echoTD(arRemainedAmount, "ar")
				response.write echoTD(arTotal, "ar-full")
				if arTotal>0 then 
					response.write echoTD(round(arRemainedAmount/arTotal*100), "")
				else
					response.write echoTD(0,"")
				end if
				response.write echoTD(apRemainedAmount, "ap")
				response.write echoTD(apTotal, "ap-full")
				if apTotal>0 then 
					response.write echoTD(round(apRemainedAmount/apTotal*100), "")
				else
					response.write echoTD(0,"")
				end if
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
				<td colspan="2">Ã„⁄ »œÊ‰ „«‰œÂùÂ«</td>
				<td><%=Separate(arRemainedAmountSum)%></td>
				<td><%=Separate(arTotalAmountSum)%></td>
				<td><%if arTotalAmountSum>0 then response.write Round(arRemainedAmountSum/arTotalAmountSum*100)%></td>
				<td><%=Separate(apRemainedAmountSum)%></td>
				<td><%=Separate(apTotalAmountSum)%></td>
				<td><%if apTotalAmountSum>0 then response.write Round(apRemainedAmountSum/apTotalAmountSum*100)%></td>
			</tr>
	</table>
<%
	end if
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------

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
			response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0

  '-----------------------------------------------------------------------------------------------------
  if request("action")=" ç«Å " then
  '----   It's Print ( Crystal Reports )
  %>
	<BR>
	<BR>
	<CENTER>
	<% 	ReportLogRow = PrepareReport ("MoeenRep01.rpt", "GLAccountGLFromDateToDateFromTafsilToTafsil", GLAccount & "" & OpenGL & "" & FromDate & "" & ToDate & "" & FromTafsil& "" & ToTafsil, "/beta/dialog_printManager.asp?act=Fin") %>
	<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

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
		response.redirect "?errMsg=" & Server.URLEncode("Õ”«» „Ê—œ ‰Ÿ— [" & GLAccount & "] œ— ”«· „«·Ì Ã«—Ì (" & OpenGL & ") ÊÃÊœ ‰œ«—œ ")
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
			«“  «—ÌŒ <B><%=replace(FromDate,"/",".")%></B>  «  «—ÌŒ <B><%=replace(ToDate,"/",".")%></B><br>
			«“  ›’Ì·Ì <B><%=FromTafsil%></B>  «  ›’Ì·Ì <B><%=ToTafsil%></B>
		</TD>
		<TD width=70 >ê—œ‘ »œÂﬂ«—		</TD>
		<TD width=70 >ê—œ‘ »” «‰ﬂ«—		</TD>
		<TD width=70 >„«‰œÂ »œÂﬂ«—		</TD>
		<TD width=70 >„«‰œÂ »” «‰ﬂ«—	</TD>
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
	<TR bgcolor="eeeeee" style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
		<TD width=50  onclick='go2Page(1,1);' style="<%if abs(ord)=1 then response.write style%>"> ›’Ì·Ì			<%if abs(ord)=1 then response.write arrow%></TD>
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