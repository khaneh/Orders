<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="ַׁׂװ Ýׁזװ"
SubmenuItem=8
if not Auth(6 , 8) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<BR><BR><BR>
<%

CSR = request("CSR")
if CSR="" then CSR=1000

%>
<% 
if not Auth(6 , 9) then 
	CSR = session("id")
else
%>
<FORM METHOD=POST ACTION="?act=show">
<CENTER>ד׃ֶזב םםׁם: <select name="CSR" class=inputBut ></CENTER>
		<option value="1000">ודו (־בַױו)</option>
		<option value="2000">-------------------</option>
<% set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
Do while not RSV.eof
%>
	<option value="<%=RSV("ID")%>" <%
		if trim(RSV("ID"))=trim(CSR) then
			response.write " selected "
		end if
		%>><%=RSV("RealName")%></option>
<%
RSV.moveNext
Loop
RSV.close
%>
</select> 
<INPUT TYPE="submit" value="דװַוֿו"><BR>
<INPUT TYPE="checkbox" <% if request("showZero") = "on" then%>checked <% end if %> NAME="showZero"> הדַםװ דַהֿו וַם ױÝׁ
</form>
<%
end if 

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Send New Message
'-----------------------------------------------------------------------------------------------------
if request("act")="show" and CSR<>"" then

	if not request("showZero") = "on" then
		extraCond = "and not ARBalance=0"
	else
		extraCond = " "
	end if 

	'=========================================== All CSRs
	'====================================================
	if CSR=1000 then '----- brief Report 
'	mySQL="SELECT Accounts.CSR, COUNT(Accounts.APBalance) AS AccountsCOUNT, SUM(CONVERT(bigint, Accounts.ARBalance)) AS sumARBalance, Users.RealName FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID GROUP BY Accounts.CSR, Users.RealName ORDER BY sumARBalance"
'	Changed By Kid 840128
	mySQL="SELECT Accounts.CSR, RealName, COUNT(Accounts.APBalance) AS AccountsCOUNT, SUM((SIGN(Accounts.ARBalance) + 1) * .5 * CONVERT(bigint, Accounts.ARBalance)) AS sumPosBalance, SUM((SIGN(Accounts.ARBalance) - 1) * .5 * CONVERT(bigint, Accounts.ARBalance)) AS sumNegBalance FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID GROUP BY Accounts.CSR, Users.RealName ORDER BY sumNegBalance DESC"

	set RSM = conn.Execute (mySQL)
	%>
	<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;" align="center" Width="90%" Cellspacing="0" Cellpadding="5">
	<tbody id="AccountsTable">
			<tr bgcolor='#33AACC'>
				<td> # </td>
				<td> ד׃ֶזב םםׁם</td>
				<td> ÊÚַֿֿ ֽ׃ַָוַ</td>
				<td width='80'> דַהֿו ָ׃Êַהßַׁ</td>
				<td width='80'> דַהֿו ָֿוßַׁ</td>
			</tr>
			<%
			SA_tempCounter= 0
			while Not ( RSM.EOF) 
				AccountsCOUNT=cint(RSM("AccountsCOUNT"))
				sumPosBalance=cdbl(RSM("sumPosBalance"))
				sumNegBalance=cdbl(RSM("sumNegBalance"))
				RealName=RSM("RealName")
				CSR=cint(RSM("CSR"))

				if (cdbl(sumARBalance) >= 0 )then
					SA_tempBalanceColor="green"
				else
					SA_tempBalanceColor="red"
				end if
				
				SA_tempCounter=SA_tempCounter+1
				if (SA_tempCounter Mod 2 = 1)then
					SA_tempColor="#FFFFFF"
				else
					SA_tempColor="#DDDDDD"
				end if

	%>				<tr >
						<td style="border-bottom: solid 1pt black"><%=SA_tempCounter %>&nbsp;</td>
						<td style="border-bottom: solid 1pt black" class=alak2><A HREF="CSRreport.asp?act=show&CSR=<%=CSR%>"><%=RealName%></A>&nbsp;</td>
						<td style="border-bottom: solid 1pt black"><%=AccountsCOUNT%>&nbsp;</td>
						<td style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="green"><%=Separate(sumPosBalance)%></FONT>&nbsp;</td>
						<td style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="red"><%=Separate(sumNegBalance)%></FONT>&nbsp;</td>
					</tr>
	<%			
				'total = cdbl(total) + cdbl(sumARBalance)
				RSM.movenext
			wend
	%>
			<tr bgcolor='#33AACC'>
				<td align='right' colspan=4> ֲהו ַׁ ַָׁם ־זֿ הדם ׃הֿם ַָׁם ֿםַׁה וד ד׃הֿ. </td>
				<td dir=ltr><! <%=Separate(total)%> ></td>

				</td>
			</tr>
		</table><BR><BR>
<%
		response.end
	end if
	'============================================ One CSR
	'====================================================
	if CSR=2000 then '----- no csr Report 
		
		mySQL="SELECT * FROM Accounts WHERE (CSR IS NULL)"& extraCond & " ORDER BY ARBalance"
	else
		mySQL="SELECT * FROM Accounts where CSR = " & CSR & " "& extraCond & " ORDER BY ARBalance"
	end if

	set RSM = conn.Execute (mySQL)
	%>
	<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;" align="center" Width="90%" Cellspacing="0" Cellpadding="5">
	<tbody id="AccountsTable">
			<tr bgcolor='#33AACC'>
				<td> # </td>
				<td> Úהזַה ֽ׃ַָ </td>
				<td> הַד ַָׁ״ ַזב </td>
				<td width='80'> דַהֿו Ýׁזװ</td>
			</tr>
			<%
			SA_tempCounter= 0
			while Not ( RSM.EOF) 
				AccountNo=RSM("ID")
				AccountTitle=RSM("AccountTitle")
				ARBalance=cdbl(RSM("ARBalance"))
				CreditLimit=RSM("CreditLimit")
				contact1 = RSM("Dear1") & " " & RSM("FirstName1") & " " & RSM("LastName1")
				if RSM("Type") = 1 then 
					AccountTitle = AccountTitle & " (Ýׁזװהֿו) "
				end if
				if isnull(SA_totalBalance) then
					ARBalance=0
					totalBalanceText="<i>N / A</i>"
					tempBalanceColor="gray"
				else
					totalBalanceText=Separate(ARBalance)
					if (ARBalance >= 0 )then
						SA_tempBalanceColor="green"
					else
						SA_tempBalanceColor="red"
					end if
				end if
				SA_tempCounter=SA_tempCounter+1
				if (SA_tempCounter Mod 2 = 1)then
					SA_tempColor="#FFFFFF"
				else
					SA_tempColor="#DDDDDD"
				end if

				total = total + ARBalance

	%>				<tr >
						<td style="border-bottom: solid 1pt black"><%=SA_tempCounter %>&nbsp;</td>
						<td style="border-bottom: solid 1pt black"><A href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=AccountNo%>" target='_blank'><%=AccountTitle%></A>&nbsp;</td>
						<td style="border-bottom: solid 1pt black"><%=contact1%>&nbsp;</td>
						<td style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="<%=SA_tempBalanceColor%>"><%=totalBalanceText%></FONT>&nbsp;</td>
					</tr>
	<%			RSM.movenext
			wend
	%>
			<tr bgcolor='#33AACC'>
				<td align='right' colspan=3> ּדÚ</td>
				<td dir=ltr align='right' > <%=Separate(total)%>&nbsp;</td>

				</td>
			</tr>
		</table><BR><BR>
<%
end if
%>

<!--#include file="tah.asp" -->
