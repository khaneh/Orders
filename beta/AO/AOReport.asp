<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'A0 (11 [=B])
PageTitle= "ַׁׂװ ׃ַםׁ"
SubmenuItem=5
if not Auth("B" , 5) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<BR><BR><BR>
<%

selectedCustomer = request("selectedCustomer")
if selectedCustomer="" then selectedCustomer=1000

%>
<% 
if not Auth(6 , 9) then 
	selectedCustomer = session("id")
else
%>
<FORM METHOD=POST ACTION="?act=show">
<CENTER>ד׃ֶזב םםׁם: <select name="selectedCustomer" class=inputBut ></CENTER>
		<option value="1000">ודו (־בַױו)</option>
		<option value="2000">-------------------</option>
<% set RSV=Conn.Execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
Do while not RSV.eof
%>
	<option value="<%=RSV("ID")%>" <%
		if trim(RSV("ID"))=trim(selectedCustomer) then
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
<INPUT TYPE="checkbox" <% if  request("showZero") = "on" then%>checked <% end if %> NAME="showZero"> הדַםװ דַהֿו וַם ױÝׁ
</form>
<%
end if 

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Send New Message
'-----------------------------------------------------------------------------------------------------
if request("act")="show" and selectedCustomer<>"" then

	if not request("showZero") = "on" then
		extraCond = "and not AOBalance=0"
	else
		extraCond = " "
	end if 

	'=========================================== All CSRs
	'====================================================
	if selectedCustomer=1000 then '----- brief Report 
	set RSM = conn.Execute ("SELECT Accounts.CSR, COUNT(Accounts.APBalance) AS AccountsCOUNT, SUM(CONVERT(bigint, Accounts.AOBalance)) AS sumAOBalance, Users.RealName FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID GROUP BY Accounts.CSR, Users.RealName ORDER BY sumAOBalance")
	%>
	<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;"  align="center" Width="90%" Cellspacing="0" Cellpadding="5">
	<tbody id="AccountsTable">
			<tr bgcolor='#33AACC'>
				<td> # </td>
				<td> ד׃ֶזב םםׁם</td>
				<td> ÊÚַֿֿ ֽ׃ַָוַ</td>
				<td width='80'> דַהֿו ׃ַםׁ</td>
			</tr>
			<%
			SA_tempCounter= 0
			while Not ( RSM.EOF)   
				AccountsCOUNT=RSM("AccountsCOUNT")
				sumAOBalance=RSM("sumAOBalance")
				RealName=RSM("RealName")
				CSR=RSM("CSR")

				if (cdbl(sumAOBalance) >= 0 )then
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

				total = total + AOBalance

	%>				<tr >
						<td style="border-bottom: solid 1pt black"><%=SA_tempCounter %>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black" class=alak2><A  HREF="?act=show&selectedCustomer=<%=CSR%>"><%=RealName%></A>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><%=AccountsCOUNT%>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="<%=SA_tempBalanceColor%>"><%=Separate(sumAOBalance)%></FONT>&nbsp;</td>
					</tr>
	<%			
				'total = clng(total) + clng(sumAOBalance)
				RSM.movenext
			wend
	%>
			<tr bgcolor='#33AACC'>
				<td align='right' colspan=3> ֲהו ַׁ ַָׁם ־זֿ הדם ׃הֿם ַָׁם ֿםַׁה וד ד׃הֿ. </td>
				<td dir=ltr><! <%=Separate(total)%> ></td>

				</td>
			</tr>
		</table><BR><BR>
<%
		response.end
	end if
	'============================================ One CSR
	'====================================================
	if selectedCustomer=2000 then '----- no csr Report 
		
		mySQL="SELECT * FROM Accounts WHERE (CSR IS NULL)"& extraCond & " ORDER BY AOBalance"
	else
		mySQL="SELECT * FROM Accounts where CSR = " & selectedCustomer & " "& extraCond & " ORDER BY AOBalance"
	end if

	set RSM = conn.Execute (mySQL)
	%>
	<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;"  align="center" Width="90%" Cellspacing="0" Cellpadding="5">
	<tbody id="AccountsTable">
			<tr bgcolor='#33AACC'>
				<td> # </td>
				<td> Úהזַה ֽ׃ַָ </td>
				<td> הַד ַָׁ״ ַזב </td>
				<td width='80'> דַהֿו ׃ַםׁ</td>
			</tr>
			<%
			SA_tempCounter= 0
			while Not ( RSM.EOF)   
				AccountNo=RSM("ID")
				AccountTitle=RSM("AccountTitle")
				AOBalance=cdbl(RSM("AOBalance"))
				CreditLimit=RSM("CreditLimit")
				contact1 = RSM("Dear1") & " " & RSM("FirstName1") & " " & RSM("LastName1")
				if RSM("Type") = 1 then 
					AccountTitle = AccountTitle & " (Ýׁזװהֿו) "
				end if
				if isnull(SA_totalBalance) then
					AOBalance=0
					totalBalanceText="<i>N / A</i>"
					tempBalanceColor="gray"
				else
					if (AOBalance>= 0 )then
						SA_tempBalanceColor="green"
					else
						SA_tempBalanceColor="red"
					end if
					totalBalanceText=Separate(AOBalance)
				end if
				SA_tempCounter=SA_tempCounter+1
				if (SA_tempCounter Mod 2 = 1)then
					SA_tempColor="#FFFFFF"
				else
					SA_tempColor="#DDDDDD"
				end if

'				total = total + AOBalance

	%>				<tr >
						<td style="border-bottom: solid 1pt black"><%=SA_tempCounter %>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><A href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=AccountNo%>" target='_blank'><%=AccountTitle%></A>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><%=contact1%>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="<%=SA_tempBalanceColor%>"><%=totalBalanceText%></FONT>&nbsp;</td>
					</tr>
	<%			RSM.movenext
			wend
	%>
			<tr bgcolor='#33AACC'>
				<td align='right' colspan=3> ּדÚ</td>
				<td dir=ltr> <%=Separate(total)%> </td>

				</td>
			</tr>
		</table><BR><BR>
<%
end if
%>

<!--#include file="tah.asp" -->