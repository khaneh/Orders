<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= "����� ����"
SubmenuItem=7
if not Auth(7 , 7) then NotAllowdToViewThisPage()

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
<a href="?act=showRemain">����� ���ȝ��� ��������</a>
<FORM METHOD=POST ACTION="?act=show">
<CENTER>����� �����: <select name="selectedCustomer" class=inputBut ></CENTER>
		<option value="1000">��� (�����)</option>
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
<INPUT TYPE="submit" value="������"><BR>
<INPUT TYPE="checkbox" <% if  request("showZero") = "on" then%>checked <% end if %> NAME="showZero"> ����� ����� ��� ���
</form>
<%
end if 
if request("act")="showRemain" then 
%>
<table>
	<tr>
		<td>���</td>
		<td>���� �����</td>
		<td>����� ����</td>
		<td></td>
	</tr>
<%
	set rs=Conn.execute("select * from accounts where apBalance<0")
	while not rs.eof 
%>
	<tr>
		<td><a href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rs("id")%>"><%=rs("accountTitle")%></a></td>
		<td><%=Separate(abs(CDbl(rs("apBalance"))))%></td>
		<td><%=Separate(abs(CDbl(rs("arBalance"))))%></td>
		<td><%if CDbl(rs("arBalance"))<0 then response.write " ��" else if CDbl(rs("arBalance")) >0 then response.write " ��"%></td>
	</tr>
<%		
		rs.moveNext
	wend
%>
</table>
<%
end if
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Send New Message
'-----------------------------------------------------------------------------------------------------
if request("act")="show" and selectedCustomer<>"" then

	if not request("showZero") = "on" then
		extraCond = "and not APBalance=0"
	else
		extraCond = " "
	end if 

	'=========================================== All CSRs
	'====================================================
	if selectedCustomer=1000 then '----- brief Report 
	set RSM = conn.Execute ("SELECT Accounts.CSR, COUNT(Accounts.APBalance) AS AccountsCOUNT, SUM(CONVERT(bigint, Accounts.APBalance)) AS sumAPBalance, Users.RealName FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID GROUP BY Accounts.CSR, Users.RealName ORDER BY sumAPBalance")
	%>
	<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;"  align="center" Width="90%" Cellspacing="0" Cellpadding="5">
	<tbody id="AccountsTable">
			<tr bgcolor='#33AACC'>
				<td> # </td>
				<td> ����� �����</td>
				<td> ����� ������</td>
				<td width='80'> ����� ����</td>
			</tr>
			<%
			SA_tempCounter= 0
			while Not ( RSM.EOF)   
				AccountsCOUNT=RSM("AccountsCOUNT")
				sumAPBalance=cdbl(RSM("sumAPBalance"))
				RealName=RSM("RealName")
				CSR=RSM("CSR")

				if (cdbl(sumAPBalance) >= 0 )then
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

				total = total + APBalance

	%>				<tr >
						<td style="border-bottom: solid 1pt black"><%=SA_tempCounter %>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black" class=alak2><A  HREF="?act=show&selectedCustomer=<%=CSR%>"><%=RealName%></A>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><%=AccountsCOUNT%>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="<%=SA_tempBalanceColor%>"><%=Separate(sumAPBalance)%></FONT>&nbsp;</td>
					</tr>
	<%			
				'total = cdbl(total) + cdbl(sumAPBalance)
				RSM.movenext
			wend
	%>
			<tr bgcolor='#33AACC'>
				<td align='right' colspan=3> ��� �� ���� ��� ��� ����� ���� ����� �� ����. </td>
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
		
		mySQL="SELECT * FROM Accounts WHERE (CSR IS NULL)"& extraCond & " ORDER BY APBalance"
	else
		mySQL="SELECT * FROM Accounts where CSR = " & selectedCustomer & " "& extraCond & " ORDER BY APBalance"
	end if

	set RSM = conn.Execute (mySQL)
	%>
	<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;"  align="center" Width="90%" Cellspacing="0" Cellpadding="5">
	<tbody id="AccountsTable">
			<tr bgcolor='#33AACC'>
				<td> # </td>
				<td> ����� ���� </td>
				<td> ��� ���� ��� </td>
				<td width='80'> ����� ����</td>
			</tr>
			<%
			SA_tempCounter= 0
			while Not ( RSM.EOF)   
				AccountNo=RSM("ID")
				AccountTitle=RSM("AccountTitle")
				APBalance=cdbl(RSM("APBalance"))
				CreditLimit=RSM("CreditLimit")
				contact1 = RSM("Dear1") & " " & RSM("FirstName1") & " " & RSM("LastName1")
				if RSM("Type") = 1 then 
					AccountTitle = AccountTitle & " (�������) "
				end if
				if isnull(SA_totalBalance) then
					APBalance=0
					totalBalanceText="<i>N / A</i>"
					tempBalanceColor="gray"
				else
					if (APBalance>= 0 )then
						SA_tempBalanceColor="green"
					else
						SA_tempBalanceColor="red"
					end if
					totalBalanceText=Separate(APBalance)
				end if
				SA_tempCounter=SA_tempCounter+1
				if (SA_tempCounter Mod 2 = 1)then
					SA_tempColor="#FFFFFF"
				else
					SA_tempColor="#DDDDDD"
				end if

'				total = total + APBalance

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
				<td align='right' colspan=3> ���</td>
				<td dir=ltr> <%=Separate(total)%> </td>

				</td>
			</tr>
		</table><BR><BR>
<%
end if
%>

<!--#include file="tah.asp" -->