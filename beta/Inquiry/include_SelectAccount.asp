<%
'		This Include File Needs Following Variables to have values:
'
'		SA_Action				(the Action of the Suubmit Button)
'		SA_TitleOrName			(AccountTitle or Name to be searched)
'		SA_SearchAgainURL		()
'		SA_StepText				(e.g. 'ê«„ œÊ„ :')
'		SA_ActName				(e.g. 'submitsearch')
'		SA_SearchBox			(e.g. 'CustomerNameSearchBox')
'		SA_IsVendor 			(is 1 or 0)
'
'
if SA_ActName="" then SA_ActName="submitsearch"
if SA_SearchBox="" then SA_SearchBox="CustomerNameSearchBox"

		SA_RecordLimit = 20

		if SA_IsVendor  = 1  then 
			extraConditions = " and [type]=1 "
		else
			extraConditions = "  "
		end if

		Set SA_RS1 = Server.CreateObject("ADODB.Recordset")
		SA_RS1.CursorLocation=3 '---------- Alix: Because in ADOVBS_INC adUseClient=3

		SA_RS1.PageSize = SA_RecordLimit

		SA_mySQL="SELECT Accounts.*, Users.RealName AS CSRName FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID WHERE (REPLACE(Accounts.AccountTitle, ' ', '') LIKE REPLACE(N'%"& sqlSafe(SA_TitleOrName) & "%', ' ', ''))"& extraConditions & " ORDER BY Accounts.AccountTitle"

'		SA_RS1.Open SA_mySQL ,Conn,adOpenStatic,,adcmdcommand 
		SA_RS1.Open SA_mySQL ,Conn,3,,adcmdcommand 
		FilePages = SA_RS1.PageCount

		Page=1
		if isnumeric(Request.QueryString("p")) then
		  if (clng(Request.QueryString("p")) <= FilePages) and clng(Request.QueryString("p") > 0) then
			Page=clng(Request.QueryString("p"))
		  end if
		end if

		if not SA_RS1.eof then
			SA_RS1.AbsolutePage=Page
		end if

		if (SA_RS1.eof) then 
		' Not Found
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("ç‰Ì‰ Õ”«»Ì ÅÌœ« ‰‘œ.<br><a href='../CRM/AccountEdit.asp?act=getAccount'>Õ”«» ÃœÌœø</a>")
		end if %><br>
		<div dir='rtl'><B><%=SA_StepText%></B>
		</div><br>
<!-- «‰ Œ«» Õ”«» --> 
		<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;" align="center" Width="90%" Cellspacing="0" Cellpadding="5">
		<tbody id="AccountsTable">
				<tr bgcolor='#33AACC'>
					<td style="width:10px;"> # </td>
					<td><INPUT TYPE="radio" NAME="O" checked disabled></td>
					<td> ⁄‰Ê«‰ Õ”«» </td>
					<td> ‰«„ ‘—ﬂ </td>
					<td> ‰«„ —«»ÿ «Ê· </td>
					<td width='50px'> „”∆Ê· ÅÌêÌ—Ì </td>
					<td width='70px'> „«‰œÂ ›—Ê‘</td>
					<td width='70px'> „«‰œÂ Œ—Ìœ</td>
				</tr>
	<%		
			SA_tempCounter= (Page - 1 ) * SA_RecordLimit
			while Not ( SA_RS1.EOF)   and (SA_RS1.AbsolutePage = Page)
				'SA_OldAccountNo=SA_RS1("ACCID")
				SA_AccountNo=SA_RS1("ID")
				SA_AccountTitle=SA_RS1("AccountTitle")
				SA_totalBalance=SA_RS1("ARBalance")
				SA_totalBalanceAP=SA_RS1("APBalance")
				SA_totalBalanceAO=SA_RS1("AOBalance")
				if SA_RS1("Type") = 1 then 
					SA_AccountTitle = SA_AccountTitle & " (›—Ê‘‰œÂ) "
				end if
				if isnull(SA_totalBalance) then
					SA_totalBalance=0
					SA_totalBalanceText="<i>N / A</i>"
					SA_tempBalanceColor="gray"
				else
					SA_totalBalanceText=Separate(SA_totalBalance)
					if (cdbl(SA_totalBalance) >= 0 )then
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

	%>				<tr>
						<td style="border-bottom: solid 1pt black"><%=SA_tempCounter %>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black;"><INPUT TYPE="radio" NAME="selectedCustomer" Value="<%=SA_RS1("ID")%>" <% if SA_RS1("status")="3" then %> disabled <% end if %> onclick="okToProceed=true;setColor(this)" ></td>
						<td  style="border-bottom: solid 1pt black">
						<%if Auth(1 , 1) then %>
							<A href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=SA_AccountNo%>" target='_blank'><%=SA_AccountTitle%></A>  <% if SA_RS1("status")="3" then %> <FONT COLOR="red"><B>„”œÊœ</B></FONT> <% end if %> 
						<%else
							response.write SA_AccountTitle
						end if%>
						&nbsp;</td>
						<td  style="border-bottom: solid 1pt black;"><%=SA_RS1("ûCompanyName")%>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><%=SA_RS1("Dear1")& " " & SA_RS1("FirstName1")& " " & SA_RS1("LastName1")%>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><B><FONT Color='blue'><%=SA_RS1("CSRName")%></FONT></B>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="<%=SA_tempBalanceColor%>"><%=SA_totalBalanceText%></FONT>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="black"><%=Separate(SA_totalBalanceAP)%></FONT>&nbsp;</td>
					</tr>
	<%			SA_RS1.movenext
			wend
	%>
			<tr bgcolor='#33AACC'>
				<td align='right' colspan=3>
					<INPUT class="GenButton" TYPE="submit" name="submitSelection" value="«‰ Œ«»" onclick="<%=SA_Action%>">

	<SCRIPT LANGUAGE="JavaScript">
	<!--
	if (! window.dialogArguments){// if this is NOT a modal window
		document.write('&nbsp;&nbsp; <INPUT class="GenButton" TYPE="button" name="submitSelection" value="ÃœÌœ" onclick="window.location=\'../CRM/AccountEdit.asp?act=getAccount\'">');
	}
	//-->
	</SCRIPT>
				</td>
				<td align='center' colspan="7" dir="RTL">&nbsp;<span id="sa_links">
	<%
			if FilePages>1 then
				response.write "’›ÕÂ &nbsp;"
				for i=1 to FilePages 
					if i = page then %>
						[<%=i%>]&nbsp;
<%					else%>
						<A HREF="?p=<%=i%>&<%=SA_SearchBox%>=<%=SA_TitleOrName%>&act=<%=SA_ActName%>"><%=i%></A>&nbsp;
<%					end if 
				next 
			end if
%>
				</span></td>
			</tr>
		</tbody>
		</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
function setColor(obj)
{
	for(i=0; i<document.all.selectedCustomer.length; i++)
		{
		theTR = document.all.selectedCustomer[i].parentNode.parentNode
		theTR.setAttribute("bgColor","#C3DBEB")
		}
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor","#FFFFFF")
}
i=0;

do {
	radio1=document.getElementsByName("selectedCustomer")[i];
	i++;
} while (radio1.disabled && i<document.getElementsByName("selectedCustomer").length)
if (!radio1.disabled){
	radio1.checked=true;
	okToProceed=true;
	setColor(radio1);
	if (window.dialogArguments){// if this is a modal window
		if (document.all.selectedCustomer.length==<%=SA_RecordLimit%>)
			document.getElementById("sa_links").innerText=" ⁄œ«œ ÃÊ«» Â« »Ì‘ — «“ «Ì‰ «”  ... Ã” ÃÊ —« „ÕœÊœ  — ﬂ‰Ìœ";
		else
			document.getElementById("sa_links").innerText="";
	}else{
		radio1.focus();
	}
}
//-->
</SCRIPT>