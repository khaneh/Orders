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

		if createDateFrom <> "" then 
			extraConditions = extraConditions & " AND CreatedDate >= '" & createDateFrom &"' "
		end if
		if createDateTo <> "" then 
			extraConditions = extraConditions & " AND CreatedDate <= '" & createDateTo &"' "
		end if
		if accountGroup <> "-1" then 
			extraConditions = extraConditions & " AND Accounts.ID IN (select account from accountGroupRelations where accountGroup=" & accountGroup & ")"
		end if
		if isPostable = "yes" then 
			extraConditions = extraConditions & " AND (Postable1=1 OR Postable2=1) "
		end if
		if isPostable = "no" then 
			extraConditions = extraConditions & " AND (Postable1=0 OR Postable2=0) "
		end if
		
		if lastInvoiceDateFrom <> "" then 
			extraConditions = extraConditions & " AND Accounts.ID IN (select customer from invoices where voided=0 and issued=1 and IsReverse=0 group by customer having max(issuedDate) >='" & lastInvoiceDateFrom &"') "
		end if
		if lastInvoiceDateTo <> "" then 
			extraConditions = extraConditions & " AND Accounts.ID IN (select customer from invoices where voided=0 and issued=1 and IsReverse=0 group by customer having max(issuedDate) <='" & lastInvoiceDateTo &"') "
		end if
		salesConditions = ""
		if salesInvoiceDateFrom <> "" then 
			salesConditions = salesConditions & " AND Invoices.IssuedDate >='" & salesInvoiceDateFrom & "' "
		end if
		if salesInvoiceDateTo <> "" then 
			salesConditions = salesConditions & " AND Invoices.IssuedDate <='" & salesInvoiceDateTo & "' "
		end if
		Ord =	request("Ord")	
		select case Ord
			case "1":
				order="AccountTitle"
			case "-1":
				order="AccountTitle DESC"
			case "2":
				order="companyName"
			case "-2":
				order="companyName DESC"
			case "3":
				order="Dear1,FirstName1,LastName1"
			case "-3":
				order="Dear1,FirstName1,LastName1 DESC"
			case "4":
				order="CSRName"
			case "-4":
				order="CSRName DESC"
			Case "5":
				order = "sumDigital"
			Case "-5":
				order = "sumDigital DESC"
			case "6":
				order="sumOffset"
			case "-6":
				order="sumOffset DESC"
			case "7":
				order="sumOther"
			case "-7":
				order="sumOther DESC"
			case else:
				order="AccountTitle"
				Ord=1
		end select	
		
		Set SA_RS1 = Server.CreateObject("ADODB.Recordset")
		SA_RS1.CursorLocation=3 '---------- Alix: Because in ADOVBS_INC adUseClient=3

		SA_RS1.PageSize = SA_RecordLimit

		'SA_mySQL="SELECT Accounts.*, Users.RealName AS CSRName FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID WHERE (REPLACE(Accounts.AccountTitle, ' ', '') LIKE REPLACE(N'%"& sqlSafe(SA_TitleOrName) & "%', ' ', ''))"& extraConditions & " ORDER BY Accounts.AccountTitle"
		SA_mySQL="SELECT Accounts.*, Users.RealName AS CSRName,digital.Price as sumDigital,offset.Price as sumOffset,other.Price as sumOther FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID left outer join (select Invoices.Customer, SUM(cast(InvoiceLines.Price as float)) as Price from Invoices inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice inner join InvoiceItemCategoryRelations on InvoiceLines.Item=InvoiceItemCategoryRelations.InvoiceItem where InvoiceItemCategoryRelations.InvoiceItemCategory=3 and Invoices.Voided=0 and Invoices.Approved=1 and Invoices.Issued=1 "& salesConditions &" group by Invoices.Customer) offset on Accounts.ID=offset.Customer left outer join (select Invoices.Customer, SUM(cast(InvoiceLines.Price as float)) as Price from Invoices inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice inner join  InvoiceItemCategoryRelations on InvoiceLines.Item=InvoiceItemCategoryRelations.InvoiceItem where InvoiceItemCategoryRelations.InvoiceItemCategory=4 and Invoices.Voided=0 and Invoices.Approved=1 and Invoices.Issued=1 "& salesConditions &" group by Invoices.Customer) digital on Accounts.ID=digital.Customer LEFT OUTER JOIN (select Invoices.Customer, SUM(cast(InvoiceLines.Price as float)) as Price from Invoices inner join InvoiceLines on Invoices.ID=InvoiceLines.Invoice inner join InvoiceItemCategoryRelations on InvoiceLines.Item=InvoiceItemCategoryRelations.InvoiceItem where InvoiceItemCategoryRelations.InvoiceItemCategory NOT IN (3,4) and Invoices.Voided=0 and Invoices.Approved=1 and Invoices.Issued=1 "& salesConditions &" group by Invoices.Customer) other on Accounts.ID=other.Customer WHERE (REPLACE(Accounts.AccountTitle, ' ', '') LIKE REPLACE(N'%"& sqlSafe(SA_TitleOrName) & "%', ' ', ''))"& extraConditions & " ORDER BY " & order
'response.write "<br>"&sa_mySQL
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
		end if 
		if ord<0 then
			style="background-color: #33CC99;"
			arrow="<br><span style='font-family:webdings'>6 6 6</span>"
		else
			style="background-color: #33CC99;"
			arrow="<br><span style='font-family:webdings'>5 5 5</span>"
		end if
		%><br>
		<div dir='rtl'><B><%=SA_StepText%></B>
		</div><br>
<!-- «‰ Œ«» Õ”«» --> 
		<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;" align="center" Width="90%" Cellspacing="0" Cellpadding="5">
		<tbody id="AccountsTable">
				<tr bgcolor='#33AACC' style="cursor:hand;" title=" — Ì» ‰„«Ì‘">
					<td style="width:10px;"> # </td>
					<td><INPUT TYPE="radio" NAME="O" checked disabled></td>
					<td onclick='go2Page(1,-1);' style="<%if abs(ord)=1 then response.write style%>"> ⁄‰Ê«‰ Õ”«» <%if abs(ord)=1 then response.write arrow%></td>
					<td onclick='go2Page(1,-2);' style="<%if abs(ord)=2 then response.write style%>"> ‰«„ ‘—ﬂ <%if abs(ord)=2 then response.write arrow%></td>
					<td onclick='go2Page(1,-3);' style="<%if abs(ord)=3 then response.write style%>"> ‰«„ —«»ÿ «Ê· <%if abs(ord)=3 then response.write arrow%></td>
					<td width='50px' onclick='go2Page(1,-4);' style="<%if abs(ord)=4 then response.write style%>"> „”∆Ê· ÅÌêÌ—Ì <%if abs(ord)=4 then response.write arrow%></td>
					<td width='70px' onclick='go2Page(1,-5);' style="<%if abs(ord)=5 then response.write style%>"> ›—Ê‘ œÌÃÌ «·<%if abs(ord)=5 then response.write arrow%></td>
					<td width='70px' onclick='go2Page(1,-6);' style="<%if abs(ord)=6 then response.write style%>"> ›—Ê‘ «›” <%if abs(ord)=6 then response.write arrow%></td>
					<td width='70px' onclick='go2Page(1,-7);' style="<%if abs(ord)=7 then response.write style%>"> «·»«ﬁÌ<%if abs(ord)=7 then response.write arrow%></td>
				</tr>
	<%		
			SA_tempCounter= (Page - 1 ) * SA_RecordLimit
			while Not ( SA_RS1.EOF)   and (SA_RS1.AbsolutePage = Page)
				'SA_OldAccountNo=SA_RS1("ACCID")
				SA_AccountNo=SA_RS1("ID")
				SA_AccountTitle=SA_RS1("AccountTitle")
				SA_digital=SA_RS1("sumDigital")
				SA_offset=SA_RS1("sumOffset")
				SA_other=SA_RS1("sumOther")
				if SA_RS1("Type") = 1 then 
					SA_AccountTitle = SA_AccountTitle & " (›—Ê‘‰œÂ) "
				end if
				if isnull(SA_digital) then SA_digital="-"
				if isnull(SA_offset) then SA_offset="-"
				if isnull(SA_other) then SA_other="-"
				SA_tempCounter=SA_tempCounter+1
	%>				<tr>
						<td style="border-bottom: solid 1pt black"><%=SA_tempCounter %>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black;"><INPUT TYPE="radio" NAME="selectedCustomer" Value="<%=SA_RS1("ID")%>" <% if SA_RS1("status")="3" then %> disabled <% end if %> onclick="okToProceed=true;setColor(this)" ></td>
						<td  style="border-bottom: solid 1pt black">
						<%if Auth(1 , 1) then %>
							<A href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=SA_AccountNo%>&searchtype=advanced&createDateFrom=<%=createDateFrom%>&createDateTo=<%=createDateTo%>&accountGroup=<%=accountGroup%>&isPostable=<%=isPostable%>&lastInvoiceDateFrom=<%=lastInvoiceDateFrom%>&lastInvoiceDateTo=<%=lastInvoiceDateTo%>&submitButton=<%=submitButton%>" target='_blank'><%=SA_AccountTitle%></A>  <% if SA_RS1("status")="3" then %> <FONT COLOR="red"><B>„”œÊœ</B></FONT> <% end if %> 
						<%else
							response.write SA_AccountTitle
						end if%>
						&nbsp;</td>
						<td  style="border-bottom: solid 1pt black;"><%=SA_RS1("ûCompanyName")%>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><%=SA_RS1("Dear1")& " " & SA_RS1("FirstName1")& " " & SA_RS1("LastName1")%>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black"><B><FONT Color='blue'><%=SA_RS1("CSRName")%></FONT></B>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="black"><%=Separate(SA_digital)%></FONT>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="black"><%=Separate(SA_offset)%></FONT>&nbsp;</td>
						<td  style="border-bottom: solid 1pt black; direction:LTR; text-align:right;"><FONT COLOR="black"><%=Separate(SA_other)%></FONT>&nbsp;</td>
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
						<A HREF="?p=<%=i%>&<%=SA_SearchBox%>=<%=SA_TitleOrName%>&act=<%=SA_ActName%>&createDateFrom=<%=createDateFrom%>&createDateTo=<%=createDateTo%>&accountGroup=<%=accountGroup%>&isPostable=<%=isPostable%>&lastInvoiceDateFrom=<%=lastInvoiceDateFrom%>&lastInvoiceDateTo=<%=lastInvoiceDateTo%>&submitButton=<%=submitButton%>&ord=<%=ord%>&salesInvoiceDateFrom=<%=salesInvoiceDateFrom%>&salesInvoiceDateTo=<%=salesInvoiceDateTo%>"><%=i%></A>&nbsp;
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
function go2Page(p,ord) {
	if(ord==0){
		ord=<%=Ord%>;
	}
	else if(ord==<%=Ord%>){
		ord= 0-ord;
	}
	//document.getElementsByName('Page')[0].value=p;
	//document.getElementsByName('Ord')[0].value=ord;
	//document.forms[0].submit();
	str = "?p="+p+"&<%=SA_SearchBox%>=<%=SA_TitleOrName%>&act=<%=SA_ActName%>&createDateFrom=<%=createDateFrom%>&createDateTo=<%=createDateTo%>&accountGroup=<%=accountGroup%>&isPostable=<%=isPostable%>&lastInvoiceDateFrom=<%=lastInvoiceDateFrom%>&lastInvoiceDateTo=<%=lastInvoiceDateTo%>&submitButton=<%=submitButton%>&ord="+ord+"&salesInvoiceDateFrom=<%=salesInvoiceDateFrom%>&salesInvoiceDateTo=<%=salesInvoiceDateTo%>"
	window.location = str;
}
//-->
</SCRIPT>
