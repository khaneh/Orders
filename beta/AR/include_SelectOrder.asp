<%
'		This Include File Needs Following Variables to have values:
'
'		SO_Action				(the Action of the submit button onclick)
'		SO_Customer				(AccountTitle or Name to be searched)
'		SO_StepText				(e.g. '░гЦ оФЦ : гДйнгх сщгят ЕгМ ЦяхФьЕ')
'
'
%>
		<div dir='rtl'><B><%=SO_StepText%></B>
		</div><br>
<!-- гДйнгх сщгят ЕгМ ЦяхФьЕ -->
<%
	
	SO_mySQL="SELECT * From Accounts WHERE (ID='"& SO_Customer & "')"
	Set SO_RS1 = conn.Execute(SO_mySQL)
	if (SO_RS1.eof) then
		response.write "Error ! No such customer"
		response.end
	else
		SO_AccountTitle=SO_RS1("AccountTitle")
	end if

'	SO_mySQL="SELECT * From Orders WHERE (Customer='"& SO_Customer & "') ORDER BY ID"
'	SO_mySQL="SELECT Orders.ID, Orders.CreatedDate FROM Orders LEFT OUTER JOIN InvoiceOrderRelations ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (InvoiceOrderRelations.Invoice IS NULL) AND (Orders.Customer = '"& SO_Customer & "') ORDER BY Orders.ID"
' Changed By Kid 82/08/18 
'-----------SAM change this
	'SO_mySQL="SELECT Orders.* From Orders LEFT OUTER JOIN InvoiceOrderRelations ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (Customer='"& SO_Customer & "') AND (Closed=0) GROUP BY Orders.ID, Orders.CreatedDate, Orders.Closed, Orders.Customer, Orders.CreatedBy HAVING COUNT(InvoiceOrderRelations.Invoice) < 1 ORDER BY Orders.ID"
	'----------SAM change this on 13 Mar 2011
	'SO_mySQL="SELECT Orders.* From Orders LEFT OUTER JOIN InvoiceOrderRelations ON Orders.ID = InvoiceOrderRelations.[Order] LEFT OUTER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (Orders.Customer='" & SO_Customer & "') AND (Orders.Closed=0) AND (ISNULL(Invoices.Voided,0) = 0) GROUP BY Orders.ID, Orders.CreatedDate, Orders.Closed, Orders.Customer, Orders.CreatedBy HAVING COUNT(InvoiceOrderRelations.Invoice) < 1 ORDER BY Orders.ID"
	SO_mySQL="SELECT * From Orders WHERE (Customer='" & SO_Customer & "') AND (isClosed=0) and ID not in (SELECT InvoiceOrderRelations.[Order] FROM InvoiceOrderRelations INNER JOIN Orders ON Orders.ID = InvoiceOrderRelations.[Order] LEFT OUTER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (ISNULL(Invoices.Voided,0) = 0) AND Orders.Customer='" & SO_Customer & "' GROUP BY InvoiceOrderRelations.[Order] HAVING COUNT(InvoiceOrderRelations.Invoice) > 0) ORDER BY Orders.ID" & thisOrder
	
	Set SO_RS1 = conn.Execute(SO_mySQL)

	if (SO_RS1.eof) then 		' Not Found	%>
		<table class="RcpTable" align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'>ЕМ█ сщгятМ │Мог Дто<br></td></tr></table><br>
<%	else
%>
	<br>
	<TABLE class="RcpTable" align="center" border="1" cellspacing="1" cellpadding="5" dir="RTL">
		<tr bgcolor='#DDDDEE'>
			<td align='center' colspan="3">сщгят ЕгМ ЦяхФь хЕ <br>'<%=SO_AccountTitle%>'<br>ъЕ чхАг щгъйФя ДтоЕ гДо</td>
		</tr>
		<tr bgcolor='#C3C3FF'>
			<td align='center' width="30"> <input type="checkbox" disabled checked> </td>
			<td align='center' width="70"> тЦгяЕ │М░МяМ</td>
			<td align='center' width="80"> йгяМн </td>
		</tr>
<%		SO_tempCounter=0
		while Not (SO_RS1.EOF)
			SO_tempCounter=SO_tempCounter+1
			if (SO_tempCounter Mod 2 = 1)then
				SO_tempColor="#FFFFFF"
			else
				SO_tempColor="#DDDDEE"
			end if
%>				<tr bgcolor='<%=SO_tempColor%>'>
					<td align='center'><input type="checkbox" name="selectedOrders" value="<%=SO_RS1("ID")%>">&nbsp;</td>
					<td dir='ltr' align='center'><%=Link2Trace(SO_RS1("ID"))%>&nbsp;</td>
					<td dir='ltr' align='center'><%=SO_RS1("CreatedDate")%>&nbsp;</td>
				</tr>
<%			SO_RS1.movenext
		wend
%>
		<tr bgcolor='#C3C3FF'>
			<td align='center' colspan="4"><input name="SO_SelectButton" class="GenButton" type="submit" value="гДйнгх" onclick="<%=SO_Action%>" >&nbsp;<input name="selectedCustomer" type="hidden" value="<%=SO_Customer%>"></td>
		</tr>
	</TABLE>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	//document.all.SO_SelectButton.focus();
	//-->
	</SCRIPT>
<%
	end if
%>
