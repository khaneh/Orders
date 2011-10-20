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
'	SO_mySQL="SELECT Orders.ID, Orders.CreatedDate FROM Orders LEFT OUTER JOIN InvoiceOrderRelations ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (InvoiceOrderRelations.Invoice IS NOT NULL) AND (Orders.Customer = '"& SO_Customer & "') ORDER BY Orders.ID"
' Changed By Kid 82/09/16
	SO_mySQL="SELECT * From Orders WHERE (Customer='"& SO_Customer & "') AND (Closed=1) ORDER BY ID"
	Set SO_RS1 = conn.Execute(SO_mySQL)

	if (SO_RS1.eof) then 		' Not Found	%>
		<table class="RcpTable" align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'>ЕМ█ сщгятМ │Мог Дто<br></td></tr></table><br>
<%	else
%>
	<br>
	<TABLE class="RcpTable" align="center" border="1" cellspacing="1" cellpadding="5" dir="RTL">
		<tr bgcolor='#DDDDEE'>
			<td align='center' colspan="3">сщгят ЕгМ ЦяхФь хЕ <br>'<%=SO_AccountTitle%>'<br>ъЕ чхАг щгъйФя тоЕ гДо</td>
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