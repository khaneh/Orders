<%
'		This Include File Needs Following Variables to have values:
'
'		SO_Action				(the Action of the submit button onclick)
'		SO_Order				(one Order ID )
'		SO_StepText				(e.g. '░гЦ оФЦ : гДйнгх сщгят ЕгМ ЦяхФьЕ')
'
'
%>
		<div dir='rtl'><B><%=SO_StepText%></B>
		</div><br>
<!-- гДйнгх сщгят ЕгМ ЦяхФьЕ -->
<%
	
	SO_mySQL="SELECT Orders.Customer, Accounts.AccountTitle FROM Orders INNER JOIN Accounts ON Orders.Customer = Accounts.ID WHERE (Orders.ID = '"& SO_Order & "')"
	Set SO_RS1 = conn.Execute(SO_mySQL)
	if (SO_RS1.eof) then
		response.redirect "?errmsg=" & Server.URLEncode("█ДМД сщгятМ ФлФо Догяо")
	else
		SO_Customer=SO_RS1("Customer")
		SO_AccountTitle=SO_RS1("AccountTitle")
	end if

'	SO_mySQL="SELECT * From Orders WHERE (Customer='"& SO_Customer & "') ORDER BY ID"
'	SO_mySQL="SELECT Orders.ID, Orders.CreatedDate FROM Orders LEFT OUTER JOIN InvoiceOrderRelations ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (InvoiceOrderRelations.Invoice IS NULL) AND (Orders.Customer = '"& SO_Customer & "') ORDER BY Orders.ID"
' Changed By Kid 82/08/18 
	SO_mySQL="SELECT * From Orders WHERE (Customer='"& SO_Customer & "') AND (Closed=0) ORDER BY ID"
	Set SO_RS1 = conn.Execute(SO_mySQL)
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
<%	SO_tempCounter=0
	SO_Orderfound=false
	while Not (SO_RS1.EOF)
		SO_tempCounter=SO_tempCounter+1
		if (SO_tempCounter Mod 2 = 1)then
			SO_tempColor="#FFFFFF"
		else
			SO_tempColor="#DDDDEE"
		end if
		SO_tmpOrderID=SO_RS1("ID")
%>		<tr bgcolor='<%=SO_tempColor%>'>
			<td align='center'><input type="checkbox" name="selectedOrders" value="<%=SO_tmpOrderID%>" <%if clng(SO_tmpOrderID)=clng(SO_Order) then SO_Orderfound=true : response.write "checked"%>>&nbsp;</td>
			<td dir='ltr' align='center'><%=Link2Trace(SO_tmpOrderID)%>&nbsp;</td>
			<td dir='ltr' align='center'><%=SO_RS1("CreatedDate")%>&nbsp;</td>
		</tr>
<%		SO_RS1.movenext
	wend
	SO_RS1.close
	if (SO_tempCounter=0) then 		' Not Found	ANY order%>
		<tr>
			<td align='center' colspan="4">
				<table class="RcpTable" align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'>ЕМ█ сщгятМ │Мог Дто<br></td></tr></table>
			</td>
		</tr>
<%	end if

	SO_mySQL="SELECT Invoices.Approved, Invoices.Issued, Invoices.ID FROM InvoiceOrderRelations INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (InvoiceOrderRelations.[Order] = '" & SO_Order & "') AND (Invoices.Voided = 0)"

	Set SO_RS1 = conn.Execute(SO_mySQL)
	if not(SO_RS1.eof) then 
		%><tr>
				<td align='center' colspan="4">
					<table class="RcpTable" align='center' cellpadding='5' style="border:2 solid #333366">
					<%
		do while not(SO_RS1.eof)
			SO_FoundInvoice=SO_RS1("ID")
			if SO_RS1("Issued") then 
				SO_tpDesc="<B>угоя тоЕ</B>"
				SO_tmpColor="#FF6666"
			elseif SO_RS1("Approved") then 
				SO_tpDesc="<B>йгММо тоЕ</B>"
				SO_tmpColor="Yellow"
			else
				SO_tpDesc="<B>йгММо ДтоЕ</B>"
				SO_tmpColor="#FFFFBB"
			end if
	%>
			
					<tr><td bgcolor='<%=SO_tmpColor%>' dir='rtl' align='center'>
						йФлЕ!<br>сщгятМ ъЕ тЦг лсйлФ ъяоМо чхАг оя М≤ щгъйФя <%=SO_tpDesc%> ФлФо огяо<br>
						<A HREF="AccountReport.asp?act=showInvoice&invoice=<%=SO_FoundInvoice%>">ДЦгМт щгъйФя ЦяхФьЕ (<%=SO_FoundInvoice%>)</A>
						</td>
					</tr>
					
		
<%	SO_RS1.movenext
		loop
		%></table>
				</td>
			</tr>
			<%
	end if
%>
		<tr bgcolor='#C3C3FF'>
			<td align='center' colspan="4"><input name="SO_SelectButton" class="GenButton" type="submit" value="гДйнгх" onclick="<%=SO_Action%>" >&nbsp;<input name="selectedCustomer" type="hidden" value="<%=SO_Customer%>"></td>
		</tr>
	</TABLE>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	if (! window.dialogArguments){// if this is NOT a modal window
		document.all.SO_SelectButton.focus();
	}
	//-->
	</SCRIPT>
