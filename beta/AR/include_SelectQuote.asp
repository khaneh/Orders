<%
'		This Include File Needs Following Variables to have values:
'
'		SQ_Action				(the Action of the submit button onclick)
'		SQ_Customer				(AccountTitle or Name to be searched)
'		SQ_StepText				(e.g. 'Çã Ïæã : ÇäÊÎÇÈ ÇÓÊÚáÇã åÇí ãÑÈæØå')
'
'

function Link2TraceQuote(QuoteNo)
	Link2TraceQuote = "<A HREF='../order/Inquiry.asp?act=show&quote="& QuoteNo & "' target='_balnk'>"& QuoteNo & "</A>"
end function
%>
		<div dir='rtl'><B><%=SQ_StepText%></B>
		</div><br>
<!-- ÇäÊÎÇÈ ÇÓÊÚáÇã åÇí ãÑÈæØå -->
<%
	
	SQ_mySQL="SELECT * From Accounts WHERE (ID='"& SQ_Customer & "')"
	Set SQ_RS1 = conn.Execute(SQ_mySQL)
	if (SQ_RS1.eof) then
		response.write "Error ! No such customer"
		response.end
	else
		SQ_AccountTitle=SQ_RS1("AccountTitle")
	end if

	SQ_mySQL="SELECT * From Quotes WHERE (Customer='"& SQ_Customer & "') AND (Closed=0) ORDER BY ID"
	Set SQ_RS1 = conn.Execute(SQ_mySQL)

	if (SQ_RS1.eof) then 		' Not Found	%>
		<table class="RcpTable" align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'>åí ÇÓÊÚáÇãí íÏÇ äÔÏ<br></td></tr></table><br>
<%	else
%>
	<br>
	<TABLE class="RcpTable" align="center" border="1" cellspacing="1" cellpadding="5" dir="RTL">
		<tr bgcolor='#DDDDEE'>
			<td align='center' colspan="3">ÇÓÊÚáÇã åÇí ãÑÈæØ Èå <br>'<%=SQ_AccountTitle%>'<br>ßå ÞÈáÇ ÝÇßÊæÑ äÔÏå ÇäÏ</td>
		</tr>
		<tr bgcolor='#C3C3FF'>
			<td align='center' width="30"> <input type="checkbox" disabled checked> </td>
			<td align='center' width="70"> ÔãÇÑå ííÑí</td>
			<td align='center' width="80"> ÊÇÑíÎ </td>
		</tr>
<%		SQ_tempCounter=0
		while Not (SQ_RS1.EOF)
			SQ_tempCounter=SQ_tempCounter+1
			if (SQ_tempCounter Mod 2 = 1)then
				SQ_tempColor="#FFFFFF"
			else
				SQ_tempColor="#DDDDEE"
			end if
%>				<tr bgcolor='<%=SQ_tempColor%>'>
					<td align='center'><input type="checkbox" name="selectedQuotes" value="<%=SQ_RS1("ID")%>">&nbsp;</td>
					<td dir='ltr' align='center'><%=Link2TraceQuote(SQ_RS1("ID"))%>&nbsp;</td>
					<td dir='ltr' align='center'><%=SQ_RS1("CreatedDate")%>&nbsp;</td>
				</tr>
<%			SQ_RS1.movenext
		wend
%>
		<tr bgcolor='#C3C3FF'>
			<td align='center' colspan="4"><input name="SQ_SelectButton" class="GenButton" type="submit" value="ÇäÊÎÇÈ" onclick="<%=SQ_Action%>" >&nbsp;<input name="selectedCustomer" type="hidden" value="<%=SQ_Customer%>"></td>
		</tr>
	</TABLE>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	//document.all.SQ_SelectButton.focus();
	//-->
	</SCRIPT>
<%
	end if
%>