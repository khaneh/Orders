<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
PageTitle= "ç«Å çò "
SubmenuItem=1
if not Auth("A" , 1) then NotAllowdToViewThisPage()

sendTo = session("id")
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<style>
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<%
	ID = cint(request("PaidChequesID"))
	if ID > 0 then
		mySQL="SELECT PaidCheques.*, Accounts.AccountTitle, Bankers.Name as BankOfOrigin FROM PaidCheques INNER JOIN Bankers ON PaidCheques.Banker = Bankers.ID INNER JOIN Payments ON PaidCheques.Payment = Payments.ID INNER JOIN Accounts ON Payments.Account = Accounts.ID WHERE (PaidCheques.ID = "& ID & ")"
		Set RS1 = conn.Execute(mySQL)
		if not(RS1.eof) then
%>
			<br><br>
			<FORM METHOD=POST ACTION="?PaidChequesID=<%=ID%>&act=Print">
				<TABLE align='center'>
				<TR>
					<TD> «—ÌŒ çﬂ </TD>
					<TD> : </TD>
					<TD><%=RS1("chequeDate")%></TD>
				</TR>
				<TR>
					<TD>„»·€</TD>
					<TD> : </TD>
					<TD><%=Separate(RS1("Amount"))%></TD>
				</TR>
				<TR>
					<TD>œ— ÊÃÂ</TD>
					<TD> : </TD>
					<TD><input type='text' name='AccountTitle' value='<% if request.form("AccountTitle") = "" then response.write(RS1("AccountTitle")) else response.write(request.form("AccountTitle"))%>' size='80'></TD>
				</TR>
				<TR>
					<TD>»«» </TD>
					<TD> : </TD>
					<TD><input type='text' name='Description' value='<% if request.form("Description") = "" then response.write(RS1("Description")) else response.write(request.form("Description"))%>' size='80'></TD>
				</TR>
					<TD>ÕÊ«·Â ﬂ—œ </TD>
					<TD> : </TD>
					<TD><input type='checkbox' name='havaleh' <% if request.form("havaleh") = "on" then response.write("checked") else response.write(request.form("havaleh")) %>></TD>
				</TR>
				<TR>
					<TD>‘„«—Â çò</TD>
					<TD> : </TD>
					<TD><%=RS1("ChequeNo")%></TD>
				</TR>
				<TR>
					<TD> «—ÌŒ —”Ìœ</TD>
					<TD> : </TD>
					<TD><input align='right' style="direction:ltr;" type='text' name='reciptDate' value='<% if request.form("reciptDate") = "" then response.write(RS1("chequeDate")) else response.write(request.form("reciptDate"))%>' size='80'></TD>
				</TR>
				<TR>
					<TD>»«‰ò</TD>
					<TD> : </TD>
					<TD>
						<%=RS1("BankOfOrigin")%>
						<input type='hidden' name='banker' value='<%=RS1("Banker")%>'>
					</TD>
				</TR>
				<TR>
					<TD colspan='3' align='center'>
						<br/>
						<input class='GenButton' name='action' type='submit' value='ç«Å »— êÂ çò'>
						<input class='GenButton' name='action' type='submit' value='ç«Å »—êÂ —”Ìœ'>
						
					</TD>
				</TR>
				<TR>
					<TD colspan='3' align='center'>
						<br/>
						<A HREF="/beta/AP/AccountReport.asp?act=showPayment&payment=<%=RS1("Payment")%>">»«“ê‘  »Â ’›ÕÂ Å—œ«Œ </A>
					</TD>
				</TR>
				</TABLE>
			</FORM>

<% 
		End if
		RS1.close
	end if
	if request("act") = "Print" then
		if request.form("action") = "ç«Å »—êÂ —”Ìœ" then 
			AccountTitle =	request.form("AccountTitle")
			Description =	request.form("Description")
			reciptDate =	request.form("reciptDate")
			ID =			cint(request("PaidChequesID"))
	%>
			<BR>
			<BR>
			<CENTER>
			<% 	ReportLogRow = PrepareReport ("chequeRecipt.rpt", "IDaccountTitleDescriptionreciptDate", ID & "" & accountTitle & "" & Description & "" & reciptDate , "/beta/dialog_printManager.asp?act=Fin") %>
			<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

			</CENTER>

			<BR><iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no ></iframe>
	<%	else 'request.form("action") = "ç«Å »— êÂ çò" then 
			AccountTitle =	request.form("AccountTitle")
			Description =	request.form("Description")
			reciptDate =	request.form("reciptDate")
			ID =			cint(request("PaidChequesID"))
			banker =		request.form("Banker")
			if request.form("havaleh") = "on" then 
				havaleh = 1
			else 
				havaleh = 0
			end if
			'response.write(havaleh)
			'response.end
	%>
			<BR>
			<BR>
			<CENTER>
			<% 	ReportLogRow = PrepareReport ("chequePrint" & banker & ".rpt", "IDaccountTitleDescriptionhavaleh", ID & "" & accountTitle & "" & Description & "" & havaleh , "/beta/dialog_printManager.asp?act=Fin") %>
			<INPUT TYPE="button" value=" ç«Å " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">

			</CENTER>

			<BR><iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:410; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no ></iframe>
	<%	end if
	end if
%>
 <!--#include file="tah.asp" -->