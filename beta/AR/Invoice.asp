<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="‰„«Ì‘ ›«ﬂ Ê—"
SubmenuItem=1
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
function ShowErrorMessage(msg)
	response.write "<table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> Œÿ« ! <br>"& msg & "<br></td></tr></table><br>"
end function

function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/TraceOrder.asp?act=show&order="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

%>
<style>
	Table { font-size: 9pt;}
	.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0FFF0; text-align:right;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
	.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=0;
//-->
</SCRIPT>
<%

if request("act")="submitsearch" then
'	response.redirect "AccountReport.asp?act=showInvoice&invoice="	 & InvoiceID
	if isnumeric(request("invoice")) then
		response.redirect "AccountReport.asp?act=showInvoice&invoice=" & request("invoice")
	elseif trim(request("query")) = "" then
		Conn.close
		response.redirect "?errmsg=" & Server.URLEncode("Œ«·Ì ”—ç ‰„Ìﬂ‰Ì„!")
	end if
	
	if isnumeric(request("query")) then
		'User has entered an ORDER NUMBER
		OrderID=clng(request("query"))
		mySQL="SELECT InvoiceOrderRelations.Invoice FROM InvoiceOrderRelations INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (InvoiceOrderRelations.[Order] = '"& OrderID & "') AND (Invoices.IsReverse = 0) AND (Invoices.Voided = 0)"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open mySQL, Conn, 3, 3
		if rs.eof then
			rs.close
			mySQL="SELECT InvoiceOrderRelations.Invoice FROM InvoiceOrderRelations INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (InvoiceOrderRelations.[Order] = '"& OrderID & "') AND (Invoices.IsReverse = 0)"
			rs.open mySQL, Conn, 3, 3
			if not rs.eof then
				InvoiceID=rs("Invoice")
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID
			else
				Conn.close
				response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ”›«—‘ œ— ÂÌç ›«ﬂ Ê—Ì ÅÌœ« ‰‘œ.")
			end if
		else
			if rs.RecordCount>1 then
				tempWriteAnd=""
				invoiceList="«Ì‰ ”›«—‘ œ— ç‰œ ›«ﬂ Ê— “Ì— „ÊÃÊœ «” :<br>"
				Do While not rs.eof 
					invoiceList=invoiceList & tempWriteAnd & "<A HREF='AccountReport.asp?act=showInvoice&invoice="& rs("invoice") &"' target='_blank'>" & rs("invoice") & "</A>"
					tempWriteAnd=" Ê "
					rs.moveNext
				Loop 
				invoiceList=invoiceList & "<br>çÂ ﬂ«— ﬂ‰Ì„ø"
				response.write "<br><br>"
				call showAlert (invoiceList ,CONST_MSG_ALERT) 
			else
				InvoiceID=rs("Invoice")
				Conn.close
				response.redirect "AccountReport.asp?act=showInvoice&invoice=" & InvoiceID
			end if
		end if
	else
		'User has entered an ACCOUNT NAME
		SA_TitleOrName=request("query")
		SA_Action="return true;"
		SA_SearchAgainURL="InvoiceInput.asp"
		SA_StepText="ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<FORM METHOD=POST ACTION="?act=showInvoices">
		<!--#include File="include_SelectAccount.asp"-->
		</FORM>
<%
	end if
elseif request("act")="showInvoices" then
	cusID=request("selectedCustomer")
	if cusID <> "" and isnumeric(cusID) then
		mySQL="SELECT TOP 1 AccountTitle FROM Accounts WHERE (ID = '"& clng(cusID) & "')"
		set RS1=Conn.execute(mySQL)
		AccountTitle = RS1("AccountTitle")
%>
		<br><br>
		<!--#include file="include_CustomerInvoices.asp" -->
<%
	end if
else
%>
	<br>
	<br>
	<FORM METHOD=POST ACTION="?act=submitsearch">
	<div dir='rtl'><B><FONT SIZE="" COLOR="red"> &nbsp;‰„«Ì‘ ›«ﬂ Ê—: </FONT><BR>Ã” ÃÊ »—«Ì ‰«„ Õ”«» Ì« ‘„«—Â ”›«—‘</B>
		<INPUT TYPE="text" NAME="query" onfocus="document.getElementsByName('invoice')[0].value='';">&nbsp;
		<INPUT class="GenButton" TYPE="submit" Name="submitShow" value="Ã” ÃÊ">&nbsp;&nbsp;&nbsp;‘„«—Â ›«ﬂ Ê—:
		<INPUT style="font-family:Tahoma;width:100px;" TYPE="text" NAME="invoice">&nbsp;
	</div>
	</FORM>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» »« ‘„«—Â ”›«—‘-->
	<hr width="90%" align="center">

	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.getElementsByName("query")[0].focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
<!--#include file="tah.asp" -->