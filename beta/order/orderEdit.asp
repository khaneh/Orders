<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ÊÌ—«Ì‘ ”›«—‘"
SubmenuItem=2
if not Auth(2 , 2) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------- Order Search Form
'-----------------------------------------------------------------------------------------------------

radif = request("radif")


if radif <> "" then
	if isNumeric(radif) then
		radif=clng(radif)
'		if radif=0 then radif=""
	else
		response.write "<br><br>"
		call showAlert("‘„«—Â ”›«—‘ «‘ »«Â «”  ·ÿ›« œÊ»«—Â Ê«—œ ﬂ‰Ìœ.", CONST_MSG_ERROR )
		radif=""
	end if
end if

if radif="" then
%>
<br><br>
<FORM METHOD=POST ACTION="?e=y">
<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" style="font-family:Tahoma,arial; font-size: 8pt;">
<TR>
	<TD>‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ:</TD>
	<TD><INPUT Name="Radif" TYPE="text" maxlength="6" size="5" tabIndex="1" dir="LTR"></TD>
	<TD><INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" style="width:100px;font-family:Tahoma,arial; font-size: 8pt;"></TD>
</TR>
</TABLE>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.Radif.focus();
//-->
</script>
<%
	response.end
end if


'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------- Submit Changes In An Order
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")=" «ÌÌœ" AND radif<>"" then

	orderType=request.form("OrderType")
	if orderType="" OR not isNumeric(OrderType) then
		conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
	else
		set RS1=Conn.Execute ("SELECT Name FROM OrderTraceTypes WHERE (ID='"& orderType & "') AND (IsActive=1)")
		if RS1.eof then
			conn.close
			response.redirect "?errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œ...<BR>")
		else
			orderTypeName=RS1("Name")
		end if
		RS1.close
	end if

	set RS1=Conn.Execute ("SELECT Name FROM OrderTraceStatus WHERE (IsActive=1) and ID = " & request.form("Vazyat"))
	statusName = RS1("name")

	set RS1=Conn.Execute ("SELECT Name FROM OrderTraceSteps WHERE (IsActive=1) and ID = " & request.form("Marhale"))
	stepName = RS1("name")

	set RS3=Conn.Execute ("SELECT * FROM Orders WHERE (ID = "& radif & ")")

OrderDate =		sqlSafe(request.form("OrderDate"))
OrderTime =		sqlSafe(request.form("OrderTime"))
ReturnDate =	sqlSafe(request.form("ReturnDate"))
ReturnTime =	sqlSafe(request.form("ReturnTime"))
CompanyName	=	sqlSafe(request.form("CompanyName"))
CustomerName =	sqlSafe(request.form("CustomerName"))
Telephone =		sqlSafe(request.form("Telephone"))
OrderTitle =	sqlSafe(request.form("OrderTitle"))
Vazyat =		sqlSafe(request.form("Vazyat"))
Marhale =		sqlSafe(request.form("Marhale"))
SalesPerson =	sqlSafe(request.form("SalesPerson"))
Qtty =			sqlSafe(request.form("Qtty"))
Size =			sqlSafe(request.form("Size"))
SimplexDuplex =	sqlSafe(request.form("SimplexDuplex"))
Price =			sqlSafe(request.form("Price"))

	mySql="UPDATE orders_trace SET order_date= N'"& OrderDate & "', order_time= N'"& OrderTime & "', return_date= N'"& ReturnDate & "', return_time= N'"& ReturnTime & "', company_name= N'"& CompanyName & "', customer_name= N'"& CustomerName & "', telephone= N'"& Telephone & "', order_title= N'"& OrderTitle & "', order_kind= N'"& orderTypeName & "', Type= '"& orderType & "', vazyat= N'"& statusName & "', status= "& Vazyat & ", step= "& Marhale & ",  marhale= N'"& stepName & "', salesperson= N'"& SalesPerson & "', qtty= N'"& Qtty & "', paperSize= N'"& Size & "', SimplexDuplex= N'"& SimplexDuplex & "', Price= N'"& Price & "' , LastUpdatedDate=N'"& shamsitoday() & "' , LastUpdatedTime=N'"& currentTime10() & "', LastUpdatedBy=N'"& session("ID")& "'  WHERE (radif_sefareshat= N'"& radif & "')"	
	conn.Execute mySql
	response.write radif &" UPDATED<br>"
	response.write "<A HREF='orderEdit.asp'>Back</A>"

	if RS3("Customer") <> request.form("CustomerID") then
		mySql="UPDATE Orders SET Customer='"& request.form("CustomerID") & "' WHERE (ID= N'"& radif & "')"	
		conn.Execute mySql
	end if

	if request.form("Marhale")="10" then
		call InformCSRorderIsReady(radif , RS3("CreatedBy"))
	end if

	if not request.form("relatedApprovedInvoiceID")="0" then
		call UnApproveInvoice(request.form("relatedApprovedInvoiceID") , request.form("relatedApprovedInvoiceBy"))
	end if

	response.redirect "../order/TraceOrder.asp?act=show&order=" & radif
	
end if

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------  View And Edit Order's Details
'-----------------------------------------------------------------------------------------------------
set RS2=Conn.Execute ("SELECT orders_trace.*, Accounts.ID AS AccID, Accounts.AccountTitle FROM Orders INNER JOIN Accounts ON Orders.Customer = Accounts.ID INNER JOIN orders_trace ON Orders.ID = orders_trace.radif_sefareshat WHERE (orders_trace.radif_sefareshat='"& radif & "')")

if RS2.eof then %>
	<FORM METHOD=POST ACTION="?"><BR><BR>
	<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" align="center" style="font-family:Tahoma,arial; font-size: 8pt;">
	<TR>
		<TD>‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ:</TD>
		<TD><INPUT Name="Radif" TYPE="text" maxlength="6" size="6" tabIndex="1" dir="LTR"></TD>
		<TD><INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" style="width:100px;font-family:Tahoma,arial; font-size: 8pt;"></TD>
	</TR>
	</TABLE>
	</FORM>
	<br>
	<%	
	call showAlert("«„ﬂ«‰ ‰œ«—œ ﬂÂ Â„çÌ‰ ”›«—‘Ì ÊÃÊœ œ«‘ Â »«‘œ.", CONST_MSG_ERROR )
	response.end 
else
	CustomerID=RS2("AccID")
end if

if RS2("salesperson")<>session("csrName") then 
	response.write "<br>"
	call showAlert("”›«—‘ êÌ—‰œÂ «Ì‰ ”›«—‘ ‘„« ‰Ì” Ìœ..<BR>ﬁ»· «“ ÊÌ—«Ì‘ »« ”›«—‘ êÌ—‰œÂ Â„«Â‰ê ò‰Ìœ.", CONST_MSG_ALERT ) 
end if
%>
<style>
	Table { font-size: 8pt;}
	INPUT { font-family:Tahoma,arial; font-size: 8pt;}
	SELECT { font-family:Tahoma,arial; font-size: 8pt;}
</style>
<font face="tahoma">
<%
'mySQL="SELECT Invoices.id, Invoices.ApprovedBy FROM Orders LEFT OUTER JOIN Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice ON Orders.ID = InvoiceOrderRelations.[Order] WHERE (ISNULL(Invoices.approved, 0) = 1) AND (Orders.ID ='"& request("radif") & "')"
'Changed By kid 820903

relatedApprovedInvoiceID = 0

mySQL="SELECT Invoices.Issued, Invoices.Approved, Invoices.ApprovedBy, Invoices.ID FROM InvoiceOrderRelations INNER JOIN Invoices ON InvoiceOrderRelations.Invoice = Invoices.ID WHERE (InvoiceOrderRelations.[Order] = '" & radif & "') AND (Invoices.Voided = 0)"
set RS3=Conn.Execute (mySQL)
if not(RS3.eof) then 
	FoundInvoice=RS3("ID")
	if RS3("Issued") then 
		Conn.Close
		response.redirect "../order/TraceOrder.asp?act=show&order=" & radif & "&errMsg=" & Server.URLEncode("Ìﬂ ›«ﬂ Ê—  ’«œ— ‘œÂ „— »ÿ »« »« «Ì‰  ”›«—‘ ‘„«—Â ÅÌœ« ﬂ—œÂ «Ì„.(‘„«—Â ›«ﬂ Ê—: <A HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& FoundInvoice & "' target='_blank'>" & FoundInvoice & "</a>)<BR> ·–«”  ﬂÂ «„ﬂ«‰Ì »—«Ì  €ÌÌ— œ— «Ì‰ ”›«—‘ ÊÃÊœ ‰œ«—œ.")
	elseif RS3("Approved") then 
		tmpDesc="<B> «ÌÌœ ‘œÂ</B>"
		tmpColor="Yellow"
		relatedApprovedInvoiceID = RS3("id")
		relatedApprovedInvoiceBy = RS3("ApprovedBy")
		response.write "<br>"
		call showAlert("Ìﬂ ›«ﬂ Ê— <b> «ÌÌœ ‘œÂ</b> »« ‘„«—Â <A HREF='../AR/AccountReport.asp?act=showInvoice&invoice="& FoundInvoice &"' target='_blank'>" & FoundInvoice & "</A> »—«Ì «Ì‰ ”›«—‘ ÊÃÊœ œ«—œ <BR>ﬂÂ »«  €ÌÌ— ”›«—‘  Ê”ÿ ‘„« «“ Õ«·   «ÌÌœ Œ«—Ã ŒÊ«Âœ ‘œ." , CONST_MSG_ALERT )  
	else
		tmpDesc=""
		tmpColor="#FFFFBB"
		response.write "<br>"
'		call showAlert(" ÊÃÂ!<br>”›«—‘Ì ﬂÂ ‘„« Ã” ÃÊ ﬂ—œÌœ ﬁ»·« œ— Ìò ›«ﬂ Ê— <B> «ÌÌœ ‰‘œÂ</B> ÊÃÊœ œ«—œ<br><A HREF='../AR/InvoiceEdit.asp?act=editInvoice&invoice=" & FoundInvoice & "'>«’·«Õ ›«ﬂ Ê— „—»ÊÿÂ (" & FoundInvoice & ")</A>", CONST_MSG_INFORM ) 
		call showAlert(" ÊÃÂ!<br>”›«—‘Ì ﬂÂ ‘„« Ã” ÃÊ ﬂ—œÌœ ﬁ»·« œ— Ìò ›«ﬂ Ê— <B> «ÌÌœ ‰‘œÂ</B> ÊÃÊœ œ«—œ<br><A HREF='../AR/AccountReport.asp?act=showInvoice&invoice=" & FoundInvoice & "'>‰„«Ì‘ ›«ﬂ Ê— „—»ÊÿÂ (" & FoundInvoice & ")</A>", CONST_MSG_INFORM ) 
	end if
end if 
%>

<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>
<FORM METHOD=POST ACTION="?" onSubmit="return checkValidation();">
	
<INPUT TYPE="hidden" name="relatedApprovedInvoiceBy" value="<%=relatedApprovedInvoiceBy%>">
<INPUT TYPE="hidden" name="relatedApprovedInvoiceID" value="<%=relatedApprovedInvoiceID%>">
<br>
<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center">
<TR bgcolor="black">
	<TD align="left"><FONT COLOR="YELLOW">Õ”«»:</FONT></TD>
	<TD align="right" colspan=5 height="25px">
		<span id="customer" style="color:yellow;"><%' after any changes in this span "./Customers.asp" must be revised%>
			<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>"><span><%=customerID & " - "& RS2("AccountTitle")%></span>.
		</span>
		<INPUT class="GenButton" TYPE="button" value=" €ÌÌ—" onClick="selectCustomer();">
	</TD>
</TR>
<TR bgcolor="black">
	<TD align="left"><FONT COLOR="YELLOW">‘„«—Â ”›«—‘:</FONT></TD>
	<TD align="right">
		<!-- Radif -->
		<INPUT TYPE="text" disabled maxlength="6" size="5" tabIndex="1" dir="LTR" value="<%=RS2("radif_sefareshat")%>">
		<INPUT TYPE="hidden" NAME="Radif" value="<%=RS2("radif_sefareshat")%>">
	</TD>
	<TD align="left"><FONT COLOR="YELLOW"> «—ÌŒ:</FONT></TD>
	<TD><TABLE border="0">
		<TR>
			<TD dir="LTR">
			<INPUT disabled TYPE="text" maxlength="10" size="8"  value="<%=RS2("order_date")%>">
			<INPUT TYPE="hidden" NAME="OrderDate" value="<%=RS2("order_date")%>">
			</TD>
			<TD dir="RTL"><FONT COLOR="YELLOW"><%="ø‘‰»Â"%></FONT></TD>
		</TR>
		</TABLE></TD>
	<TD align="left"><FONT COLOR="YELLOW">”«⁄ :</FONT></TD>
	<TD align="right">
	<INPUT disabled TYPE="text" maxlength="5" size="3" dir="LTR" value="<%=RS2("order_time")%>">
	<INPUT TYPE="hidden" NAME="OrderTime" value="<%=RS2("order_time")%>"></TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left">‰«„ ‘—ﬂ :</TD>
	<TD align="right">
		<!-- CompanyName -->
		<INPUT TYPE="text" NAME="CompanyName" maxlength="50" size="25" tabIndex="2"  value="<%=RS2("company_name")%>"></TD>
	<TD align="left">„Ê⁄œ  ÕÊÌ·:</TD>
	<TD><TABLE border="0">
		<TR>
			<TD dir="LTR"><INPUT TYPE="text" NAME="ReturnDate"  onblur="acceptDate(this)" maxlength="10" size="8" tabIndex="5" onKeyPress="return maskDate(this);" value="<%=RS2("return_date")%>"></TD>
			<TD dir="RTL">(?‘‰»Â)</TD>
		</TR>
		</TABLE></TD>
	<TD align="left">”«⁄   ÕÊÌ·:</TD>
	<TD align="right"><INPUT TYPE="text" NAME="ReturnTime" maxlength="6" size="3" tabIndex="6" dir="LTR" onKeyPress="return maskTime(this);" value="<%=RS2("return_time")%>"></TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left">‰«„ „‘ —Ì:</TD>
	<TD align="right">
		<!-- CustomerName -->
		<INPUT TYPE="text" NAME="CustomerName" maxlength="50" size="25" tabIndex="3" value="<%=RS2("customer_name")%>"></TD>
	<TD align="left">‰Ê⁄ ”›«—‘:</TD>
	<TD>
		<SELECT NAME="OrderType" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="7">
<%
		thisOrderType=RS2("Type")
		set RS_TYPE=Conn.Execute ("SELECT ID, Name FROM OrderTraceTypes WHERE (IsActive=1) ORDER BY ID")
		Do while not RS_TYPE.eof	
%>
			<OPTION value="<%=RS_TYPE("ID")%>" <%if thisOrderType=RS_TYPE("ID") then response.write "selected" %> ><%=RS_TYPE("Name")%></option>
<%
			RS_TYPE.moveNext
		loop
		RS_TYPE.close
		set RS_TYPE = nothing
%>		
		</SELECT></TD>
	<TD align="left">”›«—‘ êÌ—‰œÂ:</TD>
	<TD><INPUT NAME="SalesPerson" Type="TEXT"value="<%=RS2("salesperson")%>" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 100px' tabIndex="88" readonly>
	</TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left"> ·›‰:</TD>
	<TD align="right">
		<!-- Telephone -->
		<INPUT TYPE="text" NAME="Telephone" maxlength="50" size="25" tabIndex="4" value="<%=RS2("telephone")%>"></TD>
	<TD align="left">⁄‰Ê«‰ ﬂ«— œ«Œ· ›«Ì·:</TD>
	<TD align="right" colspan="4"><INPUT TYPE="text" NAME="OrderTitle" maxlength="255" size="50" tabIndex="9" value="<%=RS2("order_title")%>"></TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD align="left">„—Õ·Â:</TD>
	<TD colspan="2"><SELECT NAME="Marhale" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 140px' tabIndex="13" >
	<%
	set RS_STEP=Conn.Execute ("SELECT * FROM OrderTraceSteps WHERE (IsActive=1)")
	Do while not RS_STEP.eof	
	%>
		<OPTION value="<%=RS_STEP("ID")%>" <%if RS2("step")=RS_STEP("ID") then response.write "selected" %> ><%=RS_STEP("name")%></option>
		<%
		RS_STEP.moveNext
	loop
	RS_STEP.close
	set RS_STEP = nothing
	%>
	</SELECT></TD>
	<TD align="left">Ê÷⁄Ì :</TD>
	<TD colspan="2"><SELECT NAME="Vazyat" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 100px' tabIndex="14" >
	<%
	set RS_STATUS=Conn.Execute ("SELECT * FROM OrderTraceStatus WHERE (IsActive=1)")
	Do while not RS_STATUS.eof	
	%>
		<OPTION value="<%=RS_STATUS("ID")%>" <%if RS2("status")=RS_STATUS("ID") then response.write "selected" %> ><%=RS_STATUS("Name")%></option>
		<%
		RS_STATUS.moveNext
	loop
	RS_STATUS.close
	set RS_STATUS = nothing
	%>
	</SELECT></TD>
</TR>
<TR bgcolor="#CCCCCC">
	<TD colspan="6">

	<TABLE align="center" width="50%" border="0">
	<TR>
		<TD><INPUT TYPE="submit" Name="Submit" tabIndex="16" Value=" «ÌÌœ" onFocus="noNextField=true;" onBlur="noNextField=false;" style="width:100px;"></TD>
		<TD>&nbsp;</TD>
		<TD align="left"><INPUT TYPE="button" Name="Cancel" tabIndex="17" Value="«‰’—«›" style="width:100px;" onClick="window.location='orderEdit.asp';"></TD>
	</TR>
	</TABLE>

	<BR>
	<TABLE cellspacing=0 width=80% align=center>
	<TR bgcolor=white>
		<TD> «—ÌŒ</TD>
		<TD>”«⁄ </TD>
		<TD>„—Õ·Â</TD>
		<TD>Ê÷⁄Ì </TD>
		<TD>‰«„ À»  ﬂ‰‰œÂ</TD>
	</TR>
	<%
	set RS_STEP=Conn.Execute ("SELECT OrderTraceLog.*, Users.RealName FROM OrderTraceLog INNER JOIN Users ON OrderTraceLog.InsertedBy = Users.ID WHERE (OrderTraceLog.[Order] = "& radif  & ") order by OrderTraceLog.ID")
	Do while not RS_STEP.eof	
	%>
	<TR>
		<TD  style="border-bottom: solid 1pt black" dir=ltr align=right><%=RS_STEP("InsertedDate")%> </TD>
		<TD  style="border-bottom: solid 1pt black" dir=ltr align=right>(<%=RS_STEP("InsertedTime")%>)</TD>
		<TD  style="border-bottom: solid 1pt black"><%=RS_STEP("StepText")%></TD>
		<TD  style="border-bottom: solid 1pt black"><%=RS_STEP("StatusText")%></TD>
		<TD  style="border-bottom: solid 1pt black"><%=RS_STEP("RealName")%></TD>
	</TR>
		<%
		RS_STEP.moveNext
	loop
	RS_STEP.close
	set RS_STEP = nothing
	%>
	</TABLE>
	</TD>
</TR>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.CompanyName.focus();

var noNextField = false;

function documentKeyDown() {
	var theKey = event.keyCode;
	var o = window.event.srcElement;
	if (theKey == 9) { 
		if (noNextField)			// Don't Move
			if(event.shiftKey){
				noNextField=false;
				event.keyCode=9
			}
			else
			return false; 
		else {						// send focus to next box

			if (false){
			}
			else if(o.name == 'Radif'){
// Radif 
				if (o.value.length<5){
					return false;
				}
			}
			else if(o.name == 'ReturnTime'){
// Return Time
				var soutput = o.value;
				var len = o.value.length; 
				if (len==0){
					return true;
				}
				else if (len==3){
					soutput = soutput + '00'
					o.value = soutput;
				}
				else if (len==4){
					var tempString =soutput.charAt(len-1);
					soutput = soutput.substring(0,len-1) + '0' + tempString
					o.value = soutput;
				}
				else if(len!=5)
					return false;
			}
			event.keyCode=9
			return true;
		}
	}
}

document.onkeydown = documentKeyDown;

function checkValidation(){
	return true;
}

function selectCustomer(){
	document.all.tmpDlgArg.value="#"
	document.all.tmpDlgTxt.value="‰«„ Õ”«»Ì —« ﬂÂ „Ì ŒÊ«ÂÌœ Ã” ÃÊ ﬂ‰Ìœ Ê«—œ ﬂ‰Ìœ:"
	window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
	if (document.all.tmpDlgTxt.value !="") {
		window.showModalDialog('../AR/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogWidth:780px; dialogHeight:500px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
		if (document.all.tmpDlgArg.value!="#"){
			Arguments=document.all.tmpDlgArg.value.split("#")
			theSpan=document.getElementById("customer");
			theSpan.getElementsByTagName("input")[0].value=Arguments[0];
			theSpan.getElementsByTagName("span")[0].innerText=Arguments[0]+" - "+Arguments[1];
		}
	}
}

//-->
</SCRIPT>
</TABLE>
<br>
</font>

<!--#include file="tah.asp" -->