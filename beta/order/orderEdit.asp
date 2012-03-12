<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle="ÊÌ—«Ì‘ ”›«—‘"
SubmenuItem=2
if not Auth(2 , 2) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
	.mySection{border: 1px #F90 dashed;margin: 15px 10px 0 15px;}
	.myRow{border: 2px #F05 dashed;margin: 10px 0 10px 0;padding: 0 3px 5px 0;}
	.exteraArea{border: 1px #33F dotted;margin: 5px 0 0 5px;padding: 0 3px 5px 0;}
	.myLabel {margin: 0 3px 0 0;white-space: nowrap;}
	.myProp {font-weight: bold;color: #40F; margin: 0 3px 0 3px;}
	div.btn label{background-color:yellow;color: blue;padding: 3px 30px 3px 30px;cursor: pointer;}
	div.btn{margin: -5px 250px 0px 5px;}
	div.btn img{margin: 0px 20px -5px 0;cursor: pointer;}
</style>
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
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œÖ<BR>")
	else
		set RS1=Conn.Execute ("SELECT Name FROM OrderTraceTypes WHERE (ID='"& orderType & "') AND (IsActive=1)")
		if RS1.eof then
			conn.close
			response.redirect "?errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ ”›«—‘<br>«ÿ·«⁄«  Ê«—œ ‰‘œÖ<BR>")
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
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
	set rs=Conn.Execute("select * from OrderTraceTypes where id=" & request("orderType"))
	if rs.eof then 
		conn.close
		response.redirect "?errmsg=" & server.urlEncode("‰Ê⁄ ”›«—‘ —«  ⁄ÌÌ‰ ﬂ‰Ìœ")
	end if
	orderTypeID=rs("id")
	orderTypeName=rs("name")
	set orderProp = server.createobject("MSXML2.DomDocument")
	if rs("property")<>"" then 
		orderProp.loadXML(rs("property"))
		myXML = fetchKeyValues()
	end if
	rs.close
	set rs=nothing
	
		
function fetchKeyValues()
	key="---first---"
	thisRow="<?xml version=""1.0""?><keys>"
	for each tmp in orderProp.selectNodes("//key")
		if key<>tmp.parentNode.nodeName then 
			key=tmp.parentNode.nodeName
			thisRow = thisRow & "<" & key & ">"
			hasValue=0
			For Each myKey In orderProp.SelectNodes("/keys/" & key & "/key")
				thisName = myKey.GetAttribute("name")
				id=0
				for each value in request.form(thisName)
					if value <> "" then 
						thisRow = thisRow & "<key name=""" & thisName & """ id=""" 
						if myKey.GetAttribute("type") = "check" then
							thisRow = thisRow & mid(value,4)
						else
							thisRow = thisRow & id 
						end if
						thisRow = thisRow & """>" & value & "</key>"
						hasValue=hasValue +1
					end if
					id=id+1
				next
			Next
			if hasValue>0 then 
				thisRow = thisRow & "</" & key & ">"
			else
				thisRow = Replace(thisRow,"<" & key & ">","")
			end if
		end if
	Next
	thisRow = thisRow & "</keys>"
	fetchKeyValues = thisRow 
end function



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

	mySql="UPDATE orders_trace SET order_date= N'"& OrderDate & "', order_time= N'"& OrderTime & "', return_date= N'"& ReturnDate & "', return_time= N'"& ReturnTime & "', company_name= N'"& CompanyName & "', customer_name= N'"& CustomerName & "', telephone= N'"& Telephone & "', order_title= N'"& OrderTitle & "', order_kind= N'"& orderTypeName & "', Type= '"& orderType & "', vazyat= N'"& statusName & "', status= "& Vazyat & ", step= "& Marhale & ",  marhale= N'"& stepName & "', salesperson= N'"& SalesPerson & "', qtty= N'"& Qtty & "', paperSize= N'"& Size & "', SimplexDuplex= N'"& SimplexDuplex & "', Price= N'"& Price & "' , LastUpdatedDate=N'"& shamsitoday() & "' , LastUpdatedTime=N'"& currentTime10() & "', LastUpdatedBy=N'"& session("ID")& "', property=N'" & myXML & "' WHERE (radif_sefareshat= N'"& radif & "')"	
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
	call showAlert("”›«—‘ êÌ—‰œÂ «Ì‰ ”›«—‘ ‘„« ‰Ì” Ìœ..<BR>ﬁ»· «“ ÊÌ—«Ì‘ »« ”›«—‘ êÌ—‰œÂ Â„«Â‰ê ﬂ‰Ìœ.", CONST_MSG_ALERT ) 
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
'		call showAlert(" ÊÃÂ!<br>”›«—‘Ì ﬂÂ ‘„« Ã” ÃÊ ﬂ—œÌœ ﬁ»·« œ— Ìﬂ ›«ﬂ Ê— <B> «ÌÌœ ‰‘œÂ</B> ÊÃÊœ œ«—œ<br><A HREF='../AR/InvoiceEdit.asp?act=editInvoice&invoice=" & FoundInvoice & "'>«’·«Õ ›«ﬂ Ê— „—»ÊÿÂ (" & FoundInvoice & ")</A>", CONST_MSG_INFORM ) 
		call showAlert(" ÊÃÂ!<br>”›«—‘Ì ﬂÂ ‘„« Ã” ÃÊ ﬂ—œÌœ ﬁ»·« œ— Ìﬂ ›«ﬂ Ê— <B> «ÌÌœ ‰‘œÂ</B> ÊÃÊœ œ«—œ<br><A HREF='../AR/AccountReport.asp?act=showInvoice&invoice=" & FoundInvoice & "'>‰„«Ì‘ ›«ﬂ Ê— „—»ÊÿÂ (" & FoundInvoice & ")</A>", CONST_MSG_INFORM ) 
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
	<TD align="left" >‰Ê⁄ ”›«—‘:</TD>
	<TD>
		<SELECT NAME="OrderType" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold; width: 100px' tabIndex="7" onchange="alert(' ÊÃÂ œ«‘ Â »«‘Ìœ ﬂÂ  €ÌÌ— ‰Ê⁄ ”›«—‘ »«⁄À «Œ ·«· œ— Ã“∆Ì«  ”›«—‘ ŒÊ«Âœ ‘œ. ¬Ì« „”Ê·Ì  ¬‰—« „ÌùÅ–Ì—Ìœø');">
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
	<TD colspan="2"><SELECT NAME="Vazyat" style='font-family: tahoma,arial ; font-size: 8pt; font-weight: bold; width: 100px' tabIndex="14">
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
<tr bgcolor="#CCCCCC">
	<td colspan="6">
		<script type="text/javascript" src="/js/jquery-1.7.min.js"></script>
		<script type="text/javascript">
		function disGroup (e){
			groupName=$(e).parent(".mySection").attr("groupname");
			if (e.checked) {
				$(e).parent(".mySection").children('[name^="'+groupName+'"]').prop("disabled", false);
			} else {
				$(e).parent(".mySection").children('[name^="'+groupName+'"]').prop("disabled", true);
				$(e).parent(".mySection").children('[name$="disBtn"]').prop("disabled", false);
			}
		}
		
		$(document).ready(function () {
			$('[name$=-addValue]').hide();
		});
		function cloneRow(key){
			maxID = $("#" + key.replace(/\//gi,"-") + "-maxID").val();
			newRow = $("#" + key.replace(/\//gi,"-") + "-"+maxID).clone().attr('id', key.replace(/\//gi,"-") + "-" + (parseInt(maxID)+1));
			$('input:checkbox',newRow).each(function (){
				if ($(this).val().substr(0,3)=='on-')
					$(this).val('on-'+(parseInt(maxID)+1));
			});
			
			newRow.appendTo("#extreArea" + key.replace(/\//gi,"-") );
			$("#" + key.replace(/\//gi,"-") + "-maxID").val(parseInt(maxID)+1);
		}
		function removeRow(key){
			maxID = parseInt($("#" + key.replace(/\//gi,"-") + "-maxID").val());
			if (maxID>0) {
				$("#" + key.replace(/\//gi,"-") + "-"+maxID).remove();
			}
			$("#" + key.replace(/\//gi,"-") + "-maxID").val(maxID-1);
		}
		function checkOther(e){
			if ($(e).val()==-1 || $(e).val().substr(0,6)=="other:") {
				if ($(e).find("option:selected").text()=="”«Ì—") {
					$(e).next().val("„ﬁœ«— —« Ê«—œ ﬂ‰Ìœ");
				} else {
					$(e).next().val(
						$(e).find("option:selected").text());
				}
				$(e).next().show();
				$(e).next().focus();
			} else {
				$(e).next().hide();
			}
		}
		function addOther(e){
			if ($(e).val()!="" && $(e).val()!="„ﬁœ«— —« Ê«—œ ﬂ‰Ìœ" && $(e).val()!="‰„Ìù‘Â ﬂÂ Œ«·Ì »«‘Â!"){
				$(e).prev().find("option:selected").text($(e).val());
				$(e).prev().find("option:selected").val('other:'+$(e).val());
				$(e).hide();
			} else {
				$(e).val("‰„Ìù‘Â ﬂÂ Œ«·Ì »«‘Â!");
				$(e).focus();
			}
		}
		</script>
		<div>Ã“∆Ì«  ”›«—‘</div>
<%
	set rs=Conn.Execute("select * from OrderTraceTypes where id="&rs2("type"))
	set typeProp = server.createobject("MSXML2.DomDocument")
	set orderProp = server.createobject("MSXML2.DomDocument")
	
	if rs2("property")<>"" then orderProp.loadXML(rs2("property"))
	if rs("property")<>"" then typeProp.loadXML(rs("property"))
	set rs=nothing

' 	for each item in typeProp.SelectNodes("//key")
' 		response.write item.parentNode.nodeName & ": " & item.xml & "<br>"
' 	next
' 	response.end
sub showKeyEdit(key)
	'key="/keys/printing/key"
	oldGroup="---first---"
	oldLabel="---first---"
	thisRow = "<div class='myRow'>"
	maxID=0
	oldID=0
	for each mykey in orderProp.SelectNodes(key)
		id=myKey.GetAttribute("id")
		if maxID<id then maxID=id
	next
	thisRow = thisRow & "<input type='hidden' value='" & maxID & "' id='" & Replace(key,"/","-") & "-maxID'>"
'	response.write "<div class='myRow'><div class='exteraArea' id='" & Replace(key,"/","-") & "-0'>"
	for id = 0 to maxID
		thisRow = thisRow & "<div class='exteraArea' id='" & Replace(key,"/","-") & "-" & id & "'>"
' 		thisRow = thisRow & "<img title='Õ–› «Ì‰ Œÿ' src='/images/cancelled.gif' onclick='$(this).parent().remove();'>"
		For Each myKey In typeProp.SelectNodes(key)
		  thisType = myKey.GetAttribute("type")
		  thisName = myKey.GetAttribute("name")
		  thisLabel= myKey.GetAttribute("label")
		  thisGroup= myKey.GetAttribute("group")
		  set thisValue= orderProp.SelectNodes(key & "[@id='" & id & "' and @name='" & thisName & "']")
		  hasValue=false
		  if thisValue.length>0 then hasValue=true
' 		  response.write hasValue & "<br>"
		  if (oldGroup<>thisGroup and oldID=id and oldGroup <> "---first---") then thisRow = thisRow &  "</div>"
		  if oldGroup<>thisGroup or oldID<>id then 
			thisRow = thisRow & "<div class='mySection' groupName='" & thisGroup & "'>"
			if myKey.GetAttribute("disable")="1" then 
				thisRow = thisRow & "<input type='checkbox' name='" & thisName & "-disBtn' onchange='disGroup(this);'"
				if hasValue then 
					thisRow = thisRow & " checked='checked'"
					disText=""
				else
					disText=" disabled='disabled' "
				end if	
				thisRow = thisRow & ">"
			else
				disText=""
			end if
		  end if
		  if oldLabel<>thisLabel then thisRow = thisRow &  "<label class='myLabel'>" & thisLabel & "</label>"
			
		  select case thisType
		  	case "option"
		  		thisRow = thisRow &  "<select " & disText & " style='margin:0;padding:0;' name='" & thisName & "'>"
		  		for each myOption in myKey.getElementsByTagName("option")
		  			thisRow = thisRow & "<option value='" & myOption.text & "'"
		  			if hasValue then 
		  				if thisValue(0).text=myOption.text then thisRow = thisRow & " selected='selected' "
		  			end if
		  			thisRow = thisRow & ">" & myOption.GetAttribute("label") & "</option>"
			  	next
			  	thisRow = thisRow & "</select>"
			case "option-other"
				thisRow = thisRow & "<select " & disText & " style='margin:0;padding:0;' name='" & thisName & "' onchange='checkOther(this);'>"
		  		for each myOption in myKey.getElementsByTagName("option")
		  			thisRow = thisRow & "<option value='" & myOption.text & "'"
		  			if hasValue then 
		  				if thisValue(0).text=myOption.text then thisRow = thisRow & " selected='selected' "
		  			end if
		  			thisRow = thisRow & ">" & myOption.GetAttribute("label") & "</option>"
			  	next
			  	if hasValue then 
				  	if left(thisValue(0).text,6)="other:" then 
				  		thisRow = thisRow & "<option value='" & thisValue(0).text & "' selected='selected'>" & mid(thisValue(0).text,7) & "</option></select>"
				  	else
				  		thisRow = thisRow & "<option value='-1'>”«Ì—</option></select>"
				  	end if
				else
					thisRow = thisRow & "<option value='-1'>”«Ì—</option></select>"
				end if
			  	thisRow = thisRow & "<input type='text' name='" & thisName & "-addValue' onblur='addOther(this);'>"
			case "text"
				thisRow = thisRow & "<input " & disText & " type='text' style='margin:0;padding:0;' size='" & myKey.text & "' name='" & thisName & "' "
				if hasValue then thisRow = thisRow & "value='" & thisValue(0).text & "'"
				thisRow = thisRow & ">"
			case "textarea"
				thisRow = thisRow & "<textarea name='" & thisName & "' style='width:600px;' cols='" & myKey.text & "'>"
				if hasValue then thisRow = thisRow & thisValue(0).text
				thisRow = thisRow & "</textarea>"
			case "check"
				thisRow = thisRow & "<input type='checkbox' value='on-" & id & "' name='" & thisName & "' "
				if hasValue then 
					if left(thisValue(0).text,2)="on" then thisRow = thisRow & "checked='checked'"
				else
					if myKey.text="checked" then thisRow = thisRow & "checked='checked'"
				end if
				thisRow = thisRow & ">"
		  end select
		  if myKey.GetAttribute("force")="yes" then thisRow = thisRow &  "<span style='color:red;margin:0 0 0 2px;padding:0;'>*</span>"
		  oldGroup=thisGroup
		  oldLabel=thisLabel
		  oldID=id
		Next
		thisRow = thisRow & "</div></div>"
	next
	thisRow = thisRow & "<div id='extreArea" &Replace(key,"/","-")& "'></div>"
	response.write thisRow 'prependTo 
	response.write "<div class='btn'><img title='«÷«›Â' src='/images/Plus-32.png' width='20px' onclick='cloneRow(""" & key & """);'><img title='Õ–› ¬Œ—Ì‰ Œÿ' src='/images/cancelled.gif' onclick='removeRow(""" & key & """);'></div></div>"
end sub
	
	oldTmp="---first---"
	for each tmp in typeProp.selectNodes("//key")
		if oldTmp<>tmp.parentNode.nodeName then 
			oldTmp=tmp.parentNode.nodeName
			call showKeyEdit("/keys/" & oldTmp & "/key")
		end if
	next
	
	
%>
		
	</td>
</tr>
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