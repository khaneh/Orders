<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="Ê—Êœ ›«ﬂ Ê—"
SubmenuItem=1
if not (Auth(6 , 1) OR Auth(6 , 4)) then NotAllowdToViewThisPage()
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
	.InvRowInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0FFF0; direction:LTR; text-align:right;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
	.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var selectedRow=-1;
function selectAndDelete(rowNo){
	if (selectedRow==rowNo){
		delAndShiftUp(rowNo);
		document.getElementById("InvoiceLines").getElementsByTagName("TR")[selectedRow].getElementsByTagName("TD")[0].setAttribute("bgColor","");
		selectedRow=-1
		calcAndCheck()
	}
	else{
		if (selectedRow != -1)
			document.getElementById("InvoiceLines").getElementsByTagName("TR")[selectedRow].getElementsByTagName("TD")[0].setAttribute("bgColor","");
		document.getElementById("InvoiceLines").getElementsByTagName("TR")[rowNo].getElementsByTagName("TD")[0].setAttribute("bgColor","red");
		selectedRow=rowNo;
	}
}
function delAndShiftUp(rowNo){
	for (i=rowNo;i<9;i++){
		document.getElementsByName("Descriptions")[i].value=document.getElementsByName("Descriptions")[i+1].value
		document.getElementsByName("Qttys")[i].value=document.getElementsByName("Qttys")[i+1].value
		document.getElementsByName("Units")[i].value=document.getElementsByName("Units")[i+1].value
		document.getElementsByName("Fees")[i].value=document.getElementsByName("Fees")[i+1].value
		document.getElementsByName("Prices")[i].value=document.getElementsByName("Prices")[i+1].value
	}
}
function setPrice(src){
	if (src.name=="Qttys" || src.name=="Fees"){
		src.value=val2txt(txt2val(src.value))
		rowNo=src.parentNode.parentNode.rowIndex;
		tmpFee=txt2val(document.getElementsByName("Fees")[rowNo].value);
		tmpQtty=txt2val(document.getElementsByName("Qttys")[rowNo].value);
		tmpPrice= tmpFee * tmpQtty;
		document.getElementsByName("Prices")[rowNo].value = val2txt(parseInt(tmpPrice));
	}
	else if (src.name=="Discount"){
		src.value=val2txt(txt2val(src.value))
	}
	calcAndCheck()
}
function calcAndCheck(){
	var total=0;
	var payable;
	for(i=0;i<10;i++){
		total+=txt2val(document.getElementsByName("Prices")[i].value)
	}
	document.getElementsByName("TotalPrice")[0].value=val2txt(total)
	payable=total-txt2val(document.getElementsByName("Discount")[0].value);
	document.getElementsByName("Payable")[0].value=val2txt(payable)
	if(document.getElementsByName("Payable")[0].value==document.getElementsByName("MustBe")[0].value){
		document.getElementsByName("Payable")[0].style.backgroundColor='#00FF00';
	}
	else{
		document.getElementsByName("Payable")[0].style.backgroundColor='#FF0000';
	}
}
//-->
</SCRIPT>
<%
response.end()
if request("act")="getPrintForm" then
	if isnumeric(request("invoice")) then
		InvoiceID=clng(request("invoice"))
		mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("›«ﬂ Ê— ÅÌœ« ‰‘œ.")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if

	customerID=		RS1("Customer")
	creationDate=	RS1("CreatedDate")
	totalPrice=		cdbl(RS1("totalPrice"))
	totalDiscount=	cdbl(RS1("totalDiscount"))+cdbl(RS1("totalReverse"))
	totalReceivable=cdbl(RS1("totalReceivable"))
	totalVat =		cdbl(RS1("totalVat"))
	Voided=			RS1("Voided")
	Issued=			RS1("Issued")
	Approved=		RS1("Approved")
	isReverse=		RS1("IsReverse")
	IsA=			RS1("IsA")
	InvoiceNo=		RS1("Number")

	if isReverse then
		'Check for permission for EDITTING Rev. Invoice
		if not Auth(6 , 5) then NotAllowdToViewThisPage()
		itemTypeName="›«ò Ê— »—ê‘ "
		HeaderColor="#FF9900"
	else
		'Check for permission for EDITTING Invoice
		if not Auth(6 , 3) then NotAllowdToViewThisPage()
		itemTypeName="›«ò Ê—"
		HeaderColor="#C3C300"
	end if

	if Voided then
		Conn.close
		response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— »«ÿ· ‘œÂ «” .")
	elseif Issued then
		if Auth(6 , "A") then
			' Has the Priviledge to change the Invoice
			response.write "<BR>"
			call showAlert ("«Ì‰ ›«ﬂ Ê— ’«œ— ‘œÂ «” .<br>Â—ç‰œ òÂ ‘„« «Ã«“Â œ«—Ìœ òÂ «Ì‰ ›«ò Ê— —«  €ÌÌ— »œÂÌœ<br>„”Ê·Ì  „‘ò·«  «Õ „«·Ì —« „Ì Å–Ì—Ìœø",CONST_MSG_INFORM) 
		else
			Conn.close
			response.redirect "AccountReport.asp?act=showInvoice&invoice="& InvoiceID & "&errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— ’«œ— ‘œÂ «” .")
		end if
	end if 

	mySQL="SELECT ID,AccountTitle,Address1,Tel1,PostCode1,EconomicalCode FROM Accounts WHERE (ID='"& customerID & "')"
	Set RS1 = conn.Execute(mySQL)
	customerName=RS1("AccountTitle")
	customerAddress=RS1("Address1")
	customerTel=RS1("Tel1")
	customerPostCode=RS1("PostCode1")
	customerEcCode=RS1("EconomicalCode")

	RS1.close

'	response.end

	creationDate=shamsiToday()
'	creationTime=Hour(creationTime)&":"&Minute(creationTime)
'	if instr(creationTime,":")<3 then creationTime="0" & creationTime
'	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)

%>
<!-- Ê—Êœ «ÿ·«⁄«  ›«ﬂ Ê— -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
	<FORM METHOD=POST ACTION="?act=submitPrintForm">
	<tr bgcolor='#C3C300'>
		<td align="left"><TABLE>
			<TR>   
				<TD align="left"> «—ÌŒ:</td>
				<TD dir="LTR">
					<INPUT class="InvGenInput" NAME="InvoiceDate" TYPE="text" maxlength="10" size="10" value="<%=CreationDate%>"></td>
					<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>">
				<TD dir="RTL"><%=weekdayname(weekday(date))%></td>
			</TR>  
			</TABLE>
		</td>
	</tr>
	<tr bgcolor='#C3C300'>
		<td>
			<TABLE border="0" width="100%" >
			<TR>   
				<TD align="left">‰«„ Œ—Ìœ«—:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="CustomerName" rows="1" cols="50"><%=CustomerName%></TEXTAREA></TD>
				<TD align="left">‘„«—Â «ﬁ ’«œÌ:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;direction:LTR;text-align:left;" NAME="customerEcCode" TYPE="text" size="10" value="<%=customerEcCode%>"></TD>
			</TR>  
			<TR>   
				<TD align="left">‰‘«‰Ì:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerAddress" rows="2" cols="50"><%=customerAddress%></TEXTAREA></TD>
				<TD align="left">ﬂœ Å” Ì:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;" NAME="customerPostCode" TYPE="text" size="10" value="<%=customerPostCode%>"></TD>
			</TR>  
			<TR>   
				<TD align="left"> ·›‰:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerTel" rows="1" cols="50"><%=customerTel%></TEXTAREA></TD>
				<TD align="left">‘—«Ìÿ ›—Ê‘:</TD>
				<TD align=left><INPUT TYPE="radio" NAME="Cash" Value="1">‰ﬁœÌ  <INPUT TYPE="radio" NAME="Cash" Value="0" checked>€Ì—‰ﬁœÌ</TD>
			</TR>  
			</TABLE>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TR bgcolor='#CCCC88'>
				<TD align='center' width="25px"> # </td>
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" value="‘—Õ ”›«—‘" size="65" ></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" value=" ⁄œ«œ" size="6"></TD> <!--S A M-->  
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="Ê«Õœ ﬂ«·«" size="10"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="ﬁÌ„  Ê«Õœ" size="15"></TD>
				<td><INPUT class="InvHeadInput" readonly TYPE="text" value="„«·Ì« " size="7"></td><!--S A M-->
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" Value="„»·€" size="20"></TD>      
			</TR>
			</TABLE></div>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TBODY id="InvoiceLines">
<%		
		Dim Descriptions(10)
		Dim Qttys(10)
		Dim Units(10) ' Allways empty!
		Dim Fees(10)
		Dim Prices(10)
		i=0
		mySQL="SELECT * FROM InvoiceLines left outer join invoiceItems on InvoiceLines.item = invoiceItems.ID WHERE (Invoice='"& InvoiceID & "') "
		Set RS1 = conn.Execute(mySQL)
		Do While NOT RS1.eof AND i<10
			i=i+1
			Descriptions(i)=RS1("Description")
			Qttys(i)=Separate(RS1("AppQtty"))
			if RS1("AppQtty") <> 0 then 
				Fees(i)=Separate(RS1("Price")/RS1("AppQtty")) 
			else 
				Fees(i)="0"
			end if
			Prices(i)=Separate(RS1("Price"))
			RS1.moveNext
		Loop
		RS1.close

		for i=1 to 10
%>
			<TR bgcolor='#F0F0F0'>
				<TD align='center' width="25px" onclick="selectAndDelete(this.parentNode.rowIndex);"><%=i%></TD>
				<TD><INPUT class="InvRowInput2" TYPE="text" Name="Descriptions" value="<%=Descriptions(i)%>" size="65"></TD>
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Qttys" value="<%=Qttys(i)%>" size="6" onBlur="setPrice(this);"></TD>   
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Units" Value="<%=Units(i)%>" size="10"></TD>        
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Fees" Value="<%=Fees(i)%>" size="15" onBlur="setPrice(this);"></TD>        
				<td><INPUT class="InvRowInput"  TYPE="text" Name="Vat" value="<%Vat(i)%>" size="7" onBlur="this.value=val2txt(txt2val(this.value));calcAndCheck();"></td>
				<TD><INPUT class="InvRowInput3" TYPE="text" Name="Prices" Value="<%=Prices(i)%>" size="20" onBlur="this.value=val2txt(txt2val(this.value));calcAndCheck();"></TD>      
			</TR>
<%
		next
%>
			</Tbody>
			</TABLE>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td colspan="10"><div>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#CCCC88">
			<TR bgcolor='#CCCC88'>
				<TD align='center' width="24px"></td>
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="65"></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="Ã„⁄ ﬂ·:" style="text-align:left;" size="15"></TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalPrice" Value="<%=Separate(totalPrice)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>
				<TD align='center' width="24px"></td>
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="65"></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value=" Œ›Ì›:" style="text-align:left;" size="15"></TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" TYPE="text" Name="Discount" Value="<%=Separate(totalDiscount)%>" size="20" onBlur="setPrice(this);"></TD>      
			</TR>
			<tr bgcolor="#CCCC88">
				<TD align='center' width="24px"></td>
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="65"></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>  
				<TD><INPUT class="InvHeadInput" readonly TYPE="text" Value="„«·Ì«  »— «—“‘ «›“ÊœÂ:" style="text-align:left" size="15"></TD>
				<TD style="border:1 solid black;"><INPUT style="color:gray;" readonly TYPE="text" NAME="totalVat" value="<%=Separate(totalVat)%>" size="20"></TD>
			</td>
			<TR bgcolor='#CCCC88'>
				<TD align='center' width="24px"></td>
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="65"></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="ﬁ«»· Å—œ«Œ :" style="text-align:left;" size="15"></TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Payable" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>
				<TD></td>
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="65"></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="10"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" size="15"></TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="MustBe" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>
			</TABLE></div>
		</td>
	</tr>
	<tr>
		<td align='center'><INPUT class="InvGenButton" TYPE="button" value="«‰’—«›" onclick="calcAndCheck();"></td>
	</tr>
	</table>
	</FORM>
	<br>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
//		document.getElementsByName("Items")[0].focus();
		calcAndCheck();
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
<!--#include file="tah.asp" -->