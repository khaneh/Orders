<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="ç«Å ›—„ ›«ﬂ Ê—"
SubmenuItem=8
if not Auth(9 , 8) then NotAllowdToViewThisPage()
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<%
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
	.InvRowInput4 { font-family:tahoma; font-size: 9pt; border: none; background-color: #FFD3A8; direction:LTR; text-align:right;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput4 { font-family:tahoma; font-size: 9pt; border: none; background-color: #FF9900; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right;}
	.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
	.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
/* var selectedRow=-1; */
function selectAndDelete(td){
	console.log($(td).html());
	var tr = $(td).closest("tr");
	if ($(td).prop("bgColor")=="red"){
		$(td).prop("bgColor","");
		for(trNext=tr;trNext.next().size()>0;trNext=trNext.next()){
			trNext.find("input[name=Descriptions]").val(trNext.next().find("input[name=Descriptions]").val());
			trNext.find("input[name=Qttys]").val(trNext.next().find("input[name=Qttys]").val());
			trNext.find("input[name=Fees]").val(trNext.next().find("input[name=Fees]").val());
			trNext.find("input[name=Prices]").val(trNext.next().find("input[name=Prices]").val());
			trNext.find("input[name=Dis]").val(trNext.next().find("input[name=Dis]").val());
			trNext.find("input[name=Vat]").val(trNext.next().find("input[name=Vat]").val());
		}
		$($("#InvoiceLines tr")[9]).find("input[name=Descriptions]").val("");
		$($("#InvoiceLines tr")[9]).find("input[name=Qttys]").val("");
		$($("#InvoiceLines tr")[9]).find("input[name=Fees]").val("");
		$($("#InvoiceLines tr")[9]).find("input[name=Prices]").val("");
		$($("#InvoiceLines tr")[9]).find("input[name=Dis]").val("");
		$($("#InvoiceLines tr")[9]).find("input[name=Vat]").val("");
	} else {
		$("#InvoiceLines tr").find("td:first").prop("bgColor","");
		$(td).prop("bgColor","red");
	}
	calcAndCheck();
}

function setPrice(src){
	var name = $(src).attr("name");
	var tr = $(src).closest("tr");
	if (name == "Qttys" || name == "Fees" || name == "Dis"){
		$(src).val(echoNum(getNum($(src).val())));
		var tmpFee = getNum(tr.find("input[name=Fees]").val());
		var tmpQtty = getNum(tr.find("input[name=Qttys]").val());
		var vat = getNum(tr.find("input[name=Vat]").val());
		var dis = getNum(tr.find("input[name=Dis]").val());		
		var tmpPrice= tmpFee * tmpQtty - dis;
		//alert(val2txt(tmpPrice * .03));
		if (vat > 0){
			tr.find("input[name=Vat]").val(echoNum(Math.floor(tmpPrice * parseFloat($('input[name=VatRate]').val()/100))));
		}
		tr.find("input[name=Prices]").val(echoNum(parseInt(tmpPrice)));
	}

	calcAndCheck();
}
function calcAndCheck(){
	var total=0;
	var dis=0;
	var payable=0;
	var vat=0;
	var checkVat=false;
/* 	console.log("calcAndCheck"); */
	$("#InvoiceLines tr").each(function(i){
		total += getNum($(this).find("[name=Prices]").val());
		dis += getNum($(this).find("[name=Dis]").val());
		vat += getNum($(this).find("[name=Vat]").val());
	});
	if (getNum($("input[name=totalVat]").val()) !=vat){
//	console.log('calc vat: '+ vat + ' get vat: ' + getNum($("input[name=totalVat]").val()))
		$("input[name=message]").val("«Œÿ«—! Ã„⁄ „«·Ì«  «‘ »«Â „Ì»«‘œ.");
		$("input[name=totalVat]").css("backgroundColor","#FF4848");
		checkVat = false;
	} else {
		$("input[name=message]").val("");
		$("input[name=totalVat]").css("backgroundColor","#F0F0F0");
		checkVat = true;
	}
	$("input[name=TotalPrice]").val(echoNum(total + vat));
	$("input[name=Discount]").val(echoNum(dis));
	payable = total + vat;
	$("input[name=Payable]").val(echoNum(payable));
	if (getNum($("input[name=Payable]").val()) == getNum($("input[name=MustBe]").val())){
		$("input[name=Payable]").css("backgroundColor","#00FF00");
		if (checkVat)
			return true;
		else
			return false;
	} else {
		$("input[name=Payable]").css("backgroundColor","#FF0000");
		return false;
	}
}
function autoComplete(src){
	if (src.value==""){
//		alert(event.keyCode);
		switch (event.keyCode){
		case 1576: // »
			src.value="»—ê";
			event.keyCode=0;
			break;
		case 1608: // Ê
			src.value="Ê—ﬁ";
			event.keyCode=0;
			break;
		}
	}
	else if (src.value=="”") {
		switch (event.keyCode){
		case 1585: // —
			src.value="”—Ì";
			event.keyCode=0;
			break;
		case 1575: // «
			src.value="”«‰ Ì„ —";
			event.keyCode=0;
			break;
		}
	}
}
function combine()
{
	var priceVat = 0;
	var priceNoVat = 0;
	var totalVat = 0;
	var totalDisOnVat = 0;
	var totalDisNoVat = 0;
	var rfd = 0;
	$("#InvoiceLines tr").each(function(i){
		if (getNum($(this).find("input[name=Vat]").val())==0){
			if (getNum($(this).find("input[name=Prices]").val()) <= 0){
				rfd += getNum($(this).find("input[name=Dis]").val());
			} else {
				totalDisNoVat += getNum($(this).find("input[name=Dis]").val());
				priceNoVat += getNum($(this).find("input[name=Fees]").val()) * getNum($(this).find("input[name=Qttys]").val());
			}
		} else {
			priceVat += getNum($(this).find("input[name=Fees]").val()) * getNum($(this).find("input[name=Qttys]").val());
			totalVat += getNum($(this).find("input[name=Vat]").val());
			if (getNum($(this).find("input[name=Prices]").val()) <= 0){
				rfd += getNum($(this).find("input[name=Dis]").val());
			} else {
				totalDisOnVat += getNum($(this).find("input[name=Dis]").val());
			}
		}
		$(this).find("input[name=Descriptions]").val("");
		$(this).find("input[name=Qttys]").val("");
		$(this).find("input[name=Units]").val("");
		$(this).find("input[name=Fees]").val("");
		$(this).find("input[name=Prices]").val("");
		$(this).find("input[name=Vat]").val("");
		$(this).find("input[name=Dis]").val("");
	});
	
	var tr;
	if (priceNoVat > 0){
		tr = $($("#InvoiceLines tr")[1]);
		tr.find("input[name=Descriptions]").val("ﬂ«€– Ê „ﬁÊ« ÃÂ  ç«Å");
		tr.find("input[name=Qttys]").val("1");
		tr.find("input[name=Units]").val("");
		tr.find("input[name=Fees]").val(echoNum(priceNoVat));
		tr.find("input[name=Prices]").val(echoNum(priceNoVat - totalDisNoVat));
		tr.find("input[name=Vat]").val("0");
		tr.find("input[name=Dis]").val(echoNum(totalDisNoVat));
	}
	tr = $($("#InvoiceLines tr")[0]);
	tr.find("input[name=Descriptions]").val("Œœ„«  ç«Å");
	tr.find("input[name=Qttys]").val("1");
	tr.find("input[name=Units]").val("");
	tr.find("input[name=Fees]").val(echoNum(priceVat));
	tr.find("input[name=Prices]").val(echoNum(priceVat - totalDisOnVat));
	tr.find("input[name=Vat]").val(echoNum(totalVat));
	tr.find("input[name=Dis]").val(echoNum(totalDisOnVat));
	tr = $($("#InvoiceLines tr")[2]);
	//---------------------------
	tr.find("input[name=Descriptions]").val(" Œ›Ì› —‰œ ›«ﬂ Ê—");
	tr.find("input[name=Qttys]").val("1");
	tr.find("input[name=Units]").val("");
	tr.find("input[name=Fees]").val("0");
	tr.find("input[name=Prices]").val(echoNum(-rfd));
	tr.find("input[name=Vat]").val("0");
	tr.find("input[name=Dis]").val(echoNum(rfd));
	calcAndCheck();
}
//-->
</SCRIPT>
<%
Dim Descriptions(10)
Dim Qttys(10)
Dim Units(10) ' Always empty!
Dim Fees(10)
Dim Prices(10)
Dim Vat(10)
Dim Dis(10)

if request("act")="submitsearch" then
	if isnumeric(request("invoice")) then
		InvoiceID=clng(request("invoice"))
		mySQL="SELECT * FROM InvoicePrintForms WHERE (InvoiceID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?act=getPrintForm&invoice=" & InvoiceID
		else
			conn.close
			response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&msg=" & Server.URLEncode("ﬁ»·« »—«Ì «Ì‰ ›«ﬂ Ê— ›—„ ç«ÅÌ «ÌÃ«œ ‘œÂ «” .<br>«ﬂ‰Ê‰ ¬‰ —« „‘«ÂœÂ „Ì ﬂ‰Ìœ Ê „Ì  Ê«‰Ìœ «’·«Õ ﬂ‰Ìœ.")
		end if
	else
		if request("query")<>"" then
			response.redirect "?msg=" & Server.URLEncode("»»Œ‘Ì‰ ›⁄·« ‘„«—Â ”›«—‘ Ã” ÃÊ ‰„Ì ﬂ‰Ì„.")
		else
			response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
		end if
	end if
elseif request("act")="getPrintForm" then
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
	set rsVat=conn.Execute("select isnull(GLs.Vat,0) as Vat from ARItems inner join GLs on ARItems.GL=GLs.ID where Type=1 AND Link="&InvoiceID)
	customerID=		RS1("Customer")
	totalPrice=		cdbl(RS1("totalPrice"))
	totalDiscount=	cdbl(RS1("totalDiscount"))+cdbl(RS1("totalReverse"))
	totalReceivable=cdbl(RS1("totalReceivable"))
	totalVat =		cdbl(RS1("totalVat")) ' S A M
	Voided=			RS1("Voided")
	Issued=			RS1("Issued")
	Approved=		RS1("Approved")
	isReverse=		RS1("IsReverse")
	IsA=			RS1("IsA")
	InvoiceNo=		RS1("Number")
	VatRate=		rsVat("Vat")
	rsVat.close
	if Voided then
		Conn.close
		response.redirect "?errmsg=" & Server.URLEncode("«Ì‰ ›«ﬂ Ê— »«ÿ· ‘œÂ «” .")
	elseif isReverse then
		Conn.close
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— „⁄ »— ‰Ì” .")
	end if 

	mySQL="SELECT ID,IsPersonal,AccountTitle,CompanyName,Address1,Tel1,PostCode1,EconomicalCode, NorRCode FROM Accounts WHERE (ID='"& customerID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1("IsPersonal") then
		customerName=RS1("AccountTitle")
	else
		customerName=RS1("CompanyName")
	end if
	customerAddress=RS1("Address1")
	customerTel=RS1("Tel1")
	customerPostCode=RS1("PostCode1")
	customerEcCode=RS1("EconomicalCode")
	NorRCode = RS1("NorRCode")

	RS1.close

	creationDate=shamsiToday()
'	creationTime=Hour(creationTime)&":"&Minute(creationTime)
'	if instr(creationTime,":")<3 then creationTime="0" & creationTime
'	if len(creationTime)<5 then creationTime=Left(creationTime,3) & "0" & Right(creationTime,1)

%>
<!-- Ê—Êœ «ÿ·«⁄«  ›—„ ›«ﬂ Ê— -->
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
					<INPUT TYPE="hidden" NAME="InvoiceID" value="<%=InvoiceID%>">
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
				<TD><TEXTAREA readonly style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="CustomerName" rows="1" cols="50"><%=CustomerName%></TEXTAREA></TD>
				<TD align="left">‘„«—Â «ﬁ ’«œÌ:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;direction:LTR;text-align:left;" NAME="customerEcCode" TYPE="text" size="10" value="<%=customerEcCode%>"></TD>
			</TR>  
			<TR>   
				<TD align="left" rowspan="2">‰‘«‰Ì:</TD>
				<TD rowspan="2"><TEXTAREA readonly style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerAddress" rows="2" cols="50"><%=customerAddress%></TEXTAREA></TD>
				<TD align="left">ﬂœ Å” Ì:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;" NAME="customerPostCode" TYPE="text" size="10" value="<%=customerPostCode%>"></TD>
			</TR>  
			<tr>
				<td align='left'>‘„«—Â À» /‘„«—Â „·Ì:</td>
				<td><input class='InvGenInput' style='border:1 solid black;width:100px;' name="NorRCode" type='text' size='10' value='<%=NorRCode%>'></td>
			</tr>
			<TR>   
				<TD align="left"> ·›‰:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerTel" rows="1" cols="50"><%=customerTel%></TEXTAREA></TD>
				<TD align="left">‘—«Ìÿ ›—Ê‘:</TD>
				<TD align=left><INPUT TYPE="radio" NAME="CashPayment" Value="1">‰ﬁœÌ  <INPUT TYPE="radio" NAME="CashPayment" Value="0" checked>€Ì—‰ﬁœÌ</TD>
			</TR>  
			</TABLE>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td><div>
		<input type='hidden' id='VatRate' name='VatRate' value='<%=VatRate%>'>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TR bgcolor='#CCCC88'>
				<TD align='center' width="25px"> # </td>
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" value="‘—Õ ”›«—‘" size="65" ></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" value=" ⁄œ«œ" size="5"></TD> <!--S A M-->   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="Ê«Õœ ﬂ«·«" size="8"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="ﬁÌ„  Ê«Õœ" size="8"></TD>        
				<TD><Input class="InvHeadInput" readOnly type="text" Value=" Œ›Ì›" size="8"><TD>
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" Value="„»·€" size="13"></TD>      
				<TD><INPUT class="InvHeadInput4" readOnly TYPE="text" Value="„«·Ì« " size="8"></TD>
			</TR>
			</TABLE></div>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TBODY id="InvoiceLines">
<%		
		i = 0
		totalNoVat = 0
		totalSumDis = 0

		totalVatAfter9Vat = 0
		totalSumAfter9noVat = 0
		totalSumAfter9Vat = 0
		totalDisAfter9noVat = 0
		totalDisAfter9Vat = 0
		after9Vat=False
		after9noVat=false
		mySQL="SELECT * FROM InvoiceLines left outer join invoiceItems on InvoiceLines.item = invoiceItems.ID WHERE (Invoice="& InvoiceID & ")"
		Set RS1 = conn.Execute(mySQL)
		Do While NOT RS1.eof 
			i = i + 1
			If i < 10 Then 
				item =	cdbl(RS1("Item")) 
		'	if item = 39999 then 
		'		RS1.moveNext
		'	else
				'i=i+1			
				Descriptions(i)=RS1("Description")
				Qttys(i)=Separate(RS1("AppQtty"))
				Vat(i) = RS1("Vat")
				
				if RS1("AppQtty") <> 0 then 
					Fees(i)=Separate((RS1("Price") - RS1("Reverse"))/RS1("AppQtty")) 
				else 
					Fees(i)="0"
				end if
				Prices(i)=Separate(RS1("Price") - RS1("Reverse"))
				Dis(i)=Separate(RS1("Discount"))
			Else 
				If RS1("Vat")=0 Then 
					totalSumAfter9noVat = RS1("Price") + totalSumAfter9noVat
					totalDisAfter9noVat = RS1("Discount") + totalDisAfter9noVat
					after9noVat = true
				Else
					totalVatAfter9Vat = RS1("Vat") + totalVatAfter9Vat
					totalSumAfter9Vat = RS1("Price") - RS1("Reverse") + totalSumAfter9Vat
					totalDisAfter9Vat = RS1("Discount") + totalDisAfter9Vat
					after9Vat = true
				End If 
			End If 
			totalNoVat = totalNoVat + cdbl(RS1("Price") - RS1("Reverse"))
			totalSumDis = totalSumDis + CDbl(RS1("Discount"))
			RS1.moveNext
		'	end if
		Loop
		RS1.close
		If totalSumDis = 0 Then 
			Dis(1) = totalDiscount
		End if
		If (after9noVat or after9Vat) Then 
			'clean & insert two line for other line of my invoice
			'first step is check line 8,9 for vat
			If Vat(8) = 0 Then 
				totalSumAfter9noVat = totalSumAfter9noVat + Prices(8)
				totalDisAfter9noVat = totalDisAfter9noVat + Dis(8)
			Else
				totalVatAfter9Vat = totalVatAfter9Vat + Vat(8)
				totalSumAfter9Vat = totalSumAfter9Vat + Prices(8)
				totalDisAfter9Vat = totalDisAfter9Vat + Dis(8)
			End If
			If Vat(9) = 0 Then 
				totalSumAfter9noVat = totalSumAfter9noVat + Prices(9)
				totalDisAfter9noVat = totalDisAfter9noVat + Dis(9)
			Else
				totalVatAfter9Vat = totalVatAfter9Vat + Vat(9)
				totalSumAfter9Vat = totalSumAfter9Vat + Prices(9)
				totalDisAfter9Vat = totalDisAfter9Vat + Dis(9)
			End If
			'next step is make pare miter for print :)
			Descriptions(8) = "”«Ì— «ﬁ·«„ ›«ﬂ Ê—"
			Prices(8) = totalSumAfter9Vat
			Dis(8) = totalDisAfter9Vat
			Vat(8) = totalVatAfter9Vat
			Fees(8) = Prices(8)
			Qttys(8) = 1
			Descriptions(9) = "”«Ì— «ﬁ·«„ »œÊ‰ „«·Ì«  ›«ﬂ Ê—"
			Prices(9) = totalSumAfter9noVat
			Dis(9) = totalDisAfter9noVat
			Vat(9) = 0
			Fees(9) = Prices(9)
			Qttys(9) = 1
		End If 
			
		for i=1 to 10
			%>
			<TR bgcolor="#F0F0F0">
				<TD align="center" width="25px" onclick="selectAndDelete(this);"><%=i%></TD>
				<TD><INPUT class="InvRowInput2" TYPE="text" Name="Descriptions" value="<%=Descriptions(i)%>" size="65"></TD>
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Qttys" value="<%=Qttys(i)%>" size="5" onBlur="setPrice(this);"></TD>   
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Units" Value="<%=Units(i)%>" size="8" onKeyPress="autoComplete(this);"></TD>       
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Fees" Value="<%=Fees(i)%>" size="8" onBlur="setPrice(this);"></TD>        
				<TD><input class="InvRowInput" type="text" name="Dis" value="<%=Dis(i)%>" Size="8" onBlur="setPrice(this);"></TD>
				<TD><INPUT readonly class="InvRowInput3" TYPE="text" Name="Prices" Value="<%=Separate(Prices(i) - Dis(i))%>" size="13" onBlur="this.value=echoNum(getNum(this.value));calcAndCheck();"></TD>
				<TD><input readonly class='InvRowInput4' type='text' name='Vat' value='<%=Vat(i)%>' size='8'></TD>
			</TR>
<%
		next
%>
			</Tbody>
			</TABLE>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td colspan="10">
		<div>
			<TABLE align='left' Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#CCCC88">
			<TR bgcolor='#CCCC88'>      
				<TD align='left'>Ã„⁄:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalNoVatPrice" Value="<%=Separate(totalNoVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'><!--S A M--> 
				<TD align='left'>„«·Ì«  »— «—“‘ «›“ÊœÂ:</TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="totalVat" Value="<%=Separate(totalVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>      
				<TD align='left'>Ã„⁄ ﬂ·:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalPrice" Value="<%=Separate(totalPrice)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>        
				<TD align='left'> Œ›Ì›:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Discount" Value="<%=Separate(totalDiscount)%>" size="20"></TD>      
			</TR>
			</TR>
			<TR bgcolor='#CCCC88'>
				<TD align='right'><input class="InvHeadInput" type='text' readonly name='message' value='' size='30'></TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Payable" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>
			
			<TR bgcolor='#CCCC88'>
				<TD align='left'>ﬁ«»· Å—œ«Œ :</TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="MustBe" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>			
			</TABLE>
		</div>
		</td>
	</tr>
	<tr>
		<td align='center'>
			<INPUT class="InvGenButton" TYPE="submit" value=" À»  " onclick="return calcAndCheck();">
			<input class='InvGenButton' type='button' Value='œÊ ŒÿÌ ‘œ‰ ›«ﬂ Ê—' onclick='combine();'>
		</td>
	</tr>
	</table>
	</FORM>
	<br>
	<SCRIPT LANGUAGE="JavaScript">
$(document).ready(function(){
		calcAndCheck();
});
	</SCRIPT>
<%
'------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------
elseif request("act")="submitPrintForm" then

	InvoiceID=clng(request.form("InvoiceID"))
		
	if isnumeric(InvoiceID) then
		mySQL="SELECT * FROM InvoicePrintForms WHERE (InvoiceID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if not RS1.eof then
			conn.close
			response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&errmsg=" & Server.URLEncode("ﬁ»·« »—«Ì «Ì‰ ›«ﬂ Ê— ›—„ ç«ÅÌ «ÌÃ«œ ‘œÂ «” .")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if

	creationDate=		shamsiToday()

	InvoiceDate=		request.form("InvoiceDate")
	CustomerName=		request.form("CustomerName")
	customerEcCode=		request.form("customerEcCode")
	customerAddress=	request.form("customerAddress")
	customerPostCode=	request.form("customerPostCode")
	customerTel=		request.form("customerTel")
	NorRCode =			request.Form("NorRCode")
	if request.form("CashPayment") then
		CashPayment=1
	else
		CashPayment=0
	end if
	
	TotalPrice=	text2value(request.form("TotalPrice"))
	Discount=	text2value(request.form("Discount"))
	Payable=	text2value(request.form("Payable"))
	tmpPay = Payable
	PayableText=ConvertIT(tmpPay)
	
	MustBe=text2value(request.form("MustBe"))
	totalVat = text2value(request.form("totalVat"))

	if Payable<>MustBe then
		response.redirect "?act=getPrintForm&invoice="& InvoiceID & "&errMsg=" & Server.URLEncode("Ã„⁄ ›«ﬂ Ê— ’ÕÌÕ ‰Ì” .")
	end if

	mySQL="INSERT INTO InvoicePrintForms (CreatedDate, CreatedBy, InvoiceID, PaperNo, [Date], Name, Address, Tel, EcCode, PostCode, TotalPrice, Discount, Payable, PayableText, CashPayment, Vat, NorRCode) VALUES (N'"& creationDate & "', '"& session("ID") & "', '"& InvoiceID & "', 0, N'"& InvoiceDate & "', N'"& CustomerName & "', N'"& customerAddress & "', N'"& customerTel & "', N'"& customerEcCode & "', N'"& customerPostCode & "', '"& TotalPrice & "', '"& Discount & "', '"& Payable & "', N'"& PayableText & "', '"& CashPayment & "', '" & totalVat & "','" & NorRCode & "'); SELECT @@Identity AS NewInvoicePrintForm"
	Set RS1 = conn.Execute(mySQL).NextRecordSet
	NewInvoicePrintForm = RS1("NewInvoicePrintForm")

	RS1.close
	
	for i=request.form("Descriptions").count to 1 step -1
		if request.form("Descriptions")(i)<>"" OR request.form("Qttys")(i)<>"" OR request.form("Units")(i)<>"" OR request.form("Fees")(i)<>"" OR request.form("Prices")(i)<>"" then exit for
	next
	lastIndex = i

	for i=1 to lastIndex
		theDescription =	request.form("Descriptions")(i)
		theQtty =			text2value(request.form("Qttys")(i))
		theUnit =			request.form("Units")(i)
		theFee =			text2value(request.form("Fees")(i))
		thePrice =			text2value(request.form("Prices")(i))
		theVat =			text2value(request.form("Vat")(i))
		theDis =			text2Value(request.Form("Dis")(i))

		mySQL="INSERT INTO InvoicePrintFormLines (InvoicePrintForm, Description, Qtty, Unit, Fee, Price, Vat, Discount) VALUES ('"& NewInvoicePrintForm & "', N'" & theDescription & "', " & theQtty & ", N'" & theUnit & "', " & theFee & ", " & thePrice & ", " & theVat &", " & theDis & ")"
		conn.Execute(mySQL)
	next 

		response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  ›—„ ›«ﬂ Ê— À»  ‘œ.")

elseif request("act")="submitPrintFormEdit" then

	InvoiceID=clng(request.form("InvoiceID"))
		
	if isnumeric(InvoiceID) then
		mySQL="SELECT ID FROM InvoicePrintForms WHERE (InvoiceID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&errmsg=" & Server.URLEncode("ﬁ»·« »—«Ì «Ì‰ ›«ﬂ Ê— ›—„ ç«ÅÌ «ÌÃ«œ ‰‘œÂ «” .")
		end if
		InvoicePrintFormID=RS1("ID")
		RS1.close
	else
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if

	editDate=		shamsiToday()

	InvoiceDate=		request.form("InvoiceDate")
	CustomerName=		request.form("CustomerName")
	customerEcCode=		request.form("customerEcCode")
	customerAddress=	request.form("customerAddress")
	customerPostCode=	request.form("customerPostCode")
	customerTel=		request.form("customerTel")
	NorRCode =			request.Form("NorRCode")
	if request.form("CashPayment") then
		CashPayment=1
	else
		CashPayment=0
	end if
	
	TotalPrice=	text2value(request.form("TotalPrice"))
	Discount=	text2value(request.form("Discount"))
	Payable=	text2value(request.form("Payable"))
	totalVat=		text2value(request.form("totalVat"))
	tmpPay=Payable
	PayableText=ConvertIT(tmpPay)
	
	MustBe=text2value(request.form("MustBe"))
	if Payable<>MustBe then
		response.redirect "?act=getPrintForm&invoice="& InvoiceID & "&errMsg=" & Server.URLEncode("Ã„⁄ ›«ﬂ Ê— ’ÕÌÕ ‰Ì” .")
	end if

	mySQL="UPDATE InvoicePrintForms SET LastEditedDate=N'"& editDate & "', LastEditedBy='"& session("ID") & "', InvoiceID='"& InvoiceID & "', [Date]=N'"& InvoiceDate & "', Name=N'"& CustomerName & "', Address=N'"& customerAddress & "', Tel=N'"& customerTel & "', EcCode=N'"& customerEcCode & "', PostCode=N'"& customerPostCode & "', TotalPrice='"& TotalPrice & "', Discount='"& Discount & "', Payable='"& Payable & "', PayableText=N'"& PayableText & "', CashPayment='"& CashPayment & "', Vat='" & totalVat & "', NorRCode='" & NorRCode & "' WHERE (ID = "& InvoicePrintFormID & ")"
'response.write "<br>" & mySQL
'response.end
	conn.Execute(mySQL)
	
	mySQL="DELETE FROM InvoicePrintFormLines WHERE (InvoicePrintForm = "& InvoicePrintFormID & ")"
	conn.Execute(mySQL)

	for i=request.form("Descriptions").count to 1 step -1
		if request.form("Descriptions")(i)<>"" OR request.form("Qttys")(i)<>"" OR request.form("Units")(i)<>"" OR request.form("Fees")(i)<>"" OR request.form("Prices")(i)<>"" then exit for
	next
	lastIndex = i

	for i=1 to lastIndex

		theDescription =	request.form("Descriptions")(i)
		theQtty =			text2value(request.form("Qttys")(i))
		theUnit =			request.form("Units")(i)
		theFee =			text2value(request.form("Fees")(i))
		thePrice =			text2value(request.form("Prices")(i))
		theVat =			text2value(request.form("Vat")(i))
		theDis =			text2value(request.Form("Dis")(i))

		mySQL="INSERT INTO InvoicePrintFormLines (InvoicePrintForm, Description, Qtty, Unit, Fee, Price, Vat, Discount) VALUES ("& InvoicePrintFormID & ", N'" & theDescription & "', " & theQtty & ", N'" & theUnit & "', " & theFee & ", " & thePrice & ", " & theVat & ", " & theDis & ")"
		conn.Execute(mySQL)
	next 

		response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&msg=" & Server.URLEncode("«ÿ·«⁄«  ›—„ ›«ﬂ Ê— »Â —Ê“ ‘œ.")

elseif request("act")="showPrintForm" then
	if isnumeric(request("invoice")) then
		InvoiceID=clng(request("invoice"))
		mySQL="SELECT * FROM InvoicePrintForms WHERE (InvoiceID="& InvoiceID & ")"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?act=getPrintForm&invoice=" & InvoiceID & "&errmsg=" & Server.URLEncode("ﬁ»·« »—«Ì «Ì‰ ›«ﬂ Ê— ›—„ ç«ÅÌ «ÌÃ«œ ‰‘œÂ «” .")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â ›«ﬂ Ê— ﬁ«»· ﬁ»Ê· ‰„Ì »«‘œ.")
	end if
	set rsVat=conn.Execute("select isnull(GLs.Vat,0) as Vat from ARItems inner join GLs on ARItems.GL=GLs.ID where Type=1 AND Link="&InvoiceID)
	PrintFormID=		RS1("ID")
	InvoiceDate=		RS1("Date")

	totalPrice=			cdbl(RS1("totalPrice"))
	totalDiscount=		cdbl(RS1("Discount"))
	totalReceivable=	cdbl(RS1("Payable"))
	totalVat =			cdbl(RS1("Vat"))

	customerName=		RS1("Name")
	customerAddress=	RS1("Address")
	customerTel=		RS1("Tel")
	customerPostCode=	RS1("PostCode")
	customerEcCode=		RS1("EcCode")
	cashPayment=		RS1("cashPayment")
	NorRCode =			RS1("NorRCode")
	VatRate =			rsVat("Vat")
	RS1.close

%>
<!-- ‰„«Ì‘ «ÿ·«⁄«  ›—„ ›«ﬂ Ê— -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
	<FORM METHOD=POST ACTION="?act=submitPrintFormEdit">
	<tr bgcolor='#C3C300'>
		<td align="left"><TABLE>
			<TR>   
				<TD align="left"> «—ÌŒ:</td>
				<TD dir="LTR">
					<INPUT class="InvGenInput" NAME="InvoiceDate" TYPE="text" maxlength="10" size="10" value="<%=InvoiceDate%>"></td>
					<INPUT TYPE="hidden" NAME="customerID" value="<%=customerID%>">
					<INPUT TYPE="hidden" NAME="InvoiceID" value="<%=InvoiceID%>">
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
				<TD align="left" rowspan="2">‰‘«‰Ì:</TD>
				<TD rowspan="2"><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerAddress" rows="2" cols="50"><%=customerAddress%></TEXTAREA></TD>
				<TD align="left">ﬂœ Å” Ì:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;" NAME="customerPostCode" TYPE="text" size="10" value="<%=customerPostCode%>"></TD>
			</TR>
			<tr>
				<td align='left'>‘„«—Â À» /‘„«—Â „·Ì:</td>
				<td><input class='InvGenInput' style='border:1 solid black;width:100px;' name="NorRCode" type='text' size='10' value='<%=NorRCode%>'></td>
			</tr>
			<TR>   
				<TD align="left"> ·›‰:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerTel" rows="1" cols="50"><%=customerTel%></TEXTAREA></TD>
				<TD align="left">‘—«Ìÿ ›—Ê‘:</TD>
				<TD align=left><INPUT TYPE="radio" NAME="CashPayment" Value="1" <%if cashPayment then response.write "checked"%>>‰ﬁœÌ  <INPUT TYPE="radio" NAME="CashPayment" Value="0" <%if not cashPayment then response.write "checked"%>>€Ì—‰ﬁœÌ</TD>
			</TR>  
			</TABLE>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td><div>
		<input type='hidden' id='VatRate' name='VatRate' value='<%=VatRate%>'>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TR bgcolor='#CCCC88'>
				<TD align='center' width="25px"> # </td>
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" value="‘—Õ ”›«—‘" size="65" ></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" value=" ⁄œ«œ" size="5"></TD>   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="Ê«Õœ ﬂ«·«" size="8"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="ﬁÌ„  Ê«Õœ" size="8"></TD>
				<TD><input class="InvHeadInput" readOnly type="text" value=" Œ›Ì›" size="8"></TD>
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" Value="„»·€" size="13"></TD> 
				<TD><INPUT class="InvHeadInput4" readOnly TYPE="text" Value="„«·Ì« " size="8"></TD>
			</TR>
			</TABLE></div>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td>
			<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
			<TBODY id="InvoiceLines">
<%		
		i=0
		totalNoVat = 0
		totalSumDis = 0

		mySQL="SELECT * FROM InvoicePrintFormLines WHERE (InvoicePrintForm='"& PrintFormID & "')"
		Set RS1 = conn.Execute(mySQL)
		Do While NOT RS1.eof AND i<10
			i=i+1
			Descriptions(i)=	RS1("Description")
			Qttys(i)=			Separate(RS1("Qtty"))
			Fees(i)=			Separate(RS1("Fee")) 
			Units(i)=			RS1("Unit")
			Prices(i)=			Separate(RS1("Price"))
			Vat(i) =			separate(RS1("Vat"))
			Dis(i) =			separate(RS1("Discount"))
			totalSumDis =		CDbl(RS1("Discount")) + totalSumDis
			totalNoVat =		totalNoVat + cdbl(RS1("Price"))
			RS1.moveNext
		Loop
		RS1.close
		If totalSumDis = 0 Then
			Dis(1) = totalDiscount
		End if

		for i=1 to 10
			%>
			<TR bgcolor='#F0F0F0'>
				<TD align='center' width="25px" onclick="selectAndDelete(this);"><%=i%></TD>
				<TD><INPUT class="InvRowInput2" TYPE="text" Name="Descriptions" value="<%=Descriptions(i)%>" size="65"></TD>
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Qttys" value="<%=Qttys(i)%>" size="5" onBlur="setPrice(this);"></TD>   
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Units" Value="<%=Units(i)%>" size="8" onKeyPress="autoComplete(this);"></TD>        
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Fees" Value="<%=Fees(i)%>" size="8" onBlur="setPrice(this);"></TD>  
				<TD><input class="InvRowInput" type="text" name="Dis" value="<%=Dis(i)%>" size="8" onBlur="setPrice(this);"><TD>
				<TD><INPUT readonly class="InvRowInput3" TYPE="text" Name="Prices" Value="<%=Prices(i)%>" size="13" onBlur="this.value=echoNum(getNum(this.value));calcAndCheck();"></TD>    
				<TD><input readonly class='InvRowInput4' type='text' name='Vat' value='<%=Vat(i)%>' size='8'></TD>
			</TR>
<%
		next
%>
			</Tbody>
			</TABLE>
		</td>
	</tr>
	<tr bgcolor='#CCCC88'>
		<td colspan="10">
		<div>
			<TABLE align='left' Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#CCCC88">
			<TR bgcolor='#CCCC88'>      
				<TD align='left'>Ã„⁄:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalNoVatPrice" Value="<%=Separate(totalNoVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'><!--S A M--> 
				<TD align='left'>„«·Ì«  »— «—“‘ «›“ÊœÂ:</TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="totalVat" Value="<%=Separate(totalVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>      
				<TD align='left'>Ã„⁄ ﬂ·:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalPrice" Value="<%=Separate(totalPrice)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>        
				<TD align='left'> Œ›Ì›:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Discount" Value="<%=Separate(totalDiscount)%>" size="20"></TD>      
			</TR>
			</TR>
			<TR bgcolor='#CCCC88'>
				<TD align='right'><input class="InvHeadInput" type='text' readonly name='message' value='' size='30'></TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Payable" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>
			
			<TR bgcolor='#CCCC88'>
				<TD align='left'>ﬁ«»· Å—œ«Œ :</TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="MustBe" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>			
			</TABLE>
		</div>
		</td>
	</tr>
	<tr>
		<td align='center'>
			<% 	
				mySQL="SELECT * FROM Invoices WHERE (ID="& InvoiceID & ")"
				Set RS1 = conn.Execute(mySQL)
				If CInt(mid(RS1("IssuedDate"),3,2)) >= 88 Then
					reportStr = "InvoicePrintForm88.rpt"
				Else
					reportStr = "InvoicePrintForm.rpt"
				End If 
				RS1.close
				ReportLogRow = PrepareReport (reportStr, "InvoiceID", InvoiceID, "/beta/dialog_printManager.asp?act=Fin") 
			%>
			<INPUT class="InvGenButton" TYPE="submit" value=" À»   €ÌÌ—«  " onclick="return calcAndCheck();">
			<INPUT class="InvGenButton" TYPE="button" value=" ç«Å " onclick="printThisReport(this,<%=ReportLogRow%>);">
			<INPUT class="InvGenButton" TYPE="button" value=" «‰’—«› " onclick="window.location='?';">
		</td>
	</tr>
	</table>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	
//		document.getElementsByName("Items")[0].focus();
		calcAndCheck();
	
	</SCRIPT>
<%
else
%>
	<br>
	<FORM METHOD=POST ACTION="?act=submitsearch">
	<div dir='rtl'><B><FONT SIZE="" COLOR="red"> &nbsp;«‰ Œ«» ›«ﬂ Ê—: </FONT></B>
		<BR><BR>‘„«—Â ›«ﬂ Ê—:<INPUT style="font-family:Tahoma;width:100px;" TYPE="text" NAME="invoice">&nbsp;<INPUT class="GenButton" TYPE="submit" Name="submitShow" value="«œ«„Â">
		<BR>Ì« <BR>‘„«—Â ”›«—‘:<INPUT TYPE="text" NAME="query" onfocus="document.getElementsByName('invoice')[0].value='';">&nbsp;
		
	</div>
	</FORM>
	<%
	set rs=Conn.Execute("select top 50 Invoices.ID,Invoices.Customer,Invoices.IssuedDate,Users.RealName,Accounts.AccountTitle,Invoices.TotalReceivable from Invoices inner join Users on Invoices.IssuedBy=Users.ID inner join Accounts on Invoices.Customer=Accounts.ID where Invoices.Voided=0 and Invoices.Issued=1 and Invoices.IsA=1 and Invoices.ID not in (select InvoiceID from InvoicePrintForms where Voided=0) order by IssuedDate desc")
	
	do while not rs.eof
		%>
		›«ﬂ Ê— ‘„«—Â <a href='../AR/AccountReport.asp?act=showInvoice&invoice=<%=rs("ID")%>'><%=rs("ID")%></a> »Â „»·€ <%=Separate(rs("TotalReceivable"))%> „—»Êÿ »Â <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rs("Customer")%>'><%=rs("AccountTitle")%></a> œ—  «—ÌŒ <%=rs("IssuedDate")%> ’«œ— ‘œÂ Ê Â‰Ê“ ›—„ ›«ﬂ Ê— ¬‰ <a  href='InvoicePrintForm.asp?act=getPrintForm&invoice=<%=rs("ID")%>'>ç«Å</a> ‰‘œÂ.<br>
		<%
		rs.moveNext
	loop
	%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.getElementsByName("invoice")[0].focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
<!--#include file="tah.asp" -->
