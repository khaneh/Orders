<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="�ǁ ��� ������"
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
var selectedRow=-1;
function selectAndDelete(rowNo){
	if (selectedRow==rowNo){
		delAndShiftUp(rowNo);
		document.getElementById("InvoiceLines").getElementsByTagName("TR")[selectedRow].getElementsByTagName("TD")[0].setAttribute("bgColor","");
		selectedRow=-1;
		calcAndCheck();
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
		document.getElementsByName("Descriptions")[i].value=document.getElementsByName("Descriptions")[i+1].value;
		document.getElementsByName("Qttys")[i].value=document.getElementsByName("Qttys")[i+1].value;
		document.getElementsByName("Units")[i].value=document.getElementsByName("Units")[i+1].value;
		document.getElementsByName("Fees")[i].value=document.getElementsByName("Fees")[i+1].value;
		document.getElementsByName("Prices")[i].value=document.getElementsByName("Prices")[i+1].value;
		document.getElementsByName("Dis")[i].value=document.getElementsByName("Dis")[i+1].value;
		document.getElementsByName("Vat")[i].value=document.getElementsByName("Vat")[i+1].value;
	}
	document.getElementsByName("Descriptions")[9].value="";
	document.getElementsByName("Qttys")[9].value="";
	document.getElementsByName("Units")[9].value="";
	document.getElementsByName("Fees")[9].value="";
	document.getElementsByName("Prices")[9].value="";
	document.getElementsByName("Dis")[9].value="";
	document.getElementsByName("Vat")[9].value="";
}
function setPrice(src){
	if (src.name == "Qttys" || src.name == "Fees" || src.name == "Dis"){
		src.value = val2txt(txt2val(src.value))
		rowNo = src.parentNode.parentNode.rowIndex;
		tmpFee = txt2val(document.getElementsByName("Fees")[rowNo].value);
		tmpQtty = txt2val(document.getElementsByName("Qttys")[rowNo].value);
		vat = txt2val(document.getElementsByName("Vat")[rowNo].value);
		dis = txt2val(document.getElementsByName("Dis")[rowNo].value);
		
		tmpPrice= tmpFee * tmpQtty - dis;
		//alert(val2txt(tmpPrice * .03));
		if (vat > 0){
			document.getElementsByName("Vat")[rowNo].value = Math.floor(tmpPrice * txt2val(document.getElementById("VatRate").value)/100); //sam change this in 90
		}
		document.getElementsByName("Prices")[rowNo].value = val2txt(parseInt(tmpPrice));
	}

	calcAndCheck();
}
function calcAndCheck(){
	var total=0;
	var dis=0;
	var payable=0;
	var vat=0;
	var checkVat=false;
//alert('hi');
	for(i = 0; i < 10; i++){
		total += txt2val(document.getElementsByName("Prices")[i].value);
		dis += txt2val(document.getElementsByName("Dis")[i].value);
		vat += txt2val(document.getElementsByName("Vat")[i].value);
	}
	if (txt2val(document.getElementsByName("totalVat")[0].value) != vat){
		document.getElementsByName("message")[0].value = "�����! ��� ������ ������ ������.";
		document.getElementsByName("totalVat")[0].style.backgroundColor = "#FF4848";
		checkVat = false;			
	}
	else{
		document.getElementsByName("message")[0].value = "";
		document.getElementsByName("totalVat")[0].style.backgroundColor = "#F0F0F0";
		checkVat = true;
	}

	document.getElementsByName("TotalPrice")[0].value = val2txt(total + vat);
	document.getElementsByName("Discount")[0].value = val2txt(dis);
	payable = total + vat;
	//payable += txt2val(document.getElementsByName("Vat")[0].value);
	document.getElementsByName("Payable")[0].value=val2txt(payable);

	if(document.getElementsByName("Payable")[0].value==document.getElementsByName("MustBe")[0].value){
		document.getElementsByName("Payable")[0].style.backgroundColor='#00FF00';
		if (checkVat){
			return true;
		}
		else{
			return false;
		}
		
	}
	else{
		document.getElementsByName("Payable")[0].style.backgroundColor='#FF0000';
		return false;
	}
}
function autoComplete(src){
	if (src.value==""){
//		alert(event.keyCode);
		switch (event.keyCode){
		case 1576: // �
			src.value="�ѐ";
			event.keyCode=0;
			break;
		case 1608: // �
			src.value="���";
			event.keyCode=0;
			break;
		}
	}
	else if (src.value=="�") {
		switch (event.keyCode){
		case 1585: // �
			src.value="���";
			event.keyCode=0;
			break;
		case 1575: // �
			src.value="��������";
			event.keyCode=0;
			break;
		}
	}
}
function combine()
{
	priceVat = 0;
	priceNoVat = 0;
	totalVat = 0;
	totalDisOnVat = 0;
	totalDisNoVat = 0;
	rfd = 0;

	for(i = 0; i < 10; i++)
	{
		if (txt2val(document.getElementsByName("Vat")[i].value) == 0)
		{
			if (txt2val(document.getElementsByName("Prices")[i].value) <= 0)
			{
				rfd += txt2val(document.getElementsByName("Dis")[i].value);
			}
			else{
				totalDisNoVat += txt2val(document.getElementsByName("Dis")[i].value);
				priceNoVat += txt2val(document.getElementsByName("Fees")[i].value) * txt2val(document.getElementsByName("Qttys")[i].value);
			}
		}
		else
		{	
			priceVat += txt2val(document.getElementsByName("Fees")[i].value) * txt2val(document.getElementsByName("Qttys")[i].value);
			totalVat += txt2val(document.getElementsByName("Vat")[i].value);
			if (txt2val(document.getElementsByName("Prices")[i].value) <= 0)
			{
				rfd += txt2val(document.getElementsByName("Dis")[i].value);
			}
			else{
				totalDisOnVat += txt2val(document.getElementsByName("Dis")[i].value);
			}
		}
		
		document.getElementsByName("Descriptions")[i].value = "";
		document.getElementsByName("Qttys")[i].value = "";
		document.getElementsByName("Units")[i].value = "";
		document.getElementsByName("Fees")[i].value = "";
		document.getElementsByName("Prices")[i].value = "";
		document.getElementsByName("Vat")[i].value = "";
		document.getElementsByName("Dis")[i].value = "";
	}
	if (priceNoVat > 0)
	{
		document.getElementsByName("Descriptions")[1].value = "���� � ���� ��� �ǁ";
		document.getElementsByName("Qttys")[1].value = "1";
		document.getElementsByName("Units")[1].value = "";
		document.getElementsByName("Fees")[1].value = val2txt(parseInt(priceNoVat));
		document.getElementsByName("Prices")[1].value = val2txt(parseInt(priceNoVat) - parseInt(totalDisNoVat));
		document.getElementsByName("Vat")[1].value = "0";
		document.getElementsByName("Dis")[1].value = val2txt(parseInt(totalDisNoVat));
	}
	document.getElementsByName("Descriptions")[0].value = "����� �ǁ";
	document.getElementsByName("Qttys")[0].value = "1";
	document.getElementsByName("Units")[0].value = "";
	document.getElementsByName("Fees")[0].value = val2txt(parseInt(priceVat));
	document.getElementsByName("Prices")[0].value = val2txt(parseInt(priceVat) - parseInt(totalDisOnVat));
	document.getElementsByName("Vat")[0].value = val2txt(parseInt(totalVat));
	document.getElementsByName("Dis")[0].value = val2txt(parseInt(totalDisOnVat));
	//---------------------------
	document.getElementsByName("Descriptions")[2].value = "����� ��� ������";
	document.getElementsByName("Qttys")[2].value = "1";
	document.getElementsByName("Units")[2].value = "";
	document.getElementsByName("Fees")[2].value = "0";
	document.getElementsByName("Prices")[2].value = val2txt(parseInt(-rfd));
	document.getElementsByName("Vat")[2].value = "0";
	document.getElementsByName("Dis")[2].value = val2txt(parseInt(rfd));
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
			response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&msg=" & Server.URLEncode("���� ���� ��� ������ ��� �ǁ� ����� ��� ���.<br>����� �� �� ������ �� ���� � �� ������ ����� ����.")
		end if
	else
		if request("query")<>"" then
			response.redirect "?msg=" & Server.URLEncode("������ ���� ����� ����� ����� ��� ����.")
		else
			response.redirect "?errmsg=" & Server.URLEncode("����� ������ ���� ���� ��� ����.")
		end if
	end if
elseif request("act")="getPrintForm" then
	if isnumeric(request("invoice")) then
		InvoiceID=clng(request("invoice"))
		mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?errmsg=" & Server.URLEncode("������ ���� ���.")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("����� ������ ���� ���� ��� ����.")
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
		response.redirect "?errmsg=" & Server.URLEncode("��� ������ ���� ��� ���.")
	elseif isReverse then
		Conn.close
		response.redirect "?errmsg=" & Server.URLEncode("����� ������ ����� ����.")
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
<!-- ���� ������� ��� ������ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
	<FORM METHOD=POST ACTION="?act=submitPrintForm">
	<tr bgcolor='#C3C300'>
		<td align="left"><TABLE>
			<TR>   
				<TD align="left">�����:</td>
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
				<TD align="left">��� ������:</TD>
				<TD><TEXTAREA readonly style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="CustomerName" rows="1" cols="50"><%=CustomerName%></TEXTAREA></TD>
				<TD align="left">����� �������:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;direction:LTR;text-align:left;" NAME="customerEcCode" TYPE="text" size="10" value="<%=customerEcCode%>"></TD>
			</TR>  
			<TR>   
				<TD align="left" rowspan="2">�����:</TD>
				<TD rowspan="2"><TEXTAREA readonly style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerAddress" rows="2" cols="50"><%=customerAddress%></TEXTAREA></TD>
				<TD align="left">�� ����:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;" NAME="customerPostCode" TYPE="text" size="10" value="<%=customerPostCode%>"></TD>
			</TR>  
			<tr>
				<td align='left'>����� ���/����� ���:</td>
				<td><input class='InvGenInput' style='border:1 solid black;width:100px;' name="NorRCode" type='text' size='10' value='<%=NorRCode%>'></td>
			</tr>
			<TR>   
				<TD align="left">����:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerTel" rows="1" cols="50"><%=customerTel%></TEXTAREA></TD>
				<TD align="left">����� ����:</TD>
				<TD align=left><INPUT TYPE="radio" NAME="CashPayment" Value="1">����  <INPUT TYPE="radio" NAME="CashPayment" Value="0" checked>�������</TD>
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
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" value="��� �����" size="65" ></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" value="�����" size="5"></TD> <!--S A M-->   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="���� ����" size="8"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="���� ����" size="8"></TD>        
				<TD><Input class="InvHeadInput" readOnly type="text" Value="�����" size="8"><TD>
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" Value="����" size="13"></TD>      
				<TD><INPUT class="InvHeadInput4" readOnly TYPE="text" Value="������" size="8"></TD>
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
			Descriptions(8) = "���� ����� ������"
			Prices(8) = totalSumAfter9Vat
			Dis(8) = totalDisAfter9Vat
			Vat(8) = totalVatAfter9Vat
			Fees(8) = Prices(8)
			Qttys(8) = 1
			Descriptions(9) = "���� ����� ���� ������ ������"
			Prices(9) = totalSumAfter9noVat
			Dis(9) = totalDisAfter9noVat
			Vat(9) = 0
			Fees(9) = Prices(9)
			Qttys(9) = 1
		End If 
			
		for i=1 to 10
%>
			<TR bgcolor='#F0F0F0'>
				<TD align='center' width="25px" onclick="selectAndDelete(this.parentNode.rowIndex);"><%=i%></TD>
				<TD><INPUT class="InvRowInput2" TYPE="text" Name="Descriptions" value="<%=Descriptions(i)%>" size="65"></TD>
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Qttys" value="<%=Qttys(i)%>" size="5" onBlur="setPrice(this);"></TD>   
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Units" Value="<%=Units(i)%>" size="8" onKeyPress="autoComplete(this);"></TD>       
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Fees" Value="<%=Fees(i)%>" size="8" onBlur="setPrice(this);"></TD>        
				<TD><input class="InvRowInput" type="text" name="Dis" value="<%=Dis(i)%>" Size="8" onBlur="setPrice(this)"></TD>
				<TD><INPUT readonly class="InvRowInput3" TYPE="text" Name="Prices" Value="<%=Prices(i) - Dis(i)%>" size="13" onBlur="this.value=val2txt(txt2val(this.value));calcAndCheck();"></TD>
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
				<TD align='left'>���:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalNoVatPrice" Value="<%=Separate(totalNoVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'><!--S A M--> 
				<TD align='left'>������ �� ���� ������:</TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="totalVat" Value="<%=Separate(totalVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>      
				<TD align='left'>��� ��:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalPrice" Value="<%=Separate(totalPrice)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>        
				<TD align='left'>�����:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Discount" Value="<%=Separate(totalDiscount)%>" size="20"></TD>      
			</TR>
			</TR>
			<TR bgcolor='#CCCC88'>
				<TD align='right'><input class="InvHeadInput" type='text' readonly name='message' value='' size='30'></TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Payable" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>
			
			<TR bgcolor='#CCCC88'>
				<TD align='left'>���� ������:</TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="MustBe" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>			
			</TABLE>
		</div>
		</td>
	</tr>
	<tr>
		<td align='center'>
			<INPUT class="InvGenButton" TYPE="submit" value=" ��� " onclick="return calcAndCheck();">
			<input class='InvGenButton' type='button' Value='�� ��� ��� ������' onclick='combine();'>
		</td>
	</tr>
	</table>
	</FORM>
	<br>
	<SCRIPT LANGUAGE="JavaScript">
	
//		document.getElementsByName("Items")[0].focus();
		calcAndCheck();
	
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
			response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&errmsg=" & Server.URLEncode("���� ���� ��� ������ ��� �ǁ� ����� ��� ���.")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("����� ������ ���� ���� ��� ����.")
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
		response.redirect "?act=getPrintForm&invoice="& InvoiceID & "&errMsg=" & Server.URLEncode("��� ������ ���� ����.")
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

		response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&msg=" & Server.URLEncode("������� ��� ������ ��� ��.")

elseif request("act")="submitPrintFormEdit" then

	InvoiceID=clng(request.form("InvoiceID"))
		
	if isnumeric(InvoiceID) then
		mySQL="SELECT ID FROM InvoicePrintForms WHERE (InvoiceID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&errmsg=" & Server.URLEncode("���� ���� ��� ������ ��� �ǁ� ����� ���� ���.")
		end if
		InvoicePrintFormID=RS1("ID")
		RS1.close
	else
		response.redirect "?errmsg=" & Server.URLEncode("����� ������ ���� ���� ��� ����.")
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
		response.redirect "?act=getPrintForm&invoice="& InvoiceID & "&errMsg=" & Server.URLEncode("��� ������ ���� ����.")
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

		response.redirect "?act=showPrintForm&invoice=" & InvoiceID & "&msg=" & Server.URLEncode("������� ��� ������ �� ��� ��.")

elseif request("act")="showPrintForm" then
	if isnumeric(request("invoice")) then
		InvoiceID=clng(request("invoice"))
		mySQL="SELECT * FROM InvoicePrintForms WHERE (InvoiceID="& InvoiceID & ")"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			conn.close
			response.redirect "?act=getPrintForm&invoice=" & InvoiceID & "&errmsg=" & Server.URLEncode("���� ���� ��� ������ ��� �ǁ� ����� ���� ���.")
		end if
	else
		response.redirect "?errmsg=" & Server.URLEncode("����� ������ ���� ���� ��� ����.")
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
<!-- ����� ������� ��� ������ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<table Border="0" align="center" Width="100%" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855">
	<FORM METHOD=POST ACTION="?act=submitPrintFormEdit">
	<tr bgcolor='#C3C300'>
		<td align="left"><TABLE>
			<TR>   
				<TD align="left">�����:</td>
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
				<TD align="left">��� ������:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="CustomerName" rows="1" cols="50"><%=CustomerName%></TEXTAREA></TD>
				<TD align="left">����� �������:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;direction:LTR;text-align:left;" NAME="customerEcCode" TYPE="text" size="10" value="<%=customerEcCode%>"></TD>
			</TR>  
			<TR>   
				<TD align="left" rowspan="2">�����:</TD>
				<TD rowspan="2"><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerAddress" rows="2" cols="50"><%=customerAddress%></TEXTAREA></TD>
				<TD align="left">�� ����:</TD>
				<TD><INPUT class="InvGenInput" style="border:1 solid black;width:100px;" NAME="customerPostCode" TYPE="text" size="10" value="<%=customerPostCode%>"></TD>
			</TR>
			<tr>
				<td align='left'>����� ���/����� ���:</td>
				<td><input class='InvGenInput' style='border:1 solid black;width:100px;' name="NorRCode" type='text' size='10' value='<%=NorRCode%>'></td>
			</tr>
			<TR>   
				<TD align="left">����:</TD>
				<TD><TEXTAREA style="font-family:tahoma;font-size:9pt;direction:RTL;border:1 solid black;width:450px;" NAME="customerTel" rows="1" cols="50"><%=customerTel%></TEXTAREA></TD>
				<TD align="left">����� ����:</TD>
				<TD align=left><INPUT TYPE="radio" NAME="CashPayment" Value="1" <%if cashPayment then response.write "checked"%>>����  <INPUT TYPE="radio" NAME="CashPayment" Value="0" <%if not cashPayment then response.write "checked"%>>�������</TD>
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
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" value="��� �����" size="65" ></TD>       
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" value="�����" size="5"></TD>   
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="���� ����" size="8"></TD>        
				<TD><INPUT class="InvHeadInput" readOnly TYPE="text" Value="���� ����" size="8"></TD>
				<TD><input class="InvHeadInput" readOnly type="text" value="�����" size="8"></TD>
				<TD><INPUT class="InvHeadInput2" readOnly TYPE="text" Value="����" size="13"></TD> 
				<TD><INPUT class="InvHeadInput4" readOnly TYPE="text" Value="������" size="8"></TD>
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
				<TD align='center' width="25px" onclick="selectAndDelete(this.parentNode.rowIndex);"><%=i%></TD>
				<TD><INPUT class="InvRowInput2" TYPE="text" Name="Descriptions" value="<%=Descriptions(i)%>" size="65"></TD>
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Qttys" value="<%=Qttys(i)%>" size="5" onBlur="setPrice(this);"></TD>   
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Units" Value="<%=Units(i)%>" size="8" onKeyPress="autoComplete(this);"></TD>        
				<TD><INPUT class="InvRowInput"  TYPE="text" Name="Fees" Value="<%=Fees(i)%>" size="8" onBlur="setPrice(this);"></TD>  
				<TD><input class="InvRowInput" type="text" name="Dis" value="<%=Dis(i)%>" size="8" onBlur="setPrice(this);"><TD>
				<TD><INPUT readonly class="InvRowInput3" TYPE="text" Name="Prices" Value="<%=Prices(i)%>" size="13" onBlur="this.value=val2txt(txt2val(this.value));calcAndCheck();"></TD>    
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
				<TD align='left'>���:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalNoVatPrice" Value="<%=Separate(totalNoVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'><!--S A M--> 
				<TD align='left'>������ �� ���� ������:</TD>        
				<TD><INPUT class="InvHeadInput3" style="color:gray;" readOnly TYPE="text" Name="totalVat" Value="<%=Separate(totalVat)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>      
				<TD align='left'>��� ��:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput2" readOnly TYPE="text" Name="TotalPrice" Value="<%=Separate(totalPrice)%>" size="20"></TD>      
			</TR>
			<TR bgcolor='#CCCC88'>        
				<TD align='left'>�����:</TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Discount" Value="<%=Separate(totalDiscount)%>" size="20"></TD>      
			</TR>
			</TR>
			<TR bgcolor='#CCCC88'>
				<TD align='right'><input class="InvHeadInput" type='text' readonly name='message' value='' size='30'></TD>        
				<TD style="border:1 solid black;"><INPUT class="InvRowInput3" readOnly TYPE="text" Name="Payable" Value="<%=Separate(totalReceivable)%>" size="20"></TD>      
			</TR>
			
			<TR bgcolor='#CCCC88'>
				<TD align='left'>���� ������:</TD>        
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
			<INPUT class="InvGenButton" TYPE="submit" value=" ��� ������� " onclick="return calcAndCheck();">
			<INPUT class="InvGenButton" TYPE="button" value=" �ǁ " onclick="printThisReport(this,<%=ReportLogRow%>);">
			<INPUT class="InvGenButton" TYPE="button" value=" ������ " onclick="window.location='?';">
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
	<div dir='rtl'><B><FONT SIZE="" COLOR="red"> &nbsp;������ ������: </FONT></B>
		<BR><BR>����� ������:<INPUT style="font-family:Tahoma;width:100px;" TYPE="text" NAME="invoice">&nbsp;<INPUT class="GenButton" TYPE="submit" Name="submitShow" value="�����">
		<BR>�� <BR>����� �����:<INPUT TYPE="text" NAME="query" onfocus="document.getElementsByName('invoice')[0].value='';">&nbsp;
		
	</div>
	</FORM>
	<%
	set rs=Conn.Execute("select top 50 Invoices.ID,Invoices.Customer,Invoices.IssuedDate,Users.RealName,Accounts.AccountTitle,Invoices.TotalReceivable from Invoices inner join Users on Invoices.IssuedBy=Users.ID inner join Accounts on Invoices.Customer=Accounts.ID where Invoices.Voided=0 and Invoices.Issued=1 and Invoices.IsA=1 and Invoices.ID not in (select InvoiceID from InvoicePrintForms where Voided=0) order by IssuedDate desc")
	
	do while not rs.eof
		%>
		������ ����� <a href='../AR/AccountReport.asp?act=showInvoice&invoice=<%=rs("ID")%>'><%=rs("ID")%></a> �� ���� <%=Separate(rs("TotalReceivable"))%> ����� �� <a href='../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rs("Customer")%>'><%=rs("AccountTitle")%></a> �� ����� <%=rs("IssuedDate")%> ���� ��� � ���� ��� ������ �� <a  href='InvoicePrintForm.asp?act=getPrintForm&invoice=<%=rs("ID")%>'>�ǁ</a> ����.<br>
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
