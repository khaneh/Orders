<script language="JavaScript">
<!--
function setCurrentRow(rowNo){
	if (rowNo == -1) rowNo=0;
	invTable=document.getElementById("InvoiceLines");
	theTD=invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#F0F0F0');

	currentRow=rowNo;
	invTable=document.getElementById("InvoiceLines");
	theTD=invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#FFB0B0');
}
function delRow(rowNo){
	invTable=document.getElementById("InvoiceLines");
	theRow=invTable.getElementsByTagName("tr")[rowNo];
	invTable.removeChild(theRow);

	rowsCount=document.getElementsByName("Items").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.bgColor= '#F0F0F0';
		tempTD.innerText= rowNo+1;
	}
}
function addRow(){
	rowNo = currentRow
	invTable=document.getElementById("InvoiceLines");
	theRow=invTable.getElementsByTagName("tr")[rowNo];
	newRow=document.createElement("tr");
	newRow.setAttribute("bgColor", '#f0f0f0');
	newRow.setAttribute("onclick", theRow.getAttribute("onclick"));

	tempTD=document.createElement("td");
	tempTD.innerHTML=rowNo+1
	tempTD.setAttribute("align", 'center');
	tempTD.setAttribute("width", '25');
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput' TYPE='text' NAME='Items' size='3' Maxlength='6' onKeyPress='return mask(this);' onChange='return check(this);' onfocus='setCurrentRow(this.parentNode.parentNode.rowIndex);'><INPUT TYPE='hidden' name='type' value=0><INPUT TYPE='hidden' name='fee' value=0><INPUT type='hidden' name='hasVat' value=0>"

	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Descriptions' size='30'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Lengths' size='2' onBlur='setFeeQtty(this);'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Widths' size='2' onBlur='setFeeQtty(this);'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Qttys' size='3' onBlur='setFeeQtty(this);'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Sets' size='2' onBlur='setFeeQtty(this);'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput' TYPE='text' NAME='AppQttys' size='6' onBlur='setPrice(this);'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput' TYPE='text' NAME='Fees' size='7' onBlur='setPrice(this);'>"	
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT  tabIndex='9999' class='InvRowInput' TYPE='text' NAME='Prices' size='9'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput' TYPE='text' NAME='Discounts' size='7' onBlur='setPrice(this);'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT class='InvRowInput' TYPE='text' NAME='Reverses' size='5' onBlur='setPrice(this);' onfocus='setCurrentRow(this.parentNode.parentNode.rowIndex);'>"
	newRow.appendChild(tempTD);

	// S A M
	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT tabIndex='9998' class='InvRowInput' TYPE='text' NAME='Vat' size='6' readonly>"
	//tempTD.appendChild(tempTD);
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.innerHTML="<INPUT  tabIndex='9999' class='InvRowInput2' TYPE='text' NAME='AppPrices' size='9'>"
	newRow.appendChild(tempTD);

	invTable.insertBefore(newRow,theRow);
	
	rowsCount=document.getElementsByName("Items").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.bgColor= '#F0F0F0';
		tempTD.innerText= rowNo+1;
	}

	invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
}
//--------------------------------------------------------------------------------------------------------------------------------------------------
function setPrice(src){
//	alert(src.name);
	if (src.name=="Discounts" || src.name=="Reverses"){
		if (src.value.substr(src.value.length-1)=='%'){
			src.value=val2txt(txt2val(src.value));
			rowNo=src.parentNode.parentNode.rowIndex;
			src.value=val2txt(parseInt(txt2val(document.getElementsByName("Prices")[rowNo].value) * txt2val(src.value) / 100))
		}
		else{
			src.value=val2txt(txt2val(src.value));
		}
	}
	else{
		src.value=val2txt(txt2val(src.value));
	}
	isA = document.getElementsByName("IsA")[0].checked;
	//alert(isA);
	rowNo=src.parentNode.parentNode.rowIndex;
	tmpFee=txt2val(document.getElementsByName("Fees")[rowNo].value);
	tmpAppQtty=txt2val(document.getElementsByName("AppQttys")[rowNo].value);
	tmpDiscount=txt2val(document.getElementsByName("Discounts")[rowNo].value);
	tmpReverse=txt2val(document.getElementsByName("Reverses")[rowNo].value);
	tmpPrice= tmpFee * tmpAppQtty;
	tmpAppPrice=tmpPrice - tmpDiscount - tmpReverse;
	// S A M
	if (txt2val(document.getElementsByName("hasVat")[rowNo].value) == 1 && isA)
		tmpVat = tmpAppPrice * txt2val(document.getElementById("VatRate").value)/100; //sam change this in 90
	else 
		tmpVat = 0;
	tmpAppPrice += tmpVat;
	document.getElementsByName("Prices")[rowNo].value = val2txt(parseInt(tmpPrice));
	document.getElementsByName("AppPrices")[rowNo].value = val2txt(parseInt(tmpAppPrice));
	document.getElementsByName("Vat")[rowNo].value = val2txt(parseInt(tmpVat));
	
	var totalPrice = 0;
	var totalDiscount = 0;
	var totalReverse = 0;
	var totalAppPrice = 0;
	var payable = 0;
	var totalVat = 0;
	for (rowNo=0; rowNo < document.getElementsByName("Fees").length; rowNo++){
		totalPrice += parseInt(txt2val(document.getElementsByName("Prices")[rowNo].value));
		totalDiscount += parseInt(txt2val(document.getElementsByName("Discounts")[rowNo].value));
		totalReverse += parseInt(txt2val(document.getElementsByName("Reverses")[rowNo].value));
		totalAppPrice += parseInt(txt2val(document.getElementsByName("AppPrices")[rowNo].value));
		if (isA) {
			totalVat += parseInt(txt2val(document.getElementsByName("Vat")[rowNo].value));}
	}
	payable = Math.floor(totalAppPrice/1000) * 1000;
	document.all.TotalPrice.value = val2txt(totalPrice);
	document.all.TotalDiscount.value = val2txt(totalDiscount);
	document.all.TotalReverse.value = val2txt(totalReverse);
	document.all.Payable.value = val2txt(payable);
	document.all.TotalVat.value = val2txt(totalVat);

	if (totalPrice==0){
		document.all.TPDiscount.value = "- "+'% Œ›Ì›';
		document.all.TPReverse.value = "- "+'%»—ê‘ ';
	}
	else{
		document.all.TPDiscount.value = Math.round(totalDiscount/totalPrice * 100)+'% Œ›Ì›';
		document.all.TPReverse.value = Math.round(totalReverse/totalPrice * 100)+'%»—ê‘ ';
	}
}
function checkIsA(){
	isA = document.getElementsByName("IsA")[0].checked;
	var totalPrice = 0;
	var totalDiscount = 0;
	var totalReverse = 0;
	var totalAppPrice = 0;
	var payable = 0;
	var totalVat = 0;
	for (rowNo=0; rowNo < document.getElementsByName("Fees").length; rowNo++){
		tmpFee=txt2val(document.getElementsByName("Fees")[rowNo].value);
		tmpAppQtty=txt2val(document.getElementsByName("AppQttys")[rowNo].value);
		tmpDiscount=txt2val(document.getElementsByName("Discounts")[rowNo].value);
		tmpReverse=txt2val(document.getElementsByName("Reverses")[rowNo].value);
		tmpPrice= tmpFee * tmpAppQtty;
		tmpAppPrice=tmpPrice - tmpDiscount - tmpReverse;
		// S A M
		if (txt2val(document.getElementsByName("hasVat")[rowNo].value) == 1 && isA)
			tmpVat = tmpAppPrice * txt2val(document.getElementById("VatRate").value)/100; //sam change this in 90
		else 
			tmpVat = 0;
		tmpAppPrice += tmpVat;
		document.getElementsByName("Prices")[rowNo].value = val2txt(parseInt(tmpPrice));
		document.getElementsByName("AppPrices")[rowNo].value = val2txt(parseInt(tmpAppPrice));
		document.getElementsByName("Vat")[rowNo].value = val2txt(parseInt(tmpVat));
		
		// ---------------------------------------------
		totalPrice += parseInt(txt2val(document.getElementsByName("Prices")[rowNo].value));
		totalDiscount += parseInt(txt2val(document.getElementsByName("Discounts")[rowNo].value));
		totalReverse += parseInt(txt2val(document.getElementsByName("Reverses")[rowNo].value));
		totalAppPrice += parseInt(txt2val(document.getElementsByName("AppPrices")[rowNo].value));
		if (isA) {
			totalVat += parseInt(txt2val(document.getElementsByName("Vat")[rowNo].value));}
	}
	payable = Math.floor(totalAppPrice/1000) * 1000;
	document.all.TotalPrice.value = val2txt(totalPrice);
	document.all.TotalDiscount.value = val2txt(totalDiscount);
	document.all.TotalReverse.value = val2txt(totalReverse);
	document.all.Payable.value = val2txt(payable);
	document.all.TotalVat.value = val2txt(totalVat);

	if (totalPrice==0){
		document.all.TPDiscount.value = "- "+'% Œ›Ì›';
		document.all.TPReverse.value = "- "+'%»—ê‘ ';
	}
	else{
		document.all.TPDiscount.value = Math.round(totalDiscount/totalPrice * 100)+'% Œ›Ì›';
		document.all.TPReverse.value = Math.round(totalReverse/totalPrice * 100)+'%»—ê‘ ';
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------

var dialogActive=false;
var badCode = false;

var A				= 90		// DuplexFee Add-in
var ProofSimplex	= 3000
var ProofDuplex		= 4500
			  Qtt	= Array (	0 ,	 		1,		50,		150,	300 )
/*SimplexFee*/  SF	= Array (	ProofSimplex,	200,	144,	112,	100 )
/*DuplexFee*/	DF	= Array (	ProofDuplex,	SF[1]+A,	SF[2]+A,	SF[3]+A,	SF[4]+A)

//-----------------------------------------------------------------------------------------------------------------------------------------
function setFeeQtty(src){
	rowNo=src.parentNode.parentNode.rowIndex;
	itemType=parseInt(txt2val(document.getElementsByName("type")[rowNo].value));
	itemFee=document.getElementsByName("fee")[rowNo].value;
	
	if (!document.getElementsByName("Qttys")[rowNo].value == "")
	document.getElementsByName("Qttys")[rowNo].value = parseInt(document.getElementsByName("Qttys")[rowNo].value);

	if (!document.getElementsByName("Sets")[rowNo].value == "")
	document.getElementsByName("Sets")[rowNo].value = parseInt(document.getElementsByName("Sets")[rowNo].value);

	//////////////// Type =1  --->  General  //////////////////
	if (itemType==1 || itemType==5){   
		document.getElementsByName("AppQttys")[rowNo].value = parseInt(txt2val(document.getElementsByName("Qttys")[rowNo].value)) * parseInt(txt2val(document.getElementsByName("Sets")[rowNo].value)); 

		document.getElementsByName("Fees")[rowNo].value =  parseInt(txt2val(itemFee)) 
	
		if (''+document.getElementsByName("AppQttys")[rowNo].value=='NaN')
			document.getElementsByName("AppQttys")[rowNo].value = 0
	}

	//////////////// Type =2  --->  Digital  //////////////////
	if (itemType==2 && itemFee!="0"){   
		PF		= parseInt(txt2val(itemFee.substr(1)));
		tmp		= itemFee.substr(0,1);
		if (tmp == "s" ) 
			SoD = false 
		else 
			SoD = true
		Tirag	= Math.round(parseInt(txt2val(document.getElementsByName("Qttys")[rowNo].value)));
		h		= parseInt(txt2val(document.getElementsByName("Lengths")[rowNo].value));
		Price	= 0
		i		= 1
		document.getElementsByName("Widths")[rowNo].value = 30	
		if (h ==0 ) 
			{
			document.getElementsByName("Lengths")[rowNo].value=21
			h=21
			}

		/*while ( Tirag >= Qtt[ i -1] ) 
			{
			a1 = Tirag - Qtt[ i - 1 ]
			a2 = Tirag - Qtt[ i ]
			if (a2>0)	
				a3 = a1 - a2 
			else 
				a3 = a1
			if ( SoD == false ) 
				Price += ( SF[ i-1 ] + PF ) * a3
			else
				Price += ( DF[ i-1 ] + PF ) * a3

			i++
			}
		*/
		Price = ( 200 + PF ) * Tirag
		if ( SoD == false ) 
			Price = Price * 1
		else
			Price = Price * 2

			
		Price = Math.round(Price / 21 * h) 
		unitPrice = Math.round(Price / Tirag) * 10
		document.getElementsByName("Fees")[rowNo].value = unitPrice 
		
		document.getElementsByName("AppQttys")[rowNo].value = parseInt(txt2val(document.getElementsByName("Qttys")[rowNo].value)) * parseInt(txt2val(document.getElementsByName("Sets")[rowNo].value)); 
	
		if (''+document.getElementsByName("AppQttys")[rowNo].value=='NaN')
			document.getElementsByName("AppQttys")[rowNo].value = 0
		
		if (''+document.getElementsByName("Fees")[rowNo].value=='NaN')
			document.getElementsByName("Fees")[rowNo].value = 0
	}

	//////////////// Type =3  --->  Film  /////////////////////
	if (itemType==3){   
		h	= txt2val(document.getElementsByName("Lengths")[rowNo].value);
		w	= txt2val(document.getElementsByName("Widths")[rowNo].value);
		document.getElementsByName("Fees")[rowNo].value =  parseInt(txt2val(itemFee)) 
		document.getElementsByName("AppQttys")[rowNo].value =  val2txt(txt2val(document.getElementsByName("Qttys")[rowNo].value) * txt2val(document.getElementsByName("Sets")[rowNo].value) * h * w);
	}

	//////////////// Type =4  --->  Piramon  //////////////////
	if (itemType==4){   
		document.getElementsByName("Fees")[rowNo].value =  parseInt(txt2val(itemFee)) 
		//document.getElementsByName("AppQttys")[rowNo].focus();
		//document.getElementsByName("AppQttys")[rowNo].select();
	}

	setPrice(document.getElementsByName("Fees")[rowNo]);
}

//--------------------------------------------------------------------------------------------------------------------------------------------------

function mask(src){ 
	var theKey=event.keyCode;

	rowNo=src.parentNode.parentNode.rowIndex;
	invTable=document.getElementById("InvoiceLines");
	theRow=invTable.getElementsByTagName("tr")[rowNo];

	if (src.name=="Items"){
		if (theKey==13){
			event.keyCode=9
			dialogActive=true
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="‰«„ ¬Ì „Ì —« ﬂÂ „Ì ŒÊ«ÂÌœ Ã” ÃÊ ﬂ‰Ìœ Ê«—œ ﬂ‰Ìœ:"
			var myTinyWindow = window.showModalDialog('dialog_FindInvItem.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('dialog_invoiceItems.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0].value=Arguments[1];
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[1].value=Arguments[2];
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[2].value=Arguments[3];
					if (Arguments[4] == "œ«—œ") // VAT
						invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[3].value = 1;
					else 
						invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[3].value = 0;
				}
				//setFeeQtty(invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0])

				a=invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0];

				if (a){
					setFeeQtty(a)
					a.focus();
				}
			}
			dialogActive=false
		}
		else if (theKey >= 48 && theKey <= 57 ) { 
			//alert(theKey)
			//src.value=''
			return true;
		}
		else { 
			return false;
		}
	}
}

//--------------------------------------------------------------------------------------------------------------------------------------------------

function check(src){ 
	if (src.name=="Items"){
		rowNo=src.parentNode.parentNode.rowIndex;
		rowsCount=document.getElementsByName("Items").length;
		if (!dialogActive){
			if (src.value=='0'){
				if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« Õ–› ﬂ‰Ìœø")){
					delRow(rowNo);
					if (rowNo != rowsCount ){
						invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
					}else{
						invTable.getElementsByTagName("tr")[rowNo-1].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
					}
					return false;
				}
				else{
					src.focus();
				}
			}
			else {
				badCode = false;
				if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
				objHTTP.open('GET','xml2.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape( objHTTP.responseText)
				ar = tmpStr.split("#")

				if (ar[0]=="ﬂœ ﬂ«·« €·ÿ «” ")
				{
					//src.value="";
					//src.focus();
					alert("ﬂœ ﬂ«·« €·ÿ «” ");
					return false;
				}
				else{
					//document.all['A1'].innerText= objHTTP.status
					//document.all['A2'].innerText= objHTTP.statusText
					//document.all['A3'].innerText= objHTTP.responseText
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0].value = ar[0];
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[1].value = ar[1];
					invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[2].value = ar[2];
					// VAT
					if (ar[3] == "True")
						invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[3].value = 1;
					else 
						invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[1].getElementsByTagName("Input")[3].value = 0;

					setFeeQtty(invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[2].getElementsByTagName("Input")[0]);
				}

			}
		}
	}
}


function js_Link2Trace(num){
	return "<A HREF='../order/orderEdit.asp?e=n&radif="+ num + "' target='_balnk'>"+ num + "</A>"
}
function selectOrder(){
	theSpan=document.getElementById("orders");
	document.all.tmpDlgArg.value="";
	window.showModalDialog('Orders.asp?act=select&customer='+document.all.customerID.value,document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
	if (document.all.tmpDlgArg.value!="[Esc]"){
		theSpan.innerHTML="";
		Arguments=document.all.tmpDlgArg.value.split("#")
		tempWriteAnd=""
		for (i=1;i<=Arguments[0];i++){
			theSpan.innerHTML += "<input type='hidden' name='selectedOrders' value='"+Arguments[i]+"'>" + tempWriteAnd + js_Link2Trace(Arguments[i])
			tempWriteAnd=" Ê "
		}
	}
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
			theSpan.getElementsByTagName("span")[0].innerText=Arguments[1];
		}
	}
}
function submitOperations(){
	setCurrentRow(0);
	var okGo=true;
	for (rowNo=0; rowNo < document.getElementsByName("Items").length; rowNo++){
		if (document.getElementsByName('Items')[rowNo].value==''){
			delRow(rowNo);			
			rowNo=rowNo-1;
			okGo=false;
		}
	}
	if (okGo && document.getElementsByName('Items')[0]) {
		checkIsA();
		document.forms[0].submit();
	}
	else{
		alert(".ç‰œ ”ÿ— Œ«·Ì ÊÃÊœ œ«‘  ﬂÂ Õ–› ‘œ‰œ\n\n .·ÿ›« »——”Ì ﬂ‰Ìœ Ê „Ãœœ« œﬂ„Â –ŒÌ—Â —« »“‰Ìœ")
		currentRow=0;
		setCurrentRow(0);
		if (document.getElementsByName('Items')[0])
			document.getElementsByName('Items')[0].focus();
	}
}
//-->
</script>
