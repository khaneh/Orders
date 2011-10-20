<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
	response.buffer=true
	PageTitle="Å—œ«Œ  «Ì‰ —‰ Ì"
	SubmenuItem=3
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
	InvoiceID = request("Invoice")
	if not(isnumeric(InvoiceID)) then
		response.write "<br>" 
		call showAlert ("Œÿ« œ— ‘„«—Â ›«ﬂ Ê—",CONST_MSG_ERROR) 
		response.end
	end if
	InvoiceID=clng(InvoiceID)
	mySQL="SELECT * FROM Invoices WHERE (Invoices.ID ='"& InvoiceID & "')"

	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		response.write "<br>" 
		call showAlert ("ÅÌœ« ‰‘œ",CONST_MSG_ERROR) 
		response.end
	end if
	customerID=		RS1("Customer")
	totalPrice=		cdbl(RS1("totalPrice"))
	totalDiscount=	cdbl(RS1("totalDiscount"))
	totalReverse=	cdbl(RS1("totalReverse"))
	totalVat =		cdbl(RS1("totalVat"))
'	creationDate=	RS1("CreatedDate")
'	ApproveDate=	RS1("ApprovedDate")
'	issueDate=		RS1("IssuedDate")
'	VoidDate=		RS1("VoidedDate")
'	InvoiceNo=		RS1("Number") 
'	Voided=			RS1("Voided")
'	Issued=			RS1("Issued")
'	Approved=		RS1("Approved")
'	IsReverse=		RS1("IsReverse")
'	if RS1("IsA") = TRUE then IsA=1 else IsA=0
'	Creator =		RS1("Creator")
'	Approver =		RS1("Approver")
'	Issuer =		RS1("Issuer")
'	Voider =		RS1("Voider")
'	payable=request("Payable")
'	custID=request("CustID") 

	RS1.close
	TotalReceivable= totalPrice - totalDiscount - totalReverse + totalVat
	mySQL="SELECT * FROM Accounts WHERE (ID='"& customerID & "')"
	Set RS1 = conn.Execute(mySQL)
	if not RS1.EOF then
		customerName=RS1("AccountTitle")
		email1=RS1("EMail1")
	end if
	RS1.close
%>


<div style='direction:rtl;'>
Â„ò«— ê—«„Ì!
<br/>
·ÿ›« „ ‰ “Ì— —« òÅÌ ò‰Ìœ Ê »—«Ì „‘ —Ì «Ì„Ì· ò‰Ìœ°
<br/>
œ— ’Ê— Ì òÂ «Ì„Ì· „‘ —Ì «‘ »«Â À»  ‘œÂ ·ÿ›« ¬‰—«  ’ÕÌÕ ‰„«ÌÌœ.
<br/>
</div>
<%

%>
<br/>
<table>
	<tr>
		<td>‰«„ „‘ —Ì:</td>
		<td><a href='/beta/CRM/AccountInfo.asp?act=show&selectedCustomer=<%=customerID%>'><%=customerName%></a></td>
	</tr>
	<tr>
		<td>¬œ—” «Ì„Ì·:</td>
		<td><%=email1%></td>
	</tr>
	<tr>
		<td>„»·€ ﬁ«»· Å—œ«Œ :</td>
		<td align='left'><%=separate(TotalReceivable)%></td>
	</tr>
</table>
<div id='copytext' style='background-color:white;border:1px solid red;'>
<br/>
»« ”·«„
<br/>
<%=customerName%>° 
<br/>
»—«Ì Å—œ«Œ  «·ò —Ê‰ÌòÌ ›«ò Ê— ŒÊœ° »Â ‘„«—Â 
<U><%=InvoiceID%> </U>
Ê »Â „»·€ 
<U><%=separate(TotalReceivable)%></U>
·ÿ›« »— —ÊÌ ·Ì‰ò “Ì— ò·Ìò ò‰Ìœ.
<br/>
<div style='direction:ltr;'>
	<a href='http://my.pdhco.com/payment/?InvoiceID=<%=mycode(InvoiceID,"93")%>&TotalRecivable=<%=mycode(TotalReceivable,"49")%>'>http://my.pdhco.com/payment</a>
</div>
<br/>
÷„‰« ‘‰«”Â ›«ò Ê— ‘„«
<B><U><%=mycode(InvoiceID,"93")%></U></B>
„Ì »«‘œ Ê ‘‰«”Â Å—œ«Œ  ‘„«
<B><U><%=mycode(TotalReceivable,"49")%></U></B>
„Ì »«‘œ. 
<br/>
»—«Ì «ÿ„Ì‰«‰ »Ì‘ — «“ œ—”  Ê«—œ ‘œ‰ «Ì‰ «⁄œ«œ° Å” «“ À»  ‘‰«”Â Â« ‘„«—Â ›«ò Ê— Ê „»·€ Å—œ«Œ Ì —« œ— ”«Ì  „« çò ò‰Ìœ.  
<br/>
»«  ‘ò— «“ Õ”‰ «‰ Œ«» ‘„«
<br/>
Œ«‰Â ç«Å Ê ÿ—Õ
<br/><br/>
</div>

<TEXTAREA ID="holdtext" STYLE="display:none;">
</TEXTAREA>

<div align='center'>
	<input onclick='ClipBoard();' class='GenButton' width='50' type='button' name='copy' style='border:1px solid black;' value='òÅÌ'>
</div>

<%
function mycode(n,prefix)
	dim str,c,w,t
	'n = 3 * n
	str=cstr(n)
	for i=1 to 10-len(cstr(n))
		str="0"+str
	next
	'str="99"+str
	'Prefix must be 2 digits
	str =  prefix + str
	for i=1 to 12 
		if i mod 2 = 0 then w=3 else w=1
		c = c + cdbl(mid(str,i,1)) * w
		c = (10 - (c mod 10)) mod 10
	next
	str=str+cstr(c)
'str="1234567890123"
	for i=1 to 13
		if i mod 2 = 1 then t=t+mid(str,i,1)
	next
	for i=1 to 13
		if i mod 2 = 0 then t=t+mid(str,i,1)
	next
'	response.write(t)
'	response.write("<BR/>")
	mycode=t
end function

function mydecode(n)
	dim str,c,w,num
	num=cstr(n)
	str=""
	for i=1 to 6
		str=str+mid(num,i,1)+mid(num,7+i,1)
'		response.write(str)
'		response.write("<br/>")
	next
	if len(num) mod 2 = 1 then str=str+mid(num,7,1)
'response.write(str)
	num=str
	str=mid(num,1,12)
	for i=1 to 12 
		if i mod 2 = 0 then w=3 else w=1
		c = c + cdbl(mid(str,i,1)) * w
		c = (10 - (c mod 10)) mod 10
	next
	if c=cdbl(mid(num,13,1)) mod 10 and mid(str,1,2)="99" and clng(mid(str,3,10)) mod 5 = 0 then 
		mydecode=cdbl(mid(str,3,10)) / 5
	else
		mydecode="Ÿ«Â—« Œÿ«ÌÌ —Œ œ«œÂ"
	end if
end function
%>

<SCRIPT LANGUAGE="JavaScript">

function ClipBoard()
{
holdtext.innerText = copytext.innerText;
Copied = holdtext.createTextRange();
Copied.execCommand("Copy");
}
</SCRIPT>
<!--#include file="tah.asp" -->