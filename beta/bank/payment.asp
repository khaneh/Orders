<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "������ ������"
SubmenuItem=5
if not Auth("A" , 5) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------- Submit Payment
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="������" then
	VouchID = request.form("VouchID")

	if VouchID<>"" then

		response.redirect "../bank/cheq.asp?act=enterCheque&VouchID="& VouchID
		
	end if
end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- Show Voucher Details
'-----------------------------------------------------------------------------------------------------
if request("VouchID")<> "" then

	response.write "<br><br>"
	call showAlert("������� ��� ���� ���.<br><br>��� ������ ����� �� ��� ���� �� ����� .<br><br>���� ���� �� �� �� �� ���� �� ��� ���� ����� ��� ��� �� �� ����� �� ���� ���� Ȑ����. <br><br>��� ����� �� ����.<br><br>���� �� <A HREF='../AP/AccountReport.asp?act=showVoucher&voucher="& request("VouchID") & "'><B>�����</B></A> �� �����." , CONST_MSG_ALERT)
	response.end
end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------ Main
'-----------------------------------------------------------------------------------------------------
set RSS=Conn.Execute ("SELECT Vouchers.*, APItems.RemainedAmount FROM Vouchers INNER JOIN APItems ON Vouchers.id = APItems.Link WHERE (Vouchers.paid = 0) AND (Vouchers.Verified = 1) AND (APItems.Type = 6) and (APItems.RemainedAmount>0)" )
%>
<BR><BR><BR>
<FORM METHOD=POST ACTION="payment.asp">
<TABLE dir=rtl align=center width=680  cellspacing=0>
<TR bgcolor="eeeeee" >
	<TD align=center colspan=8><B>��������� ������ ����</B></TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD style="border-bottom: solid 1pt black" width=10><!A HREF="default.asp?s=1"><SMALL>#</SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black"><INPUT TYPE="checkbox" NAME="" disabled></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black" width=120><!A HREF="default.asp?s=3"><SMALL>�������</SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><!A HREF="default.asp?s=1"><SMALL>����� ������</SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><!A HREF="default.asp?s=2"><SMALL>���� </SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><!A HREF="default.asp?s=2"><SMALL>����� </SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black" width=80><!A HREF="default.asp?s=2"><SMALL>����� ���� </SMALL></A></TD>
	<TD style="border-bottom: solid 1pt black; width="><!A HREF="default.asp?s=4"><SMALL>�����</SMALL></A></TD>
</TR>
<%
tmpCounter=0
Do while not RSS.eof
	tmpCounter = tmpCounter + 1
	if tmpCounter mod 2 = 1 then
		tmpColor="#FFFFFF"
		tmpColor2="#FFFFBB"
	Else
		tmpColor="#FFFFFF"
		tmpColor2="#FFFFBB"
		'tmpColor="#DDDDDD"
		'tmpColor2="#EEEEBB"
	End if 

set RST=Conn.Execute ("SELECT * FROM Accounts WHERE (ID = "& RSS("VendorID") & ")" )

set RSF=Conn.Execute ("SELECT VoucherLines.*, PurchaseOrders.Status FROM VoucherLines INNER JOIN PurchaseOrders ON VoucherLines.RelatedPurchaseOrderID = PurchaseOrders.ID WHERE (VoucherLines.Voucher_ID = "& RSS("ID") & ")" )

DisableFlag = ""
vouchLines = ""
Do while not RSF.eof
	vouchLines = vouchLines & "<A HREF='../purchase/outServiceTrace.asp?od=" & RSF("RelatedPurchaseOrderID") & "'>" & RSF("LineTitle") & "</a><br> (�����: " & RSF("qtty") & " - ����: " & RSF("price") & ") <br> "
	if RSF("status") <> "OK" then DisableFlag=" disabled"
RSF.moveNext
Loop
%>
<TR <%=DisableFlag%>  Title="<% 
			Comment = RSS("Comment")
			if Comment<>"-" then
				response.write "�����: " & Comment
			else
				response.write "����� �����"
			end if
		%>" bgcolor="<%=tmpColor%>">
	<TD  style="border-bottom: solid 1pt black"><A target="_blank" HREF="../AP/AccountReport.asp?act=showVoucher&voucher=<%=RSS("id")%>" ><%=RSS("id")%></A></TD>
	<TD style="border-bottom: solid 1pt black"><INPUT TYPE="checkbox" NAME="VouchID" VALUE="<%=RSS("id")%>"  onclick="setColor(this)"></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><%=RST("AccountTitle")%></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><%=RSS("title")%></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><%=RSS("totalPrice")%></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><%=RSS("RemainedAmount")%></TD>
	<TD style="border-bottom: solid 1pt black; border-left: solid 1pt black"><span dir=ltr><%=RSS("CreationDate")%></span><br> (<%=RSS("CreationTime")%>)</small></TD>
	<TD style="border-bottom: solid 1pt black"> &nbsp;<%=vouchLines%></TD>
</TR>
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
<center><INPUT TYPE="submit" Name="Submit" Value="������" class="inputBut" style="width:150px;" tabIndex="14">
</form>
</center>


<SCRIPT LANGUAGE="JavaScript">
<!--
tmpColor="#FFFFFF"
tmpColor2="#FFFFBB"

function setColor(obj)
{
	for(i=0; i<document.all.VouchID.length; i++)
		{
		theTR = document.all.VouchID[i].parentNode.parentNode
		if (document.all.VouchID[i].checked)
			theTR.setAttribute("bgColor",tmpColor2)
		else
			theTR.setAttribute("bgColor",tmpColor)
		}
	//theTR = obj.parentNode.parentNode
	//theTR.setAttribute("bgColor","#FFFFFF")
}
//-->
</SCRIPT>

<!--#include file="tah.asp" -->