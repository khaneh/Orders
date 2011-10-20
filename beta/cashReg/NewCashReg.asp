<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CashRegister (9)
PageTitle="œ—Ì«› "
SubmenuItem=5
if not Auth(9 , 5) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/orderEdit.asp?e=n&radif="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

%>
<style>
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border: 2px solid #000088; padding:0; direction: RTL; width:500px;}
	.RcpMainTable Tr {Height:25px; border: 1px solid black;}
	.RcpMainTable Input { font-family:tahoma; font-size: 9pt; width:120px; border: 1px solid black; text-align:right; direction: LTR;}
	.RcpMainTable Select { font-family:tahoma; font-size: 9pt; width:120px;}
</style>
<font face="tahoma">
<%
if request("act")="submitNew" then
	if request("Cashier")="" OR request("Banker")="" then 
		response.clear
		conn.close
		response.redirect "./NewCashReg.asp?errmsg=" & server.urlencode("’‰œÊﬁœ«— Ì« „Õ· —« «‰ Œ«» ‰ﬂ—œÂ «Ìœ...")
	end if
	mySQL="SELECT * FROM CashRegisters WHERE (IsOpen=1) AND (Cashier='"& request("Cashier") & "')"
	Set RS1= conn.Execute(mySQL)
	if not RS1.eof then 
		response.clear
		conn.close
		response.redirect "./NewCashReg.asp?errmsg=" & server.urlencode("«Ì‰ ¬ﬁ«Ì ’‰œÊﬁœ«— ‘„« ﬁ»·« ÌÂ ’‰œÊﬁ »«“ œ«—‰œ!<br>«Ê· «“ Â„Â  ﬂ·Ì› «Ê‰ —Ê —Ê‘‰ ﬂ‰Ìœ.")
	end if
	OpeningAmount=text2value(request("OpeningAmount"))
	mySQL="INSERT INTO CashRegisters (Cashier, Banker, OpenDate, NameDate, OpenedBy, IsOpen, IsApproved, OpeningAmount, CashAmount, ChequeAmount, ChequeQtty, ShortOverAmount) VALUES ('" &_ 
	request("Cashier") & "', '"& request("Banker") & "', N'"& shamsiToday() & "', N'"& request("NameDate") & "', '"& session("ID") & "', '"& 1 & "', '"& 0 & "', '"& OpeningAmount & "', '"& OpeningAmount & "', '"& 0 & "', '"& 0 & "', '"& 0 & "')"
	conn.Execute(mySQL)
%><br><br>
	<TABLE width='70%' align='center'>
	<TR>
		<TD align=center bgcolor=yellow style='border: solid 1pt black'><BR><b><%="»Â ”·«„ Ì ÌÂ ’‰œÊﬁ ÃœÌœ «ÌÃ«œ ‘œ  »—Ìﬂ ⁄—÷ „Ì ﬂ‰„."%></b><BR><BR></TD>
	</TR>
	</TABLE>
<%
else%>
<BR><BR>
	<FORM METHOD=POST ACTION="?act=submitNew">
	<TABLE class='RcpMainTable' align='center'>
	<TR>
		<TD align='left'> «—ÌŒ ’‰œÊﬁ:</TD>
		<TD><INPUT TYPE="text" style="text-align:left;" Name="NameDate" value="<%=shamsiToday()%>" onblur="acceptDate(this)"></TD>
		<TD align='left'>&nbsp;</TD>
		<TD>&nbsp;</TD>
	</TR>
	<TR>
		<TD align='left'>’‰œÊﬁœ«—:</TD>
		<TD><SELECT NAME="Cashier">
			<option Value=""> -- «‰ Œ«» ﬂ‰Ìœ -- </option>
<%			mySQL="SELECT * FROM Users WHERE Display=1 ORDER BY RealName"
			Set RS1 = conn.Execute(mySQL)
			do while not (RS1.eof)
%>				<option Value="<%=RS1("ID")%>"><%=RS1("RealName")%></option>
<%
				RS1.movenext
			loop
			RS1.close
%>			</SELECT>
		</TD>
		<TD align='left'>„Õ·:</TD>
		<TD><SELECT NAME="Banker">
			<option Value=""> -- «‰ Œ«» ﬂ‰Ìœ -- </option>
<%			mySQL="SELECT * FROM Bankers WHERE (IsBankAccount <> 1)"
			Set RS1 = conn.Execute(mySQL)
			do while not (RS1.eof)
%>				<option Value="<%=RS1("ID")%>"><%=RS1("Name")%></option>
<%
				RS1.movenext
			loop
			RS1.close
%>			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD align='left'>„ÊÃÊœÌ «Ê·ÌÂ:</TD>
		<TD><INPUT TYPE="text" Name="OpeningAmount" onblur="setPrice(this);" onKeyPress="return maskNumber(this);"></TD>
		<TD align='left'>&nbsp;</TD>
		<TD>&nbsp;</TD>
	</TR>
	<TR>
		<TD colspan="4" align='center'><INPUT style="width:70px;text-align:center;" TYPE="submit" value="«ÌÃ«œ"></TD>
	</TR>
	</TABLE>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
document.all.OpeningAmount.focus();

function setPrice(src){
		src.value=val2txt(txt2val(src.value));
}
	
	//-->
	</SCRIPT>
<%
end if
conn.Close
%>
</font>
<!--#include file="tah.asp" --> 