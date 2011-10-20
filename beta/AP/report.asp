<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= " ê“«—‘ Œ—Ìœ"
SubmenuItem=3
if not Auth(7 , 3) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
fromDate = request("fromDate") 
toDate = request("toDate") 
vouchers = request("vouchers") 
payments = request("payments") 


if payments="" then
	paymentsSt = " "
else
	paymentsSt = "checked"
end if

if vouchers="" then
	vouchersSt = " "
else
	vouchersSt = "checked"
end if

if toDate=""  or fromDate="" then
	toDate = shamsiToday()
	fromDate = shamsiDate(Date()-7)
	paymentsSt = "checked"
	vouchersSt = "checked"
	flag = "first"
end if


'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------- Log FORM
'-----------------------------------------------------------------------------------------------------
%>
<BR><BR>
<FORM METHOD=POST ACTION="report.asp">
<TABLE border=0 align=center>
<TR>
	<TD colspan=5 align=center><H3>ê“«—‘ Õ”«»œ«—Ì Œ—Ìœ</H3></TD>
</TR>
<TR>
	<TD align=left>«“  «—ÌŒ</TD>
	<TD align=right><INPUT TYPE="text" NAME="fromDate" value="<%=fromDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
	<TD align=left width=20></TD>
	<TD align=left> «  «—ÌŒ</TD>
	<TD align=right><INPUT TYPE="text" NAME="toDate" value="<%=toDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
</TR>

<TR height=10>
	<TD colspan=5 align=center></TD>
</TR>

<TR>
	<TD align=left></TD>
	<TD align=right><INPUT TYPE="checkbox" NAME="vouchers" <%=vouchersSt%>> ›«ﬂ Ê—Â«</TD>
	<TD align=left width=20><INPUT TYPE="checkbox" NAME="payments" <%=paymentsSt%>></TD>
	<TD align=left>Å—œ«Œ Â«</TD>
	<TD align=left><INPUT TYPE="submit" NAME="submit" class=inputBut value="„‘«ÂœÂ"></TD>
</TR>

</TABLE>
</FORM>
<%
if flag = "first" then
	response.end
end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------- Log Rows
'-----------------------------------------------------------------------------------------------------

if request("submit")="„‘«ÂœÂ"then

	%>
	<TABLE dir=rtl align=center width=600>
	<%
	if vouchers = "on" then
		set RSS=Conn.Execute ("SELECT Accounts.AccountTitle, Vouchers.id, Vouchers.Title, Vouchers.TotalPrice, Vouchers.CreationTime, Vouchers.CreationDate, Vouchers.CreatedBy, Vouchers.verified, Vouchers.paid, Vouchers.comment, Vouchers.ImageFileName, Vouchers.VendorID, Users.RealName FROM Vouchers INNER JOIN Accounts ON Vouchers.VendorID = Accounts.ID INNER JOIN Users ON Vouchers.CreatedBy = Users.ID WHERE (Vouchers.CreationDate >= N'"& fromDate & "' and Vouchers.CreationDate <= N'"& toDate & "')")	
		%>
		<TR bgcolor="eeeeee" >
			<TD colspan=5><H4>›«ﬂ Ê—Â«</H4></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD><!A HREF="default.asp?s=1"><SMALL>⁄‰Ê«‰ ›«ﬂ Ê— </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=2"><SMALL>„»·€ ›«ﬂ Ê— </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=3"><SMALL>›—Ê‘‰œÂ</SMALL></A></TD>
			<TD><!A HREF="default.asp?s=4"><SMALL>Ê÷⁄Ì  </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=5"><SMALL> «—ÌŒ À»  </SMALL></A></TD>
		</TR>
		<%
		tmpCounter=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

		%>
		<TR bgcolor="<%=tmpColor%>" title="<% 
			Comment = RSS("Comment")
			if Comment<>"-" then
				response.write " Ê÷ÌÕ: " & Comment
			else
				response.write " Ê÷ÌÕ ‰œ«—œ"
			end if
		%>">
			<TD><A HREF="AccountReport.asp?act=showVoucher&voucher=<%=RSS("ID")%>"><%=RSS("ID")%> &nbsp;-&nbsp; <%=RSS("Title")%></TD>
			<TD><%=RSS("TotalPrice")%></A></TD>
			<TD><%=RSS("AccountTitle")%></A></TD>
			<TD><% if RSS("verified") then %> «ÌÌœ  ‘œÂ <% else %> «ÌÌœ  ‰‘œÂ<% end if %>/
				<% if RSS("paid") then %>Å—œ«Œ   ‘œÂ <% else %>Å—œ«Œ  ‰‘œÂ<% end if %>
			</TD>
			<TD><span dir=ltr><%=RSS("CreationDate")%></span><!--&nbsp;(”«⁄  <%=RSS("CreationTime")%>)--></TD>
		</TR>
			  
		<% 
		RSS.moveNext
		Loop
	end if

	if payments = "on" then
		set RSS=Conn.Execute ("SELECT Payments.CashAmount, Payments.ChequeAmount, APItems.RemainedAmount, Accounts.AccountTitle, Payments.CreatedBy, Payments.CreationDate, Payments.CreationTime, Users.RealName FROM APItems INNER JOIN Payments ON APItems.Link = Payments.id INNER JOIN Accounts ON APItems.Account = Accounts.ID INNER JOIN Users ON Payments.CreatedBy = Users.ID WHERE (Payments.CreationDate >= N'"& fromDate & "' and Payments.CreationDate <= N'"& toDate & "' and Payments.SYS = 'AP') ORDER BY Payments.ID")	
		%>
		<TR bgcolor="eeeeee" >
			<TD colspan=5><H4>Å—œ«Œ Â«</H4></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD><!A HREF="default.asp?s=1"><SMALL>„»·€ Å—œ«Œ  </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=2"><SMALL>»«ﬁÌ„«‰œÂ Å—œ«Œ  </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=3"><SMALL>œ—Ì«›  ﬂ‰‰œÂ</SMALL></A></TD>
			<TD><!A HREF="default.asp?s=4"><SMALL>Å—œ«Œ  ﬂ‰‰œÂ</SMALL></A></TD>
			<TD><!A HREF="default.asp?s=5"><SMALL> «—ÌŒ Å—œ«Œ </SMALL></A></TD>
		</TR>
		<%
		tmpCounter=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 

		%>
		<TR bgcolor="<%=tmpColor%>" >
			<TD><%=cdbl(RSS("CashAmount")) + cdbl(RSS("ChequeAmount"))%> </TD>
			<TD><%=RSS("RemainedAmount")%></A></TD>
			<TD><%=RSS("AccountTitle")%></A></TD>
			<TD><%=RSS("RealName")%></TD>
			<TD><span dir=ltr><%=RSS("CreationDate")%></span><!--&nbsp;(”«⁄  <%=RSS("CreationTime")%>)--></TD>
		</TR>
			  
		<% 
		RSS.moveNext
		Loop
	end if
	%>
	</TABLE><br>
	<%

end if
%>

<!--#include file="tah.asp" -->