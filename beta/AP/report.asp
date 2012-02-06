<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AP (7)
PageTitle= " ÒÇÑÔ ÎÑíÏ"
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
if request("Effective")="on" then 
	effective=1
else
	effective=0
end if

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
	<TD colspan=5 align=center><H3>ÒÇÑÔ ÍÓÇÈÏÇÑí ÎÑíÏ</H3></TD>
</TR>
<TR>
	<TD align=left>ÇÒ ÊÇÑíÎ</TD>
	<TD align=right><INPUT TYPE="text" NAME="fromDate" value="<%=fromDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
	<TD align=left width=20></TD>
	<TD align=left>ÊÇ ÊÇÑíÎ</TD>
	<TD align=right><INPUT TYPE="text" NAME="toDate" value="<%=toDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
</TR>

<TR height=10>
	<TD colspan=4 align=center>ÊÇÑíÎ ÈÑ ÇÓÇÓ ÊÇÑíÎ İÇßÊæÑ ÈÇÔÏ</TD>
	<td><input type="checkbox" name="Effective" <%if effective then response.write " checked='checked' " %>></td>
</TR>

<TR>
	<TD align=left></TD>
	<TD align=right><INPUT TYPE="checkbox" NAME="vouchers" <%=vouchersSt%>> İÇßÊæÑåÇ</TD>
	<TD align=left width=20><INPUT TYPE="checkbox" NAME="payments" <%=paymentsSt%>></TD>
	<TD align=left>ÑÏÇÎÊåÇ</TD>
	<TD align=left><INPUT TYPE="submit" NAME="submit" class=inputBut value="ãÔÇåÏå"></TD>
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

if request("submit")="ãÔÇåÏå"then

	%>
	<TABLE dir=rtl align=center width=600>
	<%
	if vouchers = "on" then
		if effective then 
			mySQL="SELECT Accounts.AccountTitle, Vouchers.id, Vouchers.Title, Vouchers.TotalPrice, Vouchers.CreationTime, Vouchers.CreationDate, Vouchers.CreatedBy, Vouchers.verified, Vouchers.paid, Vouchers.comment, Vouchers.ImageFileName, Vouchers.VendorID, Users.RealName FROM Vouchers INNER JOIN Accounts ON Vouchers.VendorID = Accounts.ID INNER JOIN Users ON Vouchers.CreatedBy = Users.ID WHERE (Vouchers.EffectiveDate >= N'"& fromDate & "' and Vouchers.EffectiveDate <= N'"& toDate & "')"
		else
			mySQL="SELECT Accounts.AccountTitle, Vouchers.id, Vouchers.Title, Vouchers.TotalPrice, Vouchers.CreationTime, Vouchers.CreationDate, Vouchers.CreatedBy, Vouchers.verified, Vouchers.paid, Vouchers.comment, Vouchers.ImageFileName, Vouchers.VendorID, Users.RealName FROM Vouchers INNER JOIN Accounts ON Vouchers.VendorID = Accounts.ID INNER JOIN Users ON Vouchers.CreatedBy = Users.ID WHERE (Vouchers.CreationDate >= N'"& fromDate & "' and Vouchers.CreationDate <= N'"& toDate & "')"
		end if
		set RSS=Conn.Execute (mySQL)	
		%>
		<TR bgcolor="eeeeee" >
			<TD colspan=5><H4>İÇßÊæÑåÇ</H4></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD><!A HREF="default.asp?s=1"><SMALL>ÚäæÇä İÇßÊæÑ </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=2"><SMALL>ãÈáÛ İÇßÊæÑ </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=3"><SMALL>İÑæÔäÏå</SMALL></A></TD>
			<TD><!A HREF="default.asp?s=4"><SMALL>æÖÚíÊ </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=5"><SMALL>ÊÇÑíÎ ËÈÊ </SMALL></A></TD>
		</TR>
		<%
		tmpCounter=0
		total=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
			total = total + CDbl(RSS("TotalPrice"))

		%>
		<TR bgcolor="<%=tmpColor%>" title="<% 
			Comment = RSS("Comment")
			if Comment<>"-" then
				response.write "ÊæÖíÍ: " & Comment
			else
				response.write "ÊæÖíÍ äÏÇÑÏ"
			end if
		%>">
			<TD><A HREF="AccountReport.asp?act=showVoucher&voucher=<%=RSS("ID")%>"><%=RSS("ID")%> &nbsp;-&nbsp; <%=RSS("Title")%></TD>
			<TD><%=Separate(RSS("TotalPrice"))%></A></TD>
			<TD><%=RSS("AccountTitle")%></A></TD>
			<TD><% if RSS("verified") then %>ÊÇííÏ  ÔÏå <% else %>ÊÇííÏ  äÔÏå<% end if %>/
				<% if RSS("paid") then %>ÑÏÇÎÊ  ÔÏå <% else %>ÑÏÇÎÊ äÔÏå<% end if %>
			</TD>
			<TD><span dir=ltr><%=RSS("CreationDate")%></span><!--&nbsp;(ÓÇÚÊ <%=RSS("CreationTime")%>)--></TD>
		</TR>
			  
		<% 
		RSS.moveNext
		Loop
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 
		%>
		<tr bgcolor="<%=tmpColor%>">
			<td>ÌãÚ</td>
			<td colspan="4"><%=Separate(total)%></td>
		</tr>
		<%
	end if

	if payments = "on" then
		if effective then 
			mySQL="SELECT Payments.CashAmount, Payments.ChequeAmount, APItems.RemainedAmount, Accounts.AccountTitle, Payments.CreatedBy, Payments.CreationDate, Payments.CreationTime, Users.RealName FROM APItems INNER JOIN Payments ON APItems.Link = Payments.id INNER JOIN Accounts ON APItems.Account = Accounts.ID INNER JOIN Users ON Payments.CreatedBy = Users.ID WHERE (Payments.EffectiveDate >= N'"& fromDate & "' and Payments.EffectiveDate <= N'"& toDate & "' and Payments.SYS = 'AP') ORDER BY Payments.ID"
		else
			mySQL="SELECT Payments.CashAmount, Payments.ChequeAmount, APItems.RemainedAmount, Accounts.AccountTitle, Payments.CreatedBy, Payments.CreationDate, Payments.CreationTime, Users.RealName FROM APItems INNER JOIN Payments ON APItems.Link = Payments.id INNER JOIN Accounts ON APItems.Account = Accounts.ID INNER JOIN Users ON Payments.CreatedBy = Users.ID WHERE (Payments.CreationDate >= N'"& fromDate & "' and Payments.CreationDate <= N'"& toDate & "' and Payments.SYS = 'AP') ORDER BY Payments.ID"
		end if
		set RSS=Conn.Execute (mySQL)	
		%>
		<TR bgcolor="eeeeee" >
			<TD colspan=5><H4>ÑÏÇÎÊåÇ</H4></TD>
		</TR>
		<TR bgcolor="eeeeee" >
			<TD><!A HREF="default.asp?s=1"><SMALL>ãÈáÛ ÑÏÇÎÊ </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=2"><SMALL>ÈÇŞíãÇäÏå ÑÏÇÎÊ </SMALL></A></TD>
			<TD><!A HREF="default.asp?s=3"><SMALL>ÏÑíÇİÊ ßääÏå</SMALL></A></TD>
			<TD><!A HREF="default.asp?s=4"><SMALL>ÑÏÇÎÊ ßääÏå</SMALL></A></TD>
			<TD><!A HREF="default.asp?s=5"><SMALL>ÊÇÑíÎ ÑÏÇÎÊ</SMALL></A></TD>
		</TR>
		<%
		tmpCounter=0
		total=0
		totalRemain=0
		Do while not RSS.eof
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
			total= total + cdbl(RSS("CashAmount")) + cdbl(RSS("ChequeAmount"))
			totalRemain = totalRemain + CDbl(RSS("RemainedAmount"))
		%>
		<TR bgcolor="<%=tmpColor%>" >
			<TD><%=Separate(cdbl(RSS("CashAmount")) + cdbl(RSS("ChequeAmount")))%> </TD>
			<TD><%=Separate(RSS("RemainedAmount"))%></A></TD>
			<TD><%=RSS("AccountTitle")%></A></TD>
			<TD><%=RSS("RealName")%></TD>
			<TD><span dir=ltr><%=RSS("CreationDate")%></span><!--&nbsp;(ÓÇÚÊ <%=RSS("CreationTime")%>)--></TD>
		</TR>
			  
		<% 
		RSS.moveNext
		Loop
		tmpCounter = tmpCounter + 1
		if tmpCounter mod 2 = 1 then
			tmpColor="#FFFFFF"
			tmpColor2="#FFFFBB"
		Else
			tmpColor="#DDDDDD"
			tmpColor2="#EEEEBB"
		End if 
		%>
		<tr bgcolor="<%=tmpColor%>">
			<td><%=Separate(total)%></td>
			<td><%=Separate(totalRemain)%></td>
			<td colspan="3">ÌãÚ</td>
		</tr>
		<%
	end if
	%>
	</TABLE><br>
	<%

end if
%>

<!--#include file="tah.asp" -->