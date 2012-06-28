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
	<td>
		<INPUT TYPE="submit" NAME="submit" class=inputBut value="ãÔÇåÏå">
	</td>
</TR>
<!--

<TR height=10>
	<TD colspan=4 align=center>ÊÇÑíÎ ÈÑ ÇÓÇÓ ÊÇÑíÎ İÇßÊæÑ ÈÇÔÏ</TD>
	<td><input type="checkbox" name="Effective" <%if effective then response.write " checked='checked' " %>></td>
</TR>

<TR>
	<TD align=left></TD>
	<TD align=right><INPUT TYPE="checkbox" NAME="vouchers" <%=vouchersSt%>> İÇßÊæÑåÇ</TD>
	<TD align=left width=20><INPUT TYPE="checkbox" NAME="payments" <%=paymentsSt%>></TD>
	<TD align=left>ÑÏÇÎÊåÇ</TD>
	<TD align=left></TD>
</TR>
-->
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

<TABLE dir=rtl align=center width=640 cellspacing=2 cellpadding=2 style="border:2 solid #330066;">	
	<%

	mySQL="SELECT Accounts.AccountTitle,Accounts.id as accountID, Vouchers.id, Vouchers.Title, Vouchers.TotalPrice, Vouchers.CreationTime, Vouchers.CreationDate, Vouchers.CreatedBy, Vouchers.verified, Vouchers.comment, Vouchers.ImageFileName, Vouchers.VendorID, Users.RealName, apItems.FullyApplied,apItems.RemainedAmount,EffectiveGLRows.GLDoc,Vouchers.effectiveDate,apItems.createdDate FROM Vouchers INNER JOIN Accounts ON Vouchers.VendorID = Accounts.ID INNER JOIN Users ON Vouchers.CreatedBy = Users.ID left outer join APItems on APItems.Type=6 and APItems.Link=Vouchers.id left outer join EffectiveGLRows on EffectiveGLRows.SYS='AP' and EffectiveGLRows.Link=apItems.ID WHERE (Vouchers.EffectiveDate >= N'"& fromDate & "' and Vouchers.EffectiveDate <= N'"& toDate & "') order by Vouchers.EffectiveDate"
			'mySQL="SELECT Accounts.AccountTitle, Vouchers.id, Vouchers.Title, Vouchers.TotalPrice, Vouchers.CreationTime, Vouchers.CreationDate, Vouchers.CreatedBy, Vouchers.verified, Vouchers.comment, Vouchers.ImageFileName, Vouchers.VendorID, Users.RealName, apItems.FullyApplied,apItems.RemainedAmount,EffectiveGLRows.GLDoc FROM Vouchers INNER JOIN Accounts ON Vouchers.VendorID = Accounts.ID INNER JOIN Users ON Vouchers.CreatedBy = Users.ID left outer join APItems on APItems.Type=6 and APItems.Link=Vouchers.id left outer join EffectiveGLRows on EffectiveGLRows.SYS='AP' and EffectiveGLRows.Link=apItems.ID WHERE (Vouchers.CreationDate >= N'"& fromDate & "' and Vouchers.CreationDate <= N'"& toDate & "')"
		set RSS=Conn.Execute (mySQL)	
		%>
	<TR bgcolor="eeeeee" >
		<TD colspan=6><H4>İÇßÊæÑåÇ</H4></TD>
	</TR>
	<TR bgcolor="eeeeee" style="cursor:hand;" title="ÊÑÊíÈ äãÇíÔ">
		<TD>ÊÇÑíÎ</TD>
		<TD>äÇã ÍÓÇÈ</TD>
		<TD>ÚäæÇä İÇßÊæÑ</TD>
		<TD>ãÈáÛ</TD>
		<TD>ãÇäÏå</TD>
		<TD>æÖÚíÊ</TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
	</TR>
		<%

	SumRemain=0
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
		if not IsNull(RSS("remainedAmount")) then
			SumRemain = SumRemain + CDbl(RSS("remainedAmount"))
		end if
		%>
	<TR bgcolor="<%=tmpColor%>" title="<% 
		Comment = RSS("Comment")
		if Comment<>"-" then
			response.write "ÊæÖíÍ: " & Comment
		else
			response.write "ÊæÖíÍ äÏÇÑÏ"
		end if
	%>">
		<td title="ÊÇÑíÎ ÇíÌÇÏ İÇßÊæÑ:<%=RSS("CreationDate")%>&#13; ÊÇÑíÎ ÊÇííÏ İÇßÊæÑ:<%=RSS("createdDate")%>"><%=RSS("effectiveDate")%></td>
		<TD><a href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rss("accountID")%>"><%=RSS("AccountTitle")%></A></TD>
		<TD><A HREF="AccountReport.asp?act=showVoucher&voucher=<%=RSS("ID")%>"><%=RSS("ID")%> &nbsp;-&nbsp; <%=RSS("Title")%></TD>
		<TD><%=Separate(RSS("TotalPrice"))%></A></TD>
		<td>
		<%
			if IsNull(RSS("remainedAmount")) then 
				response.write("ÊÇííÏ äÔÏå")
			else
				response.write RSS("remainedAmount")
			end if
		%>
		</td>
		<TD><% if RSS("verified") then %>ÊÇííÏ  ÔÏå <% else %>ÊÇííÏ  äÔÏå<% end if %>/
			<% if not IsNull(RSS("fullyApplied")) and RSS("fullyApplied") then %>ÑÏÇÎÊ  ÔÏå <% else %>ÑÏÇÎÊ äÔÏå<% end if %>
		</TD>
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
		<td colspan=3>ÌãÚ</td>
		<td><%=Separate(total)%></td>
		<td><%=Separate(SumRemain)%></td>
		<td></td>
	</tr>
		<%


	mySQL="SELECT Payments.id, Payments.CashAmount, Payments.ChequeAmount, APItems.RemainedAmount, Accounts.AccountTitle,Accounts.id as accountID, Payments.CreatedBy, Payments.CreationDate, Payments.CreationTime, Users.RealName,apItems.EffectiveDate FROM APItems INNER JOIN Payments ON APItems.Link = Payments.id INNER JOIN Accounts ON APItems.Account = Accounts.ID INNER JOIN Users ON Payments.CreatedBy = Users.ID WHERE (Payments.EffectiveDate >= N'"& fromDate & "' and Payments.EffectiveDate <= N'"& toDate & "' and Payments.SYS = 'AP') ORDER BY Payments.EffectiveDate"

			'mySQL="SELECT Payments.CashAmount, Payments.ChequeAmount, APItems.RemainedAmount, Accounts.AccountTitle, Payments.CreatedBy, Payments.CreationDate, Payments.CreationTime, Users.RealName FROM APItems INNER JOIN Payments ON APItems.Link = Payments.id INNER JOIN Accounts ON APItems.Account = Accounts.ID INNER JOIN Users ON Payments.CreatedBy = Users.ID WHERE (Payments.CreationDate >= N'"& fromDate & "' and Payments.CreationDate <= N'"& toDate & "' and Payments.SYS = 'AP') ORDER BY Payments.ID"

	set RSS=Conn.Execute (mySQL)	
	%>
	<TR bgcolor="eeeeee" >
		<TD colspan=6><H4>ÑÏÇÎÊåÇ</H4></TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD colspan=6 height=2 bgcolor=0></TD>
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
		<td title="ÊÇÑíÎ ÇíÌÇÏ:<%=RSS("CreationDate")%>"><%=RSS("effectiveDate")%></td>
		<TD><a href="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=rss("accountID")%>"><%=RSS("AccountTitle")%></A></TD>
		<TD title="ÊæÓØ:<%=RSS("RealName")%>">
			<a href="../AP/AccountReport.asp?act=showPayment&payment=<%=rss("id")%>">
				<%
				if cdbl(RSS("CashAmount"))>0 then response.write "äŞÏ: " & Separate(cdbl(RSS("CashAmount"))) & " "
				if cdbl(RSS("ChequeAmount"))>0 then response.write "ß: " & Separate(cdbl(RSS("ChequeAmount")))
				%>
			</a>
		</TD>
		<TD><%=Separate(cdbl(RSS("CashAmount")) + cdbl(RSS("ChequeAmount")))%> </TD>
		<TD><%=Separate(RSS("RemainedAmount"))%></A></TD>
		<TD></TD>
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
		<td colspan=3>ÌãÚ</td>
		<td><%=Separate(total)%></td>
		<td><%=Separate(totalRemain)%></td>
		<td></td>
	</tr>
	
</TABLE><br>
	<%

end if
%>

<!--#include file="tah.asp" -->