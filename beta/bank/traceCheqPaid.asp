<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "ÅÌêÌ—Ì çﬂÂ«Ì œ—Ì«› Ì"
SubmenuItem=3
if not Auth("A" , 3) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<BR><BR>
<style>
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border: 2px solid #000088; padding:0; direction: RTL; width:600px;}
	.RcpMainTable Tr {Height:25px; border: 1px solid black;}
	.RcpMainTableInput { font-family:tahoma; font-size: 9pt; width:100px; border: 1px solid gray; text-align:center; direction: LTR;}
	.RcpMainTable Select { font-family:tahoma; font-size: 9pt; width:120px;}
</style>

<CENTER>
<%

function DecsAsc (x)
  if cint(s) = cint(x) then
	if DESC=1 then
		result = "v"
	else
		result = "^"
	end if
  end if 
  DecsAsc = result
end function

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------ Submit Pay
'-----------------------------------------------------------------------------------------------------
if request.form("submit") = "Å«” ‘œ" then
	x = request.form("pay")
	
	set RSV=Conn.Execute ("SELECT PaidCheques.Amount, PaidCheques.status, PaidCheques.Banker, Bankers.Name FROM PaidCheques INNER JOIN Bankers ON PaidCheques.Banker = Bankers.ID WHERE (PaidCheques.ID = "& x & ")" ) 

	if RSV.EOF then
		response.write "<BR><BR>Œÿ«Ì ›«Ã⁄Â ¬„Ì“! <BR>»« ⁄“Ì“«‰ ŒÊœ Êœ«⁄ ﬂ‰Ìœ"
		response.write "<BR><BR>”Å” »« „œÌ— ”Ì” „  „«” »êÌ—Ìœ"
	response.write "<BR><BR><CENTER><A HREF='traceCheqPaid.asp?fromSession=y'>»—ê‘ </A></CENTER>"
		response.end
	elseif RSV("status")="1" then 
		response.write "<BR><BR>Œÿ«Ì ›«Ã⁄Â ¬„Ì“! <BR>»« ⁄“Ì“«‰ ŒÊœ Êœ«⁄ ﬂ‰Ìœ"
		response.write "<BR><BR>”Å” »« „œÌ— ”Ì” „  „«” »êÌ—Ìœ"
	response.write "<BR><BR><CENTER><A HREF='traceCheqPaid.asp?fromSession=y'>»—ê‘ </A></CENTER>"
		response.end
	end if

	Conn.Execute ("update PaidCheques set Status=1, StatusSetBy="& session("ID")& ", StatusSetDate=N'"& shamsiToday()& "' where ID="& x ) 

	Conn.Execute ("update Bankers set CurrentBalance=CurrentBalance-"& RSV("amount") & " where ID="& RSV("Banker") ) 
	
	response.write "<BR><BR>çﬂ ‘„«—Â " & x & "»Â „»·€ "& RSV("Amount") & "—Ì«·   «“ „Õ·  "& RSV("Name") & " Å«” ‘œ"
	response.write "<BR><BR><CENTER><A HREF='traceCheqPaid.asp?fromSession=y'>»—ê‘ </A></CENTER>"
	response.end
end if 
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Submit Search
'-----------------------------------------------------------------------------------------------------
if request.form("submit") = "Ã” ÃÊ" then
	Banks = request.form("Banks")
	ChequeDatesFrom = request.form("ChequeDatesFrom")
	ChequeDatesTo = request.form("ChequeDatesTo")
	status = request.form("status")
	response.cookies("OldURL") = Request.ServerVariables("HTTP_REFERER")
	'response.form = "a"
	'response.write Request.ServerVariables("HTTP_REFERER")


	myQuery ="SELECT dbo.Payments.Account AS Account, dbo.PaidCheques.*, dbo.Bankers.Name AS BankName, dbo.Bankers.ID AS BankID, dbo.Users.RealName AS CSR, dbo.RcvPayChqStatus.Name AS statusName FROM dbo.PaidCheques INNER JOIN dbo.Bankers ON dbo.PaidCheques.Banker = dbo.Bankers.ID INNER JOIN dbo.Users ON dbo.PaidCheques.CreatedBy = dbo.Users.ID INNER JOIN dbo.Payments ON dbo.PaidCheques.Payment = dbo.Payments.ID INNER JOIN dbo.RcvPayChqStatus ON dbo.PaidCheques.Status = dbo.RcvPayChqStatus.ID WHERE (1 = 1)"


	'myQuery = "SELECT Payments.Account, PaidCheques.*, Bankers.Name AS BankName, Bankers.ID AS BankID, Users.RealName AS CSR FROM PaidCheques INNER JOIN Bankers ON PaidCheques.Banker = Bankers.ID INNER JOIN Users ON PaidCheques.CreatedBy = Users.ID INNER JOIN Payments ON PaidCheques.Payment = Payments.ID WHERE 1=1 "
	
	if not ( ChequeDatesTo = "" ) then 
		myQuery  = myQuery  &  " AND (PaidCheques.ChequeDate < N'"& ChequeDatesTo & "') "
	end if 
	
	if not ( ChequeDatesFrom = "" ) then 
		myQuery  = myQuery  &  " AND (PaidCheques.ChequeDate >= N'"& ChequeDatesFrom & "')"
	end if 
	
	if not Banks = "-2" then 
		myQuery  = myQuery  & " AND (PaidCheques.Banker = "& Banks & ")"
	end if 

	if not (status = "-1" or status = "") then 
		myQuery  = myQuery  & " AND (PaidCheques.Status = "& status & ")"
	end if 
	'response.write myQuery
end if 

if myQuery = "" then 
	myQuery = "SELECT Payments.Account, PaidCheques.*, Bankers.Name AS BankName, Bankers.ID AS BankID, Users.RealName AS CSR FROM PaidCheques INNER JOIN Bankers ON PaidCheques.Banker = Bankers.ID INNER JOIN Users ON PaidCheques.CreatedBy = Users.ID INNER JOIN Payments ON PaidCheques.Payment = Payments.ID  WHERE (PaidCheques.Status = 0) "
	Banks = -2
	ChequeDatesFrom = ""
	ChequeDatesTo = ""
	status = "0"
end if

if request("fromSession") = "y" then
	myQuery = session("myQuery")
	Banks = session("Banks")
	ChequeDatesFrom = session("ChequeDatesFrom") 
	ChequeDatesTo = session("ChequeDatesTo") 
	status = session("status") 
end if

s = request("s")
if s="" then s="1"

if s="1" then
	orderBy = "PaidCheques.ID"
elseif s="2" then
	orderBy = "PaidCheques.ChequeNo"
elseif s="3" then
	orderBy = "PaidCheques.ChequeDate"
elseif s="4" then
	orderBy = "PaidCheques.amount"
elseif s="5" then
	orderBy = "PaidCheques.CreatedBy"
elseif s="6" then
	orderBy = "PaidCheques.status"
elseif s="7" then
	orderBy = "PaidCheques.Banker"
elseif s="8" then
	orderBy = "Payments.Account"
end if

Desc = request("Desc")
if Desc = "" then 
	Desc = 1
else
	Desc = 3 - Desc
end if

myQuery2  = myQuery  & " order by " & orderBy
if Desc=2 then 
	myQuery2  = myQuery2  & " DESC " 
end if

session("myQuery") = myQuery 
session("Banks") = Banks
session("ChequeDatesFrom") = ChequeDatesFrom
session("ChequeDatesTo") = ChequeDatesTo
session("status") = status


'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Search Form
'-----------------------------------------------------------------------------------------------------
%>
<FORM METHOD=POST ACTION="traceCheqPaid.asp?">
<TABLE align=center class="RcpMainTable">
<TR>
	<TD align=right>  «“  «—ÌŒ : <INPUT  class="RcpMainTableInput" dir="LTR"  TYPE="text" NAME="ChequeDatesFrom" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=ChequeDatesFrom%>"></TD>
	<TD>„—»Êÿ »Â »«‰ﬂ: 
		<select name="Banks" >
		<option value="-2">«‰ Œ«» ﬂ‰Ìœ</option>
		<option value="-2">---------------------------------</option>
		<% set RSV=Conn.Execute ("SELECT * FROM Bankers where IsBankAccount=1") 
		Do while not RSV.eof
		%>
			<option value="<%=RSV("id")%>" <% if cint(RSV("id"))=cint(Banks) then %> selected <% end if %>><%=RSV("Name")%> </option>
		<%
		RSV.moveNext
		Loop
		RSV.close
		%>
		<option value="-2">---------------------------------</option>
		<option value="-2"   <% if -2=cint(Banks) then %> selected <% end if %>>Â„Â »«‰ﬂÂ«</option>
		</select>
	</TD>
	<TD  align=center width=200><INPUT class="GenButton" TYPE="submit" name="submit" value="Ã” ÃÊ" ></TD>
</TR>
<TR>
	<TD align=right> «  «—ÌŒ :ù <INPUT  class="RcpMainTableInput" dir="LTR"  TYPE="text" NAME="ChequeDatesTo" maxlength="10" size="10" onblur="acceptDate(this)" onKeyPress="return maskDate(this);"  value="<%=ChequeDatesTo%>"></TD>
	<TD colspan=2>
		Ê÷⁄Ì  :
		<select name="status"  style="font-size:8pt">
		<option value="">«‰ Œ«» ﬂ‰Ìœ</option>
		<option value="">---------------------------------</option>
		<% set RSV=Conn.Execute ("SELECT * FROM RcvPayChqStatus where IsRcvdChqStatus = 0 order by id") 
		Do while not RSV.eof
		%>
			<option value="<%=RSV("id")%>"  <% if cint(RSV("id"))=cint(status) then %> selected <% end if %>><%=RSV("Name")%> </option>
		<%
		RSV.moveNext
		Loop
		RSV.close
		%>
		<option value="">---------------------------------</option>
		<option value="-1"   <% if -1=cint(status) then %> selected <% end if %>>Â— Ê÷⁄Ì ﬂÂ œ«—œ</option>
		</select>
	</TD>
</TR>
</TABLE><BR>
</FORM>

<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Show Report
'-----------------------------------------------------------------------------------------------------

total = 0 

'response.write "<span dir = ltr>" & myQuery
'response.end

Set RSS = conn.Execute(myQuery2)
%>
<TABLE dir=rtl align=center width=600>
<TR >
	<TD colspan=5>
		
		<BR>
	</TD>
</TR>
<TR bgcolor="eeeeee" >
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=1"><SMALL>ﬂœ  </SMALL></A> &nbsp;<% response.write DecsAsc(1) %></TD>
	<TD align=center><!A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=2"><SMALL></SMALL></A></TD>
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=2"><SMALL>‘„«—Â çﬂ</SMALL></A> &nbsp;<% response.write DecsAsc(2) %></TD>
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=3"><SMALL> «—ÌŒ çﬂ </SMALL></A> &nbsp;<% response.write DecsAsc(3) %></TD>
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=8"><SMALL> êÌ—‰œÂ </SMALL></A> &nbsp;<% response.write DecsAsc(8) %></TD>
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=4"><SMALL>„»·€ (—Ì«·)</SMALL></A> &nbsp;<% response.write DecsAsc(4) %></TD>
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=5"><SMALL> Ê”ÿ</SMALL></A> &nbsp;<% response.write DecsAsc(5) %></TD>
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=6"><SMALL>Ê÷⁄Ì  </SMALL></A> &nbsp;<% response.write DecsAsc(6) %></TD>
	<TD><A HREF="traceCheqPaid.asp?fromSession=y&Desc=<%=Desc%>&s=7"><SMALL> »«‰ﬂ </SMALL></A> &nbsp;<% response.write DecsAsc(7) %></TD>
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
<FORM METHOD=POST ACTION="traceCheqPaid.asp" onsubmit="return isEmpty()">
<TR bgcolor="<%=tmpColor%>" title="<%=RSS("Description")%>">
	<TD><% if RSS("Description")<>"" then %><font color=blue><B><% end if %><%=RSS("ID")%></B></font></TD>
	<TD align=center><% if RSS("status")=0 then %><INPUT TYPE="radio" NAME="pay" value="<%=RSS("ID")%>"><% end if %></TD>
	<TD><!A HREF="traceCheqPaid.asp?act=Detail&ID=<%=RSS("ID")%>"><A HREF="ShowCheqPaid.asp?cheqID=<%=RSS("ID")%>"><%=RSS("ChequeNo")%></A></A></TD>
	<TD><span dir=ltr><%=RSS("ChequeDate")%></span></TD>
	<TD><span dir=ltr><A HREF="../AO/AccountReport.asp?act=show&selectedCustomer=<%=RSS("Account")%>" target="_blank"><%=RSS("Account")%></A></span></TD>
	<TD><span dir=ltr><%=RSS("amount")%></span></TD>
	<TD><%=RSS("CSR")%></TD>
	<TD><small><% if RSS("status")=0 then %> Å«”  ‰‘œÂ <% else %> Å«”  ‘œÂ <% end if %></small></TD>
	<TD><%=RSS("BankName")%></TD>
</TR>
	  
<% 
total = total + RSS("amount")
RSS.moveNext
Loop
%>
<TR bgcolor="eeeeee" >
	<TD colspan=5 align=center> Ã„⁄ </TD>
	<TD colspan=4><big><%=total%></big> —Ì«·</TD>
</TR>
</TABLE><br>
<INPUT class="GenButton" TYPE="submit" name="submit" value="Å«” ‘œ"><br><br>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
<!--
function isEmpty()
{
	notFound=true;
	for (i=0;i<document.getElementsByName("pay").length;i++){
		if(document.getElementsByName("pay")[i].checked){
			notFound=false;
			x=document.getElementsByName("pay")[i].value;
		}
	}
	if ( notFound )
	{
		alert("ÂÌç çﬂÌ »—«Ì Å«” ‘œ‰ «‰ Œ«» ‰‘œÂ «” ")
		return false
	}
	else
		return confirm("çﬂ ‘„«—Â "+ x + " Å«” „Ì ‘Êœ? ")
}
//-->
</SCRIPT>
</CENTER>
<!--#include file="tah.asp" -->