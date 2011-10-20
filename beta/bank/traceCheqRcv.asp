<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "ÅÌêÌ—Ì çﬂÂ«Ì œ—Ì«› Ì"
SubmenuItem=2
if not Auth("A" , 2) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<BR><BR>
<style>
	.RcpMainTable { font-family:tahoma; font-size: 9pt; border: 2px solid #000088; padding:0; direction: RTL; width:600px;}
	.RcpMainTable Tr {Height:25px; border: 1px solid black;}
	.RcpMainTable Input { font-family:tahoma; font-size: 9pt; width:120px; border: 1px solid gray; text-align:center; direction: LTR;}
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
'-------------------------------------------------------------------------------------- change  Status
'-----------------------------------------------------------------------------------------------------
if request("act") = "changeStatus" then

	id		= request.form("id")
	Status	= request.form("Status")
	Banker = request.form("Banks")
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	response.redirect "?act=detail&ID=" & ID

	myQuery = "UPDATE ReceivedCheques SET LastStatus = "& Status & ", LastBanker = "& Banker & ", LastUpdatedDate=N'"& shamsiToday() & "', LastUpdatedBy="& session("ID")& " WHERE  (ID = "& ID & ")"

	Set RSS = conn.Execute(myQuery)
	response.redirect "?act=detail&ID=" & ID


'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Form ûAction
'-----------------------------------------------------------------------------------------------------
elseif request("act") = "actionButton" then
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	response.write "⁄„·Ì«  " & request.form("submit") &  " »—«Ì çﬂÂ«Ì “Ì— «‰Ã«„ ‘œ:<BR><BR>"
	if request.form("submit")="Ê«ê–«—Ì »Â »«‰ﬂ"  then
		if request.form("longDistance")="on" then
			status = 7
		else
			status = 5
		end if
		NewBanker= request.form("NewBanker")
	elseif request.form("submit")="«” —œ«œ"  then
		status = 4
		NewBanker= "-"
	elseif request.form("submit")="Ê’Ê·"  then
		status = 2
		NewBanker= "-"
	elseif request.form("submit")="»—ê‘ "  then
		status = 3
		NewBanker= request.form("NewBanker")

		mySQL="SELECT Users.Password, Users.Account  FROM Bankers INNER JOIN Users ON Bankers.Responsible = Users.ID WHERE (Bankers.ID = '"& NewBanker & "') AND (Users.Password = N'"& sqlSafe(request.form("ChequesNewBankerPass")) & "')"
		Set RS1=conn.execute(mySQL)
		If RS1.EOF Then
			conn.close
			response.redirect "?fromSession=y&errmsg=" & Server.URLEncode("ﬂ·„Â ⁄»Ê— êÌ—‰œÂ çﬂ Â« ‰«œ—”  «” ")
			response.end
		elseif RS1("Password")<>request.form("ChequesNewBankerPass") then
			conn.close
			response.redirect "?fromSession=y&errmsg=" & Server.URLEncode("ﬂ·„Â ⁄»Ê— êÌ—‰œÂ çﬂ Â« ‰«œ—”  «” ")
			response.end
		end if

	else
		response.write "<BR><BR><CENTER>Œÿ«Ì €Ì— „‰ Ÿ—Â!</CENTER>"
		response.end
	end if 

	for cheques = 1 to request.form("CHID").count 
		myQuery = "UPDATE ReceivedCheques SET LastStatus = "& Status 
		if not NewBanker= "-" then myQuery =  myQuery & ", LastBanker = " & NewBanker 
		myQuery =  myQuery & ", LastUpdatedDate=N'"& shamsiToday() & "', LastUpdatedBy="& session("ID")& " WHERE  (ID = "& request.form("CHID")(cheques) & ")"
		conn.Execute(myQuery)
		response.write "<li> <A HREF='?act=detail&ID="& request.form("CHID")(cheques)  & "'>‘„«—Â "&  request.form("CHID")(cheques) & "</a>"

		if status = 4 then   'ESTERDAD 

			MYSQL1 = "SELECT Receipts.SYS , Receipts.Customer AS Account , ReceivedCheques.Amount , Receipts.ID FROM ReceivedCheques INNER JOIN  Receipts ON ReceivedCheques.Receipt = Receipts.ID WHERE (ReceivedCheques.ID = "& request.form("CHID")(cheques) & ")"
			set RS1 = conn.Execute(mySQL1)
			SYS=RS1("SYS")
			Account=RS1("Account")
			Amount=RS1("Amount")
			ReceiptID=RS1("ID")
			Description = "»œÂÌ »«»  «” —œ«œ çﬂ »Â ‘„«—Â "& request.form("CHID")(cheques)
			creationDate = shamsiToday()

			MYSQL1 = "INSERT INTO "& SYS & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Description, Amount) VALUES (N'"& creationDate & "',"& session("ID") & ", 5,"& Account & ",0 ,N'"& Description & "',"& Amount & ")"
			conn.Execute(mySQL1)
			MySQL1="select ID from "& SYS & "Memo where Account="& Account & " and Amount="& Amount & " and CreatedDate=N'"& creationDate & "' order by ID DESC"
			RS1 = conn.Execute(mySQL1)
			MemoID=RS1("ID")

			GLAccount="17011"	'This must be changed...

			mySQL1="INSERT INTO "& SYS & "Items (GLAccount, GL, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, Account, RemainedAmount) VALUES ('" &_
			GLAccount & "', '"& OpenGL & "', 0, 3, '"& MemoID & "', '"& Amount & "', N'"& creationDate & "', '"& session("ID") & "', '"& Account & "', '"& Amount & "')"	
			conn.Execute(mySQL1)

			mySQL1="UPDATE Accounts SET "& SYS & "Balance = "& SYS & "Balance - "& Amount & " WHERE (ID='"& Account & "')"

			conn.Execute(mySQL1)

			'RS1.close
			Set RS1 = nothing

		end if
	next

	'  ============================================================================ 
	'  There is a Trigger in DB which sets Banker.Balance if the action is VOSOOL : 
	'	
	'	CREATE TRIGGER Trig_VosooleChq ON [dbo].[ReceivedCheques] 
	'	FOR UPDATE
	'	AS 
	'	DECLARE 
	' 		@status tinyint,
	'		@banker int,
	'		@chequeAmount int
	'		SELECT 
	'			@status = LastStatus, 
	'			@banker = LastBanker, 
	'			@chequeAmount=Amount 
	'		FROM inserted 
	'	if @status =2 
	'	UPDATE Bankers SET 
	'		CurrentBalance = CurrentBalance + @chequeAmount
	'	WHERE ID = @banker
	'  ============================================================================ 

	response.write "<BR><BR><CENTER><A HREF='?fromSession=y'>»«“ê‘  »Â ¬Œ—Ì‰ ê“«—‘  </A></CENTER><br>"

	response.end


'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------- ûcheque Details
'-----------------------------------------------------------------------------------------------------
elseif request("act") = "detail" then
	id = request("id")

	myQuery = "SELECT ReceivedCheques.*, RcvPayChqStatus.Name AS statusName, ReceivedCheques.ID AS CHID, Bankers.IsBankAccount AS IsBankAccount, Bankers.Name AS BankName, Accounts.AccountTitle AS AccountTitle, Accounts.ID AS AccountID, Receipts.ID AS RcptID FROM ReceivedCheques INNER JOIN RcvPayChqStatus ON ReceivedCheques.LastStatus = RcvPayChqStatus.ID INNER JOIN Bankers ON ReceivedCheques.LastBanker = Bankers.ID INNER JOIN Receipts ON ReceivedCheques.Receipt = Receipts.ID INNER JOIN Accounts ON Receipts.Customer = Accounts.ID WHERE (ReceivedCheques.ID = "& ID & ")"

	Set RSS = conn.Execute(myQuery)

	if RSS.EOF then
		response.write "<CENTER><BR><BR>Œÿ«! ç‰Ì‰ çﬂÌ ÊÃÊœ ‰œ«—œ</CENTER>"
	end if

	%>
	<BR><BR>
	<TABLE>
	<TR>
		<TD valign=top>
			<TABLE style="border:2pt solid white;">
			<TR bgcolor=white>
				<TD colspan=2 align=center> „‘Œ’«  çﬂ </TD>
			</TR>
			<TR>
				<TD align=left>ﬂœ :</TD>
				<TD><b><%=RSS("CHID")%></b></TD>
			</TR>
			<TR>
				<TD  align=left> ‘„«—Â çﬂ :</TD>
				<TD><b><%=RSS("ChequeNo")%></b></TD>
			</TR>
			<TR>
				<TD align=left>’«œ— ﬂ‰‰œÂ :</TD>
				<TD ><A target=_blank  HREF="../CRM/AccountInfo.asp?act=show&selectedCustomer=<%=RSS("AccountID")%>"><b><%=RSS("AccountTitle")%></b></A></TD>
			</TR>
			<TR>
				<TD align=left> «—ÌŒ çﬂ  :</TD>
				<TD> <span dir=ltr><b><%=RSS("ChequeDate")%></b></span></TD>
			</TR>
			<TR>
				<TD align=left>„»·€ (—Ì«·) :</TD>
				<TD><span dir=ltr><b><%=RSS("amount")%></b></span></TD>
			</TR>
			<TR>
				<TD align=left> Ê÷ÌÕ«  :</TD>
				<TD><b><%=RSS("Description")%></b></TD>
			</TR>
			<TR>
				<TD align=left>»«‰ﬂ :</TD>
				<TD><b><%=RSS("BankOfOrigin")%></b></TD>
			</TR>
			<TR>
				<TD align=left> Ê÷⁄Ì  :</TD>
				<TD> <b><%=RSS("statusName")%></b></TD>
			</TR>
			<TR>
				<TD align=left>„Õ·  :</TD>
				<TD> <b><%=RSS("BankName")%></b></TD>
			</TR>
			<TR bgcolor=white>
				<TD colspan=2 align=center height=2> </TD>
			</TR>
			<TR >
				<TD colspan=2 align=center ><A target=_blank HREF="../AO/AccountReport.asp?act=showReceipt&receipt=<%=RSS("RcptID")%>">„—»Êÿ »Â œ—Ì«›  ‘„«—Â <%=RSS("RcptID")%></A></TD>
			</TR>
			<!--TR bgcolor=white>
				<TD colspan=2 align=center height=2> </TD>
			</TR>
			<FORM METHOD=POST ACTION="?act=changeStatus">
			<INPUT TYPE="hidden" name="id" value="<%=RSS("CHID")%>">
			<TR>
				<TD align=left> €ÌÌ— Ê÷⁄Ì </TD>
				<TD><%
					status = RSS("LastStatus")
					%>
					<select name="status" <% if not RSS("IsBankAccount") then%> disabled <% end if %> >
					<% set RSV=Conn.Execute ("SELECT * FROM RcvPayChqStatus") 
					Do while not RSV.eof
					%>
						<option value="<%=RSV("id")%>"  <% if cint(RSV("id"))=cint(status) then %> selected <% end if %>><%=RSV("Name")%> </option>
					<%
					RSV.moveNext
					Loop
					RSV.close
					%>
					</select>
					<% if not RSS("IsBankAccount") then%> <INPUT TYPE="hidden" name="status" value="<%=status%>"> <% end if %>
				</TD>
			</TR>
			<TR>
				<TD align=left> €ÌÌ— „Õ·</TD>
				<TD><% 
					Banks = RSS("LastBanker")
					%>
					<select name="Banks" <% if status<>3  and RSS("IsBankAccount") then%> disabled <% end if %>>
					<% set RSV=Conn.Execute ("SELECT * FROM Bankers ") 
					Do while not RSV.eof
					%>
						<option value="<%=RSV("id")%>" <% if RSV("id")=Banks then %> selected <% end if %>><%=RSV("Name")%> </option>
					<%
					RSV.moveNext
					Loop
					RSV.close
					%>
					</select>
					<% if status<>3 and RSS("IsBankAccount") then%> <INPUT TYPE="hidden" name="Banks" value="<%=Banks%>"> <% end if %>
				</TD>
			</TR>
			<TR>
				<TD colspan=2>
					
					<br>
					<CENTER><INPUT TYPE="submit" value="À» "></CENTER>

				</TD>
			</TR-->
			</FORM>
			</TABLE>
		</TD>
		
		<TD valign=top>
			<table style="border:2pt solid white;" dir=ltr >
			<tr bgcolor=white>
				<td colspan=3 align=center> ”«»ﬁÂ Ê÷⁄Ì  Â« </td>
			</tr>
			<% set RSV=Conn.Execute ("SELECT RcvChqTrace.Cheque, RcvPayChqStatus.Name AS StatusName, Bankers.Name AS BankerName, RcvChqTrace.SetDate FROM RcvChqTrace INNER JOIN  RcvPayChqStatus ON RcvChqTrace.Status = RcvPayChqStatus.ID INNER JOIN  Bankers ON RcvChqTrace.Banker = Bankers.ID WHERE (RcvChqTrace.Cheque = "& ID & ") ORDER BY RcvChqTrace.SetDate") 
			Do while not RSV.eof
			%>
			<tr>
				<td align=right><%=RSV("BankerName")%></td>
				<td align=right bgcolor="f1f1f1"><%=RSV("StatusName")%></td>
				<td align=right><%=RSV("SetDate")%></td>
			</tr>
				
			<%
			RSV.moveNext
			Loop
			RSV.close
			%>
			</table>
		</TD>
	</TR>
	</TABLE>

	<%
	response.write "<BR><BR><CENTER><A HREF='?fromSession=y'>»«“ê‘  »Â ¬Œ—Ì‰ ê“«—‘  </A></CENTER><br>"
	response.end
end if
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Submit Search
'-----------------------------------------------------------------------------------------------------

'beginOfQuery = "SELECT  ReceivedCheques.*, RcvPayChqStatus.Name AS statusName, ReceivedCheques.ID AS CHID, ReceivedCheques.Receipt AS ReceiptID, Bankers.IsBankAccount AS Expr1, Bankers.Name AS BankName, Receipts.SYS AS SYS FROM ReceivedCheques INNER JOIN RcvPayChqStatus ON ReceivedCheques.LastStatus = RcvPayChqStatus.ID INNER JOIN Bankers ON ReceivedCheques.LastBanker = Bankers.ID INNER JOIN Receipts ON ReceivedCheques.Receipt = Receipts.ID"


beginOfQuery = "SELECT ReceivedCheques.*, Invoices.IsA, Invoices.Number, RcvPayChqStatus.Name AS statusName, ReceivedCheques.ID AS CHID, ReceivedCheques.Receipt AS ReceiptID, Bankers.IsBankAccount AS Expr1, Bankers.Name AS BankName, Receipts.SYS AS SYS FROM ReceivedCheques INNER JOIN RcvPayChqStatus ON ReceivedCheques.LastStatus = RcvPayChqStatus.ID INNER JOIN Bankers ON ReceivedCheques.LastBanker = Bankers.ID INNER JOIN Receipts ON ReceivedCheques.Receipt = Receipts.ID LEFT OUTER JOIN ARItems ON Receipts.ID = ARItems.Link FULL OUTER JOIN Invoices FULL OUTER JOIN ARItems ARItems_1 ON Invoices.ID = ARItems_1.Link FULL OUTER JOIN ARItemsRelations ON ARItems_1.ID = ARItemsRelations.DebitARItem ON ARItems.ID = ARItemsRelations.CreditARItem WHERE (ARItems.Type = 2) AND (Receipts.ID = Receipts.ID)"

if request.form("submit") = "Ê«ê–«—Ì" then
	Banks = request.form("Banks")
	ChequeDatesFrom = request.form("ChequeDatesFrom")
	ChequeDatesTo = request.form("ChequeDatesTo")

	myQuery = beginOfQuery & " and (LastStatus=6 or LastStatus=3) "
	status = -4

	if not ( ChequeDatesTo = "" ) then 
		myQuery  = myQuery  &  " AND (ReceivedCheques.ChequeDate < N'"& ChequeDatesTo & "') "
	end if 
	
	if not ( ChequeDatesFrom = "" ) then 
		myQuery  = myQuery  &  " AND (ReceivedCheques.ChequeDate >= N'"& ChequeDatesFrom & "')"
	end if 
	
	if not Banks = "-2" then 
		myQuery  = myQuery  & " AND (ReceivedCheques.LastBanker = "& Banks & ")"
	end if 

end if 

if request.form("submit") = "Ê’Ê·" then
	Banks = request.form("Banks")
	ChequeDatesFrom = request.form("ChequeDatesFrom")
	ChequeDatesTo = request.form("ChequeDatesTo")

	myQuery = beginOfQuery & " and (LastStatus=5 or LastStatus=7) "
	status = -5

	if not ( ChequeDatesTo = "" ) then 
		myQuery  = myQuery  &  " AND (ReceivedCheques.ChequeDate < N'"& ChequeDatesTo & "') "
	end if 
	
	if not ( ChequeDatesFrom = "" ) then 
		myQuery  = myQuery  &  " AND (ReceivedCheques.ChequeDate >= N'"& ChequeDatesFrom & "')"
	end if 
	
	if not Banks = "-2" then 
		myQuery  = myQuery  & " AND (ReceivedCheques.LastBanker = "& Banks & ")"
	end if 

end if 


if request.form("submit") = "Ã” ÃÊ" then
	status = request.form("status")
	Banks = request.form("Banks")
	ChequeDatesFrom = request.form("ChequeDatesFrom")
	ChequeDatesTo = request.form("ChequeDatesTo")

	myQuery = beginOfQuery '& " and 1=1 "
	
	if not ( ChequeDatesTo = "" ) then 
		myQuery  = myQuery  &  " AND (ReceivedCheques.ChequeDate < N'"& ChequeDatesTo & "') "
	end if 
	
	if not ( ChequeDatesFrom = "" ) then 
		myQuery  = myQuery  &  " AND (ReceivedCheques.ChequeDate >= N'"& ChequeDatesFrom & "')"
	end if 
	
	if not Banks = "-2" then 
		myQuery  = myQuery  & " AND (ReceivedCheques.LastBanker = "& Banks & ")"
	end if 

	if status = "-3" then 
		myQuery  = myQuery  & " and not (LastStatus=4 or LastStatus=2) "
	elseif status = "-4" then 
		myQuery  = myQuery  & " and (LastStatus=3 or LastStatus=6) "
	else	
		if not status = "-2" then 
			myQuery  = myQuery  & " AND (ReceivedCheques.LastStatus = "& status & ")"
		end if 
	end if

	'response.write myQuery
end if 

'if myQuery = "" then 
'	myQuery = beginOfQuery & " where not (LastStatus=-4)"
'	Banks = -2
'	ChequeDatesFrom = ""
'	ChequeDatesTo = ""
'	status = "-4"
'end if

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
	orderBy = "ReceivedCheques.ID"
elseif s="2" then
	orderBy = "ReceivedCheques.ChequeNo"
elseif s="3" then
	orderBy = "ReceivedCheques.ChequeDate"
elseif s="4" then
	orderBy = "ReceivedCheques.amount"
elseif s="5" then
	orderBy = "ReceivedCheques.LastStatus"
elseif s="6" then
	orderBy = "ReceivedCheques.LastBanker"
elseif s="7" then
	orderBy = "ReceivedCheques.BankOfOrigin"
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

<FORM METHOD=POST ACTION="?">
<INPUT class="GenButton" TYPE="submit" name="submit" value="Ê«ê–«—Ì" <% if cint(status)=-4 or cint(status)=3 or cint(status)=6 then %> style="background-color:#FFFFBB"<% else %> style="background-color:white"<% end if %>> 
<INPUT class="GenButton" TYPE="submit" name="submit" value="Ê’Ê·"  <% if cint(status)=5 or cint(status)=7 or cint(status)=-5 then %> style="background-color:#FFFFBB"<% else %> style="background-color:white"<% end if %>> 
<BR><BR>
<TABLE align=center class="RcpMainTable">
<TR>
	<TD align=right>  «“  «—ÌŒ : <INPUT  dir="LTR"  TYPE="text" NAME="ChequeDatesFrom" maxlength="10" size="10"onblur="acceptDate(this)" onKeyPress="return maskDate(this);" value="<%=ChequeDatesFrom%>"></TD>
	<TD>„Õ· çﬂ:
		<select name="Banks" onchange="" style="font-size:8pt">
		<option value="-2">«‰ Œ«» ﬂ‰Ìœ</option>
		<option value="-2">---------------------------------</option>
		<% set RSV=Conn.Execute ("SELECT * FROM Bankers ") 
		Do while not RSV.eof
		%>
			<option value="<%=RSV("id")%>" <% if cint(RSV("id"))=cint(Banks) then %> selected <% end if %>><%=RSV("Name")%> </option>
		<%
		RSV.moveNext
		Loop
		RSV.close
		%>
		<option value="-2">---------------------------------</option>
		<option value="-2"   <% if -2=cint(Banks) then %> selected <% end if %>>Â— Ã« Â”  </option>
		</select>
	</TD>
	<TD  align=center><INPUT class="GenButton" TYPE="submit" name="submit" value="Ã” ÃÊ" ></TD>
</TR>
<TR>
	<TD align=right> «  «—ÌŒ :ù <INPUT  dir="LTR"  TYPE="text" NAME="ChequeDatesTo" maxlength="10" size="10"  onKeyPress="return maskDate(this);" onblur="acceptDate(this)"  value="<%=ChequeDatesTo%>"></TD>
	<TD colspan=2>
		Ê÷⁄Ì  :ù 
		<select name="status"  style="font-size:8pt">
		<option value="-2">«‰ Œ«» ﬂ‰Ìœ</option>
		<option value="-2">---------------------------------</option>
		<% set RSV=Conn.Execute ("SELECT * FROM RcvPayChqStatus where IsRcvdChqStatus = 1 order by Name") 
		Do while not RSV.eof
		%>
			<option value="<%=RSV("id")%>"  <% if cint(RSV("id"))=cint(status) then %> selected <% end if %>><%=RSV("Name")%> </option>
		<%
		RSV.moveNext
		Loop
		RSV.close
		%>
		<option value="-2">---------------------------------</option>
		<option value="-2"   <% if -2=cint(status) then %> selected <% end if %>>Â— Ê÷⁄Ì ﬂÂ œ«—œ</option>
		<option value="-3"   <% if -3=cint(status) then %> selected <% end if %>>œ— Ã—Ì«‰</option>
		<option value="-4"   <% if -4=cint(status) then %> selected <% end if %>>¬„«œÂ Ê«ê–«—Ì</option>
		<option value="-5"   <% if -5=cint(status) then %> selected <% end if %>>¬„«œÂ Ê’Ê·</option>
		</select>
	</TD>
</TR>
</TABLE><BR>
</FORM>
<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------- Show Report
'-----------------------------------------------------------------------------------------------------

if not myQuery = "" then 

	total = 0 

	Set RSS = conn.Execute(myQuery2)
	%>
	<FORM METHOD=POST ACTION="?act=actionButton">
	<TABLE dir=rtl align=center width=600>
	<TR >
		<TD colspan=5>
			
			<BR>
		</TD>
	</TR>
	<TR bgcolor="eeeeee" >
		<TD><A HREF="?fromSession=y&Desc=<%=Desc%>&s=1"><% if cint(status)=-4 or cint(status)=3 or cint(status)=6 or cint(status)=5  or cint(status)=7 or cint(status)=-5  then  %><INPUT TYPE="checkbox" NAME="" disabled><% end if %><SMALL>ﬂœ</SMALL></A> &nbsp;<% response.write DecsAsc(1) %></TD>
		<TD><!A HREF="?fromSession=y&Desc=<%=Desc%>&s=0"><SMALL>‰Ê⁄</SMALL></A> &nbsp;<% response.write DecsAsc(0) %></TD>
		<TD><A HREF="?fromSession=y&Desc=<%=Desc%>&s=2"><SMALL>‘„«—Â çﬂ</SMALL></A> &nbsp;<% response.write DecsAsc(2) %></TD>
		<TD><A HREF="?fromSession=y&Desc=<%=Desc%>&s=3"><SMALL> «—ÌŒ çﬂ </SMALL></A> &nbsp;<% response.write DecsAsc(3) %></TD>
		<TD><A HREF="?fromSession=y&Desc=<%=Desc%>&s=4"><SMALL>„»·€ (—Ì«·)</SMALL></A> &nbsp;<% response.write DecsAsc(4) %></TD>
		<TD><A HREF="?fromSession=y&Desc=<%=Desc%>&s=5"><SMALL>Ê÷⁄Ì  </SMALL></A> &nbsp;<% response.write DecsAsc(5) %></TD>
		<TD><A HREF="?fromSession=y&Desc=<%=Desc%>&s=6"><SMALL> „Õ· </SMALL></A> &nbsp;<% response.write DecsAsc(6) %></TD>
		<TD><A HREF="?fromSession=y&Desc=<%=Desc%>&s=7"><SMALL> »«‰ﬂ </SMALL></A> &nbsp;<% response.write DecsAsc(7) %></TD>
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

	'Set RSS2 = conn.Execute("SELECT Invoices.IsA, Invoices.Number FROM Receipts INNER JOIN ARItems ON Receipts.ID = ARItems.Link INNER JOIN ARItemsRelations ON ARItems.ID = ARItemsRelations.CreditARItem INNER JOIN ARItems ARItems_1 ON ARItemsRelations.DebitARItem = ARItems_1.ID INNER JOIN Invoices ON ARItems_1.Link = Invoices.ID WHERE (ARItems.Type = 2) and Receipts.id ="& RSS("ReceiptID"))

	%>
	<TR bgcolor="<%=tmpColor%>" title="<%=RSS("Description")%>">
		<TD><% if cint(status)=-4 or cint(status)=3 or cint(status)=6 or cint(status)=5 or cint(status)=7 or cint(status)=-5 then  %><INPUT TYPE="checkbox" NAME="CHID" VALUE="<%=RSS("CHID")%>"  onclick="setColor(this)"><% end if %><A HREF="?act=detail&ID=<%=RSS("CHID")%>"><%=RSS("CHID")%></A></TD>
		
		<% if not isnull(RSS("IsA")) then 
			if RSS("IsA") then 	%>
				<TD align=center style="border:1pt solid red"><span dir=ltr>«·›</span></TD>
			<% else %>
				<TD align=center style="border:1pt solid blue"><span dir=ltr>»</span></TD>
			<% end if %>
		<% else %>
			<TD><span dir=ltr>	</span></TD>
		<% end if %>
		<TD align=center><A HREF="?act=detail&ID=<%=RSS("CHID")%>"><%=RSS("ChequeNo")%></A></TD>
		<TD><span dir=ltr><%=RSS("ChequeDate")%></span></TD>
		<TD><span dir=ltr><%=RSS("amount")%></span></TD>
		<TD><%=RSS("statusName")%></TD>
		<TD><%=RSS("BankName")%></TD>
		<TD><%=RSS("BankOfOrigin")%></TD>
	</TR>
		  
	<% 
	total = total + RSS("amount")
	RSS.moveNext
	Loop
	%>
	<TR bgcolor="eeeeee" >
		<TD colspan=4 align=center> Ã„⁄ </TD>
		<TD colspan=4><big><%=total%></big> —Ì«·</TD>
	</TR>
	<TR >
		<TD colspan=8 align=center><BR>
			<% if cint(status)=-4 or cint(status)=3 or cint(status)=6 then %>
			 <INPUT class="GenButton" TYPE="submit" name="submit" value="Ê«ê–«—Ì »Â »«‰ﬂ" >
				<select class="GenButton"  name="NewBanker" onchange="">
				<% set RSV=Conn.Execute ("SELECT * FROM Bankers where isBankAccount=1 ") 
				Do while not RSV.eof
				%>
					<option value="<%=RSV("id")%>"><%=RSV("Name")%> </option>
				<%
				RSV.moveNext
				Loop
				RSV.close
				%>
				</select>
				<INPUT TYPE="checkbox" NAME="longDistance"> ‘Â—” «‰
				<BR><BR>
		
			<INPUT class="GenButton" TYPE="submit" name="submit" value="«” —œ«œ"  onclick="return confirm(' ÊÃÂ! «Ì‰ ⁄„·Ì«  »—ê‘  Ì« «»ÿ«· Å–Ì— ‰„Ì »«‘œ. ¬Ì« «“ ’Õ  ¬‰ „ÿ„∆‰Ìœ øø')">  
			<!--INPUT class="GenButton" TYPE="submit" name="submit" value="Ê«ê–«—Ì »Â ’‰œÊﬁ" > 
				<select name="Banks" onchange="" class="GenButton" >
				<% set RSV=Conn.Execute ("SELECT CashRegisters.ID, CashRegisters.NameDate, Users.RealName FROM CashRegisters INNER JOIN Users ON CashRegisters.Cashier = Users.ID WHERE (CashRegisters.IsOpen = 1)") 
				Do while not RSV.eof
				%>
					<option value="<%=RSV("id")%>"><%=RSV("NameDate")%> (<%=RSV("RealName")%>)</option>
				<%
				RSV.moveNext
				Loop
				RSV.close
				%>
				</select-->
			<% elseif cint(status)=5 or cint(status)=7 or cint(status)=-5 then %>
			<INPUT class="GenButton" TYPE="submit" name="submit" value="»—ê‘ "  onclick="return formValidation();"> 
				<select class="GenButton"  name="NewBanker" onchange="">
				<% set RSV=Conn.Execute ("SELECT Users.RealName, Bankers.ID , Bankers.Name FROM Bankers INNER JOIN Users ON Bankers.Responsible = Users.ID WHERE (Bankers.IsBankAccount = 0) ") 
				Do while not RSV.eof
				%>
					<option value="<%=RSV("id")%>"> »Â <%=RSV("Name")%> (<%=RSV("RealName")%>)</option>
				<%
				RSV.moveNext
				Loop
				RSV.close
				%>
				<!--option value="-2">---------------------------------</option>
				<option value="-2"   <% if -2=cint(Banks) then %> selected <% end if %>>Â— Ã« Â”  </option-->
				</select>
				<INPUT TYPE="password" NAME="ChequesNewBankerPass" onkeyDown="return myKeyDownHandler();" onKeyPress="return myKeyPressHandler();">

				<BR><BR>
			<INPUT class="GenButton" TYPE="submit" name="submit" value="Ê’Ê·" onclick="return confirm(' ÊÃÂ! «Ì‰ ⁄„·Ì«  »—ê‘  Ì« «»ÿ«· Å–Ì— ‰„Ì »«‘œ. ¬Ì« «“ ’Õ  ¬‰ „ÿ„∆‰Ìœ øø')"> 
			<% end if  %>
		</TD>
	</TR>
	</TABLE><br>
	</FORM>
<% end if %>
</CENTER>
<SCRIPT LANGUAGE="JavaScript">
<!--
tmpColor="#FFFFFF"
tmpColor2="#FFFFBB"

function setColor(obj)
{
ii=parseInt(obj.id) 
if(obj.checked)
	{
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor",tmpColor2)
	}
else
	{
	theTR = obj.parentNode.parentNode
	theTR.setAttribute("bgColor",tmpColor)
	}
}

var tempKeyBuffer;
function myKeyDownHandler(){
	tempKeyBuffer=window.event.keyCode;
}
function myKeyPressHandler(){
//	alert (tempKeyBuffer)
	if (tempKeyBuffer>=65 && tempKeyBuffer<=90){
		window.event.keyCode=tempKeyBuffer+32;
	}
	else if(tempKeyBuffer==186){
		window.event.keyCode=59;
	}
	else if(tempKeyBuffer==188){
		window.event.keyCode=44;
	}
	else if(tempKeyBuffer==190){
		window.event.keyCode=46;
	}
	else if(tempKeyBuffer==191){
		window.event.keyCode=47;
	}
	else if(tempKeyBuffer==192){
		window.event.keyCode=96;
	}
	else if(tempKeyBuffer>=219 && tempKeyBuffer<=221){
		window.event.keyCode=tempKeyBuffer-128;
	}
	else if(tempKeyBuffer==222){
		window.event.keyCode=39;
	}
}

function formValidation(){

	if (document.all.ChequesNewBankerPass)
	if (!document.all.ChequesNewBankerPass.value){
		alert("ﬂ·„Â ⁄»Ê— —« Ê«—œ ﬂ‰Ìœ");
		document.all.ChequesNewBankerPass.focus();
		return false;
	}
	return true;
}
//-->
</SCRIPT>

<!--#include file="tah.asp" -->
