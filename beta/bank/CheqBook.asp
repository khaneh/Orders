<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "œ› — çﬂ"
SubmenuItem=8
if not Auth("A" , 8) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {vertical-align:top;padding:5;border:1pt solid gray;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTableTitle {background-color: #CCCCEE; text-align: center; font-weight:bold;}
	.RepTableHeader {background-color: #BBBBFF; text-align: center; font-weight:bold;}
	.RepTableFooter {background-color: #BBBBBB; direction: LTR; }
	.RepTR0 {background-color: #DDDDDD;}
	.RepTR1 {background-color: #FFFFFF;}
	.GenInput {width:70px; font-family:tahoma; font-size: 9pt; border: 1 solid black; text-align:left; direction: LTR;}
	.RepTextArea {font-family:tahoma;font-size:8pt;width:300px;height:30px;border:none;backGround:Transparent;}
	.RepROInputs {font-family:tahoma;font-size:8pt;width:70px;border:none;backGround:Transparent;}
</STYLE>
<%
'------------------------------------------------------------------------------
'------------------------------------------------------------------ Submit Memo
'------------------------------------------------------------------------------
if request("act")="submitMemo" then
	Account=	request.form("Account")
	GLAccount=	clng(request.form("GLAccount"))


	if GLAccount="" then
		response.redirect "?errMsg="&Server.URLEncode("Œÿ« œ— Ê—ÊœÌ !")
	end if

' ---
' 	Checking input
' ---

	GLMemoDate = request.form("GLMemoDate")

	'---- Checking wether EffectiveDate is valid in current open GL
	if (GLMemoDate < session("OpenGLStartDate")) OR (GLMemoDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!<br> «—ÌŒ Ê«—œ ‘œÂ œ— „ÕœÊœÂ ”«· „«·Ì Ã«—Ì ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	if Account<>"" then
		Account= clng(Account)
		bySubSys=1
	else
		if request.form("CheqNos").count > 0 then
			if request.form("Accounts")(1) <> "" then 
				SubSystemNeeded=True
			else
				SubSystemNeeded=False
			end if
		end if

		ContradictionFound=False
		for i=2 to request.form("CheqNos").count 
			if request.form("Accounts")(i) <> "" then 
				thisSubSystemNeeded=True
			else
				thisSubSystemNeeded=False
			end if
			if NOT thisSubSystemNeeded = SubSystemNeeded then
				ContradictionFound=True
				exit for
			end if
		Next
		if ContradictionFound then
			response.redirect "?errMsg="&Server.URLEncode("‰«Â„«Â‰êÌ œ— «ÿ·«⁄« !<br><br>Ê«ê–«—Ì »Â Õ”«» Â«Ì «‘Œ«’ —« Ãœ« «“ »ﬁÌÂ «‰Ã«„ œÂÌœ.")
		end if
		if SubSystemNeeded then
			bySubSys=1
		else
			bySubSys=0
		end if
	end if
' ---

	creationDate=shamsiToday()

	mySQL="SELECT ISNULL(MAX(GLDocID),0) AS LastMemo FROM GLDocs WHERE GL='"& OpenGL & "'"
	Set RS1=conn.Execute (mySQL)
	GLMemoNo = RS1("LastMemo") + 1

	'---- Creating a new GLDoc 
	mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy, BySubSystem, IsTemporary) VALUES ("& openGL & " , "& GLMemoNo & ", N'"& GLMemoDate & "' , N'"& creationDate & "', "& session("ID") & ", "& bySubSys & ", 1);SELECT @@Identity AS NewGLDoc"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	GLDoc = RS1("NewGLDoc")
	RS1.close
	'---- 

	if Account="" then
	'
	' When the first line doesn't need a SubSystem Memo
	'
		for i=1 to request.form("CheqNos").count 
			thisAccount=	request.form("Accounts")(i)
			thisGLAccount=	request.form("GLAccounts")(i)
			thisDescription=sqlSafe(request.form("Descriptions")(i))
			thisAmount=		request.form("Amounts")(i)

			if request.form("IsCredit")(i) then 
				thisIsCredit=1 
				thisIsDebit=0 
			else 
				thisIsDebit=1 
				thisIsCredit=0
			end if

			thisCheqNo=request.form("CheqNos")(i)
			thisCheqDate=request.form("CheqDates")(i)

			if thisAccount<>"" then
			'
			'	The seccond line needs a SubSystem Memo
			'	so, we wait for SYS and LINK of the Memo before we insert either of the 2 lines
			'
				' --------------------------------------------------
				' -----------	Creating a Memo and an Item for this
				' --------------------------------------------------
				' -----------	Cheque Return («” —œ«œ çﬂ)
				' **************************************************
				mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& thisGLAccount & ")"
				Set RS1=Conn.execute(mySQL)
				if RS1.eof then
					'Using default reason (sys: AO, reason: Misc.)
					Reason=6
				else
					Reason=	cint(RS1("ItemReason"))
				end if
				RS1.close

				mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
				Set RS1=Conn.execute(mySQL)
				sys=			RS1("Acron")
				ReasonName =	RS1("Name")
				RS1.close

				'*** Memo Type=5 means 'Cheque Return'
				mySQL="INSERT INTO "& sys & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& GLMemoDate & "' , "& session("ID") & ", 5, "& thisAccount & ", "& thisIsCredit & ", "& thisAmount & ", N'"& thisDescription & "');SELECT @@Identity AS NewMemo"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theMemo = RS1("NewMemo")
				RS1.close

				'*** Type = 3 means Item is a Memo
				mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, Reason, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, GL_Update) VALUES ('" &_
				GLAccount & "', "& OpenGL &", '"& thisGLAccount & "', '"& thisAccount & "', N'"& GLMemoDate & "', "& thisIsCredit & ", 3, "& theMemo & ", "& Reason & ", "& thisAmount & ", N'"& creationDate & "', "& session("ID") & ", "& thisAmount & ", 0);SELECT @@Identity AS NewItem"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theItem = RS1("NewItem")
				RS1.close
				
				if thisIsCredit then
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& thisAmount & "' WHERE (ID='"& thisAccount & "')"
				else
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& thisAmount & "' WHERE (ID='"& thisAccount & "')"
				end if
				conn.Execute(mySQL)

				' **************************************************
				' ------ End of Creating a Memo and an Item for this
				' --------------------------------------------------

				'First Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Amount, Description, IsCredit, Ref1, Ref2, SYS, Link) VALUES ( "&	GLDoc & ", "& GLAccount & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsDebit & ", N'"& thisCheqNo &"', N'"& thisCheqDate & "', '"& sys & "', '"& theItem & "')"
				conn.Execute(mySQL)

				'Seccond Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Tafsil, Amount, Description, IsCredit, Ref1, Ref2, SYS, Link) VALUES ( "&	GLDoc & ", "& thisGLAccount & ", "& thisAccount & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsCredit & ", N'"& thisCheqNo &"', N'"& thisCheqDate &"', '"& sys & "', '"& theItem & "')"
				conn.Execute(mySQL)
			else
			'
			'	Both Lines are only at GL Level
			'
				'First Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Amount, Description, IsCredit, Ref1, Ref2) VALUES ( "&	GLDoc & ", "& GLAccount & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsDebit & ", N'"& thisCheqNo &"', N'"& thisCheqDate &"')"
				conn.Execute(mySQL)

				'Seccond Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Amount, Description, IsCredit, Ref1, Ref2) VALUES ( "& GLDoc & ", "& thisGLAccount & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsCredit & ", N'"& thisCheqNo &"', N'"& thisCheqDate &"')"
				conn.Execute(mySQL)
			end if
		Next
	else
	'
	' When the first line needs a SubSystem Memo
	'
		for i=1 to request.form("CheqNos").count 
			thisAccount=	request.form("Accounts")(i)
			thisGLAccount=	request.form("GLAccounts")(i)
			thisDescription=sqlSafe(request.form("Descriptions")(i))
			thisAmount=		request.form("Amounts")(i)

			if request.form("IsCredit")(i) then 
				thisIsCredit=1 
				thisIsDebit=0 
			else 
				thisIsDebit=1 
				thisIsCredit=0
			end if

			thisCheqNo=request.form("CheqNos")(i)
			thisCheqDate=request.form("CheqDates")(i)

			if thisAccount="" then
			'
			'	The seccond line doesn't needs a SubSystem Memo
			'	we insert first line , find SYS and LINK of the Memo and then insert the seccond lines
			'
				' --------------------------------------------------
				' -----------	Creating a Memo and an Item for this
				' --------------------------------------------------
				' -----------	Cheque Return («” —œ«œ çﬂ)
				' **************************************************
				mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& GLAccount & ")"
				Set RS1=Conn.execute(mySQL)

				if RS1.eof then
					'Using default reason (sys: AO, reason: Misc.)
					Reason=6
				else
					Reason=	cint(RS1("ItemReason"))
				end if
				RS1.close

				mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
				Set RS1=Conn.execute(mySQL)
				sys=			RS1("Acron")
				ReasonName =	RS1("Name")
				RS1.close

				' * Note: This is affecting the first line info (Account) so it must be the opposite of isCredit
				'*** Memo Type=5 means 'Cheque Return'
				mySQL="INSERT INTO "& sys & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& GLMemoDate & "' , "& session("ID") & ", 5, "& Account & ", "& thisIsDebit & ", "& thisAmount & ", N'"& thisDescription & "');SELECT @@Identity AS NewMemo"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theMemo = RS1("NewMemo")
				RS1.close

				'*** Type = 3 means A*Item is a Memo
				mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, Reason, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, GL_Update) VALUES ('" &_
				thisGLAccount & "', "& OpenGL &", '"& GLAccount & "', '"& Account & "', N'"& GLMemoDate & "', "& thisIsDebit & ", 3, "& theMemo & ", "& Reason & ", "& thisAmount & ", N'"& creationDate & "', "& session("ID") & ", "& thisAmount & ", 0);SELECT @@Identity AS NewItem"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theItem = RS1("NewItem")
				RS1.close

				' * Note: This is affecting the first line info (Account) so it must be the opposite of isCredit
				if thisIsDebit then
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& thisAmount & "' WHERE (ID='"& Account & "')"
				else
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& thisAmount & "' WHERE (ID='"& Account & "')"
				end if
				conn.Execute(mySQL)

				' **************************************************
				' ------ End of Creating a Memo and an Item for this
				' --------------------------------------------------

				'First Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Tafsil, Amount, Description, IsCredit, Ref1, Ref2, SYS, Link) VALUES ( "&	GLDoc & ", "& GLAccount & ", "& Account & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsDebit & ", N'"& thisCheqNo &"', N'"& thisCheqDate &"', '"& sys & "', '"& theItem & "')"
				conn.Execute(mySQL)


				'Seccond Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Amount, Description, IsCredit, Ref1, Ref2, SYS, Link) VALUES ( "&	GLDoc & ", "& thisGLAccount & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsCredit & ", N'"& thisCheqNo &"', N'"& thisCheqDate &"', '"& sys & "', '"& theItem & "')"
				conn.Execute(mySQL)
			else
			'
			'	Both Lines need a SubSystem Memo
			'
				' --------------------------------------------------------------
				' -----------	Creating a Memo and an Item for the first line
				' --------------------------------------------------------------
				' -----------	Transfer Memo («⁄·«„ÌÂ «‰ ﬁ«·)
				' **************************************************
				mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& GLAccount & ")"
				Set RS1=Conn.execute(mySQL)

				if RS1.eof then
					'Using default reason (sys: AO, reason: Misc.)
					Reason=6
				else
					Reason=	cint(RS1("ItemReason"))
				end if
				RS1.close

				mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
				Set RS1=Conn.execute(mySQL)
				sys=			RS1("Acron")
				ReasonName =	RS1("Name")
				RS1.close

				' * Note: This is affecting the first line info (Account) so it must be the opposite of isCredit
				'*** Memo Type=7 means 'Transfer Memo'
				mySQL="INSERT INTO "& sys & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& GLMemoDate & "' , "& session("ID") & ", 7, "& Account & ", "& thisIsDebit & ", "& thisAmount & ", N'"& thisDescription & "');SELECT @@Identity AS NewMemo"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theMemo = RS1("NewMemo")
				RS1.close

				'*** Type = 3 means A*Item is a Memo
				mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, Reason, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, GL_Update) VALUES ('" &_
				thisGLAccount & "', "& OpenGL &", '"& GLAccount & "', '"& Account & "', N'"& GLMemoDate & "', "& thisIsDebit & ", 3, "& theMemo & ", "& Reason & ", "& thisAmount & ", N'"& creationDate & "', "& session("ID") & ", "& thisAmount & ", 0);SELECT @@Identity AS NewItem"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theItem = RS1("NewItem")
				RS1.close

				' * Note: This is affecting the first line info (Account) so it must be the opposite of isCredit
				if thisIsDebit then
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& thisAmount & "' WHERE (ID='"& Account & "')"
				else
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& thisAmount & "' WHERE (ID='"& Account & "')"
				end if
				conn.Execute(mySQL)

				firstItem=theItem
				firstMemo=theMemo
				firstSys=sys

				' **************************************************
				' ------ End of Creating a Memo and an Item for the first line
				' --------------------------------------------------------------


				' --------------------------------------------------------------
				' -----------	Creating a Memo and an Item for the seccond line
				' --------------------------------------------------------------
				' -----------	Transfer Memo («⁄·«„ÌÂ «‰ ﬁ«·)
				' **************************************************
				mySQL = "SELECT ItemReason FROM AXItemReasonGLAccountRelations WHERE (GL = "& openGL & ") AND (GLAccount = "& thisGLAccount & ")"
				Set RS1=Conn.execute(mySQL)

				if RS1.eof then
					'Using default reason (sys: AO, reason: Misc.)
					Reason=6
				else
					Reason=	cint(RS1("ItemReason"))
				end if
				RS1.close

				mySQL="SELECT * FROM AXItemReasons WHERE (ID="& Reason & ")"
				Set RS1=Conn.execute(mySQL)
				sys=			RS1("Acron")
				ReasonName =	RS1("Name")
				RS1.close

				'*** Memo Type=7 means 'Transfer Memo'
				mySQL="INSERT INTO "& sys & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"& GLMemoDate & "' , "& session("ID") & ", 7, "& thisAccount & ", "& thisIsCredit & ", "& thisAmount & ", N'"& thisDescription & "');SELECT @@Identity AS NewMemo"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theMemo = RS1("NewMemo")
				RS1.close

				'*** Type = 3 means A*Item is a Memo
				mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, Reason, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, GL_Update) VALUES ('" &_
				GLAccount & "', "& OpenGL &", '"& thisGLAccount & "', '"& thisAccount & "', N'"& GLMemoDate & "', "& thisIsCredit & ", 3, "& theMemo & ", "& Reason & ", "& thisAmount & ", N'"& creationDate & "', "& session("ID") & ", "& thisAmount & ", 0);SELECT @@Identity AS NewItem"
				set RS1 = Conn.execute(mySQL).NextRecordSet
				theItem = RS1("NewItem")
				RS1.close

				if thisIsCredit then
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& thisAmount & "' WHERE (ID='"& thisAccount & "')"
				else
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& thisAmount & "' WHERE (ID='"& thisAccount & "')"
				end if
				conn.Execute(mySQL)

				seccondItem=theItem
				seccondMemo=theMemo
				seccondSys=sys

				' **************************************************
				' ------ End of Creating a Memo and an Item for the first line
				' --------------------------------------------------------------

				'---------------------------------------------
				'------------------------------------ Relation
				if thisIsCredit then
					mySQL="INSERT INTO InterMemoRelation (FromItemType, FromItemLink, ToItemType, ToItemLink) VALUES ('"& firstSys & "', "& firstMemo & ", '"& seccondSys & "', "& seccondMemo & ")"
				else
					mySQL="INSERT INTO InterMemoRelation (FromItemType, FromItemLink, ToItemType, ToItemLink) VALUES ('"& seccondSys & "', "& seccondMemo & ", '"& firstSys & "', "& firstMemo & ")"
				end if
				conn.Execute(mySQL)
				'----------------------------- End of Relation
				'---------------------------------------------

				'First Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Tafsil, Amount, Description, IsCredit, Ref1, Ref2, SYS, Link) VALUES ( "&	GLDoc & ", "& GLAccount & ", "& Account & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsDebit & ", N'"& thisCheqNo &"', N'"& thisCheqDate &"', '"& firstSys & "', '"& firstItem & "')"
				conn.Execute(mySQL)


				'Seccond Line:
				mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Tafsil, Amount, Description, IsCredit, Ref1, Ref2, SYS, Link) VALUES ( "&	GLDoc & ", "& thisGLAccount & ", "& thisAccount & ", "& thisAmount & ", N'"& thisDescription & "', "& thisIsCredit & ", N'"& thisCheqNo &"', N'"& thisCheqDate &"', '"& seccondSys & "', '"& seccondItem & "')"
				conn.Execute(mySQL)

			end if
		Next
		
	end if

	conn.close
	response.redirect "../accounting/GLMemoDocShow.asp?id="&GLDoc&"&msg="& Server.URLEncode("”‰œ À»  ‘œ")
'------------------------------------------------------------------------------
'-------------------------------------------------------------------- Make Memo
'------------------------------------------------------------------------------
elseif request("act")="makeMemo" then

	Account=request.form("Account")
	GLAccount=request.form("GLAccount")

	if Account="" and GLAccount="" then
		response.redirect "?errMsg="&Server.URLEncode("›—„ Œ«·Ì «” ")
	end if

	GLMemoDate = request.form("GLMemoDate")
%>
	<br><br>
	<FORM METHOD=POST ACTION="?act=submitMemo" onsubmit="if (document.all.GLMemoDate.value==''){document.all.GLMemoDate.focus();return false;}else{return acceptDate(document.all.GLMemoDate)}">
	<INPUT TYPE="hidden" Name="Account" Value='<%=Account%>'>
	<INPUT TYPE="hidden" Name="GLAccount" Value='<%=GLAccount%>'>
	<table class="RepTable" align='center' width='90%'>
	<tr class="RepTableHeader">
		<td colspan='3' align='left'> «—ÌŒ ”‰œ:</td>
		<td colspan='2' align='right' ><INPUT class="GenInput" NAME="GLMemoDate" tabIndex="1" TYPE="text" maxlength="10" size="10" value="" onBlur="acceptDate(this)"></td>
	</tr>
	<tr class="RepTableTitle">
		<td>#</td>
		<td>Õ”«»</td>
		<td>‘—Õ</td>
		<td>»œÂﬂ«—</td>
		<td>»” «‰ﬂ«—</td>
	</tr>
<%
	tempCounter=0
	for i=1 to request.form("GLAccounts").count
	  if request.form("GLAccounts")(i)<>0 then
		tempCounter=tempCounter+1
		firstAccount=request.form("FirstAccounts")(i)
		if firstAccount=0 then firstAccount=""
		firstGLAccount=request.form("FirstGLAccounts")(i)

		thisAccount=request.form("Accounts")(i)
		if thisAccount=0 then thisAccount=""
		thisGLAccount=request.form("GLAccounts")(i)
		thisDescription=request.form("Descriptions")(i)
		thisAmount=request.form("Amounts")(i)
%>
		<tr class='RepTR<%=tempCounter MOD 2%>'>
			<td rowspan=2><%=tempCounter%></td>
			<td><%=firstGLAccount&firstAccount%></td>
			<td><%=thisDescription%></td>
			<%if request.form("IsCredit")(i) then%>
				<td><%=thisAmount%></td>
				<td>&nbsp;</td>
			<%else%>
				<td>&nbsp;</td>
				<td><%=thisAmount%></td>
			<%end if%>
		</tr>
		<tr class='RepTR<%=tempCounter MOD 2%>'>
			<td><%=thisGLAccount&thisAccount%>
				<INPUT TYPE="hidden" Name="FirstAccounts" Value='<%=FirstAccount%>'>
				<INPUT TYPE="hidden" Name="FirstGLAccounts" Value='<%=FirstGLAccount%>'>
				<INPUT TYPE="hidden" Name="Accounts" Value='<%=thisAccount%>'>
				<INPUT TYPE="hidden" Name="GLAccounts" Value='<%=thisGLAccount%>'>
				<INPUT TYPE="hidden" Name="Descriptions" Value='<%=thisDescription%>'>
				<INPUT TYPE="hidden" Name="Amounts" Value='<%=text2value(thisAmount)%>'>
				<INPUT TYPE="hidden" Name="IsCredit" Value='<%=request.form("IsCredit")(i)%>'>
				<INPUT TYPE="hidden" Name="CheqNos" Value='<%=request.form("CheqNos")(i)%>'>
				<INPUT TYPE="hidden" Name="CheqDates" Value='<%=request.form("CheqDates")(i)%>'>
				</td>
			<td><%=thisDescription%></td>
			<%if request.form("IsCredit")(i) then%>
				<td>&nbsp;</td>
				<td><%=thisAmount%></td>
			<%else%>
				<td><%=thisAmount%></td>
				<td>&nbsp;</td>
			<%end if%>
		</tr>
<%
	  end if
	Next
%>
	<tr class="RepTableHeader">
		<td colspan=5><INPUT TYPE="submit" Value="–ŒÌ—Â"class='GenInput' style='text-align:center;'></td>
	</tr>
	</table>
	</FORM>
<%
'------------------------------------------------------------------------------
'------------------------------------------------------------- Show Cheque Book
'------------------------------------------------------------------------------
elseif request("act")="showBook" then
	Response.CacheControl="no-cache"
	Response.AddHeader "pragma", "no-cache"
	glList=""
	GLAccount=0
	if inStr(request("GLAccount"),",")>1 then 
		glList=replace(request("GLAccount"),","," ")
	else
		GLAccount=request("GLAccount")
	end if
	
	Account=request("Account")
	FromDate=	request("FromDate")
	ToDate=		request("ToDate")
	displayMode=request("displayMode")
	
	if displayMode="" then displayMode=0

	if request("ShowRemained")="on" then
		ShowRemained=1
	else
		ShowRemained=0
	end if

	if GLAccount="" then
		response.redirect "?errMsg="&Server.URLEncode("›—„ Œ«·Ì «” ")
	end if
	
	if Account="" or not isnumeric(Account) then
		Tafsil = ""
	else
		Tafsil = clng(Account)
	end if

'	Changed By Kid 830812 for using new (partial) stored procedure 

	mySQL="EXEC proc_CheqBook "& GLAccount & ", '"& Tafsil & "', '"& FromDate & "', '"& ToDate & "', "& openGL & ", "& ShowRemained & ", " &  displayMode & ", '" &  glList & "'"
	 
	if FromDate="" AND ToDate="" then
		pageTitle="»Â ’Ê—  ﬂ«„·"
	elseif FromDate="" then
		pageTitle="«“ «» œ«  «  «—ÌŒ " & replace (ToDate,"/",".")
	elseif ToDate="" then
		pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  « «‰ Â« "
	else
		pageTitle="«“  «—ÌŒ "& replace (FromDate,"/",".") & "  «  «—ÌŒ " & replace (ToDate,"/",".")
	end if
'response.write mySQL
	
	Set RS1=Conn.Execute(mySQL).NextRecordset().NextRecordset()
	if not RS1.eof then
		tempCounter=0
		totalCredit=0
		totalDebit=0
%>		<br>
		<input type="hidden" Name='tmpDlgArg' value=''>
		<FORM METHOD=POST ACTION="?act=makeMemo">
		<input type="hidden" Name='Account' value='<%=Account%>'>
		<input type="hidden" Name='GLAccount' value='<%=GLAccount%>'>
		<table class="RepTable" align='center'>
		<tbody id='Cheqs'>
		<tr class="RepTableHeader">
			<td colspan="6" style="font-size:14pt;" dir="LTR"><%if GLAccount=0 then response.write glList else response.write GLAccount%><%if Account<>"" then response.write " - " & Account%><span style="font-size:9pt;"><br><br><%=pageTitle%></span></td>
		</tr>
		<tr>
			<td colspan="5">Ã” ÃÊ »—«Ì çﬂ:<INPUT TYPE="text" class="GenInput" NAME="searchChq" Value="<%=request("cheq")%>" onKeyPress="return handleSearch();"></td>
			<td align=left>
			<% 	
			values = GLAccount&"" 
			if Account="" then values = values & "0" else values = values  & Account
			values  = values  &"" & FromDate &"" &ToDate&""&openGL&""&ShowRemained
			ReportLogRow = PrepareReport ("CheqBook.rpt", "GLAccountTafsilFromDateToDateGLShowRemained", values , "/beta/dialog_printManager.asp?act=Fin") %>
			<INPUT Class="GenButton" style="border:1 solid green;" TYPE="button" TYPE="button" value=" ç«Å " onclick="printThisReport(this,<%=ReportLogRow%>);">
			</td>
		</tr>
		<tr class="RepTableHeader">
			<td>#</td>
			<td> «—ÌŒ ”‰œ</td>
			<td>‘—Õ</td>
			<td>»œÂﬂ«—</td>
			<td>»” «‰ﬂ«—</td>
			<td> «—ÌŒ çﬂ</td>
		</tr>
<%		Do while not RS1.eof
			FirstGLAcc=	RS1("GLAccount")
			FirstAcc=	RS1("Tafsil")
			if isnull(FirstAcc) then FirstAcc=0
			Ref1=		RS1("Ref1")
			Ref2=		RS1("Ref2")
			IsCredit=	RS1("IsCredit")
			Description=RS1("Description")
			GLDocDate=	RS1("GLDocDate")
			Amount=		RS1("Amount")

			if IsCredit then
				totalCredit = totalCredit + cdbl(Amount)
			else
				totalDebit = totalDebit + cdbl(Amount)
			end if

			if NOT (Ref1="" AND Ref2="") then tempCounter=tempCounter+1

%>			<tr class='RepTR<%=tempCounter MOD 2%>' onclick="return getData(this);">

				<td><%if tempCounter>0 then response.write tempCounter%>
					<INPUT TYPE="hidden" Name="FirstGLAccounts" Value="<%=FirstGLAcc%>">
					<INPUT TYPE="hidden" Name="FirstAccounts" Value="<%=FirstAcc%>">
					<INPUT TYPE="hidden" Name="GLAccounts" Value="0">
					<INPUT TYPE="hidden" Name="Accounts" Value="0"></td>
				<td dir=LTR>
					<INPUT readonly Name="GLDocDates" class="RepROInputs" Value="<%=GLDocDate%>"></td>
				<td>
					<TextArea readonly Name="Descriptions" class="RepTextArea"><%=replace(Description,"/",".")%></TextArea></td>
<%			if IsCredit then%>
				<td>
					<INPUT TYPE="hidden" Name="IsCredit" Value="<%=IsCredit%>">
				</td>
				<td>
					<INPUT readonly Name="Amounts" class="RepROInputs" Value="<%=Separate(Amount)%>">
				</td>
<%			else%>			
				<td>
					<INPUT readonly Name="Amounts" class="RepROInputs" Value="<%=Separate(Amount)%>">
				</td>
				<td>
					<INPUT TYPE="hidden" Name="IsCredit" Value="<%=IsCredit%>">
				</td>
<%			end if%>
				<td dir=LTR>
					<INPUT readonly Name="CheqDates" class="RepROInputs" Value="<%=Ref2%>">
					<INPUT type="hidden" readonly Name="CheqNos" class="RepROInputs" Value="<%=Ref1%>"></td>
			</tr>
<%			RS1.MoveNext
		Loop
		tempCounter = tempCounter + 1
%>
		<tr class='RepTR<%=tempCounter MOD 2%>'>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td><span style="width:300px;text-align:left;overflow:auto;text-overflow:ellipsis;"><B>Ã„⁄ :</B></span></td>
			<td><INPUT readonly class="RepROInputs" Value="<%=Separate(totalDebit)%>"></td>
			<td><INPUT readonly class="RepROInputs" Value="<%=Separate(totalCredit)%>"></td>
			<td>&nbsp;</td>
		</tr>
		<tr class="RepTableHeader">
			<td colspan="6"><INPUT TYPE="Button" Value=" «œ«„Â " onclick="submit();"></td>
		</tr>
		</tbody>
		</FORM>
		</table>
		<SCRIPT LANGUAGE="JavaScript">
		<!--

		function documentKeyDown() {
			var theKey = window.event.keyCode;
			var obj = window.event.srcElement;
			if (theKey == 114) { 
				if (obj.name=="searchChq"){document.all.searchChq.select();};
				window.event.keyCode=0;
				document.all.searchChq.focus();
				return false;
			}
		}
		document.onkeydown = documentKeyDown;

		document.all.searchChq.focus();
		//handleSearch();
		//-->
		</SCRIPT>
		<br>
<%	else
		RS1.Close
		conn.close
		response.redirect "?errMsg="&Server.URLEncode("ÂÌç çﬂÌ ÅÌœ« ‰‘œ")
	end if
	RS1.Close
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function handleSearch(){
		var theKey=event.keyCode;
		if (theKey==13){
			searchNo=document.getElementsByName("searchChq")[0].value
			theTable=document.getElementById("Cheqs")
			cheqsVector=document.getElementsByName("CheqNos")
			for(i=0; i<cheqsVector.length; i++){
				if (cheqsVector[i].value==searchNo){
					theTable.getElementsByTagName("TR")[i+2].scrollIntoView();
					return getData(theTable.getElementsByTagName("TR")[i+3])
				}
			}
			
		}
		else if (theKey<48 || theKey>57) // don't accept non-digits
			return false
	}

	function getData(src)
	{
		myIndex=src.rowIndex-3;
		if (src.getElementsByTagName("TD")[0].innerText==" ") return false; 
		for(i=0; i<src.getElementsByTagName("TD").length; i++){
			src.getElementsByTagName("TD")[i].setAttribute("bgColor","yellow")
		}
		document.all.tmpDlgArg.value=document.getElementsByName("Descriptions")[myIndex].value + " („»·€: " + document.getElementsByName("Amounts")[myIndex].value +")<br><br>"
		window.showModalDialog('dialog_GetDestination.asp',document.all.tmpDlgArg,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		if (document.all.tmpDlgArg.value !="") {
			Arguments=document.all.tmpDlgArg.value.split("#")
			document.getElementsByName("GLAccounts")[myIndex].value=Arguments[0]
			document.getElementsByName("Accounts")[myIndex].value=Arguments[1]
			if (Arguments[1]==0)
				src.setAttribute("title","»Â Õ”«» " + Arguments[0] +" ê–«‘ Â „Ì ‘Êœ.")
			else
				src.setAttribute("title","»Â Õ”«» "+Arguments[1] + " - " + Arguments[0] +" ê–«‘ Â „Ì ‘Êœ.")
		}
		else{
			document.getElementsByName("GLAccounts")[myIndex].value="0"
			document.getElementsByName("Accounts")[myIndex].value="0"
			for(i=0; i<src.getElementsByTagName("TD").length; i++){
				src.getElementsByTagName("TD")[i].setAttribute("bgColor","")
			}
			src.setAttribute("title","")
		}
		return false;
	}
	//-->
	</SCRIPT>
<%
'------------------------------------------------------------------------------
'------------------------------------------------------------------ Find Cheque
'------------------------------------------------------------------------------
elseif request("act")="findCheq" then
	chqNo = replace(sqlSafe(request("cheque")),".","")
	amount= replace(sqlSafe(request("amount")),".","")
	if isnumeric(chqNo) then
		chqNo=cdbl(chqNo)
		mySQL="SELECT GLDocs.GLDocDate, GLDocs.GLDocID, GLDocs.ID AS GLDoc, GLRows.* FROM GLRows INNER JOIN GLDocs ON GLRows.GLDoc = GLDocs.ID WHERE (GLRows.deleted = 0) AND (CAST(GLRows.Ref1 AS float) = '"& chqNo & "') AND (GLDocs.IsRemoved = 0) AND (GLDocs.deleted = 0) AND (GLDocs.GL = '"& openGL & "') AND NOT(IsTemporary = 0 AND IsFinalized = 0) ORDER BY GLDocs.GLDocDate, GLRows.ID"
	elseif isnumeric(amount) then
		amount=cdbl(amount)
		mySQL="SELECT GLDocs.GLDocDate, GLDocs.GLDocID, GLDocs.ID AS GLDoc, GLRows.* FROM GLRows INNER JOIN GLDocs ON GLRows.GLDoc = GLDocs.ID WHERE (GLRows.deleted = 0) AND (GLRows.Ref1 <> '') AND (GLRows.Amount = '"& amount & "') AND (GLDocs.IsRemoved = 0) AND (GLDocs.deleted = 0) AND (GLDocs.GL = '"& openGL & "') ORDER BY GLDocs.GLDocDate, GLRows.ID"
	else
		conn.close
		response.redirect "?errMsg="&Server.URLEncode("‘„«—Â çﬂ „⁄ »— ‰Ì” .")
	end if

	Set RS1=Conn.execute(mySQL)
	if RS1.eof then
		conn.close
'		response.redirect "?errMsg="&Server.URLEncode("«Ì‰ ‘„«—Â çﬂ ÅÌœ« ‰‘œ.")
		response.redirect "?errMsg="&Server.URLEncode("çﬂ »« «Ì‰ „‘Œ’«  ÅÌœ« ‰‘œ.")
	else
%>
		<br>
		<table class="RepTable" align='center' width='90%'>
		<tr class="RepTableTitle">
			<td>#</td>
			<td> «—ÌŒ</td>
			<td>”‰œ</td>
			<td>Õ”«»</td>
			<td>‘—Õ</td>
			<td>»œÂﬂ«—</td>
			<td>»” «‰ﬂ«—</td>
		</tr>
<%
		tempCounter=0
		Do While NOT RS1.eof
			tempCounter=tempCounter+1
			GLDocDate =		RS1("GLDocDate")
			GLDoc =			RS1("GLDoc")
			GLDocNo =		RS1("GLDocID")
			GLAccount =		RS1("GLAccount")
			Account =		RS1("Tafsil")
			Description =	RS1("Description")
			Ref1 =			RS1("Ref1")
			Ref2 =			RS1("Ref2")
			Amount =		RS1("Amount")
			IsCredit =		RS1("IsCredit")

			credit = ""
			debit = ""
			if IsCredit  then 
				credit = Amount 
				totalCredit = totalCredit + cdbl(Amount)
			else
				debit = Amount 
				totalDebit = totalDebit + cdbl(Amount)
			end if
%>
			<tr class='RepTR<%=tempCounter MOD 2%>'>
				<td><%=tempCounter%></td>
				<td style='width:60;direction:LTR;'><%=GLDocDate%></td>

				<td><A HREF="../accounting/GLMemoDocShow.asp?id=<%=GLDoc%>" Target="_blank"><%=GLDocNo%></A></td>
				<td style="cursor:hand;" onclick="window.location='?act=showBook&Account=<%=Account%>&GLAccount=<%=GLAccount%>&cheq=<%=Ref1%>';"><B><%=GLAccount&Account%></B></td>
				<td><%=Description%></td>
				<%if IsCredit then%>
					<td>&nbsp;</td>
					<td><%=Separate(Amount)%></td>
				<%else%>
					<td><%=Separate(Amount)%></td>
					<td>&nbsp;</td>
				<%end if%>
			</tr>
<%
			RS1.moveNext
		Loop
%>
		<tr class="RepTableTitle">
			<td colspan=7 height=30>&nbsp;</td>
		</tr>
		</table>
<%
	end if 
'------------------------------------------------------------------------------
'---------------------------------------------------------------------- default
'------------------------------------------------------------------------------
else
%>
	<BR>
	<Center>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>

	<FORM METHOD=POST ACTION="?act=showBook" onsubmit="return checkValidation();">
	<TABLE>
	<TR>
		<TD colspan=2 align=center> ›’Ì·Ì</TD>
		<TD align=center>„⁄Ì‰</TD>
	</TR>
	<TR>
		<TD colspan=2><INPUT dir="LTR"  TYPE="text" NAME="Account" ID="Account" maxlength="6" value="<%=Account%>" onKeyPress='return mask(this);' onBlur='check(this);' style="width:250px;border:solid 1pt black" tabindex=2></TD>
		<TD><INPUT dir="LTR"  TYPE="text" NAME="GLAccount" id="GLAccount" maxlength="50" value="<%=GLAccount%>" onKeyPress='return mask(this);' onBlur='check(this);' style="width:100px;border:solid 1pt black" tabindex=1></TD>
	</TR>
	<TR>
		<TD colspan=2><TextArea NAME="AccountName" id="AccountName" readonly style="width:250px;height:40px;font-family:tahoma;font-size:9pt;background:transparent; border:solid 1px white"><%=AccountName%></TextArea></TD>
		<TD><TextArea NAME="GLAccountName" id="GLAccountName" readonly  style="width:100px;height:40px;font-family:tahoma;font-size:9pt;background:transparent; border:solid 1px white;"><%=GLAccountName%></TextArea></TD>
	</TR>
	<tr>
		<td><input type='radio' name='displayMode' value=0 CHECKED><span>‰„«Ì‘ ⁄«œÌ</span></td>
		<td><input type='radio' name='displayMode' value=1><span>‰„«Ì‘ —Ê“«‰Â</span></td>
		<td><input type='radio' name='displayMode' value=2><span>‰„«Ì‘ „«Â«‰Â</span></td>
	</tr>
	<TR>
		<TD align=left><INPUT TYPE="checkbox" NAME="ShowRemained" tabindex=5> „«‰œÂ ﬁ»· ‰„«Ì‘ œ«œÂ ‘Êœ.</TD>
		<TD align=left>«“  «—ÌŒ</TD>
		<TD><INPUT dir="LTR" class="GenInput" NAME="FromDate" id="FromDate" tabIndex="3" TYPE="text" maxlength="10" size="10" value="" onBlur="acceptDate(this);if(this.value!=''){document.all.ShowRemained.checked=true;}else{document.all.ShowRemained.checked=false;}" style="width:100px;"></TD>
	</TR>
	<TR>
		<TD align=center><INPUT TYPE="Submit" Value=" «œ«„Â " tabindex=6></TD>
		<TD align=left> «  «—ÌŒ</TD>
		<TD><INPUT dir="LTR" class="GenInput" NAME="ToDate" id="ToDate" tabIndex="4" TYPE="text" maxlength="10" size="10" value="" onBlur="acceptDate(this)" style="width:100px;"></TD>
	</TR>
	<TR>
	</TR>
	</FORM>
	</TABLE>

	<hr width='80%'>
	<FORM METHOD=POST ACTION="?act=findCheq">
	<TABLE>
	<TR>
		<TD align=center colspan=2>Ã” ÃÊÌ çﬂ</TD>
	</TR>
	<TR>
		<TD align=center>‘„«—Â çﬂ</TD>
		<TD><INPUT dir="LTR"  TYPE="text" NAME="cheque" maxlength="10" onKeyPress='return mask(this);' onChange='check(this);' style="width:100px;border:solid 1pt black" tabindex=10></TD>
	</TR>
	<TR>
		<TD align=center>„»·€</TD>
		<TD><INPUT dir="LTR"  TYPE="text" NAME="amount" maxlength="10" onKeyPress='return mask(this);' onChange='check(this);' style="width:100px;border:solid 1pt black" tabindex=10></TD>
	</TR>
	</TABLE>
	<INPUT TYPE="Submit" Value=" Ã” ÃÊ " tabindex=11 onkeyDown='if(event.keyCode==9) {document.all.GLAccount.focus(); return false;}'>
	</FORM>
	</Center>

	<SCRIPT LANGUAGE="JavaScript">
	<!--
	document.all.GLAccount.focus();

	var dialogActive=false;


	function mask(src){ 
		var theKey=event.keyCode;
		if (theKey==32){
			event.keyCode=9
			document.all.tmpDlgArg.value="#"
			if (src.name=='GLAccount'){
				document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì „⁄Ì‰:"
				dialogActive=true
				window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				dialogActive=false
				if (document.all.tmpDlgTxt.value !="") {
					dialogActive=true
					window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					dialogActive=false
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						document.all.GLAccountName.value=Arguments[1];
					}
				}
			}
			else if (src.name=='Account') {
				document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì  ›’Ì·Ì:"
				dialogActive=true
				window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				dialogActive=false
				if (document.all.tmpDlgTxt.value !="") {
					dialogActive=true
					window.showModalDialog('../AR/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogWidth:780px; dialogHeight:500px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					dialogActive=true
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						document.all.AccountName.value=Arguments[1];
					}
				}
			}
		}
		else if ((theKey >= 48 && theKey <= 57) || theKey==13 || theKey==44 ) // [0]-[9] and [Enter] are acceptible
			return true;
		else
			return false;
	}

	function check(src){ 
//--------------------SAM FireFox Compatibility-------------------------------
		if (!dialogActive){
			badCode = false;
			if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
			if (src.name=='GLAccount' && src.value.indexOf(",")=='-1') {
				objHTTP.open('GET','../accounting/xml_GLAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape(objHTTP.responseText)
				document.all.GLAccountName.value=tmpStr;
				//if (tmpStr == 'ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ')
				//	src.value='';
			}
			else if (src.name=='Account') {
				objHTTP.open('GET','../accounting/xml_CustomerAccount.asp?id='+src.value,false)
				objHTTP.send()
				tmpStr = unescape(objHTTP.responseText)
				document.all.AccountName.value=tmpStr;
			}
		}
	}
	function checkValidation(){
	  try{
		box=document.all.GLAccount;
		check(box);
		if (box.value==''){
			box.style.backgroundColor="red";
			alert("ÂÌç Õ”«»Ì «‰ Œ«» ‰‘œÂ");
			box.style.backgroundColor="";
			box.focus();
			return false;
		}
		box=document.all.FromDate;
		if (box.value!='' && box.value!=null){
			if (!acceptDate(box))
				return false;
		}
		box=document.all.ToDate;
		if (box.value!='' && box.value!=null){
			if (!acceptDate(box))
				return false;
		}
	  }catch(e){
			alert("Œÿ«Ì €Ì— „‰ Ÿ—Â");
			return false;
	  }
	}

	//-->
	</SCRIPT>
<%
end if
%>

<!--#include file="tah.asp" -->
