<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= "”‰œ „—ﬂ»"
SubmenuItem=2
if not Auth(8 , "F") then NotAllowdToViewThisPage()


'---------------------------------------------
'---------------------------- ShowErrorMessage
'---------------------------------------------
function ShowErrorMessage(msg)
	response.write "<table align='center' cellpadding='5'><tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> Œÿ« ! <br>"& msg & "<br></td></tr></table><br>"
end function

Sub MakeSubsystemItem()
	'--- variables used:
	'	Conn
	'	Account 
	'	GLAccount
	'	GLMemoDate 
	'	IsCredit 
	'	Amount 
	'	theDescription 
	'	creationDate
	'
	'	Sys		(Must have value before)
	'	theItem	(Must have value before)

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
	'*** Memo Type=8 means 'Compound Memo'
	mySQL="INSERT INTO "& sys & "Memo (CreatedDate, CreatedBy, Type, Account, IsCredit, Amount, Description) VALUES (N'"&_
	GLMemoDate & "' , "& session("ID") & ", 8, "& Account & ", "& IsCredit & ", "& Amount & ", N'"& theDescription & "');SELECT @@Identity AS NewMemo"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	theMemo = RS1("NewMemo")
	RS1.close

	' * Note: there is no FirstGLAccount because this is a compound memo
	'*** Type = 3 means A*Item is a Memo
	mySQL="INSERT INTO "& sys & "Items (GLAccount, GL, Account, EffectiveDate, IsCredit, Type, Link, Reason, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, GL_Update) VALUES ('" &_
	GLAccount & "', "& OpenGL &", '"& Account & "', N'"& GLMemoDate & "', "& IsCredit & ", 3, "& theMemo & ", "& Reason & ", "& Amount & ", N'"& creationDate & "', "& session("ID") & ", "& Amount & ", 0);SELECT @@Identity AS NewItem"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	theItem = RS1("NewItem")
	RS1.close

	'*** Updating Accout's balance in the subsystem:
	if IsCredit then
		mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& Amount & "' WHERE (ID='"& Account & "')"
	else
		mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& Amount & "' WHERE (ID='"& Account & "')"
	end if
	conn.Execute(mySQL)
End Sub
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<style>
	Table { font-size: 9pt;}
	.InvRowInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; direction:LTR; width:100%;}
	.InvHeadInput { font-family:tahoma; font-size: 9pt; border: none; background-color: #CCCC88; text-align:center;}
	.InvRowInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: transparent; text-align:right; width:100%;}
	.InvRowInput3 { font-family:tahoma; font-size: 9pt; border: 1px solid gray; background-color: #F0F0F0; text-align:right; width:100%;cursor:hand;}
	.InvHeadInput2 { font-family:tahoma; font-size: 9pt; border: none; background-color: #AACC77; text-align:center;}
	.InvHeadInput3 { font-family:tahoma; font-size: 9pt; border: none; background-color: #F0F0F0; text-align:right; direction: right-to-left;}
	.InvGenInput  { font-family:tahoma; font-size: 9pt; border: none; }
	.InvGenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.GLInput1 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:RTL;}
	.GLInput2 { font-family:tahoma; font-size: 9pt; border: 1px solid black; direction:LTR; text-align:right; }
	.GLTable1 { border: none; direction:RTL; border:1px dashed red ; width:685;} 
	.GLTable1 tr {height:20; background-color: #CCCC88; }
	.GLTable2 { border: none; direction:RTL;} 
	.GLTable2 tr {height:20; text-align:center; background-color: #F0F0F0; }
	.GLTable2 td {border-bottom: 1px solid black; border-right: 1px solid black;}
	.GLTable3 tr {background-color: #F0F0F0}
	
	.CompMemoEdit td {background-color: #CCCCCC; font-size:10pt;}

	.GLTable5 { font-family:tahoma; font-size: 10pt; background-color: #66CCCC; border:2 solid #225555; padding:0; }
	.GLTable5 TH { background-color: #66BBCC; border: 0; font-size:16pt;border-bottom: 2 solid #225555;}
	.GLTable5TR  td { background-color: #EEEEEE; border: 0; font-size:10pt;}
	.GLTable5TR1 td{ background-color: #99DDDD; border: 0; }
	.GLTable5TR2 td{ background-color: #66CCCC; border: 0; }

	.GLTR1 { font-family:tahoma; font-size: 9pt; height:30; text-align:center; vertical-align:top; background-color: #C3C300; }
	.GLTR2 { height:20; text-align:center; background-color: #C3C300; }
	.GLTD1 { font-family:tahoma; font-size: 9pt; height:20; text-align:center; }
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var okToProceed=false;
var currentRow=0;
var IsTaraz = false
//-->
</SCRIPT>
<%

ON ERROR RESUME NEXT
	DocID=	clng(request("DocID"))
	if Err.Number<>0 then Err.clear:DocID=0
	id=		clng(request("id"))
	if Err.Number<>0 then Err.clear:id=0
ON ERROR GOTO 0

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Input a new GL Memo Doc
'-----------------------------------------------------------------------------------------------------
'xml_GLAccount.asp
if request("act")="" then
if session("IsClosed")="True" then 
	response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
end if
%>
<!-- Ê—Êœ «ÿ·«⁄«  ”‰œ -->
	<br>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<BR>
		<FORM METHOD=POST ACTION="?act=submitMemo" onsubmit="return checkValidation()">
		<table class="GLTable1" align="center" Cellspacing="0" Cellpadding="0">
			<tr class="GLTR1" align="center" Cellspacing="1" Cellpadding="0">
			<TD colspan="10"><TABLE Border="0" Width="100%" Cellspacing="1" Cellpadding="2" Dir="RTL">
				<TR>
				<TD align="left">œ› — :</TD>
				<TD>	<%=OpenGLName%>
				</TD>
				<TD align="left"> «—ÌŒ ”‰œ :</TD>
				<TD>	
					<INPUT class="GLInput2" NAME="GLMemoDate" TYPE="text" maxlength="10" size="10" value="<%=shamsiToday()%>" onblur="acceptDate(this)" >
				</TD>
				<TD align="left">‘„«—Â ”‰œ :</TD>
				<TD>	
<%				mySQL="SELECT ISNULL(MAX(GLDocID),0) AS LastMemo FROM GLDocs WHERE GL='"& OpenGL & "'"
				Set RS1=conn.Execute (mySQL)
				LastMemo = RS1("LastMemo")
%>
					<INPUT class="GLInput2" NAME="GLMemoNo" TYPE="text" maxlength="10" size="10" value="<%=LastMemo+1%>">
				</TD>
				</TR></TABLE></TD>
			</tr>
			<tr class="GLTR2">
				<TD colspan="10" align=right><div>

				<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" >
				<tr>
					<td style="width:26; border-right:none;"> # </td>
					<td style="width:60;"> ›’Ì·Ì</td>
					<td style="width:40;">„⁄Ì‰</td>
					<td style="width:300;">‘—Õ</td>
					<td style="width:80;">»œÂﬂ«—</td>
					<td style="width:80;">»” «‰ﬂ«—</td>
				</tr>
				</TABLE></div></TD>
			</tr>
			<tr>
			<td colspan="10">
			<div style="overflow:auto; height:250px; width:*;">
				<TABLE Border="0" Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="GLTable3">
				<Tbody id="GLrows">
				<tr bgcolor='#F0F0F0' onclick="setCurrentRow(this.rowIndex);" >
					<td colspan="6">
						<INPUT class="InvGenButton" TYPE="button" value="«÷«›Â" onkeyDown="if(event.keyCode==9) {setCurrentRow(this.parentNode.parentNode.rowIndex); return false;};" onClick="addRow();">
					</td>
				</tr>
				</Tbody></TABLE>
			</div>
			</td>
			</tr>
			<tr bgcolor='#CCCC88'>
				<td align=left>
				<B><span id="tarazDiv"> </span></B>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				»œÂﬂ«—  : <input type=text name="totalDebit" style="border:none; background:none;" value=0>
				&nbsp;&nbsp;&nbsp;
				»” «‰ﬂ«— : <input type=text name="totalCredit" style="border:none; background:none;" value=0>
				</td>
			</tr>
			</table>
		<br> 
		<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
		<tr>
			<td align='center'><INPUT class="InvGenButton" TYPE="submit" name="submit" value="–ŒÌ—Â"></td>
			<td align='center'><INPUT class="InvGenButton" TYPE="submit" name="submit" value="–ŒÌ—Â »Â ’Ê—  Ì«œœ«‘ " onclick="saveDraft()"> </td>
		</tr>
		</TABLE>
		</FORM>

<%
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Submit Edit GL Memo Doc
'-----------------------------------------------------------------------------------------------------

elseif request("act")="submitEditMemo" then
		if request.form("submit")="Õ–›" then
			response.write "asdfasdf"
			response.end

		end if

	ON ERROR RESUME NEXT
		if request.form("submit")="Õ–›" then
			IsTemporary = 0
		else
			IsTemporary = 1
		end if

		GLMemoDate=		sqlSafe(request.form("GLMemoDate"))
		GLMemoNo=		clng(request.form("GLMemoNo"))
		totalDebit =	cdbl(text2value(request.form("totalDebit")))
		totalCredit=	cdbl(text2value(request.form("totalCredit")))
		GLDoc =			clng(request.form("GLDoc"))

		TotalItemCount	=	request.form("GLAccounts").count 

		ReDim Accounts(TotalItemCount)
		ReDim GLAccounts(TotalItemCount)
		ReDim Descriptions(TotalItemCount)
		ReDim Amounts(TotalItemCount)
		ReDim IsCredits(TotalItemCount)
		ReDim Syss(TotalItemCount)
		ReDim Links(TotalItemCount)
		ReDim Functions(TotalItemCount)

		for i=1 to TotalItemCount
			Functions(i) =		request.form("Functions")(i)
			Accounts(i) =		clng(text2value(request.form("Accounts")(i)))
			GLAccounts(i) =		clng(text2value(request.form("GLAccounts")(i)))
			Descriptions(i) =	sqlSafe(request.form("Descriptions")(i))
			Syss(i) =			sqlSafe(request.form("Syss")(i))
			Links(i) =			sqlSafe(request.form("Links")(i))
			credit		=		cdbl(text2value(request.form("credits")(i)))	
			debit		=		cdbl(text2value(request.form("debits")(i)))	
			if credit <> "" and credit <> "0" then 
				Amounts(i) = credit 
				IsCredits(i)= 1
			else
				Amounts(i) = debit 
				IsCredits(i)= 0
			end if
		next	

		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«!")
		end if
	ON ERROR GOTO 0

	creationDate = shamsiToday()
	voidDate = creationDate 'for voided Memos

	'---- Checking wether EffectiveDate is valid in current open GL
	if (GLMemoDate < session("OpenGLStartDate")) OR (GLMemoDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?act=editDoc&id="& GLMemoNo & "&errMsg=" & Server.URLEncode("Œÿ«!  «—ÌŒ Ê«—œ ‘œÂ „ ⁄·ﬁ »Â ”«· „«·Ì Ã«—Ì ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then 
		Conn.close
		response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----

	'---- Marking old GLDoc and its GLRows as DELETED
	conn.Execute("UPDATE GLRows SET deleted = 1 WHERE (GLDoc = "& GLDoc & ")")
	conn.Execute("UPDATE GLDocs SET deleted = 1 WHERE (ID = "& GLDoc & ")")
	'---- 
	
	'---- Creating a new GLDoc 
	mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy, IsTemporary, IsCompound) VALUES ("& openGL & " , "& GLMemoNo & ", N'"& GLMemoDate & "' , N'"& creationDate & "', "& session("ID") & ", "& IsTemporary & ",1);SELECT @@Identity AS NewGLDoc"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	GLDoc = RS1 ("NewGLDoc")
	RS1.close
	'---- 

	'---- Inserting new GLRows 
	for i=1 to TotalItemCount
		if Functions(i)="V" then
			if Accounts(i) <> 0 then
			' Void Sybsystem Memo

			  '***--- Voiding A*Item of Memo (Very Similar to sub voidMemo in ../AO/MemoVoid.asp ) ---
			  '*** Type = 3 means the Item is a Memo
			  '***
			    sys = Syss(i)
				voidedItem = Links(i) '** Links(i) contains A*Item id from GLRows
				Account = Accounts(i)
				Amount = Amounts(i)
				IsCredit = IsCredits(i)

				'*********  Finding the A*Memo
				mySQL="SELECT Link FROM "& sys & "Items WHERE (Type=3) AND (ID='"& voidedItem & "')" 
				Set RS1=conn.Execute(mySQL)
				MemoID=RS1("Link")

				'-------------------SAM
				'mySQL = "UPADTE "& sys & "Items SET GL_Update=0 WHERE (Type=3) AND (ID='"& voidedItem & "')" 
				'set RS1=conn.Execute(mySQL)
				
				'*********  Void the A*Memo
				mySQL="UPDATE "& sys & "Memo SET Voided=1, VoidedDate=N'"& voidDate & "', VoidedBy='"& session("ID") & "' WHERE (ID='"& MemoID & "')"
				conn.Execute(mySQL)
				 
				'*********  Finding other Items related to this Item
				if isCredit then
					mySQL="SELECT ID AS RelationID, Debit"& sys & "Item, Amount FROM "& sys & "ItemsRelations WHERE (Credit"& sys & "Item = '"& voidedItem & "')"
					Set RS1=conn.Execute(mySQL)
					Do While not (RS1.eof)
						'*********  Adding back the amount in the relation, to the *DEBIT* Item ...
						conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Debit"& sys & "Item") & "')")

						'*********  Deleting the relation
						conn.Execute("DELETE FROM "& sys & "ItemsRelations WHERE ID='"& RS1("RelationID") & "'")
						
						RS1.movenext
					Loop
				else ' (isCredit==false)
					mySQL="SELECT ID AS RelationID, Credit"& sys & "Item, Amount FROM "& sys & "ItemsRelations WHERE (Debit"& sys & "Item = '"& voidedItem & "')"
					Set RS1=conn.Execute(mySQL)
					Do While not (RS1.eof)
						'*********  Adding back the amount in the relation, to the *CREDIT* Item ...
						conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Credit"& sys & "Item") & "')")

						'*********  Deleting the relation
						conn.Execute("DELETE FROM "& sys & "ItemsRelations WHERE ID='"& RS1("RelationID") & "'")
						
						RS1.movenext
					Loop
				end if

				'*********  Voiding A*Item 
				'conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& voidedItem & "')")
				' ------------- SAM ---------
				conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& voidedItem & "')")
				conn.Execute("UPDATE "& sys & "Items SET GL_Update=0 WHERE ID="&voidedItem)
				'----------------------------
				'**************************************************************
				'*				Affecting Account's A* Balance  
				'**************************************************************
				if isCredit then
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& Amount & "' WHERE (ID='"& Account & "')"
				else
					mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& Amount & "' WHERE (ID='"& Account & "')"
				end if

				conn.Execute(mySQL)
				
				'***
			  '***--- End of Voiding A*Item of Memo --------------------------------------------------
			end if
		else
			if Accounts(i) <> 0 and Functions(i)="I" then
			' New Sybsystem Memo
				GLAccount = GLAccounts(i)
				Account = Accounts(i)
				Amount = Amounts(i)
				IsCredit = IsCredits(i)
				theDescription = Descriptions(i)
				Sys = Syss(i)
				theItem =Links(i)

				' --------------------------------------------------------------
				' -----	Creating a Memo and an A*Item for this line
				' **************************************************

				Call MakeSubsystemItem()

				' **************************************************
				' ------ End of Creating Memo and A*Item for this line
				' --------------------------------------------------------------

				Syss(i) = Sys 

				Links(i) = theItem
			end if
			
			Syss(i) = "'" & Syss(i) & "'"
			if Accounts(i) = 0 then 
				Accounts(i) = "NULL"
				Syss(i) = "NULL"
				Links(i) = "NULL"
			end if

			mySQL="INSERT INTO GLRows ( GLDoc, GLAccount,  Tafsil, Amount, Description, Sys, Link, IsCredit) VALUES ( "&_
			 GLDoc & ", "& GLAccounts(i) & ", "& Accounts(i) & ", "& Amounts(i) & ", N'"& Descriptions(i) & "', "& Syss(i) & ", "& Links(i) & ", "& IsCredits(i) & ")"
			 response.write mySQL
			 'response.end

			conn.Execute(mySQL)
			

		end if
	next	
	'---- 
	response.redirect "GLMemoDocShow.asp?id="& GLDoc & "&msg=" & Server.URLEncode(" €ÌÌ—«  À»  ‘œ.")

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Submit new GL Memo Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitMemo" then

	GLMemoDate=request.form("GLMemoDate")
	GLMemoNo=request.form("GLMemoNo")
	totalDebit = text2value(request.form("totalDebit"))
	totalCredit = text2value(request.form("totalCredit"))
	creationDate = shamsiToday()
	if request.form("submit")="–ŒÌ—Â »Â ’Ê—  Ì«œœ«‘ " then
		IsTemporary = 0
	else
		IsTemporary = 1
	end if

	
	if GLMemoNo="" or not(isnumeric(GLMemoNo)) then
		ShowErrorMessage("Œÿ«! ‘„«—Â ”‰œ Ê«—œ ‰‘œÂ «” ")
		response.end
	end if

	'---- Checking wether EffectiveDate is valid in current open GL
	if (GLMemoDate < session("OpenGLStartDate")) OR (GLMemoDate > session("OpenGLEndDate")) then
		Conn.close
		response.write "<BR><BR><BR><CENTER>Œÿ«! ”‰œ Ê«—œ ‘œÂ „ ⁄·ﬁ »Â ”«· „«·Ì Ã«—Ì ‰Ì” .</CENTER>"
		response.end
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----

	WarningMsg=""
	Set RS3=Conn.Execute ("SELECT GLDocID FROM GLDocs WHERE (GLDocID='"& GLMemoNo & "') AND (GL='"& OpenGL & "')")
	if not RS3.eof then
		Set RS3= Conn.Execute("SELECT Max(GLDocID) AS MaxGLDocID FROM GLDocs WHERE (GL='"& OpenGL & "')")
		GLMemoNo=RS3("MaxGLDocID")+1
		WarningMsg="‘„«—Â ”‰œ  ﬂ—«—Ì »Êœ.<br>”‰œ »« ‘„«—Â <B>"& GLMemoNo & "</B> À»  ‘œ."
	end if
	RS3.Close

	mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy,  IsTemporary, IsCompound) VALUES ("& OpenGL & " , "& GLMemoNo & ", N'"& GLMemoDate & "' , N'"& creationDate & "', "& session("id") & ", "& IsTemporary & ", 1);SELECT @@Identity AS NewGLDoc"
	set RS1 = Conn.execute(mySQL).NextRecordSet
	GLDoc = RS1 ("NewGLDoc")
	RS1.close


	for i=1 to request.form("Accounts").count 
		Account = text2value(request.form("Accounts")(i))
		GLAccount = text2value(request.form("GLAccounts")(i))
		theDescription = request.form("Descriptions")(i)
		debit = text2value(request.form("debits")(i))
		credit = text2value(request.form("credits")(i))
		Sys = "NULL"
		theItem ="NULL"

		if credit <> "" and credit <> "0" then 
			Amount = credit 
			IsCredit = 1
		else
			Amount = debit 
			IsCredit = 0
		end if

		if amount = "" then amount = 0

		if Account <> 0 then
			' --------------------------------------------------------------
			' -----	Creating a Memo and an A*Item for this line
			' **************************************************

			Call MakeSubsystemItem()

			' **************************************************
			' ------ End of Creating Memo and A*Item for this line
			' --------------------------------------------------------------
			sys = "'" & sys & "'"

		end if


		' --------------------------------------------------------------
		' ----- Creating the GLRow for this line
		if Account=0 then 
			Account = "NULL"
			Sys = "NULL"
			theItem ="NULL"
		end if

		mySQL="INSERT INTO GLRows (GLDoc, Tafsil, GLAccount, Amount, Description, SYS, Link, IsCredit) VALUES ("&_
			GLDoc & ", " & Account & ", " & GLAccount & ", " & Amount & ", N'" & theDescription & "', " & sys & ", " & theItem & ", " & IsCredit & ")"

		Conn.Execute(mySQL)
		' **************************************************
		' ----- End of Creating the GLRow for this line
		' --------------------------------------------------------------

	next	
	conn.close

	response.redirect "GLMemoDocShow.asp?id="& GLDoc &"&msg="& Server.URLEncode("”‰œ »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ.") &"&errmsg="& Server.URLEncode(WarningMsg)

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Edit a Compund Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="editDoc" then


	mySQL="SELECT GLDocs.*, GLRows.GLAccount, GLRows.Tafsil, GLRows.Amount, GLRows.Description, GLRows.SYS, GLRows.Link, GLRows.IsCredit, GLAccounts.Name FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc INNER JOIN GLAccounts ON GLRows.GLAccount = GLAccounts.ID WHERE (GLDocs.Deleted = 0) AND (GLDocs.GLDocID="& id & ") AND (GLDocs.isCompound=1) AND (GLDocs.GL = "& OpenGL & ") AND (GLAccounts.GL = "& OpenGL & ") ORDER BY GLRows.ID"
	set RS1=conn.execute(mySQL)

	if RS1.eof then
		response.write "<BR><BR><BR><CENTER>Œÿ«! ç‰Ì‰ ”‰œ „—ﬂ»Ì œ— «Ì‰ ”«· „«·Ì ÊÃÊœ ‰œ«—œ</CENTER>"
		response.end
	end if 

	GLDocID = RS1("GLDocID")
	Creator = RS1("CreatedBy")

	classType=1


	if RS1("IsRemoved") OR RS1("deleted") then 
		response.write "<BR><BR><BR><CENTER>Œÿ«! «Ì‰ ”‰œ ﬁ»·« Õ–› ‘œÂ «” .</CENTER>"
		response.end
	elseif RS1("BySubSystem") then 
		response.write "<BR><BR><BR><CENTER>Œÿ«! «Ì‰ Ìﬂ ”‰œ « Ê„« Ìﬂ «” .</CENTER>"
		response.end
	elseif RS1("IsFinalized") then 
		response.write "<BR><BR><BR><CENTER>Œÿ«! «Ì‰ ”‰œ ﬁ»·« ﬁÿ⁄Ì ‘œÂ «” .</CENTER>"
		response.end
	elseif RS1("IsChecked") then 
		response.write "<BR><BR><BR><CENTER>Œÿ«! «Ì‰ ”‰œ ﬁ»·« »——”Ì ‘œÂ «” .</CENTER>"
		response.end
	elseif RS1("IsTemporary") then 
		if Not (session("ID") = Creator OR Auth( 8 , 7 ) )then
			ShowErrorMessage("‘„« „Ã«“ »Â ÊÌ—«Ì‘ «Ì‰ ”‰œ ‰Ì” Ìœ.")		
			conn.close
			response.end
		end if	
		statusString = " „Êﬁ  "
		status = "Temporary"
		classType=1

	else
		if Not (session("ID") = Creator OR Auth( 8 , 8 ) )then
			ShowErrorMessage("‘„« „Ã«“ »Â ÊÌ—«Ì‘ «Ì‰ ”‰œ ‰Ì” Ìœ.")		
			conn.close
			response.end
		end if	
		statusString = " Ì«œœ«‘  "
		status = "Draft"
		classType=3
	end if 
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----

	isBySubSystem = false

	if RS1("BySubSystem") then
		isBySubSystem = true
		if RS1("IsRemoved") then
			classType=2
		else
			classType=4
		end if
	end if

	isCompound = false

	if RS1("IsCompound") then
		isCompound = true
		if RS1("IsRemoved") then
			classType=2
		else
			classType=5
		end if
	end if 

	%>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>

	<FORM METHOD=POST ACTION="?act=submitEditMemo" onsubmit="return checkValidation()">
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="20">
			<td colspan=2></td>
		</tr>
		<tr height="10">
			<td width="250"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="20">
			<td colspan=2></td>
		</tr>
	</table>

	<table id="MemoTable" Border="0" align=center Cellspacing="1" Cellpadding="5" Dir="RTL" class="GLTable5">
	<tr>
		<th colspan=7>
			”‰œ ‘„«—Â <%=RS1("GLDocID")%> <INPUT TYPE='hidden' name='GLMemoNo' value='<%= RS1("GLDocID")%>'>
		</th>
	</tr>
	<tr>
		<th colspan=7 style="padding:0">
		<TABLE width="100%" border=0 Cellspacing="1" Cellpadding="10" Dir="RTL">
		<TR class="GLTable5TR2">
			<TD>‘„«—Â ⁄ÿ›: <%= RS1("id")%><INPUT TYPE='hidden' name='GLDoc' value='<%= RS1("ID")%>'></TD>
			<TD align=center>Ê÷⁄Ì : <%=statusString%></TD>
			<TD align=left> «—ÌŒ ”‰œ: <span dir=ltr><%= RS1("GLDocDate")%><INPUT TYPE='hidden' name='GLMemoDate' value='<%= RS1("GLDocDate")%>'></span></TD>
		</TR>
		</TABLE>
		</th>
	</tr>
	<tr class="GLTable5TR1">
		<td style="width:20;">
			<div id="mnuDivHome" style="width:100%;height:100%;position:relative;cursor:hand;">
				<div id="mnuDiv" style = 'position:absolute;top:-4;left:-3; width:100; height:20; background-color:#DDDDDD; border: 1px solid red; visibility:hidden;' onmouseout="hideMenu();" onmouseover="keepMenu();"> 
					<IMG SRC='../images/s_undo.gif' WIDTH='20' HEIGHT='20' BORDER='0' ALT='«‰’—«›' onclick="undoRow();" onmouseover="keepMenu();"> 
					<IMG SRC='../images/s_delete.gif' WIDTH='20' HEIGHT='20' BORDER='0' ALT='«»ÿ«· «Ì‰ Œÿ' onclick="voidRow();" onmouseover="keepMenu();"> 
					<IMG SRC='../images/s_insert.gif' WIDTH='20' HEIGHT='20' BORDER='0' ALT='œ—Ã Ìﬂ Œÿ' onclick="insertRow();" onmouseover="keepMenu();">
					<IMG SRC='../images/s_edit.gif' WIDTH='20' HEIGHT='20' BORDER='0' ALT='ÊÌ—«Ì‘ «Ì‰ Œÿ' onclick="editRow();" onmouseover="keepMenu();"> 
				</div>
			</div>
		</td>
		<td style="width:20;"> # </td>
		<td style="width:50;"> ›’Ì·Ì</td>
		<td style="width:45;">„⁄Ì‰</td>
		<td style="width:300;">‘—Õ</td>
		<td style="width:80;">»œÂﬂ«—</td>
		<td style="width:80;">»” «‰ﬂ«—</td>
	</tr>
<%
	i=0
	Do while not RS1.eof
		i = i + 1
		GLAccount = RS1("GLAccount")
		accTitle = RS1("name")
		theDescription = RS1("Description")
		Amount = Separate(RS1("Amount"))
		IsCredit = RS1("IsCredit")
		Tafsil = RS1("Tafsil")
		Sys =	RS1("Sys")
		ItemLink =	RS1("Link")

		if len(sys)>1 then  'if isBySubSystem then
			theLink = "ShowItem.asp?sys=" & sys & "&Item=" & ItemLink
		else
			theLink = "about:blank"
		end if

		credit = ""
		debit = ""
		if IsCredit then 
			credit = Amount 
			totalCredit = totalCredit + cdbl(Amount)
		else
			debit = Amount 
			totalDebit = totalDebit + cdbl(Amount)
		end if

		%>
		<tr class="GLTable5TR">
			<td ><div align=center style="width:100%;height:100%;position:relative;cursor:hand;" onmouseover="showMenu(this);">[+]</div></td>
			<td > <%=i%> 
				<INPUT TYPE='hidden' name='functions' value=''>
				<INPUT TYPE='hidden' name='Syss' value='<%=Sys%>'>
				<INPUT TYPE='hidden' name='Links' value='<%=ItemLink%>'>
			</td>
			<!---------------------SAM----------------->
			<td dir=ltr><INPUT class='InvRowInput2' TYPE='text' readonly NAME='Accounts'   value='<%=Tafsil%>'    style='width:50;' ></td>
			<td dir=ltr><INPUT class='InvRowInput2' TYPE='text' readonly NAME='GLAccounts' value='<%=GLAccount%>' style='width:45;' ></td>
			<td dir=rtl><INPUT class='InvRowInput2' TYPE='text' readonly NAME='Descriptions' value='<%=theDescription%>' style='width:300;' ></td>
			<td dir=ltr><INPUT class='InvRowInput2' TYPE='text' readonly NAME='debits'		value='<%=debit%>' style='width:80;' ></td>
			<td dir=ltr><INPUT class='InvRowInput2' TYPE='text' readonly NAME='credits'	value='<%=credit%>' style='width:80;' ></td>
		</tr>
		<%
	RS1.movenext
	loop	
%>
	<tr class="GLTable5TR1">
		<td style="width:20;" ><div align=center style="width:100%;height:100%;position:relative;cursor:hand;" onmouseover="showMenu(this);">[+]</div></td>
		<td></td>
		<td></td>
		<td></td>
		<td><DIV style="float:right;font-weight:bold;"><span id="tarazDiv"> </span></DIV><div style="float:left;">Ã„⁄</div></td>
		<td><input type="text" name="totalDebit" style="border:none; background:none; width:100%;" value="<%=Separate(totalDebit)%>"></td>
		<td><input type="text" name="totalCredit" style="border:none; background:none;width:100%;" value="<%=Separate(totalCredit)%>"></td>
	</tr>
	</TABLE>
	<TABLE Border="0" Cellspacing="5" Cellpadding="0" Dir="RTL" align='left'>
	<tr>
		<td align='center'><INPUT class="InvGenButton" TYPE="submit" name="submit" value="–ŒÌ—Â"></td>
		<!--td align='center'><INPUT class="InvGenButton" TYPE="submit" name="submit" value="Õ–›" style="border-color:red;" onclick="return confirm('¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ”‰œ —« Õ–› ﬂ‰Ìœø')"> </td-->
		<td align='center'><INPUT class="InvGenButton" TYPE="button" value="«‰’—«›" onclick="history.back()"></td>
	</tr>
	</TABLE>
	</FORM>

<SCRIPT LANGUAGE="JavaScript">
<!--
var mnuCountr=0;
var mnuDiv = document.getElementById("mnuDiv");
function showMenu(src){
	if (! src.getElementsByTagName('DIV')[0]){
		mnuCountr+=2;
		mnuDiv.style.visibility="visible";
		src.appendChild(mnuDiv);
	}
}
function keepMenu(){
	//alert('keep');
	if (mnuCountr<4)
		mnuCountr+=2;
}
function hideMenu(){
	if (mnuCountr>0){
		mnuCountr-=1;
		setTimeout('hideMenu()',2000);
		return;
	}
	mnuCountr=0;
	mnuDiv.style.visibility='hidden';
	document.getElementById("mnuDivHome").appendChild(mnuDiv);
}

function insertRow(){
	theRow = mnuDiv.parentNode.parentNode.parentNode
	
	mnuCountr=0;
	hideMenu();
	
	newRow=document.createElement("tr");

	tempTD=document.createElement("td");
	tempTD.innerHTML="<div align=center style='width:100%;height:100%;position:relative;cursor:hand;' onmouseover='showMenu(this);'><IMG SRC='../images/s_insert.gif' WIDTH='20' HEIGHT='20' BORDER='0' ALT='”ÿ— œ—Ã ‘œÂ' ></div>"
	newRow.appendChild(tempTD);


	tempTD=document.createElement("td");
	tempTD.innerHTML="<INPUT TYPE='hidden' name='functions' value='I'><INPUT TYPE='hidden' name='Syss' value=''><INPUT TYPE='hidden' name='Links' value=''>";
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='Accounts' maxlength=6 onKeyPress='return mask(this);' onBlur='check(this);' >"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='GLAccounts' maxlength=5 onKeyPress='return mask(this);' onBlur='check(this);' >"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='Descriptions'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='debits' onblur='setPrice(this)'  onKeyPress='return onlyNumber(this);'  onfocus='this.value=txt2val(this.value);this.select()' >"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='credits' onblur='setPrice(this)' onKeyPress='return onlyNumber(this);'   onfocus='this.value=txt2val(this.value);this.select()' >"
	newRow.appendChild(tempTD);

	newRow.className="CompMemoEdit";

	theRow.parentNode.insertBefore(newRow,theRow)

	newRow.getElementsByTagName('TD')[2].getElementsByTagName('INPUT')[0].focus();

}

function undoRow(){
	theRow = mnuDiv.parentNode.parentNode.parentNode
	tds = theRow.getElementsByTagName('TD')

	mnuCountr=0;
	hideMenu();

	func = tds[1].getElementsByTagName('INPUT')[0];

	if (!func ){
		return;
	}
	else if (func.value=='I'){
		removeRow(theRow);
		memoIsTaraz();
		return;
	}
	else if (func.value!='E' && func.value!='V' ){
		alert('«Ì‰ Ìﬂ ”ÿ— ÊÌ—«Ì‘ ‘œÂ ‰Ì” .')
		return;
	}

	theRow.className="GLTable5TR";

	tds[0].getElementsByTagName('DIV')[0].innerHTML="[+]";

	func.value='';
	if(tds[1].getElementsByTagName('DIV')[0])
		tds[1].removeChild(tds[1].getElementsByTagName('DIV')[0]);

	tds[2].innerHTML="<INPUT class='InvRowInput2' TYPE='text' readonly NAME='Accounts'   value='" + tds[2].getElementsByTagName('INPUT')[0].defaultValue + "' >"

	tds[3].innerHTML="<INPUT class='InvRowInput2' TYPE='text' readonly NAME='GLAccounts' value='" + tds[3].getElementsByTagName('INPUT')[0].defaultValue + "' >"

	tds[4].innerHTML="<INPUT class='InvRowInput2' TYPE='text' readonly NAME='Descriptions' value='" + tds[4].getElementsByTagName('INPUT')[0].defaultValue + "' >"

	tds[5].innerHTML= "<INPUT class='InvRowInput2' TYPE='text' readonly NAME='debits'	value='" + tds[5].innerText + tds[5].getElementsByTagName('INPUT')[0].defaultValue + "' >"

	tds[6].innerHTML="<INPUT class='InvRowInput2' TYPE='text' readonly NAME='credits'	value='" + tds[6].innerText + tds[6].getElementsByTagName('INPUT')[0].defaultValue + "' >"

	memoIsTaraz();

}	

function editRow(){
	theRow = mnuDiv.parentNode.parentNode.parentNode
	tds = theRow.getElementsByTagName('TD')

	mnuCountr=0;
	hideMenu();
	
	func = tds[1].getElementsByTagName('INPUT')[0];


	if (!func || tds[2].getElementsByTagName('INPUT')[0].value){
		alert('«Ì‰ ”ÿ— —« ‰„Ì  Ê«‰Ìœ ÊÌ—«Ì‘ ﬂ‰Ìœ.')
		return;
	}

	theRow.className="CompMemoEdit";

	tds[0].getElementsByTagName('DIV')[0].innerHTML="<IMG SRC='../images/s_edit.gif' WIDTH='20' HEIGHT='20' BORDER='0' ALT='”ÿ— ÊÌ—«Ì‘ ‘œÂ' >";

	tds[1].getElementsByTagName('INPUT')[0].value='E';
	
	tds[3].innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='GLAccounts' maxlength=5 style='width:45;border:1px solid gray;cursor:hand;' onKeyPress='return mask(this);' onBlur='check(this);' Value='" + tds[3].getElementsByTagName('INPUT')[0].value + "'>"

	tds[4].innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='Descriptions' maxlength=250 style='100%;border:1px solid gray;' Value='" + tds[4].getElementsByTagName('INPUT')[0].value + "'>"

	tds[5].innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='debits'  style='border:1px solid gray;' onblur='setPrice(this)'  onKeyPress='return onlyNumber(this);'  onfocus='this.value=txt2val(this.value);this.select()' Value='" + tds[5].getElementsByTagName('INPUT')[0].value + "' >"

	tds[6].innerHTML="<INPUT class='InvRowInput3' TYPE='text' NAME='credits' style='border:1px solid gray;' onblur='setPrice(this)'  onKeyPress='return onlyNumber(this);'  onfocus='this.value=txt2val(this.value);this.select()' Value='" + tds[6].getElementsByTagName('INPUT')[0].value + "' >"

	tds[3].getElementsByTagName('INPUT')[0].focus();

}

function removeRow(theRow){

	//theRow = mnuDiv.parentNode.parentNode.parentNode
	tds = theRow.getElementsByTagName('TD')

	//mnuCountr=0;
	//hideMenu();
	
	func = tds[1].getElementsByTagName('INPUT')[0];

	if (!func ){
		alert('«Ì‰ ”ÿ— —« ‰„Ì  Ê«‰Ìœ Õ–› ﬂ‰Ìœ.')
		return;
	}
	else if (func.value=='E' || func.value==''){
		alert('«Ì‰ ”ÿ— —« ‰„Ì  Ê«‰Ìœ Õ–› ﬂ‰Ìœ.')
		return;
	}
	theRow.parentNode.removeChild(theRow);

}

function voidRow(){
	theRow = mnuDiv.parentNode.parentNode.parentNode
	tds = theRow.getElementsByTagName('TD')

	mnuCountr=0;
	hideMenu();

	func = tds[1].getElementsByTagName('INPUT')[0];

	if (!func || func.value!=''){
		alert('«Ì‰ ”ÿ— —« ‰„Ì  Ê«‰Ìœ »«ÿ· ﬂ‰Ìœ.')
		return;
	}

	theRow.className="CompMemoEdit";

	tds[0].getElementsByTagName('DIV')[0].innerHTML="<IMG SRC='../images/s_delete.gif' WIDTH='20' HEIGHT='20' BORDER='0' ALT='”ÿ— »«ÿ· ‘œÂ' >";

	func.value="V"
	tds[1].innerHTML+="<div style='position:relative;'><div style='position:absolute;left:-610;top:-15;width:610;'><hr style='color:red;'></div></div>";
	
	tds[5].innerHTML = tds[5].getElementsByTagName('INPUT')[0].value + "<INPUT TYPE='hidden' name='debits' Value='"+tds[5].getElementsByTagName('INPUT')[0].value+"'>"

	tds[6].innerHTML = tds[6].getElementsByTagName('INPUT')[0].value + "<INPUT TYPE='hidden' name='credits' Value='"+tds[6].getElementsByTagName('INPUT')[0].value+"'>"

	memoIsTaraz()

}

//-->
</SCRIPT>
<%
end if

'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------

conn.Close
%>
</font>
<% if request("act")="" OR request("act")="editDoc" then%>

<script language="JavaScript">
<!--

function setCurrentRow(rowNo){
	if (rowNo == -1) rowNo=0;
	invTable=document.getElementById("GLrows");
	theTD=invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#F0F0F0');

	currentRow=rowNo;
	invTable=document.getElementById("GLrows");
	theTD=invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[0];
	theTD.setAttribute("bgColor", '#FFB0B0');
}
function delRow(rowNo){
	invTable=document.getElementById("GLrows");
	theRow=invTable.getElementsByTagName("tr")[rowNo];
	invTable.removeChild(theRow);

	rowsCount=document.getElementsByName("Accounts").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.bgColor= '#F0F0F0';
		tempTD.innerText= rowNo+1;
	}
}
function addRow(){
	rowNo = currentRow
	invTable=document.getElementById("GLrows");
	theRow=invTable.getElementsByTagName("tr")[rowNo];
	newRow=document.createElement("tr");
	newRow.setAttribute("bgColor", '#f0f0f0');
	newRow.setAttribute("onclick", theRow.getAttribute("onclick"));

	tempTD=document.createElement("td");
	tempTD.innerHTML=rowNo+1
	tempTD.setAttribute("align", 'center');
	tempTD.setAttribute("width", '26');
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.setAttribute("width", '60');
	tempTD.innerHTML="<INPUT class='InvRowInput'  TYPE='text' NAME='Accounts'   maxlength=6 onKeyPress='return mask(this);' onBlur='check(this);' style='width:100%;' >"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("width", '40');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='GLAccounts' maxlength=5 onKeyPress='return mask(this);' onBlur='check(this);' style='width:100%;' >"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("width", '300');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='Descriptions' style='width:100%;'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.setAttribute("width", '80');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='debits' style='width:100%;' onblur='setPrice(this)'  onKeyPress='return onlyNumber(this);'  onfocus='this.value=txt2val(this.value);this.select()'>"
	newRow.appendChild(tempTD);

	tempTD=document.createElement("td");
	tempTD.setAttribute("dir", 'LTR');
	tempTD.setAttribute("width", '80');
	tempTD.innerHTML="<INPUT class='InvRowInput2' TYPE='text' NAME='credits' style='width:100%;'  onblur='setPrice(this)' onKeyPress='return onlyNumber(this);'   onfocus='setCurrentRow(this.parentNode.parentNode.rowIndex);this.value=txt2val(this.value);this.select()'>"
	newRow.appendChild(tempTD);


	invTable.insertBefore(newRow,theRow);
	
	rowsCount=document.getElementsByName("Accounts").length;
	for (rowNo=0; rowNo < rowsCount ; rowNo++){
		tempTD=invTable.getElementsByTagName("tr")[rowNo].getElementsByTagName("td")[0]
		tempTD.bgColor= '#F0F0F0';
		tempTD.innerText= rowNo+1;
	}

	invTable.getElementsByTagName("tr")[currentRow].getElementsByTagName("td")[1].getElementsByTagName("Input")[0].focus();
//	document.all.dddd2.innerText=invTable.innerHTML
}

function saveDraft(){ 
	if (document.all.totalDebit.value != 0 || document.all.totalCredit.value != 0)
			IsTaraz = true
}

function checkValidation(){ 
	memoIsTaraz()
	if (IsTaraz==true)
		return true
	else{
		alert("Œÿ«! ”‰œ  —«“ ‰Ì” ")
		return false
	}
}

var dialogActive=false;

function mask(src){ 
	var theKey=event.keyCode;
	if (src.name=="Accounts"){
		if (theKey==32){
			event.keyCode=0;
			dialogActive=true;
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì  ›’Ì·Ì:"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false;
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('../ar/dialog_selectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:300px; dialogWidth:600px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					src.title=Arguments[1];
				}
			}
//				src.parentNode.nextSibling.getElementsByTagName("INPUT")[0].focus();
		}
		else if (theKey >= 48 && theKey <= 57 ) 
			return true;
		else
			return false;
	}
	else if (src.name=="GLAccounts"){
		if (theKey==32){
			event.keyCode=0;
			dialogActive=true;
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì œ› — ﬂ·:"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false;
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")
					src.value=Arguments[0];
					src.title=Arguments[1];
				}
			}
//				src.parentNode.nextSibling.getElementsByTagName("INPUT")[0].focus();
		}
		else if (theKey >= 48 && theKey <= 57 ) // [0]-[9] are acceptible
			return true;
		else
			return false;
	}
	else if (src.name=="debits" || src.name=="credits" || src.name=="Refs1"){
		if (theKey < 48 || theKey > 57) { // [0]-[9] are acceptible
			return false;
		}
		return true;
	}
	else if (src.name=="Descriptions" || src.name=="Refs2"){
		if (theKey==13){  // [Enter] 
			return false;
		}
		return true;
	}
}

function onlyNumber(src){ 
	var theKey=event.keyCode;
	if (theKey==13){  // [Enter] 
		return true;
	}
	else if (theKey < 48 || theKey > 57) { // 0-9 are acceptible
		return false;
	}
}

function areYouSureToExit()
{
	a= confirm("are you sure?")
	return a
}

function setPrice(src){
	src.value=val2txt(txt2val(src.value));
	//rowNo=src.parentNode.parentNode.rowIndex;
	if (src.name=="credits" && src.value!=0) 
		src.parentNode.previousSibling.getElementsByTagName('INPUT')[0].value = ""
	if (src.name=="debits" && src.value!=0) 
		src.parentNode.nextSibling.getElementsByTagName('INPUT')[0].value = ""
	if (src.value ==0) src.value=""
	if (src.value ==" ") src.value=""

	memoIsTaraz();

}

function check(src){ 
	if (src.name=="Accounts"){
		if (!dialogActive){
			if (window.XMLHttpRequest) {
			var objHTTP=new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		}
			objHTTP.open('GET','xml_CustomerAccount.asp?id='+src.value,false)
			objHTTP.send()
			tmpStr = unescape( objHTTP.responseText)
			src.title=tmpStr;
			if (tmpStr=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ"){
				src.value="";
			}
		}
	}
	else if(src.name=="GLAccounts"){
		if (!dialogActive){
			if (window.XMLHttpRequest) {
			var objHTTP=new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		}
			objHTTP.open('GET','xml_GLAccount.asp?id='+src.value,false)
			objHTTP.send()
			tmpStr = unescape( objHTTP.responseText)

			src.title=tmpStr;
			if (tmpStr=="ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ"){
				src.value="";
			}
		}
	}
	else if(src.name=="debits" || src.name=="credits"){
		src.value=val2txt(txt2val(src.value));
		if (''+src.value=="NaN" || src.value=="0") src.value = ""
		//rowNo=src.parentNode.parentNode.rowIndex;

		if (src.value!=""){
			if (src.name=="credits") 
				src.parentNode.previousSibling.getElementsByTagName('INPUT')[0].value = ""
			else
				src.parentNode.nextSibling.getElementsByTagName('INPUT')[0].value = ""
		}

		memoIsTaraz();
		
	}
}
function memoIsTaraz(){ 
	var totalCredit = 0;
	var totalDebit = 0;
	rowsCnt=document.getElementsByName("debits").length
	for (rowNo=0; rowNo < rowsCnt; rowNo++){
		if (document.getElementsByName("credits")[rowNo].type!='hidden')
		totalCredit += parseInt(txt2val(document.getElementsByName("credits")[rowNo].value));
		if (document.getElementsByName("debits")[rowNo].type!='hidden')
		totalDebit += parseInt(txt2val(document.getElementsByName("debits")[rowNo].value));
	}
	document.all.totalCredit.value = val2txt(totalCredit);
	document.all.totalDebit.value = val2txt(totalDebit);
	if (totalDebit == totalCredit && totalCredit != 0){
		IsTaraz = true
		document.all.tarazDiv.innerHTML = "<FONT COLOR='#008833'>”‰œ  —«“ «” </FONT>"
	}
	else{
		IsTaraz = false
		document.all.tarazDiv.innerHTML = "<FONT COLOR='#FF3300'>”‰œ  —«“ ‰Ì” </FONT>"
	}

}

setPrice(document.all.totalDebit)
setPrice(document.all.totalCredit)
if (parseInt(document.all.totalDebit.value)==0) IsTaraz=false;

//-->
</SCRIPT>

<%end if%>
<!--#include file="tah.asp" -->
