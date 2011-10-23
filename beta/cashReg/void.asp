<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<%
'CashRegister (9)
act = request("act")

if act = "voidReceipt" then
	ON ERROR RESUME NEXT
		ReceiptID = clng(request("Receipt"))
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("‘„«—Â œ—Ì«›  „⁄ »— ‰Ì” .")
		end if
	ON ERROR GOTO 0

	'----------
	' Findng Receipt 
	'----------
	mySQL = "SELECT ISNULL(CashAmount, 0) AS CashAmount, ISNULL(ChequeAmount, 0) AS ChequeAmount, ChequeQtty, SYS, Voided FROM Receipts WHERE (ID = '"& ReceiptID & "')"
	Set RS1= conn.Execute(mySQL)

	if RS1.eof then
		conn.close
		response.redirect "top.asp?errmsg=" & Server.URLEncode("Œÿ«! <br> ç‰Ì‰ œ—Ì«› Ì ÊÃÊœ ‰œ«—œ.")
	else
		VoidDate=		ShamsiToday()
		Voided=			RS1("Voided")
		CashAmount=		cdbl(RS1("CashAmount"))
		ChequeAmount=	cdbl(RS1("ChequeAmount"))
		ChequeQtty=		RS1("ChequeQtty")
		SYS=			RS1("SYS")
		RS1.close

		if Voided then
			conn.close
			response.redirect "../"& SYS& "/AccountReport.asp?act=showReceipt&receipt="& ReceiptID & "&errmsg=" & Server.URLEncode("Œÿ«!Õ–› «Ì‰ œ—Ì«›  „„ﬂ‰ ‰Ì” <br><br>ﬁ»·« »«ÿ· ‘œÂ «” .")
		end if
'		mySQL="SELECT CashRegisterLines.ID AS CashRegLine, ISNULL(Receipts.CashAmount, 0) AS CashAmount, ISNULL(Receipts.ChequeAmount, 0) AS ChequeAmount, Receipts.ChequeQtty, Receipts.SYS, CashRegisterLines.CashReg, CashRegisterLines.Voided, CashRegisters.IsOpen, CashRegisters.Cashier FROM CashRegisters INNER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg INNER JOIN Receipts ON CashRegisterLines.Link = Receipts.ID WHERE (CashRegisterLines.Type = 1) AND (CashRegisterLines.Link = '"& ReceiptID & "')"
'		Changed By kid 821228

		'----------
		'	Finding the CashRegister Line for this Receipt
		'	(Type=1 means Receipt in CashRegisterLineTypes)
		'----------
		mySQL = "SELECT CashRegisterLines.ID AS CashRegLine, CashRegisterLines.isA, CashRegisterLines.CashReg, CashRegisters.IsOpen, CashRegisterLines.Voided, CashRegisters.Cashier FROM CashRegisters INNER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg WHERE (CashRegisterLines.Type = 1) AND (CashRegisterLines.Link = '"& ReceiptID & "')"
		Set RS1= conn.Execute(mySQL)
		if NOT RS1.eof then
			' There is a CashRegister Line for this Receipt
			HasCashRegLine =	true
			CashRegLine =		RS1("CashRegLine")
			CashReg =			RS1("CashReg")
			IsOpen =			RS1("IsOpen")
			Cashier =			RS1("Cashier")
			isA =				RS1("isA")
		else
			' There is NOT a CashRegister Line for this Receipt
			HasCashRegLine =	false
		end if
		RS1.close
		'------------Check cash sitll open before void receipt
		if not isOpen then 
			conn.close
			response.redirect "top.asp?errmsg=" & Server.URLEncode("«Ì‰ œ—Ì«›  „ ⁄·ﬁ »Â ’‰œÊﬁÌ «”  ﬂÂ ﬁ»·« »” Â ‘œÂ!<br>Å” ‰„Ìù Ê«‰Ìœ ¬‰—« »«ÿ· ﬂ‰Ìœ")
		end if

		if NOT ( Auth(9 , 7) OR (IsOpen AND Cashier=session("ID")) ) then
			' Doesn't Have the Priviledge to VOID the RECEIPT/PAYMENT 
			conn.close
			response.redirect "top.asp?errmsg=" & Server.URLEncode("Œÿ«!<br>‘„« „ÃÊ“ ·«“„ »—«Ì «»ÿ«· «Ì‰ œ—Ì«›  —« ‰œ«—Ìœ.")
		end if

	  '**************************** Voiding (AR/AP/AO)Item of Receipt  ****************
	  '***
		'*********  Finding the (AR/AP/AO)Item of Receipt 
		'*** Type = 2 means (AR/AP/AO)Item is a Receipt
		mySQL="SELECT ID, Account FROM "& SYS & "Items WHERE (Type=2) AND (Link = '"& ReceiptID & "')"
		Set RS1=conn.Execute(mySQL)
		ItemID = RS1("ID")
		Account = RS1("Account")

		'*********  Finding other (AR/AP/AO)Items related to this Item
		mySQL="SELECT ID AS RelationID, Debit"& SYS & "Item, Amount FROM "& SYS & "ItemsRelations WHERE (Credit"& SYS & "Item = '"& ItemID & "')"
		Set RS1=conn.Execute(mySQL)
		do while not (RS1.eof)
			'*********  Adding back the amount in the relation, to the debit (AR/AP/AO)Item ...
			conn.Execute("UPDATE "& SYS & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Debit"& SYS & "Item") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM "& SYS & "ItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		loop

		'*********  Voiding the (AR/AP/AO)Item
		conn.Execute("UPDATE "& SYS & "Items SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& ItemID & "')")

		'**************************************************************
		'*			Decreasing Account's (AR/AP/AO) Balance  
		'*			('because we have removed a CREDIT Item )
		'**************************************************************
		mySQL="UPDATE Accounts SET "& SYS & "Balance = "& SYS & "Balance - '"& CashAmount + ChequeAmount & "'  WHERE (ID ='"& Account & "')"
		conn.Execute(mySQL)
		
	  '***
	  '***---------------- End of Voiding (AR/AP/AO)Item  ----------------
		

		' ######################################################### 
		'				SET ALL RECEIVED CHEQUES STATUS 
		'				TO  RETURNED
		' #########################################################
		' Status = 4  means Returned
		mySQL="UPDATE ReceivedCheques SET LastStatus = 4 , LastUpdatedDate =N'"& VoidDate & "', LastUpdatedBy = '"& session("ID") & "' WHERE (Receipt = '"& ReceiptID & "')"
		conn.Execute(mySQL)

		' ######################################################### 
		'					VOID the RECEIPT ...
		' #########################################################
		mySQL="UPDATE Receipts SET Voided=1,VoidedBy='"& session("ID") & "',VoidedDate=N'"& VoidDate & "' WHERE (ID = '"& ReceiptID & "')"
		conn.Execute(mySQL)

		if HasCashRegLine then
			' ######################################################### 
			'					VOID the CASH REG LINE  ...
			' #########################################################
			mySQL="UPDATE CashRegisterLines SET Voided = 1 WHERE (ID = '"& CashRegLine & "')"
			conn.Execute(mySQL)
				' ######################################################### 
				'				CHANGING CASH REGISTER BALANCE 
				'				(decreasing)
				' #########################################################
				if isA then 
					mySQL="UPDATE CashRegisters SET CashAmountA = CashAmountA - '"& CashAmount & "', ChequeAmount = ChequeAmount - '"& ChequeAmount & "', ChequeQtty = ChequeQtty - '"& ChequeQtty & "' WHERE (ID ='"& CashReg & "')"
				else
					mySQL="UPDATE CashRegisters SET CashAmountB = CashAmountB - '"& CashAmount & "', ChequeAmount = ChequeAmount - '"& ChequeAmount & "', ChequeQtty = ChequeQtty - '"& ChequeQtty & "' WHERE (ID ='"& CashReg & "')"
				end if
				conn.Execute(mySQL)
		end if
	end if

	conn.close

	if HasCashRegLine then
		response.redirect "CashRegReport.asp"
	else
		response.redirect "../"& SYS & "/AccountReport.asp?act=showReceipt&receipt=" & ReceiptID
	end if

elseif act = "voidPayment" then

	ON ERROR RESUME NEXT
		PaymentID = clng(request("Payment"))
		if Err.Number<>0 then
			Err.clear
			conn.close
			response.redirect "top.asp?errMsg=" & Server.URLEncode("‘„«—Â Å—œ«Œ  „⁄ »— ‰Ì” .")
		end if
	ON ERROR GOTO 0

	'----------
	' Findng Payment 
	'----------
	mySQL = "SELECT ISNULL(CashAmount, 0) AS CashAmount, ISNULL(ChequeAmount, 0) AS ChequeAmount, SYS, Voided FROM Payments WHERE (ID = '"& PaymentID & "')"
	Set RS1= conn.Execute(mySQL)

	if RS1.eof then
		conn.close
		response.redirect "top.asp?errmsg=" & Server.URLEncode("Œÿ«! <br> ç‰Ì‰ Å—œ«Œ Ì ÊÃÊœ ‰œ«—œ.")
	else
		VoidDate=		ShamsiToday()
		Voided=			RS1("Voided")
		CashAmount=		cdbl(RS1("CashAmount"))
		ChequeAmount=	cdbl(RS1("ChequeAmount"))
		SYS=			RS1("SYS")
		RS1.close

		if Voided then
			conn.close
			response.redirect "../"& SYS& "/AccountReport.asp?act=showPayment&payment="& PaymentID & "&errmsg=" & Server.URLEncode("Œÿ«!Õ–› «Ì‰ Å—œ«Œ  „„ﬂ‰ ‰Ì” <br><br>ﬁ»·« »«ÿ· ‘œÂ «” .")
		end if
		'----------
		'	Finding the CashRegister Line for this Payment
		'	(Type=2 means Payment in CashRegisterLineTypes)
		'----------
		mySQL = "SELECT CashRegisterLines.ID AS CashRegLine,CashRegisterLines.isA, CashRegisterLines.CashReg, CashRegisters.IsOpen, CashRegisterLines.Voided, CashRegisters.Cashier FROM CashRegisters INNER JOIN CashRegisterLines ON CashRegisters.ID = CashRegisterLines.CashReg WHERE (CashRegisterLines.Type = 2) AND (CashRegisterLines.Link = '"& PaymentID & "')"
		Set RS1= conn.Execute(mySQL)

		if NOT RS1.eof then
			' There is a CashRegister Line for this Payment
			HasCashRegLine =	true
			CashRegLine =		RS1("CashRegLine")
			CashReg =			RS1("CashReg")
			IsOpen =			RS1("IsOpen")
			Cashier =			RS1("Cashier")
			isA =				RS1("isA")
		else
			' There is NOT a CashRegister Line for this Payment
			HasCashRegLine =	false
		end if
		RS1.close
		'------------Check cash sitll open before void payment
		if not isOpen then 
			conn.close
			response.redirect "top.asp?errmsg=" & Server.URLEncode("«Ì‰ Å—œ«Œ  „ ⁄·ﬁ »Â ’‰œÊﬁÌ «”  ﬂÂ ﬁ»·« »” Â ‘œÂ!<br>Å” ‰„Ìù Ê«‰Ìœ ¬‰—« »«ÿ· ﬂ‰Ìœ")
		end if

		if NOT ( Auth(9 , 7) OR (IsOpen AND Cashier=session("ID")) ) then
			' Doesn't Have the Priviledge to VOID the RECEIPT/PAYMENT 
			conn.close
			response.redirect "top.asp?errmsg=" & Server.URLEncode("Œÿ«!<br>‘„« „ÃÊ“ ·«“„ »—«Ì «»ÿ«· «Ì‰ Å—œ«Œ  —« ‰œ«—Ìœ.")
		end if

	  '**************************** Voiding (AR/AP/AO)Item of Payment  ****************
	  '***
		'*********  Finding the (AR/AP/AO)Item of Payment 
		'*** Type = 5 means (AR/AP/AO)Item is a Payment
		mySQL="SELECT ID, Account FROM "& SYS & "Items WHERE (Type=5) AND (Link = '"& PaymentID & "')"
		Set RS1=conn.Execute(mySQL)
		ItemID = RS1("ID")
		Account = RS1("Account")

		'*********  Finding other (AR/AP/AO)Items related to this Item
		mySQL="SELECT ID AS RelationID, Credit"& SYS & "Item, Amount FROM "& SYS & "ItemsRelations WHERE (Debit"& SYS & "Item = '"& ItemID & "')"
		Set RS1=conn.Execute(mySQL)
		do while not (RS1.eof)
			'*********  Adding back the amount in the relation, to the debit (AR/AP/AO)Item ...
			conn.Execute("UPDATE "& SYS & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Credit"& SYS & "Item") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM "& SYS & "ItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		loop

	  '******************  Voiding the (AR/AP/AO)Item	*****************
	  '***
		conn.Execute("UPDATE "& SYS & "Items SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& ItemID & "')")

		'**************************************************************
		'*			Increasing Account's (AR/AP/AO) Balance  
		'*			('because we have removed a DEBIT Item )
		'**************************************************************
		mySQL="UPDATE Accounts SET "& SYS & "Balance = "& SYS & "Balance + '"& CashAmount + ChequeAmount & "'  WHERE (ID ='"& Account & "')"
		conn.Execute(mySQL)
		
	  '***
	  '***---------------- End of Voiding (AR/AP/AO)Item ----------------
		

		' ######################################################### 
		'					VOID the PAYMENT ...
		' #########################################################
		mySQL="UPDATE Payments SET Voided=1,VoidedBy='"& session("ID") & "',VoidedDate=N'"& VoidDate & "' WHERE (ID = '"& PaymentID & "')"
		conn.Execute(mySQL)

		if HasCashRegLine then
			' ######################################################### 
			'					VOID the CASH REG LINE  ...
			' #########################################################
			mySQL="UPDATE CashRegisterLines SET Voided=1 WHERE (ID = '"& CashRegLine & "')"
			conn.Execute(mySQL)
				' ######################################################### 
				'				CHANGING CASH REGISTER BALANCE 
				'				(increasing)
				' #########################################################
				if isA then 
					mySQL="UPDATE CashRegisters SET CashAmountA = CashAmountA + "& CashAmount & " WHERE (ID ="& CashReg & ")"
				else
					mySQL="UPDATE CashRegisters SET CashAmountB = CashAmountB + "& CashAmount & " WHERE (ID ="& CashReg & ")"
				end if
				'mySQL="UPDATE CashRegisters SET CashAmount = CashAmount + '"& CashAmount & "' WHERE (ID ='"& CashReg & "')"
				conn.Execute(mySQL)
		end if
	end if

	conn.close

	if HasCashRegLine then
		response.redirect "CashRegReport.asp"
	else
		response.redirect "../"& SYS & "/AccountReport.asp?act=showPayment&payment=" & PaymentID
	end if
else
	conn.close
	response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ«.")
end if
%> 