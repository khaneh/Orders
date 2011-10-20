<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<%
if request("act")="VoidMemo" then
	Sys=request("Sys")
	MemoID=request("Memo")
	notAllowed=true
	if sys="AR" then
		if Auth(6 , "H") then ' Has the Priviledge to VOID the AR_MEMO
			notAllowed=false
		end if
	elseif sys="AP" then
		if Auth(7 , 8 ) then ' Has the Priviledge to VOID the AP_MEMO
			notAllowed=false
		end if
	elseif sys="AO" then
		if Auth("B" , 6 ) then ' Has the Priviledge to VOID the AO_MEMO
			notAllowed=false
		end if
	else
		conn.close 
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ« œ— ‰Ê⁄ “Ì—”Ì” „")
	end if

	if notAllowed then
		response.redirect "top.asp?errMsg=" & Server.URLEncode("‘„« „Ã«“ »Â «»ÿ«· «⁄·«„ÌÂ ‰Ì” Ìœ")
	end if

	if MemoID="" or not(isnumeric(MemoID)) then
		conn.close 
		response.redirect "top.asp?errMsg=" & Server.URLEncode("Œÿ« œ— ‘„«—Â «⁄·«„ÌÂ")
	end if
	MemoID=clng(MemoID)

	mySQL="SELECT * FROM "& sys & "Memo WHERE (ID='"& MemoID & "')"
	Set RS1 = conn.Execute(mySQL)
	if RS1.eof then
		conn.close 
		response.redirect "top.asp?errMsg=" & Server.URLEncode("ÅÌœ« ‰‘œ ")
	else
		voided=			RS1("Voided")
		voidedDate=		RS1("VoidedDate")
		Account=		RS1("Account")
		isCredit=		RS1("isCredit")
		Amount=			RS1("Amount")
		memoType=		RS1("Type")

		if voided then
		'-- Memo already voided
			conn.close 
			response.redirect "../"& sys & "/AccountReport.asp?act=showMemo&sys="& sys & "&memo="& MemoID & "&errMsg=" & Server.URLEncode("«Ì‰ «⁄·«„ÌÂ ﬁ»·«  œ—  «—ÌŒ <span dir='LTR'>"& voidedDate & "</span> »«ÿ· ‘œÂ «” ")
		elseif memoType=8 then
		'-- This memo is a part of a Compound Memo and cannot be voided
			conn.close 
			response.redirect "../"& sys & "/AccountReport.asp?act=showMemo&sys="& sys & "&memo="& MemoID & "&errMsg=" & Server.URLEncode("«Ì‰ «⁄·«„ÌÂ »Œ‘Ì «“ Ìﬂ ”‰œ „—ﬂ» «” .<BR> »—«Ì «»ÿ«· ¬‰ »Â »Œ‘ ÊÌ—«Ì‘ ”‰œ „—ﬂ» „—«Ã⁄Â ﬂ‰Ìœ.")
		end if

	end if

	voidDate = shamsiToday()

	' ----
	' Voiding Memo

	call voidMemo

	' End of voiding memo
	' ----

	message=""

  if memoType=7 then ' It's a TRANSFER memo --> the other side must also be voided.
	
	message="&msg=" & Server.URLEncode("«Ì‰ Ìﬂ «⁄·«„ÌÂ «‰ ﬁ«· »Êœ ﬂÂ »«ÿ· ‘œ.<br><br>")

	firstMemoID	= MemoID
	firstSYS	= sys
	if isCredit then
		mySQL="SELECT FromItemType, FromItemLink FROM InterMemoRelation WHERE (ToItemType = '"& sys & "') AND (ToItemLink = '"& MemoID & "')"
		Set RS1=conn.Execute(mySQL)
		MemoID	= RS1("FromItemLink")
		sys		= RS1("FromItemType")
	else
		mySQL="SELECT ToItemType, ToItemLink FROM InterMemoRelation WHERE (FromItemType = '"& sys & "') AND (FromItemLink = '"& MemoID & "')"
		Set RS1=conn.Execute(mySQL)
		MemoID	= RS1("ToItemLink")
		sys		= RS1("ToItemType")
	end if
	RS1.close()
	
	select case sys
	case "AR":
		subSys=" «“ ”Ì” „ ›—Ê‘ "
	case "AP":
		subSys=" «“ ”Ì” „ Œ—Ìœ "
	case "AO":
		subSys=" «“ ”«Ì— "
	end select 
	message= message & Server.URLEncode("»Â Œ«ÿ— «»ÿ«· «Ì‰ «⁄·«„ÌÂ° «⁄·«„ÌÂ <a target='_blank' href='" & "../"& sys & "/AccountReport.asp?act=showMemo&sys="& sys & "&Memo="& MemoID &"'>"& MemoID &"</a> <b>"& subSys & "</b>‰Ì“ »«Ìœ »«ÿ· „Ì ‘œ.<br>")

	mySQL="SELECT * FROM "& sys & "Memo WHERE (ID='"& MemoID & "')"
	Set RS1 = conn.Execute(mySQL)
	if not RS1.eof then
		voided=			RS1("Voided")
		voidedDate=		RS1("VoidedDate")
		Account=		RS1("Account")
		isCredit=		RS1("isCredit")
		Amount=			RS1("Amount")
		memoType=		RS1("Type")
		RS1.close()

		if voided then
			conn.close 
			response.redirect "../"& sys & "/AccountReport.asp?act=showMemo&sys="& sys & "&memo="& MemoID & "&errMsg=" & Server.URLEncode("«Ì‰ «⁄·«„ÌÂ ﬁ»·«  œ—  «—ÌŒ <span dir='LTR'>"& voidedDate & "</span> »«ÿ· ‘œÂ «” .")
		end if

		' ----
		' Voiding Related Memo

		call voidMemo

		' End of voiding related memo
		' ----

		message= message & Server.URLEncode("<br>«⁄·«„ÌÂ œÊ„ Â„ »«ÿ· ‘œ.")
	else
		' Other side Memo Not Found
		RS1.close()
		message= message & Server.URLEncode("<br><FONT COLOR='red'>Ê·Ì «⁄·«„ÌÂ „—»ÊÿÂ ÅÌœ« ‰‘œ.</FONT>")
	end if

'   We should NOT delete the relation between VOIDED memos ... they are not DELETED, just VOIDED
'	
'	if isCredit then
'		mySQL="DELETE FROM InterMemoRelation WHERE (ToItemType = '"& sys & "') AND (ToItemLink = '"& MemoID & "')"
'		conn.Execute(mySQL)
'	else
'		mySQL="DELETE FROM InterMemoRelation WHERE (FromItemType = '"& sys & "') AND (FromItemLink = '"& MemoID & "')"
'		conn.Execute(mySQL)
'	end if

	MemoID =	firstMemoID
	sys  =		firstSYS
  end if

	response.redirect "../"& sys & "/AccountReport.asp?act=showMemo&sys="& sys & "&Memo="& MemoID & message
end if

sub voidMemo

	mySQL="UPDATE "& sys & "Memo SET Voided=1, VoidedDate=N'"& voidDate & "', VoidedBy='"& session("ID") & "' WHERE (ID='"& MemoID & "')"
	conn.Execute(mySQL)

  '**************************** Voiding A*Item of the MEMO ****************
  '*** Type = 3 means the Item is a Memo
  '***
	'*********  Finding the Item of the Memo
	mySQL="SELECT ID FROM "& sys & "Items WHERE (Type = 3) AND (Link='"& MemoID & "')"
	Set RS1=conn.Execute(mySQL)
	voidedItem=RS1("ID")
	'*********  Finding other Items related to this Item
	if isCredit then
		mySQL="SELECT ID AS RelationID, Debit"& sys & "Item, Amount FROM "& sys & "ItemsRelations WHERE (Credit"& sys & "Item = '"& voidedItem & "')"
		Set RS1=conn.Execute(mySQL)
		Do While not (RS1.eof)
			'*********  Adding back the amount in the relation, to the debit Item ...
			conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Debit"& sys & "Item") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM "& sys & "ItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		Loop
	else
		mySQL="SELECT ID AS RelationID, Credit"& sys & "Item, Amount FROM "& sys & "ItemsRelations WHERE (Debit"& sys & "Item = '"& voidedItem & "')"
		Set RS1=conn.Execute(mySQL)
		Do While not (RS1.eof)
			'*********  Adding back the amount in the relation, to the credit Item ...
			conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Credit"& sys & "Item") & "')")

			'*********  Deleting the relation
			conn.Execute("DELETE FROM "& sys & "ItemsRelations WHERE ID='"& RS1("RelationID") & "'")
			
			RS1.movenext
		Loop
	end if

	'*********  Voiding A*Item 
	conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& voidedItem & "')")

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
	'***---------------- End of  Voiding A*Item of Memo ----------------

	' ----
end sub
conn.Close
%>