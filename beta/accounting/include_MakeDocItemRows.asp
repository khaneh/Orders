<%
	'
	'	This Include file is for use in MakeDoc.asp 
	'	only becase of improving readability of code.
	'
'	response.write sys & "," & sysDefaultGLAccount& "<br>"
Sub WriteRow()
  '-- Added By kid 830531  Emptying OTHER reasons
	if GLAccount="18001" then GLAccount=""
  '--

	if ""&Account="" then
		AccountClass="RepAccountInput2"
	else
		AccountClass="RepAccountInput"
	end if

	if ""&GLAccount="" then 'OR GLAccount="18001" then
		AccountGLClass="RepGLInput"
	else
		AccountGLClass="RepGLInput2"
	end if

%>					<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
						<td>
							<INPUT class="<%=AccountClass%>" maxlength="6" TYPE="text" NAME="Account<%=tempCounter%>" value="<%=Account%>"  style="width:30pt" <%=disabled%> readonly>
						</td>
						<td>
							<INPUT class="<%=AccountGLClass%>" maxlength="5" TYPE="text" NAME="GLAccount<%=tempCounter%>" value="<%=GLAccount%>" onKeyPress="return mask(this,true);" onblur="check(this,true)" <%=disabled%>>
						</td>
						<td align=rihgt><%=AccountTitle%></td>
						<td style="padding:2px;" width=200>
							<textarea readonly Name="Description<%=tempCounter%>" style="font-family:tahoma;overflow:auto;border:none; background:transparent; width:100%; height:25px; font-size:7pt;" <%=disabled%>><%=VoidText & LineDescription%></textarea>
						</td> 
						<td dir='LTR' align='right'>&nbsp;<%=Separate(Credit)%></td>
						<td dir='LTR' align='right'><%=Separate(Debit)%>
							<INPUT TYPE="hidden" NAME="IsCredit<%=tempCounter%>" Value="<%=not(IsCredit)%>">
							<INPUT TYPE="hidden" NAME="Amount<%=tempCounter%>" Value="<%=Amount%>">
							<INPUT TYPE="hidden" NAME="Ref1<%=tempCounter%>" Value="<%=Ref1%>">
							<INPUT TYPE="hidden" NAME="Ref2<%=tempCounter%>" Value="<%=Ref2%>"></td>
					</TR>
<%
End Sub
Sub WriteFirstRow()
  '-- Added By kid 830531  Emptying OTHER reasons
	if FirstGLAccount="18001" then FirstGLAccount=""
  '--
	if ""&Account="" then
		AccountClass="RepAccountInput"
	else
		AccountClass="RepAccountInput2"
	end if

	if ""&FirstGLAccount="" then 'OR FirstGLAccount="18001" then
		AccountGLClass="RepGLInput"
		FirstGLAccount=0
	else
		AccountGLClass="RepGLInput2"
	end if

%>					<TR class='<%if tempCounter MOD 2 = 0 then response.write "RepTR1" else response.write "RepTR2"%>'>
						<td rowspan=<%=ItemLines%> title="<%if disabled="disabled" then response.write "«Ì‰ ›«ﬂ Ê— «·›Ì »ÊœÂ° ·–« »«Ìœ ›«ﬂ Ê— ¬‰—« «Ê· ç«Å ﬂ‰Ìœ" else response.write sys end if%>">
							<INPUT TYPE="checkbox" NAME="Items" Value="<%=tempCounter%>" onclick="return selectRow(this.parentNode.parentNode);" <%=disabled%>>
							<INPUT TYPE="hidden" NAME="Item<%=tempCounter%>" Value="<%=ItemID%>" >
							<INPUT TYPE="hidden" NAME="Lines<%=tempCounter%>" Value="<%=ItemLines%>">
							<INPUT TYPE="hidden" NAME="Sys<%=tempCounter%>" Value="<%=sys%>"></td>
						<td rowspan='<%=ItemLines%>' dir='LTR' align='right' valign='top' <%if VoidText <> "" then response.write " bgcolor='#FFCCCC'"%>>
							<%=effectiveDate%><br><br>(<A style="font-size:7pt;" HREF="ShowItem.asp?sys=<%=sys%>&Item=<%=ItemID%>" target="_blank"><%=VoidText & Descriptions(ItemType)%></A>)
							<br><br><%=VoidLink%></td>
						<td class="RepTD1">
							<INPUT class="<%=AccountClass%>" maxlength="6" TYPE="text" NAME="Account<%=tempCounter%>" value="<%=Account%>"  style="width:30pt" <%=disabled%> raedonly>
							</td>
						<td>
							<INPUT class="<%=AccountGLClass%>" maxlength="5" TYPE="text" NAME="GLAccount<%=tempCounter%>" value="<%=FirstGLAccount%>" onKeyPress="return mask(this,false);" onblur="check(this,false)" <%=disabled%> <%if AccountGLClass="RepGLInput2" and FirstGLAccount<>0 and sys <>"AO" then response.write " readonly"%>>
							</td>
						<td><%=AccountTitle%></td>
						<td style="padding:2px;" width=200>
							<textarea readonly Name="Description<%=tempCounter%>" style="font-family:tahoma;overflow:auto;border:none; background:transparent; width:100%; height:25px; font-size:7pt;" <%=disabled%>><%=VoidText & LineDescription%></textarea>
						</td>
						<td dir='LTR' align='right'>&nbsp;<%=Separate(Debit)%></td>
						<td dir='LTR' align='right'><%=Separate(Credit) %>
							<INPUT TYPE="hidden" NAME="IsCredit<%=tempCounter%>" Value="<%=IsCredit%>">
							<INPUT TYPE="hidden" NAME="Amount<%=tempCounter%>" Value="<%=Amount%>">
							<INPUT TYPE="hidden" NAME="Ref1<%=tempCounter%>" Value="<%=Ref1%>">
							<INPUT TYPE="hidden" NAME="Ref2<%=tempCounter%>" Value="<%=Ref2%>"></td>
					</TR>
<%
End Sub
	
	mySQL="SELECT * FROM AXItemTypes"
	Set RS1 = conn.execute(mySQL)
	do while not RS1.eof
		Descriptions(RS1("ID"))=RS1("Name")
		Rs1.movenext
	loop

	'Changed by Kid 820828 :
	OrderBy = sys & "Items.EffectiveDate, "& sys & "Items.Type, "& sys & "Items.ID"

	AdditionalConditions =""

	if request.form("FromDate")<>"" then AdditionalConditions = AdditionalConditions & "AND ("& sys & "Items.EffectiveDate >= N'"& request.form("FromDate") & "') "
	if request.form("ToDate")<>"" then AdditionalConditions = AdditionalConditions & "AND ("& sys & "Items.EffectiveDate <= N'"& request.form("ToDate") & "') "
	if request.form("ItemTypes").count>0 then 
		AdditionalConditions = AdditionalConditions & "AND ("
		myOr=""
		for i=1 to request.form("ItemTypes").count
			AdditionalConditions = AdditionalConditions & myOr & sys & "Items.Type=" & request.form("ItemTypes")(i)
			myOR="OR "
		next
		AdditionalConditions = AdditionalConditions & ") "

		if ubound(ItemLinksArray)>-1 then 
			AdditionalConditions = AdditionalConditions & "AND ("& sys & "Items.Link IN ("& ItemLinks & ") ) "
		end if

	end if

'	Changed By Kid 830216:
'	Changed By Sam 871024:
	mySQL="SELECT "& sys & "Items.*, Accounts.AccountTitle, GLAccounts.Name AS GLAccountName, "& sys & "Memo.Description AS MmoDescription, "& sys & "Memo.Type AS MemoType,  ISNULL(Receipts.CashAmount, 0) AS RcpCash, ISNULL(Receipts.ChequeAmount, 0) AS RcpChq, Receipts.ChequeQtty AS ChqQtty,  ISNULL(Payments.CashAmount, 0) AS PayCash, ISNULL(Payments.ChequeAmount, 0) AS PayChq,  ReceivedCash.Description AS CashDescription, ISNULL(ISNULL(Invoices.Number, Vouchers.Number),'-') AS Number, Invoices.IsA FROM ReceivedCash RIGHT OUTER JOIN Receipts ON ReceivedCash.Receipt = Receipts.ID RIGHT OUTER JOIN "& sys & "Items  INNER JOIN Accounts ON "& sys & "Items.Account = Accounts.ID LEFT OUTER JOIN Invoices ON " & sys & "Items.Link = Invoices.ID LEFT OUTER JOIN Vouchers ON " & sys & "Items.Link = Vouchers.ID LEFT OUTER JOIN GLAccounts ON "& sys & "Items .GLAccount = GLAccounts.ID AND "& sys & "Items .GL = GLAccounts.GL LEFT OUTER JOIN Payments ON "& sys & "Items.Link = Payments.ID ON Receipts.ID = "& sys & "Items.Link LEFT OUTER JOIN "& sys & "Memo ON "& sys & "Items.Link = "& sys & "Memo.ID WHERE ("& sys & "Items.GL_Update = 1) AND ("& sys & "Items.GL = '"& OpenGL & "') "& AdditionalConditions & " ORDER BY "& OrderBy
'response.write(mySQL)
'response.end
	Set RS1 = conn.execute(mySQL)
	
	if Not (RS1.EOF) then
%>
		<TABLE class="RepTable" width='90%' align='center'>
		<TR>
			<TD colspan=8 dir='rtl' align='center'>
			</TD>
		</TR>
		<TR>
			<TD class="RepTableTitle" colspan=8 dir='rtl' align='center'>¬Ì „ Â«Ì “Ì—”Ì” „ <B><%=sysName%></B> ﬂÂ Â‰Ê“ ”‰œ Õ”«»œ«—Ì »—«Ì ¬‰Â« ’«œ— ‰‘œÂ «” .
			</TD>
		</TR>
		<TR class="RepTableHeader">
			<TD width=5><INPUT TYPE="checkbox" onclick="selectAll(this)"></TD>
			<TD width=60> «—ÌŒ</TD>
			<TD> ›÷Ì·Ì</TD>
			<TD width=30>„⁄Ì‰</TD>
			<TD width=160>‰«„ Õ”«»</TD>
			<TD width=200>‘—Õ</TD>
			<TD width=60>»œÂﬂ«—</TD>
			<TD width=60>»” «‰ﬂ«—</TD>
		</TR>
<%
		While Not (RS1.EOF)
			disabled=""
			thisLink=clng(RS1("Link"))
			for i=0 to ubound(ItemLinksArray)
				if clng(ItemLinksArray(i))=thisLink then
					ItemLinksArray(i)=0
					'exit for
				end if
			next
			tempCounter=tempCounter+1
			Amount=cdbl(RS1("AmountOriginal"))
			Vat	= cdbl(RS1("Vat"))
			Ref1=""
			Ref2=""
			if RS1("IsCredit")=True then
				IsCredit=True 
				Credit=Amount
				Debit=""
			else
				IsCredit=false 
				Credit=""
				Debit=Amount
			end if
			if RS1("Voided")=True then
				IsCredit = not (IsCredit)
				tmp=Debit
				Debit=Credit
				Credit=tmp
				VoidText = "«»ÿ«· "
				VoidLink = "<B><A href='SubsysDocsEdit.asp?act=find&link=" & sys & RS1("ID") & "' target='_blank'>”‰œ ﬁ»·Ì</A></B>"
			else
				VoidText = ""
				VoidLink = ""
			end if

			effectiveDate=	RS1("EffectiveDate")
			ItemID=			RS1("ID")
			ItemType =		RS1("Type")
			MemoType =		RS1("MemoType")
			Account =		RS1("Account")
			FirstGLAccount =RS1("firstGLAccount")
			DestGLAccount =	RS1("GLAccount")
			DestGLAccName =	RS1("GLAccountName")
			AccountTitle =	RS1("AccountTitle")

			ItemLines=2
			LineDescription="&nbsp;"
		'****
		'****
			if ItemType=3 AND MemoType=7 then
			'Type = 3 («⁄·«„ÌÂ)
			'MemoType = 7 («‰ ﬁ«·)
				if (NOT IsCredit and not RS1("Voided")) or (IsCredit and RS1("Voided")) then
					ItemLines=2
					LineDescription=RS1("MmoDescription")
					'--------------
					Call WriteFirstRow()
					'--------------
					mySQL="SELECT ToItemType,ToItemLink FROM InterMemoRelation WHERE (FromItemType='"& sys & "') AND (FromItemLink='"& RS1("Link") & "')"
					Set RS2=Conn.Execute(mySQL)
					if not rs2.eof then 
						ToSys=RS2("ToItemType")
						ToItemLink=RS2("ToItemLink")
					else
						response.write "<br>" & mySQL
						response.end
					end if
					mySQL="SELECT "& ToSys & "Items.FirstGLAccount, "& ToSys & "Items.Account, Accounts.AccountTitle, "& ToSys & "Memo.Description FROM "& ToSys & "Items INNER JOIN Accounts ON "& ToSys & "Items.Account = Accounts.ID INNER JOIN "& ToSys & "Memo ON "& ToSys & "Items.Link = "& ToSys & "Memo.ID WHERE ("& ToSys & "Items.Type = 3) AND ("& ToSys & "Items.Link = '"& ToItemLink & "')"
					Set RS2=Conn.Execute(mySQL)
					Account =		RS2("Account")
					GLAccount =		RS2("firstGLAccount")
					AccountTitle =	RS2("AccountTitle")
					LineDescription=RS2("Description")
					RS2.Close
					'--------------
					Call WriteRow()
					'--------------
				else
					' Type = 3 , MemoType = 7 , but IsCredit=True
					' When there is a "Transfrer Memo" which is a credit item (destination of transfer action).
					' In this case we don't show the memo. (because it is shown as "debit" somewhere else)
					tempCounter = tempCounter - 1
				end if
		'****
		'****
			elseif ItemType=2 AND left(DestGLAccount,2) = "12" then
			'Type = 2 (œ—Ì«› )
			'[Ê«—Ì“ »Â »«‰ﬂ]
				ItemType=10
				ItemLines=2
				LineDescription=RS1("CashDescription")
				'--------------
				Call WriteFirstRow()
				'--------------
				Account =		""
				GLAccount =		DestGLAccount
				AccountTitle =	DestGLAccName
				' -----------
				Call WriteRow()
				'------------
		'****
		'****
			elseif ItemType=2 then
			'Type = 2 (œ—Ì«› ) 
			'[œ— ’‰œÊﬁ]
				ReceiptID=	RS1("Link")
				RcpCash=	cdbl(RS1("RcpCash"))
				ChqQtty=	RS1("ChqQtty")
				ItemLines=ChqQtty+1
				if RcpCash<>0 then ItemLines = ItemLines +1

				LineDescription="œ—Ì«›  «“ " & AccountTitle & " ÿÌ —”Ìœ " & ReceiptID
				'--------------------------
				'--------------------------
				mySQL="select * from CashRegisterLines where Type=1 and voided=0 and link=" & ReceiptID
				set RS2=conn.Execute(mySQL)
				isA = CBool(RS2("isA"))
				RS2.close
				set RS2=nothing
				if sys="AR" then
					mySQL="SELECT ISNULL(CONVERT(tinyint, Invoices.IsA), 2) AS IsA, Accounts.IsADefault, ISNULL(InvoicePrintForms.ID,0) as invPrintForm  FROM Accounts INNER JOIN ARItems ON ARItems.Account = Accounts.ID LEFT OUTER JOIN ARItemsRelations ON ARItemsRelations.CreditARItem = ARItems.ID LEFT OUTER JOIN ARItems ARItems_2 ON ARItems_2.ID = ARItemsRelations.DebitARItem LEFT OUTER JOIN Invoices ON ARItems_2.Link = Invoices.ID left outer join InvoicePrintForms on ARItems_2.Link=InvoicePrintForms.InvoiceID and InvoicePrintForms.Voided=0 WHERE (ARItems.Type = 2) AND (ARItems.Link = '"& ReceiptID & "')"
					Set RS2=Conn.Execute(mySQL)
					if cint(RS2("IsA"))=1 and RS2("invPrintForm")=0 then disabled="disabled"
					Rs2.close
					Set Rs2=Nothing
				end if
				'--------------
				Call WriteFirstRow()
				'--------------
				
				if RcpCash<>0 then
					if IsCredit then
						Credit=RcpCash
						Debit=""
					else
						Credit=""
						Debit=RcpCash
					end if
					Amount = RcpCash
					Account = ""
					' NEW cash
					if IsA then
						GLAccount = sysCashGLAccountA
						AccountTitle = sysCashGLAccountAName
					else
						GLAccount = sysCashGLAccountB
						AccountTitle = sysCashGLAccountBName
					end if
					'GLAccount = DestGLAccount
					'AccountTitle = DestGLAccName
					LineDescription="œ—Ì«›  ÊÃÂ ‰ﬁœ «“ " & RS1("AccountTitle") & " ÿÌ —”Ìœ " & ReceiptID
					'--------------
					Call WriteRow()
					'--------------
				end if
				if ChqQtty <> 0 then
					Set RS2=Conn.Execute("SELECT * FROM ReceivedCheques WHERE (Receipt='"& ReceiptID& "')")
					do while not RS2.eof
						Amount=RS2("Amount")
						if IsCredit then
							Credit=Amount
							Debit=""
						else
							Credit=""
							Debit=Amount
						end if
						Account = ""
						' NEW cash
						if IsA then
							GLAccount = sysChequeGLAccountA
							AccountTitle = sysChequeGLAccountAName
						else
							GLAccount = sysChequeGLAccountB
							AccountTitle = sysChequeGLAccountBName
						end if
						'GLAccount = DestGLAccount
						'AccountTitle = DestGLAccName
						Ref1=RS2("ChequeNo")
						Ref2=RS2("ChequeDate")
						LineDescription="çﬂ "& Ref1 & " „Ê—Œ "& Ref2 & " "& RS2("BankOfOrigin") & " œ—Ì«› Ì ÿÌ —”Ìœ " & ReceiptID & " "& RS2("Description") & " " & RS1("AccountTitle")
						'--------------
						Call WriteRow()
						'--------------
						RS2.moveNext
					loop
					RS2.close
				end if
		'****
		'****
			elseif ItemType=5 then
			'Type = 5 (Å—œ«Œ ) 
				PaymentID=	RS1("Link")
				PayCash=	cdbl(RS1("PayCash"))
				PayChq=		cdbl(RS1("PayChq"))
				IsA=false
			  if NOT(PayCash=0 AND PayChq=0) then
				ItemLines=1
				LineDescription="Å—œ«Œ  »Â " & AccountTitle & " ÿÌ —”Ìœ " & PaymentID

				if PayCash<>0 then ItemLines = ItemLines + 1

				if PayChq=0 then
				' Only Cash
					'--------------
					Call WriteFirstRow()
					'--------------
					if IsCredit then
						Credit=PayCash
						Debit=""
					else
						Credit=""
						Debit=PayCash
					end if
					Amount =		PayCash
					Account =		""
					GLAccount =		DestGLAccount
					AccountTitle =	DestGLAccName

					LineDescription="Å—œ«Œ  ÊÃÂ ‰ﬁœ »Â " & RS1("AccountTitle") & " ÿÌ —”Ìœ " & PaymentID
					'--------------
					Call WriteRow()
					'--------------
				else
				' Some Cheques + (maybe) chash
					Set RS2 = Server.CreateObject("ADODB.Recordset")
					RS2.CursorLocation=3 'Because in ADOVBS_INC adUseClient=3
					'CheqStatus = 21 means: paid cheque is still not passed
					mySQL="SELECT PaidCheques.ChequeNo, PaidCheques.ChequeDate, PaidCheques.Description, PaidCheques.Amount, BnkGLAcc.GLAccount, GLAccounts.Name AS GLAccName FROM (SELECT Banker, GLAccount, GL FROM BankerCheqStatusGLAccountRelation WHERE (GL = "& openGL & ") AND (CheqStatus = 21)) BnkGLAcc INNER JOIN GLAccounts ON BnkGLAcc.GL = GLAccounts.GL AND BnkGLAcc.GLAccount = GLAccounts.ID RIGHT OUTER JOIN PaidCheques ON BnkGLAcc.Banker = PaidCheques.Banker WHERE (PaidCheques.Payment = "& PaymentID & ")"
					RS2.Open mySQL ,Conn,3	'3=adOpenStatic
					PayChqQtty = RS2.RecordCount
					ItemLines = ItemLines + PayChqQtty 
					'--------------
					Call WriteFirstRow()
					'--------------
					if PayCash<>0 then
						if IsCredit then
							Credit=PayCash
							Debit=""
						else
							Credit=""
							Debit=PayCash
						end if
						Amount =		PayCash
						Account =		""
						GLAccount =		DestGLAccount
						AccountTitle =	DestGLAccName

						LineDescription="Å—œ«Œ  ÊÃÂ ‰ﬁœ »Â " & RS1("AccountTitle") & " ÿÌ —”Ìœ " & PaymentID
						'--------------
						Call WriteRow()
						'--------------
					end if

					Do While NOT RS2.eof
						Amount=RS2("Amount")
						if IsCredit then
							Credit=Amount
							Debit=""
						else
							Credit=""
							Debit=Amount
						end if

						Ref1=			RS2("ChequeNo")
						Ref2=			RS2("ChequeDate")
						Account =		""
						GLAccount =		RS2("GLAccount")
						AccountTitle =	RS2("GLAccName")

						LineDescription="çﬂ "& Ref1 & " „Ê—Œ "& Ref2 & " œ— ÊÃÂ " & RS1("AccountTitle") & " "& RS2("Description") & " " 
						'--------------
						Call WriteRow()
						'--------------
						RS2.moveNext
					Loop
					RS2.close
					Set RS2=Nothing
				end if

			  else
				response.write "<br><B>»Â ‰Ÿ— „Ì —”œ ﬂÂ «‘ »«Â ‘œÂ »«‘œ</B>"
			  end if
		'****
		'****
			else
			' Other Item Types
			' (Invoice / RevInvoice / Non-Transfer Memo)
				if Vat > 0 then 
					ItemLines = 3
				else 
					ItemLines = 2
				end if
	
				LineDescription=Descriptions(ItemType)
				select case ItemType
					case 1:
						number =  RS1("Number")
						if number = "" then 
							number = RS1("Link") & " ”Ì” „"
							if RS1("IsA") then disabled="disabled"
						end if
						if cdbl(RS1("GLAccount")) = 13003 then 
							LineDescription="»Â«Ì ›«ﬂ Ê— ›—Ê‘ ‘„«—Â "& number 
						else
							LineDescription="»Â«Ì ›«ﬂ Ê— ›—Ê‘ ‘„«—Â "& number & " »Â " & RS1("AccountTitle") 
						end if
						VatDescription = "„«·Ì«  »— «—“‘ «›“ÊœÂ ›«ﬂ Ê— ›—Ê‘ ‘„«—Â "& number & " «“  " & RS1("AccountTitle")
					case 3:
						LineDescription=RS1("MmoDescription")
					case 4:
						number =  RS1("Number")
						if number = "" then 
							number = RS1("Link") & " ”Ì” „"
						end if
						LineDescription="»—ê‘  «“ ›—Ê‘ »Â "& RS1("AccountTitle") & " ÿÌ ›«ﬂ Ê— »—ê‘  ‘„«—Â "& number
						VatDescription = "„«·Ì«  »—ê‘ Ì »«»  ›«ﬂ Ê— "& number & " »Â " & RS1("AccountTitle")
					case 5:
						LineDescription="Å—œ«Œ  »Â "& RS1("AccountTitle") & " ÿÌ —”Ìœ ‘„«—Â "& RS1("Link") 
						VatDescription = ""
					case 6:
						number = RS1("Number")
						if number = "-" then
							number = RS1("Link") & " ”Ì” „"
						end if
						'response.write(RS1("Account"))
						'response.end
						LineDescription = "Œ—Ìœ «“ "& RS1("AccountTitle") & " ÿÌ ›«ﬂ Ê— ‘„«—Â "& number
						VatDescription = "„«·Ì«  Å—œ«Œ Ì »«»  Œ—Ìœ «“ " & RS1("AccountTitle") & " ÿÌ ›«ﬂ Ê— " & number
				end select

				'--------------
				if ItemType = 6 then 
					LineDescription = "»Â«Ì ›«ﬂ Ê— Œ—Ìœ "& number
					set rs = Conn.Execute("select count(id) as cnt from VoucherLines where Voucher_ID=" & rs1("link"))
					ItemLines = 1 + CInt(rs("cnt"))
					if Vat > 0 then ItemLines = ItemLines + 1
				else
					if Vat > 0 then
						ItemLines = 3
					else
						ItemLines = 2
					end if
				end if
				

				Call WriteFirstRow()
				'--------------
				'tmpAccount = Account
				Account = ""
				GLAccount = RS1("GLAccount")
				AccountTitle = RS1("GLAccountName")
				Amount = Amount - Vat
				if isCredit = 0 then 
					Debit = Debit - Vat
				else 
					Credit = Credit - Vat
				end if
				' -----------
				if ItemType = 6 then 
					'LineDescription = "Œ—Ìœ «“ "& RS1("AccountTitle") & " ÿÌ ›«ﬂ Ê— ‘„«—Â "& number
					set rs=Conn.Execute("select VoucherLines.*,InventoryItems.Unit from VoucherLines left outer join PurchaseOrders on VoucherLines.RelatedPurchaseOrderID=PurchaseOrders.id left outer join InventoryItems on PurchaseOrders.TypeID=InventoryItems.ID and PurchaseOrders.IsService=0 where Voucher_ID=" & rs1("link"))
					while not rs.eof
						LineDescription = "Œ—Ìœ " & rs("qtty") & " "  
						if not IsNull(rs("unit")) then LineDescription = LineDescription & rs("unit") & " "
						LineDescription = LineDescription & RS("lineTitle") & " ÿÌ ›«ﬂ Ê— ‘„«—Â "& number
						Credit=rs("price")
						Debit=""
						Amount=rs("price")
						Call WriteRow()
						rs.moveNext	
					wend
					rs.close
					set rs=nothing
				else
					Call WriteRow()
				end if
				
				'------------ SAM
				'Account = tmpAccount
				if Vat > 0 then 
					Amount = Vat
					LineDescription = VatDescription
					if isCredit = 0 then 
						Debit = Vat
					else 
						Credit = Vat
					end if
					GLAccount = sysVatAccount
					AccountTitle = sysVatAccountName
					'------------
					call WriteRow()
					'------------
				end if
			end if

			RS1.MoveNext
		Wend
%>
			<TR>
				<TD class="RepTableFooter" colspan='8' align=center>&nbsp;<INPUT style="font-family:tahoma;border:1 solid black; width:50px;" type="submit" value="«ÌÃ«œ">&nbsp;</td>
			</TR>
		</TABLE>
		<br>
<%	end if%>
