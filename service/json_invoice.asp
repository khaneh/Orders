<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset = "utf-8"
	Response.CodePage = 65001
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
select case request("act")
	'---------------------------------------------------------------------------------------------
	case "updateLineDesc": '----------------------------------------------------------------------
	'---------------------------------------------------------------------------------------------
		set j = jsObject()
		id = cdbl(request.form("id"))
		Conn.Execute("update InvoiceLines set description=N'" & Request.form("desc") & "' where id=" & id)
		set rs=Conn.Execute("select * from InvoiceLines where id=" & id)
		j("desc")=rs("description")
		rs.close
		set rs=nothing
		'---------------------------------------------------------------------------------------------
	case "updateIsA": '-------------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------
		set j = jsObject()
		id = cdbl(request.form("id"))
		isA = CInt(Request.form("isa"))
		set rs = Conn.Execute("select * from invoices where id=" & id)
		isIssued = CBool(rs("issued"))
		isApproved = CBool(rs("approved"))
		orgIsA = CBool(rs("isA"))
		hasVat = false
		if CDbl(rs("totalVat"))>0 then hasVat = true
		errMsg=""
		if isIssued then 
			if auth(6,"N") then 
				set rs = Conn.Execute("select * from EffectiveGLRows where SYS='AR' and link in (select id from ARItems where link=" & id & " and type=1)")
				if not rs.eof then 
					if isA=1 then 
						if not hasVat then errMsg = "این فاکتور سند حسابداری دارد، لطفا ابتدا این فاکتور را از سند خارج کنید"
					else
						if hasVat then errMsg = "این فاکتور سند حسابداری دارد، لطفا ابتدا این فاکتور را از سند خارج کنید"
					end if
				end if
			else
				errMsg = "شما مجاز به تغییر نیستید!"
			end if
		elseif isApproved then 
			if not auth(6, "M") then errMsg = "شما مجاز به تغییر نیستید!"
		end if
		if errMsg="" then 
			select case isA
				case 0:
					Conn.Execute("update Invoices set isa=0 where id=" & id)
					Conn.Execute("update InvoiceLines set vat=0 where invoice=" & id)
				case 1:
					Conn.Execute("update Invoices set isa=1 where id=" & id)
					Conn.Execute("update InvoiceLines set vat=price * " & session("VatRate") & "/100 where hasVat=1 and invoice=" & id)
					
				case -1:
					Conn.Execute("update InvoiceLines set vat=0 where invoice=" & id)
					Conn.Execute("update Invoices set isa=1 where id=" & id)
	
			end select
			set rs=Conn.Execute("select * from InvoiceLines where invoice=" & id)
			totalReceivable = 0
			totalDiscount = 0
			totalReverse = 0
			totalVat = 0
			rfdID = 0
			while not rs.eof
				if rs("item")<>"39999" then 
					totalVat = totalVat + CDbl(rs("vat"))
					totalReverse = totalReverse + CDbl(rs("reverse"))
					totalDiscount = totalDiscount + CDbl(rs("discount"))
					totalReceivable = totalReceivable + CDbl(rs("price")) + CDbl(rs("vat")) + CDbl(rs("reverse")) + CDbl(rs("discount"))
				else
					rfdID=rs("id")
				end if
				rs.MoveNext
			wend		
			RFD = TotalReceivable - fix(TotalReceivable / 1000) * 1000
			TotalReceivable = TotalReceivable - RFD
			TotalDiscount = TotalDiscount + RFD
			if RFD > 0 then
				if rfdID>0 then 
					Conn.Execute("update InvoiceLines set Discount = " & RFD & " where id=" & rfdID)
				else
					mySQL="INSERT INTO InvoiceLines (Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat) VALUES ('"& id & "', '39999', N'تخفيف رند فاكتور', '0', '0', '0', '0', '0', '0', '" & RFD & "', '0', '0')"
					conn.Execute(mySQL)
				end if
				else
					if rfdID>0 then Conn.Execute("delete InvoiceLines where id=" & rfdID)
			end if
			Conn.Execute("update Invoices set totalVat= " & totalVat & ",totalReceivable=" & totalReceivable & ",totalDiscount=" & totalDiscount & " where id=" & id)
			set rs=Conn.Execute("select * from Invoices where id=" & id)
			j("isa")=rs("isa")
		end if
		j("errMsg")=errMsg
		rs.close
		set rs=nothing
		'---------------------------------------------------------------------------------------------
	case "updateNo": '--------------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------
		set j = jsObject()
		id = cdbl(request.form("id"))
		set rs=Conn.Execute("select * from Invoices where id=" & id)
		if CBool(rs("isA")) then 
			Conn.Execute("update Invoices set number=N'" & Request.form("no") & "' where id=" & id)
			set rs=Conn.Execute("select * from Invoices where id=" & id)
		end if
		j("no")=rs("number")
		rs.close
		set rs=nothing
		'---------------------------------------------------------------------------------------------
	case "approveInvoice": '--------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------
		set j = jsObject()
		j("err")=0
		if not Auth(6 , "C") then		
			j("err")=1
			j("msg")="شما مجاز به تاييد فاكتور نيستيد"
		else
			InvoiceID=request("id")
			if not(isnumeric(request("invoice"))) then
				j("err")=1
				j("msg")="خطا"
			else
				mySQL="SELECT Invoices.*,orders.isApproved FROM Invoices inner join InvoiceOrderRelations on Invoices.ID=InvoiceOrderRelations.Invoice inner join orders on InvoiceOrderRelations.[Order] = orders.id WHERE Invoices.ID="& InvoiceID
				Set RS1 = conn.Execute(mySQL)
				if not RS1.eof then
					if RS1("Voided") = True then
						j("err")=1
						j("msg")="اين فاكتور قبلا باطل شده است."
					elseif RS1("Issued") = True then
						j("err")=1
						j("msg")="اين فاكتور قبلا صادر شده است."
					elseif RS1("Approved") = True then
						j("err")=1
						j("msg")="اين فاكتور قبلا تاييد شده است."
					elseif RS1("isApproved") = False then
						j("err")=1
						j("msg")="لطفا سفارش را تاييد كنيد"
					end if
					mySQL = "SELECT COUNT(Invoice) AS OrderCount FROM (SELECT DISTINCT InvoiceOrderRelations.Invoice FROM InvoiceOrderRelations inner join Invoices on InvoiceOrderRelations.Invoice = Invoices.ID WHERE InvoiceOrderRelations.[Order] IN (SELECT [Order] FROM InvoiceOrderRelations WHERE Invoice=204133) and Invoices.Voided=0) tbl"
					set RS1= conn.Execute(mySQL)
					if RS1("OrderCount")>1 then 
						j("err")=1
						j("msg")="به اين سفارش دو فاكتور مربوط شده! لطفا به هر سفارش فقط يك فاكتور متصل كنيد."
					end if
					if j("err")<>1 then
						mySQL="UPDATE Invoices SET Approved=1, ApprovedDate=N'"& shamsiToday() & "', ApprovedBy='"& session("ID") & "' WHERE (ID='"& InvoiceID & "')"
						conn.Execute(mySQL)
						j("msg")="سفارش تاييد شد"
					end if
				end if
			end if
		end if
		conn.close
		'---------------------------------------------------------------------------------------------
	case "removeApprove": '---------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------
		set j = jsObject()
		j("err")=0
		if not Auth(6 , "G") then		
			j("err")=1
			j("msg")="شما مجاز به برداشتن تاييد فاكتور نيستيد"
		else
			InvoiceID=request("id")
			if not(isnumeric(request("invoice"))) then
				j("err")=1
				j("msg")="خطا"
			else
				mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
				Set RS1 = conn.Execute(mySQL)
				if not RS1.eof then
					if RS1("Voided") = True then
						j("err")=1
						j("msg")="اين فاكتور قبلا باطل شده است."
					elseif RS1("Issued") = True then
						j("err")=1
						j("msg")="اين فاكتور قبلا صادر شده است."
					elseif RS1("Approved") = False then
						j("err")=1
						j("msg")="اين فاكتور قبلا تاييد نشده است."
					end if
					if j("err")<>1 then
						mySQL="UPDATE Invoices SET Approved=0, ApprovedDate=null, ApprovedBy=null WHERE (ID='"& InvoiceID & "')"
						conn.Execute(mySQL)
						j("msg")="سفارش از تاييد خارج شد"
					end if
				end if
			end if
		end if
		conn.close
		'---------------------------------------------------------------------------------------------
	case "issueInvoice": '----------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------	
		set j = jsObject()
		j("err")=0
		creationDate=	shamsiToday() 
		issueDate=		SqlSafe(request("issueDate"))
		if not auth(6,"I") and issueDate="" then issueDate=creationDate
		if not Auth(6 , "D") then		
			j("err")=1
			j("msg")="شما مجاز به صدور فاكتور نيستيد"
		else
			InvoiceID=request("id")
			if not(isnumeric(request("invoice"))) then
				j("err")=1
				j("msg")="خطا"
			elseif auth(6,"I") and issueDate="" then
				j("err")=1
				j("msg")="تاريخ وارد نشده"
			elseif Not CheckDateFormat(issueDate) Then
				j("err")=1
				j("msg")="تاريخ به درستي وارد نشده"
			elseif (issueDate < session("OpenGLStartDate")) OR (issueDate > session("OpenGLEndDate")) then
				j("err")=1
				j("msg")="خطا!<br>تاريخ وارد شده معتبر نيست. <br>(در سال مالي جاري نيست)"
			elseif session("IsClosed")="True" then
				j("err")=1
				j("msg")="خطا! سال مالي جاري بسته شده و شما قادر به تغيير در آن نيستيد."
			else
				mySQL="SELECT Invoices.*,orders.isApproved FROM Invoices inner join InvoiceOrderRelations on Invoices.ID=InvoiceOrderRelations.Invoice inner join orders on InvoiceOrderRelations.[Order] = orders.id WHERE Invoices.ID="& InvoiceID
				Set RS1 = conn.Execute(mySQL)
				if not RS1.eof then
					if RS1("Voided") = True then
						j("err")=1
						j("msg")="اين فاكتور قبلا باطل شده است."
					elseif RS1("Issued") = True then
						j("err")=1
						j("msg")="اين فاكتور قبلا صادر شده است."
					elseif RS1("isApproved") = False then
						j("err")=1
						j("msg")="لطفا سفارش را تاييد كنيد"
					end if
					if j("err")<>1 then
						mySQL="UPDATE Invoices SET Issued=1, IssuedDate=N'"& issueDate & "', IssuedBy='"& session("ID") & "' WHERE (ID='"& InvoiceID & "')"
						conn.Execute(mySQL)
						if rs1("isReverse") then
							isCredit=1
							itemType=4 
						else
							isCredit=0
							itemType=1
						'----------------------- Declaring the related orders as closed --------------
						conn.Execute("UPDATE Orders SET isClosed=1 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice='" & InvoiceID & "'))")
						end if
				
						'**************************** Creating ARItem for Invoice / Reverse Invoice  ****************
						'*** Type = 1 means ARItem is an Invoice
						'*** Type = 4 means ARItem is a Reverse Invoice
					
						firstGLAccount=	"13003"	'This must be changed... (Business Debitors)
						if IsA then
							GLAccount=	"91001"	'This must be changed... (Sales A)
						else
							GLAccount=	"91002"	'This must be changed... (Sales B)
						end if
						'					
						mySQL="INSERT INTO ARItems (GLAccount, GL, FirstGLAccount, Account, EffectiveDate, IsCredit, Type, Link, AmountOriginal, CreatedDate, CreatedBy, RemainedAmount, Vat) VALUES ('" &_
						GLAccount & "', '"& OpenGL & "', '"& firstGLAccount & "', '"& RS1("Customer") & "', N'"& issueDate & "', '"& isCredit & "', '"& itemType & "', '"& InvoiceID & "', '"& invoiceFee & "', N'"& creationDate & "', '"& session("ID") & "', '"& invoiceFee & "', '" & Vat & "')"
						conn.Execute(mySQL)
					
						if isReverse then
							'*** ATTENTION: Increasing AR Balance ....
							mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& invoiceFee & "' WHERE (ID='"& RS1("Customer") & "')"
						else
							'*** ATTENTION: Decreasing AR Balance ....
							mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& invoiceFee & "' WHERE (ID='"& RS1("Customer") & "')"
						end if
						conn.Execute(mySQL)
						j("msg")="فاکتور صادر شد"
					end if
				end if
			end if
		end if
		conn.close
		'---------------------------------------------------------------------------------------------
	case "removeIssue": '-----------------------------------------------------------------------------
		'---------------------------------------------------------------------------------------------
		set j = jsObject()
		j("err")=0
		if not Auth(6 , "F") then		
			'Doesn't have the Priviledge to VOID the Invoice 
			j("err")=1
			j("msg")="شما مجاز به خروج فاكتور از صدور نيستيد"
		end if
		comment=sqlSafe(request("comment"))
		InvoiceID=request("id")
		if InvoiceID="" or not(isnumeric(InvoiceID)) then
			j("err")=1
			j("msg")="خطا در شماره فاكتور"
		end if
		InvoiceID=clng(InvoiceID)
		set rs = Conn.Execute("select * from EffectiveGLRows where SYS='AR' and link in (select id from ARItems where link=" & InvoiceID & " and type=1)")
		if not rs.EOF then
			j("err")=1
			j("msg")="این فاکتور سند حسابداری دارد، ابتدا آنرا از سند خارج کنید"
		end if
		mySQL="SELECT * FROM Invoices WHERE (ID='"& InvoiceID & "')"
		Set RS1 = conn.Execute(mySQL)
		if RS1.eof then
			j("err")=1
			j("msg")="خطا! شماره فاكتور پيدا نشد!!!"
		else
			voided=			RS1("Voided")
			issued=			RS1("Issued")
			issuedBy=		RS1("IssuedBy")
			isReverse=		RS1("IsReverse")
			customerID=		RS1("Customer")
			invoiceFee=		RS1("TotalReceivable")
			IsA =			RS1("IsA")
			if voided then
				j("err")=1
				j("msg")="اين فاكتور قبلا در تاريخ <span dir='LTR'>"& RS1("VoidedDate") & "</span> باطل شده است."
			end if
			if j("err")<>1 then 
				mySQL="UPDATE Invoices SET  Issued=0, IssuedDate=null, IssuedBy=null WHERE (ID='"& InvoiceID & "')"
				conn.Execute(mySQL)
				if isReverse then
					isCredit=1
					itemType=4 
					itemTypeName="فاكتور برگشت از فروش"
				else
					isCredit=0
					itemType=1
					itemTypeName="فاكتور"
					'---------- Declaring the related orders as Open  -------------------
					'Changed By Kid ! 840509 
					'set orders which are ONLY related to this invoice, "Open"
					'that means, orders which are related to this invoice and are NOT related to any OTHER issued invoices.
					mySQL ="UPDATE Orders SET isClosed=0 WHERE ID IN (SELECT [Order] FROM InvoiceOrderRelations WHERE (Invoice = '" & InvoiceID & "') AND ([Order] NOT IN (SELECT InvoiceOrderRelations.[ORDER] FROM Invoices INNER JOIN InvoiceOrderRelations ON Invoices.ID = InvoiceOrderRelations.Invoice WHERE (Invoices.Issued = 1) AND (Invoices.Voided = 0) AND (Invoices.isReverse = 0) AND (Invoices.ID <> '" & InvoiceID & "'))))"
					conn.Execute(mySQL)
				end if
		
				'**************************** Voiding ARItem of Invoice / Reverse Invoice ****************
				'*** Type = 1 means ARItem is an Invoice
				'*** Type = 4 means ARItem is a Reverse Invoice
				'***
				'*********  Finding the ARItem of Invoice / Reverse Invoice
				mySQL="SELECT ID FROM ARItems WHERE (Type = '"& itemType & "') AND (Link='"& InvoiceID & "')"
				Set RS1=conn.Execute(mySQL)
				voidedARItem=RS1("ID")
				'*********  Finding other ARItems related to this Item
				if isReverse then
					mySQL="SELECT ID AS RelationID, DebitARItem, Amount FROM ARItemsRelations WHERE (CreditARItem = '"& voidedARItem & "')"
					Set RS1=conn.Execute(mySQL)
					Do While not (RS1.eof)
						'*********  Adding back the amount in the relation, to the credit ARItem ...
						conn.Execute("UPDATE ARItems SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("DebitARItem") & "')")
				
						'*********  Deleting the relation
						conn.Execute("DELETE FROM ARItemsRelations WHERE ID='"& RS1("RelationID") & "'")
						
						RS1.movenext
					Loop
				else
					mySQL="SELECT ID AS RelationID, CreditARItem, Amount FROM ARItemsRelations WHERE (DebitARItem = '"& voidedARItem & "')"
					Set RS1=conn.Execute(mySQL)
					Do While not (RS1.eof)
						'*********  Adding back the amount in the relation, to the credit ARItem ...
						conn.Execute("UPDATE ARItems SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("CreditARItem") & "')")
				
						'*********  Deleting the relation
						conn.Execute("DELETE FROM ARItemsRelations WHERE ID='"& RS1("RelationID") & "'")
						
						RS1.movenext
					Loop
				end if
				
				'*********  Voiding ARItem 
				conn.Execute("UPDATE ARItems SET RemainedAmount=0, FullyApplied=0, Voided=1 WHERE (ID = '"& voidedARItem & "')")
				
				'**************************************************************
				'*				Affecting Account's AR Balance  
				'**************************************************************
				if isReverse then
					mySQL="UPDATE Accounts SET ARBalance = ARBalance - '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
				else
					mySQL="UPDATE Accounts SET ARBalance = ARBalance + '"& invoiceFee & "' WHERE (ID='"& CustomerID & "')"
				end if
				
				conn.Execute(mySQL)
				
				'***
				'***---------------- End of  Voiding ARItem of Invoice / Reverse Invoice ----------------
				
				' Sending a Message to Issuer ...
				if trim(comment)<>"" then comment = chr(13) & chr(10) & "[" & comment & "]"
				MsgTo			=	issuedBy
				msgTitle		=	"Invoice Voided"
				msgBody			=	"فاكتور فوق توسط "& session("CSRName") & " از صدور خارج شد." & comment
				RelatedTable	=	"invoices"
				relatedID		=	invoiceID
				replyTo			=	0
				IsReply			=	0
				urgent			=	1
				MsgFrom			=	session("ID")
				MsgDate			=	shamsiToday()
				MsgTime			=	currentTime10()
				Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")
				' Copying the PreInvoice Data...
				j("msg") = "فاكتور از صدور خارج شد"
		end if
	end if
	
end select
Response.Write toJSON(j)

%>