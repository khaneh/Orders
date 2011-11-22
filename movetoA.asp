<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
Response.CharSet = "windows-1256"

conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=jame;Data Source=.\sqlexpress;PWD=5tgb;"
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 180
conn.open conStr
conn.CommandTimeout = 120 

AppBgColor = "#BBBBBB"	'Other:"#99AACC"
AppFgColor =  "#C3DBEB" '"#DEEBD9"
SelectedMenuColor = "#0E5499"
UnSelectedMenuColor = "#309261"
SelectedSubMenuColor = "#DEEBD9" '"#C3DBEB" 
UnSelectedSubMenuColor = "#0E5499" ' "#609250"
TabWidth=65
ImgTabSelected="/images/tab-1.gif"
ImgTabNotSelected="/images/tab-2.gif"

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<TITLE> تبديل فاكتور به الف </TITLE>
<style>
	body { font-family: tahoma; font-size: 8pt;}
	body A { Text-Decoration : none ;}
	Input { font-family: tahoma; font-size: 9pt;}
	td { font-family: tahoma; font-size: 8pt;}
	.tt { font-family: tahoma; font-size: 10pt; color:yellow;}
	.tt2 { font-family: tahoma; font-size: 8pt; color:yellow;}
	.inputBut { font-family: tahoma; font-size: 8pt; richness: 10}
	.t7pt { font-size: 8pt;}
	.t8pt { font-size: 10pt;}
	.alak a { color: #cccccc; text-decoration: none;  font-size: 10pt;}
	.alak2 a { color: black; text-decoration: none;  font-size: 10pt; }
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; height:20px; }
</style>
</HEAD>
<BODY bgcolor=<%=AppBgColor%> topmargin=0 leftmargin=0>
<script src="/js/jquery-1.7.min.js"></script>
<%
if request("act")="" or request("act")="check" then
	if request("act")="check" then
		invoiceID=request("invoice")
		issuedDate=request("issuedDate")
		set rs = Conn.Execute("select *,cast(hasvat as int)*.04*price as thisVat from InvoiceLines where invoice=" & invoiceID)
		if rs.eof then
			response.write("bad Invoice Number!")
			response.end
		end if
		TotalPrice=0
		TotalDiscount=0
		TotalReverse=0
		TotalReceivable=0
		TotalVat=0
		hasRFD=false
		while not rs.eof
			if rs("item")="39999" then 
				hasRFD=true
				rfdID=rs("id")
			end if
			thePrice = CDbl(rs("Price"))
			theDiscount = CDbl(rs("Discount"))
			theReverse = CDbl(rs("Reverse"))
			theVat = CDbl(rs("thisVat"))
			TotalPrice	=		TotalPrice + thePrice
			TotalDiscount =		TotalDiscount + theDiscount
			TotalReverse =		TotalReverse + theReverse
			TotalReceivable =	TotalReceivable + (thePrice - theDiscount - theReverse + theVat)
			TotalVat =			TotalVat + theVat
			response.write("vat for item " & rs("item") & " is: " & theVat & "<br>")
			rs.moveNext
		wend
		rs.close
		RFD = TotalReceivable - fix(TotalReceivable / 1000) * 1000
		TotalReceivable = TotalReceivable - RFD
		TotalDiscount = TotalDiscount + RFD
		response.write("RFD is: " & RFD & "<br>")
		response.write("TotalReceivable is: " & TotalReceivable & "<br>")
		response.write("TotalDiscount is: " & TotalDiscount & "<br>")
		set rs=Conn.Execute("select * from arItems where type=1 and reason=1 and link=" & invoiceID)
		if rs.eof then 
			response.write("It's seems that invoice not issued!")
			response.end
		end if
		originalDate=rs("effectiveDate")
		arID=CLng(rs("id"))
		oldRemain=cdbl(rs("remainedAmount"))
		oldAmountOrigin=CDbl(rs("amountOriginal"))
		if (oldRemain+TotalReceivable-oldAmountOrigin) > 0 then 
			fullyApplied=0
			response.write("This Invoice don't fully applied!<br>")
			response.write("Remained: " & (oldRemain+TotalReceivable-oldAmountOrigin) & "<br>")
		else
			fulluApplied=1
			response.write("This Invoice fully applied<br>")
		end if
		rs.close
		set rs=Conn.Execute("select * from effectiveGlrows where sys='AR' and link in (select id from arItems where type=1 and reason=1 and link=" & invoiceID & ")")
		if rs.eof then 
			response.write("this Invoice has not GL Row!")
			response.end
		end if
		glAccount91002=false
		glAccount13003=false
		glAccount49010=false
		other=false
		while not rs.eof
			if rs("glAccount")="91002" then 
				glAccount91002=true
			elseif rs("glAccount")="13003" then 
				glAccount13003=true
			elseif rs("glAccount")="49010" or rs("glAccount")="49009" then
				glAccount49010=true
			else
				other=true
			end if
			
			glDocID=rs("glDocID")
			rs.moveNext
		wend
		if not glAccount91002 then 
			response.write("this invoice not has 91002 in GL Row, please check " & glDocID & " doc.<br>")
			response.end
		end if
		if not glAccount13003 then 
			response.write("this invoice not has 13003 in GL Row, please check " & glDocID & " doc.<br>")
			response.end
		end if
		if not glAccount49010 then 
			response.write("this invoice already has Vat!<br>")
		end if
		if other then 
			response.write("There is other gl account not in 13003, 49010,91002. i'm confused!")
			response.end
		end if
	end if
%>
<form action="?act=<% if request("act")="check" then response.write("submit") else response.write("check")%>" method="post">
	<table>
		<tr>
			<td>invoice id</td>
			<td><input type="text" name="invoice" value="<%=request("invoice")%>"></td>
		</tr>
		<tr>
			<td>issued date</td>
			<td><input type="text" name="issuedDate" value="<%=request("issuedDate")%>"></td>
		</tr>
		<tr>
			<td colspan="2">
				<input id="submitBT" type="submit" value="<% if request("act")="check" then response.write("submit") else response.write("check")%>">
			</td>
		</tr>
	</table>
</form>
<%
elseif request("act")="submit" then
	invoiceID=request("invoice")
	issuedDate=request("issuedDate")
	set rs = Conn.Execute("select *,cast(hasvat as int)*.04*price as thisVat from InvoiceLines where invoice=" & invoiceID)
	TotalPrice=0
	TotalDiscount=0
	TotalReverse=0
	TotalReceivable=0
	TotalVat=0
	hasRFD=false
	while not rs.eof
		if rs("item")="39999" then 
			hasRFD=true
			rfdID=rs("id")
		else
			thePrice = CDbl(rs("Price"))
			theDiscount = CDbl(rs("Discount"))
			theReverse = CDbl(rs("Reverse"))
			theVat = CDbl(rs("thisVat"))
			TotalPrice	=		TotalPrice + thePrice
			TotalDiscount =		TotalDiscount + theDiscount
			TotalReverse =		TotalReverse + theReverse
			TotalReceivable =	TotalReceivable + (thePrice - theDiscount - theReverse + theVat)
			TotalVat =			TotalVat + theVat
			Conn.Execute("update invoiceLines set vat=" & theVat & " where id=" & rs("id"))
		end if
		rs.moveNext
	wend
	rs.close
	RFD = TotalReceivable - fix(TotalReceivable / 1000) * 1000
	TotalReceivable = TotalReceivable - RFD
	TotalDiscount = TotalDiscount + RFD
	if hasRFD then 
		Conn.Execute("update invoiceLines set discount=" & RFD & " where id=" & rfdID)
	elseif RFD >0 then 
		Conn.Execute("INSERT INTO jame.dbo.InvoiceLines	(Invoice, Item, Description, Length, Width, Qtty, Sets, AppQtty, Price, Discount, Reverse, Vat, hasVat) VALUES (" & invoiceID & ", 39999, N'تخفيف رند فاكتور', 0, 0, 0, 0, 0, 0, " & RFD & ", 0, 0, 1);")
	end if
	set rs=Conn.Execute("select * from arItems where type=1 and reason=1 and link=" & invoiceID)
	originalDate=rs("effectiveDate")
	arID=CLng(rs("id"))
	oldRemain=cdbl(rs("remainedAmount"))
	oldAmountOrigin=CDbl(rs("amountOriginal"))
	if (oldRemain+TotalReceivable-oldAmountOrigin) > 0 then 
		fullyApplied=0
	else
		fulluApplied=1
	end if
	Conn.Execute("update Invoices set totalVat=" & totalVat & " ,TotalDiscount=" & TotalDiscount & " ,TotalReceivable=" & TotalReceivable & ", issuedDate='" &issuedDate & "', isA=1,originalIssuedDate='" & originalDate & "' where id=" & invoiceID)
	Conn.execute("update arItems set fullyApplied=" & fullyApplied & ", amountOriginal=" & TotalReceivable & ", effectiveDate='" & issuedDate & "',remainedAmount=" & oldRemain+TotalReceivable-oldAmountOrigin & ", vat=" & totalVat & " where id=" & arID)
	rs.close
	set rs=Conn.Execute("select * from effectiveGlrows where sys='AR' and link in (select id from arItems where type=1 and reason=1 and link=" & invoiceID & ")")
	while not rs.eof
		if rs("glAccount")="91002" then 
			Conn.Execute("update glRows set glAccount=91001, amount=" & TotalReceivable - totalVat & " where id=" & rs("id"))
		elseif rs("glAccount")="13003" then 
			Conn.Execute("update glRows set amount=" & TotalReceivable & " where id=" & rs("id"))
		elseif rs("glAccount")="49010" or rs("glAccount")="49009" then
			Conn.Execute("update glRows set glAccount=49010, amount=" & totalVat & "where id=" & rs("id"))
			hasVat=true
		end if
		
		glDoc=rs("glDoc")
		desc=replace(rs("description"),"بهاي","")
		rs.moveNext
	wend
	if not hasVat then 
		Conn.Execute("insert into glRows (GLDoc, GLAccount, Amount, Description, SYS, Link, IsCredit, deleted) VALUES (" & glDoc & ", 49010, " & totalVat & ", 'ماليات بر ارزش افزوده' + '" & desc & "', N'AR', " & arID & ", 1, 0)")
	end if
	response.write ("it's Done :) <a href='movetoA.asp'>Next</a>")
end if
%>
</body>