<%

'---------------------------------------------
'---------------------------------- Link2Trace
'---------------------------------------------
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/orderEdit.asp?e=n&radif="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

%>
<style>
	.GenTable { font-family:tahoma; font-size: 9pt; border:0; padding:0; }
	.GenInput { font-family:tahoma; font-size: 9pt; border: none; text-align:right; direction: LTR;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.GrayOutLine {border: 1px solid gray; }
	.RelInput { font-family:tahoma; font-size: 8pt; border: none; text-align:right; direction: LTR; background: transparent;}

</style>
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ 2 Select Account
'-----------------------------------------------------------------------------------------------------
if request("act")="submitsearch" then
	if request("CustomerNameSearchBox") <> "" then
		SA_TitleOrName=request("CustomerNameSearchBox")
		SA_Action="return true;"
		SA_SearchAgainURL="ItemsRelation.asp"
		SA_StepText="&nbsp;ê«„ œÊ„ : «‰ Œ«» Õ”«»"
%>
		<FORM METHOD=POST ACTION="ItemsRelation.asp?act=relate">
		<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	end if
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Relates to Items
'-----------------------------------------------------------------------------------------------------
elseif request("act")="relate" then

	TopNo = "" '"TOP 20"

	dim typeNamesArray(6) 
	typeNamesArray(1) = "›«ﬂ Ê—"
	typeNamesArray(2) = "œ—Ì«› "
	typeNamesArray(3) = "”‰œ"
	typeNamesArray(4) = "»—ê‘ "
	typeNamesArray(5) = "Å—œ«Œ "
	typeNamesArray(6) = "›«ﬂ Ê— Œ—Ìœ"


	if request("selectedCustomer") <> "" then
	Customer = request("selectedCustomer")
	sys = request("sys")
	if sys = "" then sys = "AR"
	if not fixSys = "-"  then sys = fixSys
%>
	<br>
	<CENTER>
	<% if fixSys = "-"  then %>
	<FORM METHOD="POST" ACTION="ItemsRelation.asp?act=relate">
	<INPUT TYPE="hidden" name="selectedCustomer" value="<%=Customer%>">
		<INPUT TYPE="radio" NAME="sys" value="AR" <% if sys="AR" then %>checked <% end if %> onclick="submit()"> ›—Ê‘
		<INPUT TYPE="radio" NAME="sys" value="AP" <% if sys="AP" then %>checked <% end if %> onclick="submit()"> Œ—Ìœ
		<INPUT TYPE="radio" NAME="sys" value="AO" <% if sys="AO" then %>checked <% end if %> onclick="submit()"> ”«Ì—
	</FORM>
	<% else 
		response.write "<br><br>" 
		sys = fixSys 
	end if
	%>
	</CENTER>
	<FORM METHOD="POST" ACTION="ItemsRelation.asp?act=submitRelation">
	<INPUT TYPE="hidden" name="sys" value="<%=sys%>">
	<INPUT TYPE="hidden" name="Customer" value="<%=Customer%>">
	<TABLE class="GenTable" align="center" cellspacing="2" cellpadding="0" dir="RTL">
	<tbody id="MainTable">
	<%
	
	mySQL="SELECT * From Accounts WHERE (ID='"& Customer & "')"
	Set RS1 = conn.Execute(mySQL)
	if (RS1.eof) then
		call showAlert( "Õ”«» ÊÃÊœ ‰œ«—œ." , CONST_MSG_ERROR)
		response.end
	else
		AccountTitle=RS1("AccountTitle")
	end if

	noJavaScript=false

%>
	<tr bgcolor='#DDDDFF' height="50">
		<td align='center' colspan="3" style="font-weight:bold; font-size:9pt;border:1 solid black;">„—»Êÿ »Â <br><%=AccountTitle%>
			<input type="hidden" id="selectedCreditItem" value="">
			<input type="hidden" id="selectedDebitItem" value="">		
		</td>
	</tr>
	<tr height="10">
		<td align='center' colspan="3"></td>
	</tr>
	<tr height="30">
		<td bgColor="#FFC0C0" style="border:1px solid black;" width="300">&nbsp;&nbsp;&nbsp; <b>»œÂﬂ«—Ì</b> 
		<td width="50" rowspan="4" onclick="unselect();">&nbsp;</td>
		<td bgColor="#C0EEC0" style="border:1px solid black;" width="300">&nbsp;&nbsp;&nbsp; <b>»” «‰ﬂ«—Ì</b>
		</td>
	</tr>
	<tr>
		<td valign="top">
<%
	mySQL="SELECT "& TopNo & " * From "& sys & "Items WHERE (Account='"& Customer & "' AND IsCredit=0 AND FullyApplied=0 AND Voided=0) ORDER BY EffectiveDate DESC, ID DESC"
	Set RS1 = conn.Execute(mySQL)

	if (RS1.eof) then 
		call showAlert( "ÅÌœ« ‰‘œ" , CONST_MSG_ALERT)
		noJavaScript=true
	else
	'********************   Debit Items

%>
		<TABLE class="GenTable" align="center" cellspacing="5" cellpadding="2" dir="RTL" border="0" width="100%">
		<Tbody ID="DebitsTAble" name='DebitsTable'>
<%		TotalDebit=0
		while Not (RS1.EOF)
			description = ""
			'================================== Type = 1
			'================================== 
			if RS1("Type")=1 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— " & RS1("Link") & "</a>"
			
			'================================== Type = 2
			'================================== 
			elseif RS1("Type")=2 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showReceipt&receipt="& RS1("Link") & "' target='_blank'>" & "œ—Ì«›  " & RS1("Link") & "</a>"
				Set RST = conn.Execute("select * from ReceivedCash where Receipt=" & RS1("Link"))

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = RST("Description")
				end if
			
			'================================== Type = 3
			'================================== 
			elseif RS1("Type")=3 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showMemo&sys="& sys & "&memo="& RS1("Link") & "' target='_blank'>" & "«⁄·«„ÌÂ "& RS1("Link") & "</a>"
				Set RST = conn.Execute("select * from "& sys & "memo where ID=" & RS1("Link"))

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = RST("Description")
				end if

			'================================== Type = 4
			'================================== 
			elseif RS1("Type")=4 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "»—ê‘  «“ ›—Ê‘" & RS1("Link") & "</a>"

			'================================== Type = 5
			'================================== 
			elseif RS1("Type")=5 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showPayment&payment="& RS1("Link") & "' target='_blank'>" & "Å—œ«Œ  "& RS1("Link") & "</a>"

				mySQL="SELECT ISNULL(PaidCash.Description,'') AS CashDesc, PaidCheques.Description AS CheqDesc FROM PaidCheques RIGHT OUTER JOIN Payments ON PaidCheques.Payment = Payments.ID LEFT OUTER JOIN PaidCash ON Payments.ID = PaidCash.Payment WHERE (Payments.ID = "& RS1("Link") & ")"
				Set RST = conn.Execute(mySQL)

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = RST("CashDesc")
					Do Until RST.eof
						description = description + " " + RST("CheqDesc")
						RST.moveNext
					Loop
				end if
				RST.close

			'================================== Type = 6
			'================================== 
			elseif RS1("Type")=6 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showVoucher&voucher="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— Œ—Ìœ "& RS1("Link") & "</a>"

				mySQL="SELECT Title FROM Vouchers WHERE (id = "& RS1("Link") & ")"

				Set RST = conn.Execute(mySQL)

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = replace(RST("Title"), "/", ".")
				end if
				RST.close

			'================================== Unknown
			'================================== 
			else
				sourceLink="<a href='javascript:void(0);'>" & "‰«‘‰«”" & "</a>"
			end if

%>			<tr bgcolor="white" height='70' valign='top' onclick="slctDdit(this)"><td  class="GrayOutLine">
			<input type='hidden' name='DebitItems' value='<%=RS1("ID")%>'>
			<table class='GenTable' width="300" cellspacing='4' cellpadding='1'>
				<tr bgcolor='#FFDDDD'>
					<td class="GrayOutLine"><%=sourceLink%></td>
					<td class="GrayOutLine"><span dir='LTR'><%=RS1("EffectiveDate")%></span></td>
					<td class="GrayOutLine">„»·€ ﬂ·: <INPUT class='RelInput' readonly value='<%=Separate(RS1("AmountOriginal"))%>' size='10'></td>
				</tr>
				<tr>
					<td colspan="3">
					<a disabled><%=description%></a><BR>
					<span>
					<%
					mySQL2="SELECT * From "& sys & "ItemsRelations INNER JOIN  "&sys & "Items ON "& sys & "ItemsRelations.Credit"& sys & "Item = "& sys & "Items.ID WHERE ("& sys & "ItemsRelations.Debit"& sys & "Item = "& RS1("id") & ") "
					Set RS2 = conn.Execute(mySQL2)
					while Not (RS2.EOF)
						response.write "<span disabled><input class='RelInput' size='10' readonly value=' " & separate(RS2("Amount")) & " '> " & " «“ "& typeNamesArray(cint(RS2("Type"))) & " ‘„«—Â "& RS2("Link") & "</span><br>"

						RS2.movenext
					wend
					if RS1("voided") then response.write "<b>»«ÿ· ‘œÂ</b>"
					%>
					&nbsp;</span></td>
				</tr>
				<tr>
					<td class="GrayOutLine" colspan="2">„«‰œÂ: <span><INPUT class='RelInput' readonly Name='DebitRemained' value='<%=Separate(RS1("RemainedAmount"))%>' size='10'></span></td>
					<td align="center"><span style="visibility:hidden;color:red;"><b>Å«” ‘œ</b></span></td>
				</tr>
			</table></td>
			</tr>
<%
			TotalDebit = TotalDebit + cdbl(RS1("RemainedAmount"))
			RS1.movenext
		wend
%>
		</Tbody>
		</TABLE>
<%
	end if
%>		
		</td>
		<td valign="top">
<%
	mySQL="SELECT "& TopNo & " * From "& sys & "Items WHERE (Account='"& Customer & "' AND IsCredit=1 AND FullyApplied=0 AND Voided=0) ORDER BY EffectiveDate DESC, ID DESC"

	Set RS1 = conn.Execute(mySQL)

	if (RS1.eof) then 
		call showAlert( "ÅÌœ« ‰‘œ" , CONST_MSG_ALERT)
		noJavaScript=true
	else
	'********************   Credit Items
%>
		<TABLE class="GenTable" align="center" cellspacing="5" cellpadding="2" dir="RTL" border="1" width="100%">
		<Tbody ID="CreditsTAble" name='CreditsTable'>
<%			TotalCredit=0
			while Not (RS1.EOF)
			description = ""
			'================================== Type = 1
			'================================== 
			if RS1("Type")=1 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— " & RS1("Link") & "</a>"
			
			'================================== Type = 2
			'================================== 
			elseif RS1("Type")=2 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showReceipt&receipt="& RS1("Link") & "' target='_blank'>" & "œ—Ì«›  " & RS1("Link") & "</a>"
				Set RST = conn.Execute("select * from ReceivedCash where Receipt=" & RS1("Link"))

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = RST("Description")
				end if
			
			'================================== Type = 3
			'================================== 
			elseif RS1("Type")=3 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showMemo&sys="& sys & "&memo="& RS1("Link") & "' target='_blank'>" & "«⁄·«„ÌÂ "& RS1("Link") & "</a>"
				Set RST = conn.Execute("select * from "& sys & "memo where ID=" & RS1("Link"))

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = RST("Description")
				end if

			'================================== Type = 4
			'================================== 
			elseif RS1("Type")=4 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showInvoice&invoice="& RS1("Link") & "' target='_blank'>" & "»—ê‘  «“ ›—Ê‘" & RS1("Link") & "</a>"

			'================================== Type = 5
			'================================== 
			elseif RS1("Type")=5 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showPayment&payment="& RS1("Link") & "' target='_blank'>" & "Å—œ«Œ  " & RS1("Link") & "</a>"

				mySQL="SELECT ISNULL(PaidCash.Description,'') AS CashDesc, PaidCheques.Description AS CheqDesc FROM PaidCheques RIGHT OUTER JOIN Payments ON PaidCheques.Payment = Payments.ID LEFT OUTER JOIN PaidCash ON Payments.ID = PaidCash.Payment WHERE (Payments.ID = "& RS1("Link") & ")"
				Set RST = conn.Execute(mySQL)

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = RST("CashDesc")
					Do Until RST.eof
						description = description + " " + RST("CheqDesc")
						RST.moveNext
					Loop
				end if
				RST.close

			'================================== Type = 6
			'================================== 
			elseif RS1("Type")=6 then
				sourceLink="<a style='text-decoration:none; color:green;' href='AccountReport.asp?act=showVoucher&voucher="& RS1("Link") & "' target='_blank'>" & "›«ﬂ Ê— Œ—Ìœ "& RS1("Link") & "</a>"

				mySQL="SELECT Title FROM Vouchers WHERE (id = "& RS1("Link") & ")"

				Set RST = conn.Execute(mySQL)

				if RST.EOF then
					description = "Œÿ« ! ”‰œ „—»ÊÿÂ ÊÃÊœ ‰œ«—œ"
				else
					description = replace(RST("Title"), "/", ".")
				end if
				RST.close

			'================================== Unknown
			'================================== 
			else
				sourceLink="<a href='javascript:void(0);'>" & "‰«‘‰«”" & "</a>"
			end if

%>			<tr bgcolor="white" height='70' valign='top' onclick="slctCrdit(this)"><td class="GrayOutLine">
			<input type='hidden' name='CreditItems' value='<%=RS1("ID")%>'>
			<table class='GenTable' width="300" cellspacing='4' cellpadding='1'>
				<tr bgcolor='#CCEECC'>
					<td class="GrayOutLine"><%=sourceLink%></td>
					<td class="GrayOutLine"><span dir='LTR'><%=RS1("EffectiveDate")%></span></td>
					<td class="GrayOutLine">„»·€ ﬂ·: <INPUT class='RelInput' readonly value='<%=Separate(RS1("AmountOriginal"))%>' size='10'></td>
				</tr>
				<tr>
					<td colspan="3">
					<a disabled><%=description%></a><BR>
					<span>
					<%
					mySQL2="SELECT * From "& sys & "ItemsRelations INNER JOIN  "&sys & "Items ON "& sys & "ItemsRelations.Debit"& sys & "Item = "& sys & "Items.ID WHERE ("& sys & "ItemsRelations.Credit"& sys & "Item = "& RS1("id") & ") "
					Set RS2 = conn.Execute(mySQL2)
					while Not (RS2.EOF)
						response.write "<span disabled><input class='RelInput' size='10' readonly value=' " & separate(RS2("Amount")) & " '> " & "»Â "& typeNamesArray(cint(RS2("Type"))) & " ‘„«—Â "& RS2("Link") & "</span><br>"

						RS2.movenext
					wend
					if RS1("voided") then response.write "<b>»«ÿ· ‘œÂ</b>"
					%>
					&nbsp;</span></td>
				</tr>
				<tr>
					<td class="GrayOutLine" colspan="2">„«‰œÂ: <span><INPUT class='RelInput' readonly Name='CreditRemained' value=<%=Separate(RS1("RemainedAmount"))%> size='10'></span></td>
					<td align="center"><span style="visibility:hidden;color:green;"><b>Å«” ‘œ</b></span></td>
				</tr>
			</table></td>
			</tr>
<%
				TotalCredit= TotalCredit + cdbl(RS1("RemainedAmount"))
				RS1.movenext
			wend
%>
		</Tbody>
		</TABLE>
<%
	end if
%>		
		</td>
	</tr>
	<tr height="5">
		<td dir='LTR' align='left' colspan="3"><span id="Relations"></span></td>
	</tr>
	<tr style="padding:4">
		<td class="GrayOutLine" bgColor="#FFC0C0">Ã„⁄ : <span>&nbsp;&nbsp;<INPUT class='RelInput' readonly Name='TotalDebitRemained' value='<%=Separate(TotalDebit)%>' size='10'></span></td>
		<td class="GrayOutLine" bgColor="#C0EEC0">Ã„⁄ : <span>&nbsp;&nbsp;<INPUT class='RelInput' readonly Name='TotalCreditRemained' value='<%=Separate(TotalCredit)%>' size='10'></span></td>
	</tr>
	<tr height="10">
		<td align='center' colspan="3"><input name="selectedCustomer" type="hidden" value="<%=Customer%>"></td>
	</tr>
	<tr>
		<td align='left'><input class="GenButton" type="button" style="font-family:Tahoma" value="«‰’—«›" onclick="window.location='ItemsRelation.asp';" >&nbsp;</td>
		<td align='center'>&nbsp;<input class="GenButton" type="button" style="font-family:Tahoma" value="  «ÌÌœ " onclick="this.disabled=true;submit();" >&nbsp;</td>
		<td align='right'><input class="GenButton" type="button" style="font-family:Tahoma" value="Å«ﬂ ﬂ—œ‰" onclick="window.location='ItemsRelation.asp?act=relate&selectedCustomer=<%=Customer%>'" >&nbsp;</td>
	</tr>
	</Tbody>
	</TABLE><br>
	</FORM>
<%
	end if
'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------- submits Relations
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submitRelation" then
	sys = request("sys")
	selectedCustomer = request("selectedCustomer")
'response.write request.form("CreditItems") 
'response.write "<br>"
'response.write request.form("DebitItems")
'response.write "<br>"
'response.end
	if sys = "" then sys = "AR"
	for i=1 to request.form("CreditItems").count 
		remained=text2value(request.form("CreditRemained")(i))
		mySQL="UPDATE "& sys & "Items SET RemainedAmount='"& remained & "' WHERE (ID='"& request.form("CreditItems")(i) & "')"
		conn.Execute(mySQL)	
'		response.write mySQL & "<br>"
	next
	for i=1 to request.form("DebitItems").count 
		remained=text2value(request.form("DebitRemained")(i))
		mySQL="UPDATE "& sys & "Items SET RemainedAmount='"& remained & "' WHERE (ID='"& request.form("DebitItems")(i) & "')"
		conn.Execute(mySQL)
'		response.write mySQL & "<br>"
	next

	mySQL="UPDATE "& sys & "Items SET FullyApplied='1' WHERE (RemainedAmount='0')"	
	conn.Execute(mySQL)
'	response.write mySQL & "<br>"
	
	for i=1 to request.form("CreditARItem").count 
		amount = cdbl(request.form("Amount")(i))
		mySQL="INSERT INTO "& sys & "ItemsRelations (CreatedDate, CreatedBy, Credit"& sys & "Item, Debit"& sys & "Item, Amount) VALUES (N'"& ShamsiToday() & "', '"& session("ID") & "', '"& request.form("CreditARItem")(i) &"', '"& request.form("DebitARItem")(i) &"', '"& amount &"')"
		conn.Execute(mySQL)
'		response.write mySQL & "<br>"
	next
	conn.Close
'response.write request.form("CreditARItem").count
	response.redirect "ItemsRelation.asp?sys=" & sys & "&act=relate&selectedCustomer=" & selectedCustomer
'response.end
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Tear a Relation
'-----------------------------------------------------------------------------------------------------
elseif request("act")="unrelate" then

	relation=clng(request("relation"))
	sys= ucase(request("sys"))
	if NOT (sys = "AR" OR sys = "AP" OR sys = "AO" ) then
		conn.close
		response.redirect "top.asp?errMsg="& Server.URLEncode("Œÿ«!")
	end if

	mySQL="SELECT * FROM "& sys & "ItemsRelations WHERE (ID = '"& relation & "')"
	Set RS1=conn.Execute(mySQL)
	If not RS1.eof then
		'*********  Adding back the amount in the relation, to the debit Item ...
		conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Debit"& sys & "Item") & "')")

		'*********  Adding back the amount in the relation, to the credit Item ...
		conn.Execute("UPDATE "& sys & "Items SET RemainedAmount=RemainedAmount+ '"& RS1("Amount") & "', FullyApplied=0 WHERE (ID = '"& RS1("Credit"& sys & "Item") & "')")

		'*********  Deleting the relation
		conn.Execute("DELETE FROM "& sys & "ItemsRelations WHERE (ID = '"& relation & "')")
	End if
	RS1.close

	conn.close
	response.redirect Request.ServerVariables("HTTP_REFERER")&"&msg=" &Server.URLEncode("—«»ÿÂ Õ–› ‘œ.")
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Search for account
'-----------------------------------------------------------------------------------------------------
else%>
<!-- Ã” ÃÊ »—«Ì ‰«„ Õ”«» --><BR><BR>
	<FORM METHOD=POST ACTION="ItemsRelation.asp?act=submitsearch" onsubmit="if (document.all.CustomerNameSearchBox.value=='') return false;">
	<div dir='rtl'>&nbsp;<B> ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Õ”«»</B>
		<INPUT TYPE="text" NAME="CustomerNameSearchBox">&nbsp;
		<INPUT class="GenButton" TYPE="submit" value="Ã” ÃÊ"><br>
	</div>
	</FORM>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.CustomerNameSearchBox.focus();
	//-->
	</SCRIPT>
<%
end if
conn.Close
if not (noJavaScript=true) then %>
<script language="JavaScript">
<!--
//---------------------------------------------
//---------------------------------------------
//---------------------------------------------
var slctRowClr='#FFFF00'
function slctCrdit(srcRow){
	if (srcRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0].value==0)
		return 0;
	document.getElementsByName("CreditsTable")[0].style.cursor="auto"
	document.all.selectedCreditItem.value=srcRow.rowIndex
	for (i=0;i<document.getElementsByName("CreditItems").length;i++){
		theRow=document.getElementsByName("CreditItems")[i].parentNode.parentNode;
		theRow.setAttribute("bgColor","#FFFFFF")
	}
	srcRow.setAttribute("bgColor",slctRowClr)

	if(document.all.selectedDebitItem.value!=''){

		dstRow=document.getElementsByName("DebitItems")[document.all.selectedDebitItem.value].parentNode.parentNode;
		dstSpan=dstRow.getElementsByTagName("span")[1]
		dstName=dstRow.getElementsByTagName("td")[1].innerText
		dstRemained=dstRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0]

		srcSpan=srcRow.getElementsByTagName("span")[1]
		srcName=srcRow.getElementsByTagName("td")[1].innerText
		srcRemained=srcRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0]

		suggest=Math.min(parseInt(txt2val(srcRemained.value)),parseInt(txt2val(dstRemained.value)))
		amount=parseInt(txt2val(prompt("„»·€ „Ê—œ‰Ÿ— —« Ê«—œ ﬂ‰Ìœ",suggest)))

		if (amount!='0' && amount <= suggest){
			srcRemained.value = val2txt(parseInt(txt2val(srcRemained.value)) - parseInt(amount));
			dstRemained.value = val2txt(parseInt(txt2val(dstRemained.value)) - parseInt(amount));
			srcSpan.innerHTML+="<input class='RelInput' readonly value='"+val2txt(amount)+"' size='10'>"+"»Â "+dstName +"<br>&nbsp;"
			dstSpan.innerHTML+="<input class='RelInput' readonly value='"+val2txt(amount)+"' size='10'>"+"«“ "+srcName +"<br>&nbsp;"

			document.getElementById("Relations").innerHTML+="<input type='hidden' name='CreditARItem' value='"+srcRow.getElementsByTagName("input")[0].value+"'>"
			document.getElementById("Relations").innerHTML+="<input type='hidden' name='DebitARItem' value='"+dstRow.getElementsByTagName("input")[0].value+"'>"
			document.getElementById("Relations").innerHTML+="<input type='hidden' name='Amount' value='"+amount+"'>"
			setRemained();
		}
		if (dstRemained.value =='0'){
			dstRow.getElementsByTagName("tr")[2].getElementsByTagName("span")[1].style.visibility="visible"
			document.all.selectedDebitItem.value=""
			dstRow.setAttribute("bgColor","#FFFFFF")
		}
		if (srcRemained.value =='0'){
			srcRow.getElementsByTagName("tr")[2].getElementsByTagName("span")[1].style.visibility="visible"
			document.all.selectedCreditItem.value=""
			srcRow.setAttribute("bgColor","#FFFFFF")
		}

	}
	else{
		document.getElementsByName("DebitsTable")[0].style.cursor="crosshair"
	}

}

//---------------------------------------------
//---------------------------------------------
//---------------------------------------------
function slctDdit(dstRow){
	if (dstRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0].value==0)
		return 0;

	document.getElementsByName("DebitsTable")[0].style.cursor="auto"
	document.all.selectedDebitItem.value=dstRow.rowIndex

	for (i=0;i<document.getElementsByName("DebitItems").length;i++){
		theRow=document.getElementsByName("DebitItems")[i].parentNode.parentNode;
		theRow.setAttribute("bgColor","#FFFFFF");
	}
	dstRow.setAttribute("bgColor",slctRowClr)

	if(document.all.selectedCreditItem.value!=''){

		srcRow=document.getElementsByName("CreditItems")[document.all.selectedCreditItem.value].parentNode.parentNode;
		srcSpan=srcRow.getElementsByTagName("span")[1]
		srcName=srcRow.getElementsByTagName("td")[1].innerText
		srcRemained=srcRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0]

		dstSpan=dstRow.getElementsByTagName("span")[1]
		dstName=dstRow.getElementsByTagName("td")[1].innerText
		dstRemained=dstRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0]

		suggest=Math.min(parseInt(txt2val(dstRemained.value)),parseInt(txt2val(srcRemained.value)))
		amount=parseInt(txt2val(prompt("„»·€ „Ê—œ‰Ÿ— —« Ê«—œ ﬂ‰Ìœ",suggest)))

		if (amount!='0' && amount <= suggest){
			dstRemained.value = val2txt(parseInt(txt2val(dstRemained.value)) - parseInt(amount));
			srcRemained.value = val2txt(parseInt(txt2val(srcRemained.value)) - parseInt(amount));
			dstSpan.innerHTML+="<input class='RelInput' readonly value='"+val2txt(amount)+"' size='10'>"+"«“ "+srcName +"<br>&nbsp;"
			srcSpan.innerHTML+="<input class='RelInput' readonly value='"+val2txt(amount)+"' size='10'>"+"»Â "+dstName +"<br>&nbsp;"

			document.getElementById("Relations").innerHTML+="<input type='hidden' name='CreditARItem' value='"+srcRow.getElementsByTagName("input")[0].value+"'>"
			document.getElementById("Relations").innerHTML+="<input type='hidden' name='DebitARItem' value='"+dstRow.getElementsByTagName("input")[0].value+"'>"
			document.getElementById("Relations").innerHTML+="<input type='hidden' name='Amount' value='"+amount+"'>"
			setRemained();
		}

		if (srcRemained.value =='0'){
			srcRow.getElementsByTagName("tr")[2].getElementsByTagName("span")[1].style.visibility="visible"
			document.all.selectedCreditItem.value=""
			srcRow.setAttribute("bgColor","#FFFFFF")
		}
		if (dstRemained.value =='0'){
			dstRow.getElementsByTagName("tr")[2].getElementsByTagName("span")[1].style.visibility="visible"
			document.all.selectedDebitItem.value=""
			dstRow.setAttribute("bgColor","#FFFFFF")
		}

	}
	else{
		document.getElementsByName("CreditsTable")[0].style.cursor="crosshair"
	}
}

//---------------------------------------------
//---------------------------------------------
//---------------------------------------------
function setRemained(){
	totalDebit=0
	for (i=0;i<document.getElementsByName("DebitItems").length;i++){
		theRow=document.getElementsByName("DebitItems")[i].parentNode.parentNode;
		totalDebit+= parseInt(txt2val(theRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0].value));
	}
	document.all.TotalDebitRemained.value=val2txt(totalDebit);

	totalCredit=0
	for (i=0;i<document.getElementsByName("CreditItems").length;i++){
		theRow=document.getElementsByName("CreditItems")[i].parentNode.parentNode;
		totalCredit += parseInt(txt2val(theRow.getElementsByTagName("tr")[2].getElementsByTagName("input")[0].value));
	}
	document.all.TotalCreditRemained.value=val2txt(totalCredit);
}

//---------------------------------------------
//---------------------------------------------
//---------------------------------------------
function unselect(){
	document.getElementsByName("CreditsTable")[0].style.cursor="auto"
	document.getElementsByName("DebitsTable")[0].style.cursor="auto"
	if(document.all.selectedCreditItem.value!=''){
		document.getElementsByName("CreditItems")[document.all.selectedCreditItem.value].parentNode.parentNode.setAttribute("bgColor","#FFFFFF");
		document.all.selectedCreditItem.value=''
	}
	if(document.all.selectedDebitItem.value!=''){
		document.getElementsByName("DebitItems")[document.all.selectedDebitItem.value].parentNode.parentNode.setAttribute("bgColor","#FFFFFF");
		document.all.selectedDebitItem.value=''
	}
}

//---------------------------------------------
//---------------------------------------------
//---------------------------------------------
function delRelation(src){
	srcRow=src.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode
	alert(srcRow)
}
//-->
</script>
<%else%>
<script language="JavaScript">
<!--
function slctCrdit(srcRow){}
function slctDdit(dstRow){}
function unselect(){}
//-->
</script>
<%end if%>