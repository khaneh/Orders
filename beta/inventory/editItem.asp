<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
PageTitle= " «’·«Õ ﬂ«·«"
SubmenuItem=6
if not Auth(5 , 6) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Update Item Properties
'-----------------------------------------------------------------------------------------------------
if request("act")="editPrice" then

	StartDate		= SqlSafe(request.form("StartDate"))
	EndDate			= SqlSafe(request.form("EndDate"))
	UnitPrice		= cdbl(text2value(request.form("UnitPrice")))
	autoKey			= cdbl(text2value(request.form("autoKey")))
	InventoryItem	= cdbl(text2value(request.form("InventoryItem")))

	CreatedDate		= shamsiToday()
	CreatedTime		= currentTime10()
	CreatedBy		= session("ID")

	if UnitPrice=0 then
		response.write "<br><br>"				
		call showAlert ("Œÿ«!<br><br>ﬁÌ„  —« Ê«—œ ‰ﬂ—œÂ «Ìœ.",CONST_MSG_ERROR) 
		response.end
	end if

	sql = "SELECT * FROM InventoryItemsUnitPrice WHERE (((StartDate <= '" & StartDate & "') AND (EndDate >= '" & StartDate & "')) OR ((StartDate <= '" & EndDate & "') AND (EndDate >= '" & EndDate & "')) OR ((StartDate >= '" & StartDate & "') AND (EndDate <= '" & EndDate & "')) ) AND (InventoryItem ='" & InventoryItem & "') AND (autoKey <> '" & autoKey & "')"
	set RS = conn.Execute (sql)
	if not RS.eof then
		response.redirect "?itemDetail=" & InventoryItem & "&errMsg=" & Server.URLEncode("Œÿ«!<br><b>»«“Â “„«‰Ì  œ«Œ· œ«—œ.</b><br><br>ﬁ»·« »«“Â Ì [ " & replace(RS("StartDate"),"/",".") & "  « " & replace(RS("EndDate"),"/",".") & "] Ê«—œ ‘œÂ «” . <br> ‘„« [ " & replace(StartDate,"/",".") & "  « " & replace(EndDate,"/",".") & "] Ê«—œ ﬂ—œÂ «Ìœ.")
	end if
	RS.close

	if autoKey > 0 then
		sql = "UPDATE InventoryItemsUnitPrice SET InventoryItem ='" & InventoryItem & "', UnitPrice ='" & UnitPrice & "', StartDate ='" & StartDate & "', EndDate ='" & EndDate & "', CreatedBy ='" & CreatedBy & "', CreatedDate ='" & CreatedDate & "', CreatedTime ='" & CreatedTime & "' WHERE (autoKey='" & autoKey & "')"
	else
		sql ="INSERT INTO InventoryItemsUnitPrice (InventoryItem, UnitPrice, StartDate, EndDate, CreatedBy, CreatedDate, CreatedTime) VALUES ('" & InventoryItem & "', '" & UnitPrice & "', '" & StartDate & "', '" & EndDate & "', '" & CreatedBy & "', '" & CreatedDate & "', '" & CreatedTime & "')"
	end if
	
	conn.execute(sql)

	conn.close
	response.redirect "?itemDetail=" & InventoryItem & "&msg=" & Server.URLEncode("ﬁÌ„  À»  ‘œ.")


elseif request("submit")="À»  „ÊÃÊœÌ ÃœÌœ"then

	id				= request.form("id") 
	oldQtty			= request.form("oldQtty") 
	qtty			= request.form("qtty") 
	Unit			= request.form("Unit") 
	comments		= request.form("comments") 
	ItemName		= request.form("ItemName") 

	if cdbl(oldQtty) < cdbl(qtty) then
		'---------------------- Increase Item Qtty by Hand
		'---------------------- RelatedID = -2 , type = 2
		qttyUp = qtty - oldQtty
		mySql="INSERT INTO InventoryLog (ItemID, RelatedID, logDate, Qtty, owner, CreatedBy, IsInput, type, comments) VALUES ("& id & ", -2 ,N'"& shamsiToday() & "', "& qttyUp & ", -1 , "& session("id") & ", 1 , 2, N'"& comments & "')"
		conn.Execute mySql
		response.write "<BR><BR><BR><CENTER><li> «›“«Ì‘ „ÊÃÊœÌ "& ItemName & " »Â „Ì“«‰ " & qttyUp& "&nbsp;"&Unit&" «‰Ã«„ ‘œ</CENTER>"
	end if
	
	if cdbl(oldQtty) > cdbl(qtty) then
		'---------------------- Decrease Item Qtty by Hand
		'---------------------- RelatedID = -2 , type = 2
		qttyDown = oldQtty - qtty
		mySql = "INSERT INTO InventoryLog (ItemID, RelatedID, Qtty, logDate, owner, CreatedBy, IsInput, type, comments) VALUES ("& ID & ",-2,"& qttyDown& ",N'"& shamsiToday() & "', -1 ,"& session("ID") & ", 0, 2, N'"& comments & "')"
		conn.Execute mySql
		response.write "<BR><BR><BR><CENTER><li> ﬂ«Â‘ „ÊÃÊœÌ "& ItemName & " »Â „Ì“«‰ " & qttyDown & "&nbsp;"&Unit&" «‰Ã«„ ‘œ</CENTER>"
	end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Update Item Properties
'-----------------------------------------------------------------------------------------------------
elseif request("submit")="À»  „‘Œ’«  ﬂ«·«"then

	oldCatID		= request.form("oldCatID") 
	id				= request.form("id") 
	catItem			= request.form("catItem") 
	ItemName		= request.form("ItemName") 
	OldItemID		= request.form("OldItemID") 
	ownerAcc		= request.form("ownerAcc") 
	costingMethod	= request.form("cost") 
	Minim			= request.form("Minim") 
	oldQtty			= request.form("oldQtty") 
	qtty			= request.form("qtty") 
	Unit			= request.form("Unit") 
	outByOrder		= request.form("outByOrder")
	EditDate		= shamsiToday()
	EditBy			= session("ID")

	if outByOrder="on" then
		outByOrder = 1
	else
		outByOrder =0
	end if


	if cdbl(oldQtty) < cdbl(qtty) then
		'---------------------- Request to Increase Item Qtty by Hand
		qttyUp = qtty - oldQtty
		response.write "<center><br><br>"
		response.write "<li>‘„« œ— Õ«· «›“«Ì‘ „ÊÃÊœÌ <b> "& ItemName & " </b> »Â „Ì“«‰ <b>" & qttyUp& "&nbsp;"&Unit & " </b> Â” Ìœ.<br><br> ·ÿ›« œ·Ì· «Ì‰ «’·«Õ „ÊÃÊœÌ —« Ê«—œ ﬂ‰Ìœ:" 
		%>
		<FORM METHOD=POST ACTION="">
			<INPUT TYPE="hidden" name="id" value="<%=id%>">
			<INPUT TYPE="hidden" name="oldQtty" value="<%=oldQtty%>">
			<INPUT TYPE="hidden" name="qtty" value="<%=qtty%>">
			<INPUT TYPE="hidden" name="ItemName" value="<%=ItemName%>">
			<INPUT TYPE="hidden" name="Unit" value="<%=Unit%>">
			<TEXTAREA NAME="comments" ROWS="5" COLS="20"></TEXTAREA><BR><BR>
			<INPUT TYPE="submit" name="submit" value="À»  „ÊÃÊœÌ ÃœÌœ">
		</FORM>
		</center>
		<%
	end if
	
	if cdbl(oldQtty) > cdbl(qtty) then
		'---------------------- Request to Decrease Item Qtty by Hand
		qttyDown = oldQtty - qtty
		response.write "<center><br><br>"
		response.write "<li>‘„« œ— Õ«· ﬂ«Â‘ „ÊÃÊœÌ <b> "& ItemName & " </b> »Â „Ì“«‰ <b>" & qttyDown& "&nbsp;"&Unit & " </b> Â” Ìœ.<br><br> ·ÿ›« œ·Ì· «Ì‰ «’·«Õ „ÊÃÊœÌ —« Ê«—œ ﬂ‰Ìœ:" 
		%>
		<FORM METHOD=POST ACTION="">
			<INPUT TYPE="hidden" name="id" value="<%=id%>">
			<INPUT TYPE="hidden" name="oldQtty" value="<%=oldQtty%>">
			<INPUT TYPE="hidden" name="qtty" value="<%=qtty%>">
			<INPUT TYPE="hidden" name="ItemName" value="<%=ItemName%>">
			<INPUT TYPE="hidden" name="Unit" value="<%=Unit%>">
			<TEXTAREA NAME="comments" ROWS="5" COLS="20"></TEXTAREA><BR><BR>
			<INPUT TYPE="submit" name="submit" value="À»  „ÊÃÊœÌ ÃœÌœ">
		</FORM>
		</center>
		<%
	end if
	
	response.write "<BR><BR><TABLE align=center style='border: solid 2pt black'><TR><TD>"
	response.write "<li> ﬂœ œ” Â »‰œÌ = " & catItem& "<br>"
	response.write "<li> ‰«„ ﬂ«·« = " & ItemName& "<br>"
	response.write "<li> ’«Õ» ﬂ«·« = " 
	'if ownerAcc="-1" then 
	'response.write 	"Œ«‰Â ç«Å Ê ÿ—Õ<br>"
	'else
	'response.write 	"‘„«—Â Õ”«» "& ownerAcc & "<br>"
	'end if
	response.write "<li> ﬂœ ﬂ«·« = " & OldItemID& "<br>"
	response.write "<li> ‘ÌÊÂ ﬁÌ„  ê–«—Ì = " & costingMethod& "<br>"
	response.write "<li> Õœ«ﬁ· „ÊÃÊœÌ = " & Minim& "<br>"
	response.write "<li>  ⁄œ«œ = " & oldQtty& "<br>"
	if outByOrder=1 then
		response.write "<li> Œ—ÊÃ »— «”«” ”›«—‘"
	end if
	response.write "<li> Ê«Õœ = " & Unit& "<br>"
	response.write "</TD></TR></TABLE>"

'	Added By Alix 821118 
'	-------------------------------
'	Log The Edition Before Updating
	conn.Execute("INSERT INTO InventoryItemsEditLog SELECT '"& EditDate & "' AS EditedOn, '"& EditBy & "' AS EditedBy, " & oldCatID & " as oldCategory, * FROM InventoryItems WHERE (ID = "& ID & ")")

'	End of Log
' -------------------------------

	mySql="update InventoryItems set OldItemID="& OldItemID & ", owner="& ownerAcc & ", Name=N'"& ItemName & "', Minim="& Minim & ", Unit=N'"& Unit & "', costingMethod='"& costingMethod & "', outByOrder=" & outByOrder & " where id="	& id
	conn.Execute mySql

	mySql="update InventoryItemCategoryRelations set Cat_ID="& catItem & " where Item_ID="	& id
	conn.Execute mySql

	response.write "<center><br>ﬂ«·«Ì ›Êﬁ »« „Ê›ﬁÌ  ÊÌ—«Ì‘ ‘œ.</center>"
	response.end 

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------- Edit Qtty's (BATCH)
'-----------------------------------------------------------------------------------------------------
elseif request("submit")="«’·«Õ „ÊÃÊœÌ Â«"then
	'response.write "<br>" & replace(request.form,"&","<br>")
	i=0
	for each Qtty in request.form("Qtty") 
		i=i+1
		id = request.form("id")(i)
		set RS4 = conn.Execute ("SELECT * FROM InventoryItems where id=" & id)
		if RS4.eof then 
			response.write "<br>Œÿ«Ì €Ì— „‰ Ÿ—Â. »« „œÌ— ”Ì” „  „«” »êÌ—Ìœ." 
			response.end
		else
			oldQtty = RS4("Qtty")
		end if

		if clng(oldQtty) < clng(Qtty) then
			'---------------------- Increase Item Qtty by Hand
			'---------------------- RelatedID = -2 
			qttyUp = qtty - oldQtty
			mySql="INSERT INTO InventoryLog (ItemID, RelatedID, logDate, Qtty, owner, CreatedBy, IsInput, type) VALUES ("& id & ", -2 ,N'"& shamsiToday() & "', "& qttyUp & ", -1 , "& session("id") & ", 1 , 2)"
			conn.Execute mySql
		end if
		
		if clng(oldQtty) > clng(Qtty) then
			'---------------------- Decrease Item Qtty by Hand
			'---------------------- RelatedID = -2 
			qttyDown = oldQtty - qtty
			mySql = "INSERT INTO InventoryLog (ItemID, RelatedID, Qtty, logDate, owner, CreatedBy, IsInput, type) VALUES ("& ID & ",-2,"& qttyDown& ",N'"& shamsiToday() & "', -1 ,"& session("ID") & ", 0, 2)"
			conn.Execute mySql
		end if
		
	next
end if

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------- Report page Hearder
'-----------------------------------------------------------------------------------------------------
if request("itemDetail")="" then
	%>
	<BR><BR>
	<FORM METHOD=POST ACTION="editItem.asp">
	<%
	catItem1 = request("catItem")
	%>
	<TABLE dir=rtl align=center width=600>
	<TR bgcolor="#DDDDDD" >
		<TD align=center >«‰ Œ«» ﬂ‰Ìœ:
		
			<SELECT NAME="catItem" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1" onchange="document.forms[0].submit()">
			<option value="-1">œ” Â »‰œÌ ﬂ«·« </option>
			<option value="-1">------------------</option>
			<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories")
				while not (RS4.eof) %>
					<OPTION value="<%=RS4("ID")%>"<%
					if trim(catItem1) = trim(RS4("ID")) then
					response.write " selected "
					end if
					%>>* <%=RS4("Name")%> </option>
					<%	
					RS4.MoveNext
				wend
				RS4.close
				%>
			</SELECT>
		</TD>
	</TR>
	</table>
	</FORM>
	<%
end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------- List Items of a category
'-----------------------------------------------------------------------------------------------------
if request("catItem") > 0 then
	set RSS=Conn.Execute ("SELECT InventoryItems.ID, InventoryItems.OldItemID, InventoryItems.owner, InventoryItems.Minim, InventoryItems.Name, InventoryItems.Qtty, InventoryItems.CusQtty, InventoryItems.Unit, InventoryItems.costingMethod, InventoryItemCategories.Name AS Expr1 FROM InventoryItems INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID INNER JOIN InventoryItemCategories ON InventoryItemCategoryRelations.Cat_ID = InventoryItemCategories.ID WHERE (InventoryItemCategories.ID ="& catItem1 & ") and InventoryItems.enabled=1 order by OldItemID")	
	%>
	<TABLE dir=rtl align=center width=600>
	<TR bgcolor="eeeeee">
		<TD><!A HREF="default.asp?s=1"><SMALL>ﬂœ ﬂ«·«</SMALL></A></TD>
		<TD><!A HREF="default.asp?s=2"><SMALL>‰«„ ﬂ«·«</SMALL></A></TD>
		<TD><!A HREF="default.asp?s=3"><SMALL> ⁄œ«œ</SMALL></A></TD>
		<TD><!A HREF="default.asp?s=3"><SMALL> ⁄œ«œ «—”«·Ì</SMALL></A></TD>
		<TD><!A HREF="default.asp?s=4"><SMALL>Õœ«ﬁ· „ÊÃÊœÌ</SMALL></A></TD>
		<TD><!A HREF="default.asp?s=5"><SMALL>Ê«Õœ</SMALL></A></TD>
		<TD><!A HREF="default.asp?s=6"><SMALL>‰ÕÊÂ «—“Ì«»Ì</SMALL></A></TD>
	</TR>
	<FORM METHOD=POST ACTION="">
	<INPUT TYPE="hidden" name="catItem" value="<%=catItem1%>">
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
		<TR bgcolor="<%=tmpColor%>" height=25>
			<TD align=center dir=ltr><%=RSS("OldItemID")%></TD>
			<TD><A HREF="editItem.asp?itemDetail=<%=RSS("ID")%>"># <%=RSS("Name")%></A></TD>
			<TD align=center dir=ltr>
			<% if not Auth(5 , 7) then %>
			<%=RSS("Qtty")%>
			<% else %>
			<INPUT TYPE="hidden" NAME="id" value="<%=RSS("id")%>" style="border:0pt; width:40pt; background-color:transparent">
			<INPUT TYPE="text" NAME="Qtty" value="<%=RSS("Qtty")%>" style="border:0pt; width:40pt; background-color:transparent;text-align:center">
			<% end if %>
			</TD>
			<TD align=center dir=ltr><a style="cursor:hand" onclick="window.open('cusItemDetails.asp?id=<%=RSS("ID")%>&name=<%=RSS("Name")%>','CusItemDetails','width=400,height=300')"><%=RSS("CusQtty")%></a></TD>
			<TD align=center dir=ltr><%=RSS("Minim")%></TD>
			<TD><%=RSS("Unit")%></TD>
			<TD><%=RSS("costingMethod")%></TD>
		</TR>
			 
		<% 
		RSS.moveNext
	Loop
	%>
	</TABLE><br>
	<% if Auth(5 , 7) then %>
	<CENTER><INPUT TYPE="submit" name="submit" value="«’·«Õ „ÊÃÊœÌ Â«"></CENTER>
	<% end if %>
	</FORM>
	<%
end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Details of an item
'-----------------------------------------------------------------------------------------------------
if request("itemDetail")<>"" then

	set RS3 = conn.Execute ("SELECT InventoryItemCategories.id as catID, InventoryItems.ID, InventoryItems.outByOrder, InventoryItems.OldItemID, InventoryItems.owner, InventoryItems.Name, InventoryItems.Minim, InventoryItems.Qtty, InventoryItems.Unit, InventoryItems.costingMethod FROM InventoryItems INNER JOIN InventoryItemCategoryRelations ON InventoryItems.ID = InventoryItemCategoryRelations.Item_ID INNER JOIN InventoryItemCategories ON InventoryItemCategoryRelations.Cat_ID = InventoryItemCategories.ID WHERE (InventoryItems.ID = '"& sqlSafe(request("itemDetail")) & "')")
	if RS3.EOF then
		response.write "<br><br>"
		call showAlert ("Œÿ«!<br><br>ç‰Ì‰ çÌ“Ì œ— «‰»«— ‰œ«—Ì„",CONST_MSG_ERROR) 
		response.end
	end if
	%>

	<style>
		.table12 {font-family: tahoma; font-size: 8pt; border:}
		.table12 TR {height:20px; }
		.table12 TH {direction:rtl; font-weight:normal;}
		.table12 TD {direction:ltr; }
	</style>

	<SCRIPT LANGUAGE="JavaScript">
	<!--
	var noNextField = false;
	function copyInfo(index){
		var myObj=document.getElementsByTagName("table").item('prices').getElementsByTagName("tr").item(index);

		document.all.autoKey.value=myObj.getElementsByTagName("input").item(0).value;

		document.all.StartDate.value=myObj.getElementsByTagName("td").item(0).innerText;
		document.all.EndDate.value=myObj.getElementsByTagName("td").item(1).innerText;
		document.all.UnitPrice.value=myObj.getElementsByTagName("td").item(2).innerText;

		document.all.ClearForm.style.visibility="visible";
		document.all.BtnSubmit.value="ÊÌ—«Ì‘ ﬁÌ„ ";
		document.all.StartDate.select();
	}
	function clearForm(){
		document.all.autoKey.value="";
		document.all.StartDate.value="";
		document.all.EndDate.value="";
		document.all.UnitPrice.value="";
		document.all.BtnSubmit.value="À»  ﬁÌ„ ";
		document.all.ClearForm.style.visibility="hidden";
		document.all.StartDate.select();
	}
	function hideIT()
	{
	//alert(document.all.aaa2.value)
	if(document.all.aaa2.value==2) 
		{
			document.all.aaa1.style.visibility= 'visible'
			document.all.ownerAcc.value = "<%=trim(RS3("owner"))%>"
		}
		else
		{
			document.all.aaa1.style.visibility= 'hidden'
			document.all.ownerAcc.value = "-1"
		}
	}
	//-->
	</SCRIPT>

	<BR>

	<FORM METHOD=POST ACTION="?">
	<INPUT TYPE="hidden" name="id" value="<%=RS3("ID")%>">
	<TABLE border=0 align=center>
	<TR>
		<TD colspan=2 align=center><H3>ÊÌ—«Ì‘ «ÿ·«⁄«  ﬂ«·«</H3></TD>
	</TR>
	<TR>
		<TD align=left>‰Ê⁄ ﬂ«·«</TD>
		<TD align=right>
			<INPUT TYPE="hidden" name="oldCatID" value="<%=RS3("catID")%>">
			<SELECT NAME="catItem" style='font-family: tahoma,arial ; font-size: 9pt; font-weight: bold' size="1">
			<option value="-1">œ” Â »‰œÌ ﬂ«·« </option>
			<option value="-1">------------------</option>
			<%
				set RS4 = conn.Execute ("SELECT * FROM InventoryItemCategories")
				while not (RS4.eof) %>
					<OPTION value="<%=RS4("ID")%>"<%
					if trim(RS3("catID")) = trim(RS4("ID")) then
					response.write " selected "
					end if
					%>>* <%=RS4("Name")%> </option>
					<%	
					RS4.MoveNext
				wend
				RS4.close
				%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD align=left>‰«„ ﬂ«·«</TD>
		<TD align=right><INPUT TYPE="text" NAME="ItemName" value="<%=RS3("Name")%>" size=64></TD>
	</TR>
	<TR>
		<TD align=left>ﬂœ ﬂ«·«</TD>
		<TD align=right><INPUT TYPE="hidden" NAME="ownerAcc" value="-1"><INPUT TYPE="hidden" NAME="cost" value="FIFO"><INPUT TYPE="text" NAME="OldItemID" value="<%=RS3("OldItemID")%>"></TD>
	</TR>
	<TR>
		<TD align=left>Õœ«ﬁ· „ÊÃÊœÌ</TD>
		<TD align=right><INPUT dir=ltr TYPE="text" NAME="Minim" value="<%=RS3("Minim")%>"></TD>
	</TR>
	<TR>
		<TD align=left> ⁄œ«œ ›⁄·Ì</TD>
		<TD align=right><INPUT TYPE="hidden" NAME="oldQtty" value="<%=RS3("Qtty")%>"><INPUT dir=ltr TYPE="text" NAME="qtty" value="<%=RS3("Qtty")%>"></TD>
	</TR>
	<TR>
		<TD align=left>Ê«Õœ «‰œ«“Â êÌ—Ì</TD>
		<TD align=right><INPUT TYPE="text" NAME="Unit" value="<%=RS3("Unit")%>"></TD>
	</TR>
	<TR>
		<TD align=left><INPUT TYPE="checkbox" NAME="outByOrder" <% if RS3("outByOrder") then%> checked <% end if %>></TD>
		<TD align=right>Œ—ÊÃ »— «”«” ”›«—‘</TD>
	</TR>
	<TR>
		<TD align=left></TD>
		<TD align=right height=20>
		<BR><INPUT TYPE="submit" NamE="submit" value="À»  „‘Œ’«  ﬂ«·«">
		</TD>
	</TR>
	</TABLE>
	</FORM>
	<%
	if Auth(5 , "D") then 'Set Market Price 
	%>
	<hr>
		<FORM METHOD=POST ACTION="?act=editPrice">
		<TABLE style="border:dashed 1px navy;" align=center>
		<TR>
			<TD colspan=2 align=center><H3>À»  ﬁÌ„  »«“«— / BHR</H3></TD>
		</TR>
		<TR>
			<TD align=left>‰«„ ﬂ«·«</TD>
			<TD align=right width=300><b><%=RS3("Name")%></b></TD>
		</TR>
		<TR>
			<TD colspan=2 align=center>
			<%
					set RS4 = conn.Execute ("SELECT InventoryItemsUnitPrice.*, Users.RealName AS CreatorName FROM InventoryItemsUnitPrice INNER JOIN Users ON InventoryItemsUnitPrice.CreatedBy = Users.ID WHERE (InventoryItem='" & RS3("ID") & "') ORDER BY StartDate ")
					if not RS4.eof then
			%>			
						<TABLE id="prices" class="table12" border=0 cellpadding=2 cellspacing=1 bgcolor="0">
							<TR bgcolor="#AAAAAA">
							<TH width=75>«“</TD>
							<TH width=75> «</TD>
							<TH width=75>ﬁÌ„  Ê«Õœ</TD>
							<TH width=90> Ê”ÿ</TD>
						</TR>
			<%
						while not (RS4.eof) 
							tmpCounter = tmpCounter + 1
							if tmpCounter mod 2 = 1 then
								tmpColor="#FFFFFF"
								tmpColor2="#FFFFBB"
							Else
								tmpColor="#DDDDDD"
								tmpColor2="#EEEEBB"
							End if 
			%>
							<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="copyInfo(this.rowIndex)">
								<TD><%=RS4("StartDate")%></TD>
								<TD><%=RS4("EndDate")%></TD>
								<TD><%=Separate(RS4("UnitPrice"))%></TD>
								<TD style="direction:rtl" title="<%=RS4("CreatedDate") & "(" & RS4("CreatedTime") & ")"%>">
									&nbsp;<%=RS4("CreatorName")%>
									<INPUT TYPE="hidden" value="<%=RS4("autoKey")%>">
								</TD>
							</TR>
			<%
							RS4.MoveNext
						wend
			%>			</TABLE>			
			<%
					end if
					RS4.close
			%>
			</TD>
		</TR>
		<TR>
			<TD colspan=2 align=center><HR></TD>
		</TR>
		<TR>
			<TD align=left> «—ÌŒ ‘—Ê⁄</TD>
			<TD align=right><INPUT dir=ltr TYPE="text" NAME="StartDate" value="" onblur="acceptDate(this)" ></TD>
		</TR>
		<TR>
			<TD align=left> «—ÌŒ Å«Ì«‰</TD>
			<TD align=right><INPUT dir=ltr TYPE="text" NAME="EndDate" value="" onblur="acceptDate(this)" ></TD>
		</TR>
		<TR>
			<TD align=left>ﬁÌ„  Ê«Õœ (—Ì«·)</TD>
			<TD align=right><INPUT TYPE="text" NAME="UnitPrice" value="" onblur="this.value = val2txt(txt2val(this.value))"></TD>
		</TR>
		<TR>
			<TD align=right>
				<input type="button" value="ÃœÌœ" name="ClearForm" onclick="clearForm()" style="visibility:hidden;">
			</TD>
			<TD align=right height=20>
				<INPUT TYPE="hidden" name="autoKey" value="">
				<INPUT TYPE="hidden" name="InventoryItem" value="<%=RS3("ID")%>">
				<INPUT TYPE="submit" name="BtnSubmit" value="À»  ﬁÌ„ ">
			</TD>
		</TR>
		</TABLE>
		</FORM>

		<%
	end if
end if
%>
<!--#include file="tah.asp" -->