<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="¬Ì „ Â«Ì ›«ﬂ Ê—"
SubmenuItem=6
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
if request("act")="del" then

	mySQL="DELETE FROM InvoiceItemCategoryRelations WHERE (InvoiceItem = '" & sqlSafeNoEnter(request("item")) & "') AND (InvoiceItemCategory = '" & sqlSafeNoEnter(request("cat")) & "')"
	Conn.Execute (mySQL)
	response.redirect "?act=search&search_box=" & Server.URLEncode(request("nam"))

end if

function sqlSafeNoEnter (s)
  st=s
  st=replace(St,"'","`")
  st=replace(St,chr(34),"`")
  st=replace(St,vbCrLf," ")
  sqlSafeNoEnter=st
end function

function getTypeName(t)
	result="øøøø"
	select case t
	case 0:
		result = "œ” Ì"
	case 1
		result = " ⁄œ«œ x ›—„ x ›Ì"
	case 2
		result = "„Õ«”»Â ﬁÌ„  œÌÃÌ «·"
	case 3
		result = "ÿÊ· x ⁄—÷ x  ⁄œ«œ x ›—„ x ›Ì"
	case 4
		result = "ﬁÌ„  À«» "
	case 5
		result = "Œœ„«  »Ì—Ê‰Ì"
	end select
	getTypeName = result
end function
%>
<style>
	Input { font-family: Tahoma;font-size: 8pt;height:25px;}
	TextArea { font-family: Tahoma;font-size: 9pt;}
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var noNextField = false;
var xmlDoc = null;
function loadXML(xmlFile) 
{ 
 try //Internet Explorer
 {
	xmlDoc = new ActiveXObject("Microsoft.XMLDOM")
	xmlDoc.async=false; 
	xmlDoc.onreadystatechange=verify; 
	xmlDoc.load(xmlFile); 
	xmlObj=xmlDoc.documentElement; 
 }
 catch(e) 
 {
	try // Firefox
	{
	xmlDoc = document.implementation.createDocument("","",null);
	xmlDoc.async=false; 
	xmlDoc.load(xmlFile); 
	//alert(xmlFile);
	//alert(xmlDoc.hasChildNodes());
	}
	catch (e) {alert(e.message)}
 }
 try
 {
	// xmlDoc.async=false; 
	// xmlDoc.onreadystatechange=verify; 
	 //xmlDoc.load(xmlFile); 
	 //xmlObj=xmlDoc.documentElement; 
 }
 catch(e) {alert(e.message)}
}

function verify() 
{ 
 // 0 Object is not initialized 
 // 1 Loading object is loading data 
 // 2 Loaded object has loaded data 
 // 3 Data from object can be worked with 
 // 4 Object completely initialized 
 if (xmlDoc.readyState != 4) 
 { 
   return false; 
 } 
}
function copyInfo(index){
	var myObj=document.getElementsByTagName("table")['result'].getElementsByTagName("tr")[index];
	document.all.name_box.value			= myObj.getElementsByTagName("td").item(1).innerText;
	document.all.enbl_box.checked		= document.getElementsByName("theEnbl")[index-1].checked;
	document.all.vat_box.checked		= document.getElementsByName("theVat")[index-1].checked;
	document.all.fee_box.value			= myObj.getElementsByTagName("td").item(4).innerText;
	document.all.id_box.value			= myObj.getElementsByTagName("input").item(0).value;
	document.all.type_box.selectedIndex	= myObj.getElementsByTagName("input").item(1).value;
	document.all.InvItm_box.value		= myObj.getElementsByTagName("td").item(5).innerText;
	document.all.categories_box.innerHTML="";

	loadXML('xml_InvoiceItemCategories.asp?id='+document.all.id_box.value)
	
	if(xmlDoc.hasChildNodes()){
	
		for (i = 0 ; i < xmlDoc.childNodes[1].childNodes.length ; i++){
			document.all.categories_box.innerHTML += "<A HREF='?act=del&item=" + document.all.id_box.value + "&cat=" + xmlDoc.childNodes[1].childNodes[i].getAttribute("id") + "&nam=" + escape (document.all.name_box.value) + "' title='Õ–›' style='text-decoration:none;color:red;'><B>x</B></A> " + xmlDoc.childNodes[1].childNodes[i].text + "<BR>";
		}
	}
	else{
		document.all.categories_box.innerHTML="<FONT COLOR='red'>Error</FONT>";
	}
	document.all.name_box.select();
	document.all.name_box.focus();
}
function checkValidation(){
	return true;
}
//-->
</SCRIPT>
<br>
<font face="tahoma">
<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center">
<TR bgcolor="#AACCCC">
	<TD><FORM METHOD=POST ACTION="?act=search"><TABLE>
	<TR><%'---------------------------------------------------SAM----------------------------------------------------------------%>
		<TD>‰«„:</TD>
		<TD><INPUT TYPE="text" NAME="search_box" maxlength="50" size="25" tabIndex="2" value=<%=request.form("search_box")%>></TD>
		<TD>«“ ‘„«—Â</TD>
		<TD><INPUT TYPE="text" NAME="from_box" value='<%=request.form("from_box")%>' size="10" tabIndex="3"></TD>
		<TD> « ‘„«—Â</TD>
		<TD><INPUT TYPE="text" NAME="to_box" value='<%=request.form("to_box")%>' size="10" tabIndex="4"></TD>
		<TD>›ﬁÿ ﬂœÂ«Ì ›⁄«· ‰„«Ì‘ œ«œÂ ‘Ê‰œ</TD>
		<TD><INPUT type='checkbox' <%if request.form("enable_chk") = "on" then response.write("checked='checked'")%> name='enable_chk'></TD>
		<TD><INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" style="width:100px; font-family:tahoma,arial;"></TD>
	</TR>
	</TABLE></TD></FORM>
</TR>

<TR bgcolor="#CCCCAA">
	<TD><FORM METHOD=POST ACTION="?act=add" onSubmit="return checkValidation();">
	<TABLE width="100%" >
	<TR>
		<TD>‰«„ ¬Ì „</TD>
		<TD><INPUT checked disabled TYPE="checkbox">›⁄«·</TD>
		<TD>‰Ê⁄ (ﬁÌ„  ê–«—Ì)</TD>
		<TD>ﬁÌ„ </TD>
	</TR>
	<TR>
		<TD valign='top'>
			<TEXTAREA NAME="name_box" rows="3" cols="30"></TEXTAREA>
		</TD>
		<TD valign='top'><INPUT TYPE="checkbox" NAME="enbl_box"></TD>
		<TD valign='top'>
			<SELECT NAME="type_box" style="font-family:tahoma;width:200px;">
				<OPTION Value="0" >œ” Ì</OPTION>
				<OPTION Value="1" > ⁄œ«œ x ›—„ x ›Ì</OPTION>
				<OPTION Value="2" >„Õ«”»Â ﬁÌ„  œÌÃÌ «·</OPTION>
				<OPTION Value="3" >ÿÊ· x ⁄—÷ x  ⁄œ«œ x ›—„ x ›Ì</OPTION>
				<OPTION Value="4" >ﬁÌ„  À«» </OPTION>
				<OPTION Value="5" >Œœ„«  »Ì—Ê‰Ì</OPTION>
			</SELECT>
		</TD>
		<TD valign='top'><INPUT TYPE="Text" NAME="fee_box"></TD>
	</TR>
	<TR>
		<TD>
			ﬂœ ¬Ì „: &nbsp;<INPUT TYPE="Text" NAME="id_box">
		</TD>
		<TD>ﬂœ ﬂ«·«/”—ÊÌ” „— »ÿ: </TD>
		<TD><INPUT TYPE="Text" NAME="InvItm_box" dir=LTR style="text-align:right"></TD>
		<TD align='center'>„«·Ì«  »— «—“‘ «›“ÊœÂ: <INPUT type='checkbox' name='vat_box'></TD>
	</TR>
	<TR>
		<TD colspan=4><hr></TD>
	</TR>
	<TR>
		<TD valign=top align=left>œ” Â »‰œÌ Â«Ì ›⁄·Ì:</TD>
		<TD colspan=3>
			<div id="categories_box">		
			</div>
		</TD>
	</TR>
	<TR>
		<TD align=left>«›“Êœ‰ œ” Â »‰œÌ:</TD>
		<TD>
			<SELECT NAME="cat_box" style="font-family:tahoma;width:200px;">
			<OPTION Value="0" >-- «‰ Œ«» ﬂ‰Ìœ --</OPTION>
<%
			mySQL="SELECT * FROM InvoiceItemCategories ORDER BY [Name]" 
			set RS=Conn.Execute(mySQL)
			
			Do While NOT RS.eof
%>
				<OPTION Value="<%=RS("id")%>" ><%=RS("name")%></OPTION>
<%
				RS.MoveNext
			Loop
%>			</SELECT>
		</TD>
		<TD valign='top' align=left colspan=2>
			<INPUT TYPE="submit" Name="Submit" Value="–ŒÌ—Â" style="width:100px; font-family:tahoma,arial;">
		</TD>
	</TR>
	</TABLE></TD></FORM>
</TR>
</TABLE>
<%

myCriteria= "REPLACE([name], ' ', '') LIKE REPLACE(N'%"& sqlSafeNoEnter(request("search_box")) & "%', ' ', '')"
if request("act")="add" AND request.form("name_box")<>"" AND isnumeric(request.form("id_box")) then
	
	if request.form("enbl_box")="on" then
		enable=1
	else
		enable=0
	end if

	ItemCategory =			clng(request.form("cat_box"))
	ItemID =				clng(request.form("id_box"))
	ItemName =				sqlSafeNoEnter(request.form("name_box"))
	ItemType =				sqlSafeNoEnter(request.form("type_box"))
	ItemFee =				sqlSafeNoEnter(request.form("fee_box"))
	RelatedInventoryItem =	request.form("InvItm_box")
	If request.form("vat_box")="on" Then 
		ItemVat = 1
	Else
		ItemVat = 0 
	End If 

	if isnumeric(RelatedInventoryItem) then 
		RelatedInventoryItem = clng(RelatedInventoryItem) 
	else
		RelatedInventoryItem = 0
	End if
	if RelatedInventoryItem=0 then RelatedInventoryItem=-1


	mySQL="SELECT ID FROM InvoiceItems WHERE ([ID]='"& ItemID & "')"
	set RS1=Conn.Execute (mySQL)
	If ItemVat Then 
		hasVat = 1 
	Else 
		hasVat = 0
	End If 
	if RS1.eof then
		mySQL="INSERT INTO InvoiceItems ([ID], [name], [Enabled], [Type], [Fee], [RelatedInventoryItemID], hasVat) VALUES ("&_
			ItemID & ", N'"& ItemName& "', '" & enable & "','" & ItemType & "', '" & ItemFee & "', '" & RelatedInventoryItem & "', " & hasVat & ") "
	else
		mySQL="UPDATE InvoiceItems SET [Name]=N'" & ItemName & "',[Enabled]='" & enable & "', [Type]='" & ItemType & "', [Fee]='" & ItemFee & "', [RelatedInventoryItemID]='" & RelatedInventoryItem & "', hasVat= '" & hasVat & "' WHERE ([ID]='"& ItemID & "')"
	end if
	RS1.close

	Conn.Execute (mySQL)

	if ItemCategory<>0 then

		mySQL="SELECT * FROM InvoiceItemCategoryRelations WHERE (InvoiceItem = '" & ItemID & "' AND InvoiceItemCategory='" & ItemCategory & "')"
		set RS1=Conn.Execute (mySQL)
		if RS1.eof then
			mySQL="INSERT INTO InvoiceItemCategoryRelations (InvoiceItem, InvoiceItemCategory) VALUES ('" & ItemID & "', '" & ItemCategory & "')"
			Conn.Execute (mySQL)
		end if
		RS1.close
	end if

	response.write "<B>UpdatedÖ</B><BR>"

	myCriteria= "[name] LIKE N'"& ItemName & "'"
end if
if request("act")="search" OR request("act")="add" then
	'---------------------------------------------------------------SAM------------------------------------------------------
	fromCode = 0
	toCode = 0
	isEnabled = ""
	if isnumeric(request.form("from_box")) Then fromCode = cdbl(request.form("from_box"))
	if isnumeric(request.form("to_box")) then toCode = cdbl(request.form("to_box"))
	'if fromCode < 30000 then fromCode = 30000
	if toCode > 999999 then toCode = 99999
	if request.form("enable_chk") = "on" then isEnabled = " AND (Enabled = 1) "
	'response.write(toCode)
	mySQL="SELECT * FROM InvoiceItems WHERE ("& myCriteria & ") AND ([ID] BETWEEN " & fromCode & " AND " & toCode & ")" & isEnabled & "ORDER BY [ID]"
	set RS1=Conn.Execute (mySQL)
	if not RS1.eof then
		tmpCounter=0
%>
		<center>
		<br>
		<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#555588" width="60%" id="result">
			<TR bgcolor="#CCCCFF">
				<TD style="height:50px;">ﬂœ ¬Ì „</TD>
				<TD>‰«„ ¬Ì „</TD>
				<TD>›⁄«·</TD>
				<TD width=120>‰Ê⁄ ﬁÌ„  ê–«—Ì</TD>
				<TD>ﬁÌ„ </TD>
				<TD>ﬂœ ﬂ«·«/”—ÊÌ”</TD>
				<TD>„«·Ì«  »— «—“‘ «›“ÊœÂ</TD>
			</TR>
<%
		MaxListItems=9999
		Do while (not RS1.eof AND tmpCounter < MaxListItems)
			tmpCounter = tmpCounter + 1
			if tmpCounter mod 2 = 1 then
				tmpColor="#FFFFFF"
				tmpColor2="#FFFFBB"
			Else
				tmpColor="#DDDDDD"
				tmpColor2="#EEEEBB"
			End if 
			'alert(this.getElementByTagName('td').items(0).data);
%>
			<TR bgcolor="<%=tmpColor%>" style="cursor: pointer;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="copyInfo(this.rowIndex)">
				<TD style="height:50px;"><%=RS1("ID")%>
					<input type="hidden" value="<%=RS1("ID")%>">
					<input type="hidden" value="<%=RS1("Type")%>">
					</TD>
				<TD><span id="theName"><%=RS1("name")%>&nbsp;</span></TD>
				<TD align="center" dir="LTR"><INPUT disabled TYPE="checkbox" Name="theEnbl" <%if RS1("Enabled") then response.write "checked"%>></TD>
				<TD><%=getTypeName(RS1("Type"))%>&nbsp;</TD>
				<TD dir=LTR align=right><%=RS1("Fee")%>&nbsp;</TD>
				<TD dir=LTR align=right><%=RS1("RelatedInventoryItemID")%>&nbsp;</TD>
				<TD align='center'><INPUT disabled type='checkbox' name='theVat' <%if RS1("hasVat") then response.write "checked" %>></TD>
			</TR>
<%			RS1.moveNext
		Loop
		if not RS1.eof then
%>
			<TR bgcolor="#CCCCFF" >
				<TD style="height:50px;">*</TD>
				<TD colspan="3"> ⁄œ«œ ÃÊ«» Â« “Ì«œ «” .<br> Ã” ÃÊ —« „ÕœÊœ — ﬂ‰Ìœ.</TD>
			</TR>

<%		End if
%>
		</TABLE>
		</center>
<%		
	End if
end if
Conn.Close
%>

</font>
<SCRIPT LANGUAGE="JavaScript">
<!--
document.all.search_box.focus();
//-->
</SCRIPT>
<!--#include file="tah.asp" -->
