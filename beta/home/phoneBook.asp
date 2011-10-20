<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Home (0)
PageTitle= " œ› —  ·›‰"
SubmenuItem=3
if not Auth(0 , 3) then NotAllowdToViewThisPage()


%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%
function sqlSafeNoEnter (s)
  st=s
  st=replace(St,"'","`")
  st=replace(St,chr(34),"`")
  st=replace(St,vbCrLf," ")
  sqlSafeNoEnter=st
end function
'--------------------------------------------
function echoTels(tels)
	result=""
	if not IsNull(tels) then  
		tels = Replace(Replace(Replace(Replace(tels," ",","),"-",","),"*",","),"Ê",",")
		numbers = Split(tels,",")
		for each tel in numbers
			if Len(tel)>2 then
				result = result & "<a href='#' title=' »—«Ì  „«” »« ‘„«—Â "& tel &" ﬂ·Ìﬂ ﬂ‰Ìœ ' onclick=""dial('" & tel & "','" & session("exten") & "');"">"&tel&"</a> &nbsp;"
			else 
				result = result & tel & " &nbsp;"
			end if
		NEXT
	end if
	echoTels = result
end function
'-------------------------------------------------------
%>
<style>
	Input { font-family: Tahoma;font-size: 8pt;height:25px;}
	TextArea { font-family: Tahoma;font-size: 9pt;}
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
var noNextField = false;
function copyInfo(index){
	var myObj=document.getElementsByTagName("table").item('result').getElementsByTagName("tr").item(index);
	document.all.name_box.value=myObj.getElementsByTagName("td").item(1).innerText;
	document.all.tel_box.value=myObj.getElementsByTagName("td").item(2).innerText;
	document.all.address_box.value=myObj.getElementsByTagName("td").item(3).innerText;
	document.all.id_box.value=myObj.getElementsByTagName("input").item(0).value;
	document.all.name_box.select();
	document.all.name_box.focus();
}
function checkValidation(){
	return true;
}
function dial(tel,exten){
	window.showModalDialog('dial.asp?tel='+tel+'&exten='+exten,'dialogHeight:80px; dialogWidth:140px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');

}
//-->
</SCRIPT>
<br>
<font face="tahoma">
<TABLE border="0" cellspacing="0" cellpadding="2" dir="RTL" width="700" align="center">
<TR bgcolor="#AACCCC">
	<TD><FORM METHOD=POST ACTION="phoneBook.asp?act=search"><TABLE>
	<TR>
		<TD>‰«„:</TD>
		<TD><INPUT TYPE="text" NAME="search_box" maxlength="50" size="25" tabIndex="2"></TD>
		<TD><INPUT TYPE="submit" Name="Submit" Value="Ã” ÃÊ" style="width:100px; font-family:tahoma,arial;"></TD>
	</TR>
	</TABLE></TD></FORM>
</TR>
<TR bgcolor="#CCCCAA">
	<TD><FORM METHOD=POST ACTION="phoneBook.asp?act=add" onSubmit="return checkValidation();">
	<TABLE width="80%">
	<TR>
		<TD>‰«„</TD>
		<TD> ·›‰</TD>
		<TD>¬œ—”</TD>
		<TD>&nbsp;</TD>
	</TR>
	<TR>
		<TD><TEXTAREA NAME="name_box" rows="3" cols="30" <% if not Auth(0 , 6) then %> readonly<% end if %>></TEXTAREA></TD>
		<TD dir="LTR"><TEXTAREA NAME="tel_box" rows="3" cols="30"  <% if not Auth(0 , 6) then %> readonly<% end if %>></TEXTAREA></TD>
		<TD><TEXTAREA NAME="address_box" rows="3" cols="50"  <% if not Auth(0 , 6) then %> readonly<% end if %>></TEXTAREA></TD>
		<TD valign="bottom"><INPUT TYPE="hidden" NAME="id_box">
		<% if Auth(0 , 6) then %><INPUT TYPE="submit" Name="Submit" Value="–ŒÌ—Â" style="width:100px; font-family:tahoma,arial;"><% end if %>
		</TD>
	</TR>
	</TABLE></TD></FORM>
</TR>
</TABLE>
<%
'SELECT COUNT(*) AS qtty, company_name, customer_name, telephone FROM orders_trace GROUP BY company_name, customer_name, telephone
myCriteria= "REPLACE([name], ' ', '') LIKE REPLACE(N'%"& sqlSafeNoEnter(request("search_box")) & "%', ' ', '')"
'myCriteria= "REPLACE([name], ' ', '') LIKE REPLACE(N'%"& "”·" & "%', ' ', '')"
if request("act")="add" AND not (request.form("name_box")="" AND request.form("tel_box")="") then
	if request.form("id_box") = "" then
		Conn.Execute ("INSERT INTO PhoneBook ([name], [Tel], [Address]) VALUES (N'"& sqlSafeNoEnter(request.form("name_box"))& "', N'"& sqlSafeNoEnter(request.form("tel_box"))& "', N'"& sqlSafeNoEnter(request.form("address_box"))& "') ")
		response.write "<B>Added...</B><BR>"
	else
		Conn.Execute ("UPDATE PhoneBook SET [name]=N'"& sqlSafeNoEnter(request.form("name_box"))& "',[Tel]=N'"& sqlSafeNoEnter(request.form("tel_box"))& "',[Address]=N'"& sqlSafeNoEnter(request.form("address_box"))& "' WHERE ([ID]='"& sqlSafeNoEnter(request.form("id_box"))& "')")
		response.write "<B>Updated...</B><BR>"
	end if
myCriteria= "REPLACE([name], ' ', '') LIKE REPLACE(N'%"& sqlSafeNoEnter(request("name_box")) & "%', ' ', '')"
end if
if request("act")="search" OR request("act")="add" then

	set RS1=Conn.Execute ("SELECT * FROM PhoneBook WHERE ("& myCriteria & ") ORDER BY [name]")
	if not RS1.eof then
		tmpCounter=0
%>
		<center>
		<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#555588" width="80%" id="result">
			<TR bgcolor="#CCCCFF">
				<TD style="height:50px;">-</TD>
				<TD>‰«„</TD>
				<TD> ·›‰</TD>
				<TD>¬œ—”</TD>
			</TR>
<%		Do while (not RS1.eof AND tmpCounter < 50)
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
			<TR bgcolor="<%=tmpColor%>" style="cursor: hand;" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="copyInfo(this.rowIndex)">
				<TD style="height:50px;"><%=tmpCounter%><input type="hidden" name="theID" value="<%=RS1("ID")%>"></TD>
				<TD><span id="theName"><%=RS1("name")%>&nbsp;</span></TD>
				<TD align="center" dir="LTR"><span id="theTel"><%=echoTels(RS1("Tel"))%>&nbsp;</span></TD>
				<TD><span id="theAddr"><%=RS1("Address")%>&nbsp;</span></TD>
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
if request("act")="search" then
	'SELECT COUNT(*) AS qtty, company_name, customer_name, telephone FROM orders_trace GROUP BY company_name, customer_name, telephone
	myCriteria= "REPLACE([company_name], ' ', '') LIKE REPLACE(N'%"& sqlSafeNoEnter(request("search_box")) & "%', ' ', '') OR REPLACE([customer_name], ' ', '') LIKE REPLACE(N'%"& sqlSafeNoEnter(request("search_box")) & "%', ' ', '')"
	set RS1=Conn.Execute ("SELECT company_name, customer_name, telephone FROM orders_trace WHERE ("& myCriteria & ") GROUP BY company_name, customer_name, telephone ORDER BY [customer_name]")
	if not RS1.eof then
		tmpCounter=0
%>
		<br>
		<center>
		<TABLE border="1" cellspacing="0" cellpadding="2" dir="RTL" borderColor="#558855" width="80%" id="result2">
			<TR bgcolor="#CCFFCC">
				<TD style="height:50px;">-</TD>
				<TD>‰«„</TD>
				<TD>‘—ﬂ </TD>
				<TD> ·›‰</TD>
			</TR>
<%		Do while (not RS1.eof AND tmpCounter < 50)
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
			<TR bgcolor="<%=tmpColor%>" >
				<TD style="height:20px;"><%=tmpCounter%></TD>
				<TD><span><%=RS1("customer_name")%>&nbsp;</span></TD>
				<TD><span><%=RS1("company_name")%>&nbsp;</span></TD>
				<TD align="center" dir="LTR"><span><%=echoTels(RS1("telephone"))%>&nbsp;</span></TD>
			</TR>
<%			RS1.moveNext
		Loop
		if not RS1.eof then
%>
			<TR bgcolor="#CCFFCC" >
				<TD style="height:50px;">*</TD>
				<TD colspan="3"> ⁄œ«œ ÃÊ«» Â« “Ì«œ «” .<br> Ã” ÃÊ —« „ÕœÊœ — ﬂ‰Ìœ.</TD>
			</TR>

<%		End if
%>
		</TABLE>
		</center>
<%		
	End if
elseif request("act")="dial" then 
	call makeCall(request("tel"),request("exten"))
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
