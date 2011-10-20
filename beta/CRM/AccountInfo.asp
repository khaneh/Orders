<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'CRM (1)
PageTitle="„‘ —Ì«‰"
SubmenuItem=1
if not Auth(1 , 1) then NotAllowdToViewThisPage()
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

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.GetCustTbl {font-family:tahoma; background-color: #DDDDDD; width:630; direction: RTL; }
	.GetCustTbl td {padding:2; font-size: 9pt; height:25;}
	.GetCustInp { font-family:tahoma; font-size: 9pt;}
	.CusTableHeader {background-color: #33AACC; text-align: center; font-weight:bold;}
	.CustContactTable {font-family:tahoma; width:100%; border:1 solid black; direction: RTL; background-color:#CCCCCC;}
	.CustContactTable td {padding:5;}
	.CustTable1 {font-family:tahoma; width:100%; height:100%; border:2px solid black; direction: RTL; background-color:#CCCCCC; }
	.CustTable1 a {text-decoration:none;color:#000088}
	.CustTable1 a:hover {text-decoration:underline;}
	.CustTable2 {font-family:tahoma; border:none; direction: RTL; width:100%; height:100%;}
	.CustTable3 {font-family:tahoma; border:1px solid black; direction: RTL; background-color:black; }
	.CustTable3 td {padding:5px;}
	.CustTable4 {font-family:tahoma; direction: RTL; width:100%; background-color:#C3DBEB;}
	.CustTableMenu {width:100%; border:none; direction: RTL;}
	.CustTableMenu td {border-bottom:1 solid black; height:25px; padding:5px;}
	.CustTableMenuSelected {background-color:#C3DBEB;}
	.CusTD1 {background-color: #CCCC66; text-align: left; font-weight:bold;}
	.CusTD2 {background-color: #DDDDDD; direction: LTR; text-align: right; font-size:9pt;}
	OPTION.mar{background-color:maroon; color:white;}
</STYLE>
<script language="javascript">
function dial(tel,exten){
	window.showModalDialog('../home/dial.asp?tel='+tel+'&exten='+exten,'dialogHeight:80px; dialogWidth:140px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');

}

</script>

<%
if request("act")="select" then
  if request.form("submitButton")="Ã” ÃÊ œ— Â„Â «ÿ·«⁄« " then
	if request("search") <> "" then
		SA_TitleOrName=Replace(Replace(request("search"),"Ì","Ì"),"ﬂ","ﬂ")
		SA_Action="return selectOperations();"
		SA_SearchAgainURL="AccountInfo.asp"
		SA_StepText="" '"ê«„ œÊ„ : «‰ Œ«» Õ”«»"
		SA_ActName		= "select"	
		SA_SearchBox	= "search"		

%>
		<FORM METHOD=POST ACTION="?act=show">
		<!--#include File="../AR/include_SelectAccountComplex.asp"-->
		</FORM>
<%
	end if
  '-------------------------------------------S A M-----------------------------------------
 
  elseif request("submitButton")="Ã” ÃÊ ÅÌ‘—› Â" then
  'response.write "test"
    'if request("search") <> "" then
		SA_TitleOrName=Replace(Replace(request("search"),"Ì","Ì"),"ﬂ","ﬂ")
		SA_Action="return selectOperations();"
		SA_SearchAgainURL="AccountInfo.asp"
		SA_StepText="" '"ê«„ œÊ„ : «‰ Œ«» Õ”«»"
		SA_ActName		= "select"	
		SA_SearchBox	= "search"		
		createDateFrom = request("createDateFrom")
		createDateTo = request("createDateTo")
		accountGroup = request("accountGroup")
		isPostable = request("isPostable")
		lastInvoiceDateFrom = request("lastInvoiceDateFrom")
		lastInvoiceDateTo = request("lastInvoiceDateTo")
		salesInvoiceDateFrom = request("salesInvoiceDateFrom")
		salesInvoiceDateTo = request("salesInvoiceDateTo")
		submitButton = request("submitButton")
		'response.write "test" & lastInvoiceDateFrom & "<br>" & lastInvoiceDateTo & "<br>"
%>
		<FORM METHOD=POST ACTION="?act=show">
		<!--#include File="../AR/include_SelectAccountAdvanced.asp"-->
		</FORM>
<%
	'end if
'-------------------------------------------S A M-----------------------------------------
  else
	if isnumeric(request("search")) then
		conn.close
		response.redirect "?act=show&selectedCustomer=" & request("search")
	elseif request("search") <> "" then
		SA_TitleOrName=Replace(Replace(request("search"),"Ì","Ì"),"ﬂ","ﬂ")
		SA_Action="return selectOperations();"
		SA_SearchAgainURL="AccountInfo.asp"
		SA_StepText="" '"ê«„ œÊ„ : «‰ Œ«» Õ”«»"
		SA_ActName		= "select"	
		SA_SearchBox	= "search"		

%>
		<FORM METHOD=POST ACTION="?act=show">
		<!--#include File="../AR/include_SelectAccount.asp"-->
		</FORM>
<%
	end if
  end if
elseif request("act")="show" then
	NextOf=request("NextOf")
	PrevOf=request("PrevOf")
	CusID=request("selectedCustomer")
	'---------------------------------------------S A M-------------------------------
	if request("searchtype")="advanced" then 
		createDateFrom = request("createDateFrom")
		createDateTo = request("createDateTo")
		accountGroup = request("accountGroup")
		isPostable = request("isPostable")
		lastInvoiceDateFrom = request("lastInvoiceDateFrom")
		lastInvoiceDateTo = request("lastInvoiceDateTo")
		salesInvoiceDateFrom = request("salesInvoiceDateFrom")
		salesInvoiceDateTo = request("salesInvoiceDateTo")
		submitButton = request("submitButton")
		if createDateFrom <> "" then 
			extraConditions = extraConditions & " AND CreatedDate >= '" & createDateFrom &"' "
		end if
		if createDateTo <> "" then 
			extraConditions = extraConditions & " AND CreatedDate <= '" & createDateTo &"' "
		end if
		if accountGroup <> "-1" then 
			extraConditions = extraConditions & " AND Accounts.ID IN (select account from accountGroupRelations where accountGroup=" & accountGroup & ")"
		end if
		if isPostable = "yes" then 
			extraConditions = extraConditions & " AND (Postable1=1 OR Postable2=1) "
		end if
		if isPostable = "no" then 
			extraConditions = extraConditions & " AND (Postable1=0 OR Postable2=0) "
		end if
		if lastInvoiceDateFrom <> "" then 
			extraConditions = extraConditions & " AND Accounts.ID IN (select distinct customer from Invoices where issuedDate >='" & lastInvoiceDateFrom & "' and voided=0 and issued=1 and IsReverse=0) "
		end if
		if lastInvoiceDateTo <> "" then 
			extraConditions = extraConditions & " AND Accounts.ID IN (select distinct customer from Invoices where issuedDate <='" & lastInvoiceDateTo & "' and voided=0 and issued=1 and IsReverse=0) "
		end if
	end if
	'---------------------------------------------S A M-------------------------------
	if NextOf <> "" AND isNumeric(NextOf) then
	
		NextOf=clng(NextOf)
		
		
		mySQL="SELECT TOP 1 AccountTitle FROM Accounts WHERE (ID = '"& NextOf & "')"
		
		set RS1=Conn.execute(mySQL)

		if RS1.EOF then
			response.write "<br>" 
			call showAlert ("Œÿ«Ì ⁄ÃÌ»! «Ì‰ „‘ —Ì œ— »«‰ﬂ «ÿ·«⁄«  ÅÌœ« ‰‘œ!",CONST_MSG_ERROR) 
			response.write "<br>" 
			response.end
		end if
		
		theTitle = RS1("AccountTitle")
		'mySQL="SELECT TOP 1 ID FROM Accounts WHERE (AccountTitle > N'"& theTitle & "') OR ((ID > '"& NextOf & "') AND (AccountTitle = N'"& theTitle & "')) Order BY AccountTitle, ID"
		'-------- changed by Alix (83-4-9) : be khaste mohaghegh tartibe hesaab haa az Alefbaa be Shomare Hesab tagheer kard
'---------------------------------------------S A M-------------------------------
		if request("searchtype")="advanced" then 
			mySQL="SELECT Accounts.*, Users.RealName AS CSRName FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID WHERE 1=1 "& extraConditions & " ORDER BY Accounts.AccountTitle"
		else
			mySQL="SELECT TOP 1 ID FROM Accounts WHERE (ID > '"& NextOf & "') Order BY ID"
		end if

		set RS1=Conn.execute(mySQL)
		if request("searchtype")="advanced" then
		 
			while not RS1.eof 
				if NextOf = RS1("ID") then 
					RS1.MoveNext
					if not Rs1.eof then
						CusID = RS1("ID")
						'response.write cusID
						LinkToNext="<a href=""?act=show&NextOf="& CusID & "&searchtype=advanced&createDateFrom="&createDateFrom&"&createDateTo="&createDateTo&"&accountGroup=" &accountGroup& "&isPostable=" &isPostable & "&lastInvoiceDateFrom="& lastInvoiceDateFrom &"&lastInvoiceDateTo="&lastInvoiceDateTo&"&submitButton="& submitButton &""">»⁄œÌ &gt;</a>"
					else
						CusID = NextOf
						LinkToNext="»⁄œÌ ‰œ«—œ"
					end if
				end if
				if not Rs1.eof then RS1.movenext
			wend
		else
			if not Rs1.eof then
				CusID = RS1("ID")
				LinkToNext="<a href=""?act=show&NextOf="& CusID & """>»⁄œÌ &gt;</a>"
			else
				CusID = NextOf
				LinkToNext="»⁄œÌ ‰œ«—œ"
			end if
		end if
		if request("searchtype")="advanced" then 
			LinkToPrev="<a href=""?act=show&PrevOf="& CusID & "&searchtype=advanced&createDateFrom="&createDateFrom&"&createDateTo="&createDateTo&"&accountGroup=" &accountGroup& "&isPostable=" &isPostable & "&lastInvoiceDateFrom="& lastInvoiceDateFrom &"&lastInvoiceDateTo="&lastInvoiceDateTo&"&submitButton="& submitButton &""">&lt; ﬁ»·Ì</a>"
		else
			LinkToPrev="<a href=""?act=show&PrevOf="& CusID & """>&lt; ﬁ»·Ì</a>"
		end if
	'--------------------------------------------------------------------------------
	elseif PrevOf <> "" AND isNumeric(PrevOf) then
		PrevOf = clng(PrevOf)
		mySQL="SELECT TOP 1 AccountTitle FROM Accounts WHERE (ID = '"& PrevOf & "')"
		set RS1=Conn.execute(mySQL)

		if RS1.EOF then
			response.write "<br>" 
			call showAlert ("Œÿ«Ì ⁄ÃÌ»! «Ì‰ „‘ —Ì œ— »«‰ﬂ «ÿ·«⁄«  ÅÌœ« ‰‘œ!",CONST_MSG_ERROR) 
			response.write "<br>" 
			response.end
		end if

		theTitle = RS1("AccountTitle")

		'mySQL="SELECT TOP 1 ID FROM Accounts WHERE (AccountTitle < N'"& theTitle & "') OR ((ID < '"& PrevOf & "') AND (AccountTitle = N'"& theTitle & "')) Order BY AccountTitle DESC, ID DESC"

		if request("searchtype")="advanced" then 
			mySQL="SELECT Accounts.*, Users.RealName AS CSRName FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID WHERE 1=1 "& extraConditions & " ORDER BY Accounts.AccountTitle"
		else
			mySQL="SELECT TOP 1 ID FROM Accounts WHERE (ID < '"& PrevOf & "') Order BY ID DESC"
		end if
		set RS1=Conn.execute(mySQL)
		if request("searchtype")="advanced" then 
			while not RS1.eof and break<>"exit"
				tmp=RS1("ID")
				if tmp=PrevOf then break="exit"
				RS1.MoveNext
				if RS1("ID")=PrevOf then break="exit"
			wend 
			if tmp=cusID then 
				CusID = PrevOf 
				LinkToPrev="ﬁ»·Ì ‰œ«—œ"
			else
				CusID = tmp
				LinkToPrev="<a href=""?act=show&PrevOf="& CusID & "&searchtype=advanced&createDateFrom="&createDateFrom&"&createDateTo="&createDateTo&"&accountGroup=" &accountGroup& "&isPostable=" &isPostable & "&lastInvoiceDateFrom="& lastInvoiceDateFrom &"&lastInvoiceDateTo="&lastInvoiceDateTo&"&submitButton="& submitButton &""">&lt; ﬁ»·Ì</a>"
			end if
		else
			if not Rs1.eof then
				CusID = RS1("ID")
				LinkToPrev="<a href=""?act=show&PrevOf="& CusID & """>&lt; ﬁ»·Ì</a>"
			else
				CusID = PrevOf 
				LinkToPrev="ﬁ»·Ì ‰œ«—œ"
			end if
		end if
		if request("searchtype")="advanced" then 
			LinkToNext="<a href=""?act=show&NextOf="& CusID & "&searchtype=advanced&createDateFrom="&createDateFrom&"&createDateTo="&createDateTo&"&accountGroup=" &accountGroup& "&isPostable=" &isPostable & "&lastInvoiceDateFrom="& lastInvoiceDateFrom &"&lastInvoiceDateTo="&lastInvoiceDateTo&"&submitButton="& submitButton &""">»⁄œÌ &gt;</a>"
		else
			LinkToNext="<a href=""?act=show&NextOf="& CusID & """>»⁄œÌ &gt;</a>"
		end if
	elseif CusID <> "" AND isNumeric(CusID) then
		CusID=clng(CusID)
		if request("searchtype")="advanced" then 
			LinkToNext="<a href=""?act=show&NextOf="& CusID & "&searchtype=advanced&createDateFrom="&createDateFrom&"&createDateTo="&createDateTo&"&accountGroup=" &accountGroup& "&isPostable=" &isPostable & "&lastInvoiceDateFrom="& lastInvoiceDateFrom &"&lastInvoiceDateTo="&lastInvoiceDateTo&"&submitButton="& submitButton &""">»⁄œÌ &gt;</a>"
			LinkToPrev="<a href=""?act=show&PrevOf="& CusID & "&searchtype=advanced&createDateFrom="&createDateFrom&"&createDateTo="&createDateTo&"&accountGroup=" &accountGroup& "&isPostable=" &isPostable & "&lastInvoiceDateFrom="& lastInvoiceDateFrom &"&lastInvoiceDateTo="&lastInvoiceDateTo&"&submitButton="& submitButton &""">&lt; ﬁ»·Ì</a>"
		else 
			LinkToNext="<a href=""?act=show&NextOf="& CusID & """>»⁄œÌ &gt;</a>"
			LinkToPrev="<a href=""?act=show&PrevOf="& CusID & """>&lt; ﬁ»·Ì</a>"
		end if
	else
		Conn.close
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â Õ”«» „⁄ »— ‰Ì” ")
	end if
  if CusID <> "" then
	mySQL="SELECT Users.RealName AS CSRName, Accounts.* FROM Accounts LEFT OUTER JOIN Users ON Accounts.CSR = Users.ID WHERE (Accounts.ID='"& CusID & "')"
	Set RS1=conn.execute(mySQL)
	if RS1.EOF then
		conn.close
		response.redirect "?errmsg=" & Server.URLEncode("‘„«—Â Õ”«» „⁄ »— ‰Ì” ")
	end if

AccountTitle = RS1("AccountTitle")
%>
	<TABLE class="CustTable1" cellspacing="0" cellpadding="0">
	<TR>
	<TD colspan="2" valign="top" height="30px;">
		<table class="CustTable2" cellspacing='1'>
		<tr>
			<td colspan="2" class="CusTableHeader">
				<TABLE width="100%" border="0" cellpadding="0" cellspacing="0">
				<TR class="CusTableHeader">
					<TD align="left"><%=LinkToPrev%></TD>
					<TD width='70%' align="center"><span style="font-weight:normal;font-size:13pt;"><%=RS1("AccountTitle")%></span><br><a href="javascript:void(0);" onclick="window.open ('oldCusProf.asp?acID=<%=CusID%>',null,'Width=500, height=600, scrollbars=yes, resizable=yes');" style="font-weight:normal;font-size:6pt">(«ÿ·«⁄«  ﬁ»·Ì)</a></TD>
					<TD><%=LinkToNext%></TD>
				</TR>
				</TABLE>
			</td>
		</tr>
		<tr>
			<td colspan="2">
<%	
	AccountIsDisabled = false
	if RS1("Status")<>1 then
			AccountIsDisabled = true
		select case RS1("Status")
			Case 2: StatusText = "„ Êﬁ›"
			Case 3: StatusText = "„”œÊœ"
		end select
%>		<TABLE width=70% align='center'>
		<TR>
			<TD align=center bgcolor=#FFBBBB style='border: dashed 2px Red'><BR><b><%=" ÊÃÂ!<br> «Ì‰ Õ”«» <U>"&StatusText&"</U> «” "%></b><BR><BR></TD>
		</TR>
		</TABLE>
<%
	end if
%>
			</td>
		</tr>
		</table>
	</TD>
	</TR>
	<TR>
	<TD width="100px" valign="top">

<%
tab = request("tab")
if tab = "" then 
	if session("tab")="" then
		tab = "1" 
	else
		tab = session("tab")
	end if
end if
if tab = "1" then
 css1 = "class='CustTableMenuSelected'"
elseif tab = "0" then
 css0 = "class='CustTableMenuSelected'"
elseif tab = "2" then
 css2 = "class='CustTableMenuSelected'"
elseif tab = "3" then
 css3 = "class='CustTableMenuSelected'"
elseif tab = "4" then
 css4 = "class='CustTableMenuSelected'"
elseif tab = "5" then
 css5 = "class='CustTableMenuSelected'"
elseif tab = "6" then
 css6 = "class='CustTableMenuSelected'"
else
 css1 = "class='CustTableMenuSelected'"
 tab = "1"
end if
session("tab") = tab
%>
		<TABLE width="100%" class="CustTableMenu" cellpadding="0" cellspacing="0">
			<TR><TD <%=css1%>><A HREF="AccountInfo.asp?tab=1&act=show&selectedCustomer=<%=cusID%>">„‘Œ’« </A></TD></TR>
			<TR><TD <%=css0%>><A HREF="AccountInfo.asp?tab=0&act=show&selectedCustomer=<%=cusID%>">ç«Å ›—„</A></TD><TR>
			<%if Auth(6 , 0) then %>
			<TR><TD <%=css2%>><A HREF="AccountInfo.asp?tab=2&act=show&selectedCustomer=<%=cusID%>">«ÿ·«⁄«  ›—Ê‘</A></TD></TR>
			<%end if %>
			<%if Auth(7 , 0) then %>
			<TR><TD <%=css3%>><A HREF="AccountInfo.asp?tab=3&act=show&selectedCustomer=<%=cusID%>">«ÿ·«⁄«  Œ—Ìœ</A></TD></TR>
			<%end if %>
			<%if Auth("B" , 0) then %>
			<TR><TD <%=css4%>><A HREF="AccountInfo.asp?tab=4&act=show&selectedCustomer=<%=cusID%>">«ÿ·«⁄«  ”«Ì—</A></TD></TR>
			<%end if %>
			<%if Auth(5 , 0) then %>
			<TR><TD <%=css6%>><A HREF="AccountInfo.asp?tab=6&act=show&selectedCustomer=<%=cusID%>">«ÿ·«⁄«  «‰»«—</A></TD></TR>
			<%end if %>
			<%if Auth(9 , 0) or Auth("A" , 0) then %>
			<TR><TD <%=css5%>><A HREF="AccountInfo.asp?tab=5&act=show&selectedCustomer=<%=cusID%>">œ—Ì«›  Å—œ«Œ  </A></TD></TR>
			<%end if %>
			<!--TR><TD>«ÿ·«⁄«  ”«œ”ÌÂ</TD></TR>
			<TR><TD>«ÿ·«⁄«  ”«»⁄ÌÂ</TD></TR>
			<TR><TD>«ÿ·«⁄«  À«„‰ÌÂ</TD></TR>
			<TR><TD>«ÿ·«⁄«   «”⁄ÌÂ</TD></TR>
			<TR><TD>«ÿ·«⁄«  ⁄«‘—ÌÂ</TD></TR-->
		</TABLE>
	</TD>
	<TD valign="top">
	  <TaBlE class="CustTable4" cellspacing="2" cellspacing="2">
	  <% if tab = "0" then %>
	  <!--#include file="include_CRM_PrintForm.asp"-->
	  <% elseif tab = "1" then %>
	  <!--#include file="include_CRM_PrimData.asp" -->
	  <% elseif tab = "2" then %>
	  <!--#include file="include_CRM_ARData.asp" -->
	  <% elseif tab = "3" then %>
	  <!--#include file="include_CRM_APData.asp" -->
	  <% elseif tab = "4" then %>
	  <!--#include file="include_CRM_AOData.asp" -->
	  <% elseif tab = "5" then %>
	  <!--#include file="include_CRM_Cash.asp" -->
	  <% elseif tab = "6" then %>
	  <!--#include file="include_CRM_Inv.asp" -->
	  <% end if %>
	  </TaBlE>
	</TD>
	</TR>
	</TABLE>
<%
  end if
else%>
	<br>
	<FORM METHOD=POST ACTION="?act=select" onsubmit="if ((document.all.search.value=='') && (document.getElementById('btnSearch').value!='Ã” ÃÊ ÅÌ‘—› Â')) return false;">
	<div dir='rtl'><B>ê«„ «Ê· : Ã” ÃÊ »—«Ì ‰«„ Ì« ‘„«—Â Õ”«»</B>
		<INPUT TYPE="text" NAME="search">&nbsp;
		<INPUT TYPE="submit" Name="submitButton" value="Ã” ÃÊ" id="btnSearch"> &nbsp; 
		<INPUT TYPE="submit" Name="submitButton" value="Ã” ÃÊ œ— Â„Â «ÿ·«⁄« " id='allData'> &nbsp;
		<input type="button" name="btnAdvanced" onclick="showAdvanced();" value="Ã” ÃÊ ÅÌ‘—› Â" id="btnAdvanced">
		<br>
	</div>
	<div id="advancedSearch" style='visibility:hidden;'>
		<table>
			<tr>
				<td> «—ÌŒ «ÌÃ«œ:</td>
				<td>«“</td>
				<td><input type="text" name="createDateFrom" maxlength=10 OnBlur="return acceptDate(this);"/></td>
				<td> «</td>
				<td><input type="text" name="createDateTo" maxlength=10 OnBlur="return acceptDate(this);"/></td>
			</tr>
			<tr>
				<td> «—ÌŒ ¬Œ—Ì‰ ”›«—‘:</td>
				<td>«“</td>
				<td><input type="text" name="lastInvoiceDateFrom" maxlength=10 OnBlur="return acceptDate(this);"/></td>
				<td> «</td>
				<td><input type="text" name="lastInvoiceDateTo" maxlength=10 OnBlur="return acceptDate(this);"/></td>
			</tr>
			<tr>
				<td>»«“Â Ã„⁄ ›—Ê‘:</td>
				<td>«“</td>
				<td><input type="text" name="salesInvoiceDateFrom" maxlength=10 OnBlur="return acceptDate(this);"/></td>
				<td> «</td>
				<td><input type="text" name="salesInvoiceDateTo" maxlength=10 OnBlur="return acceptDate(this);"/></td>
			</tr>
			<tr>
				<td>’‰›:</td>
				<td></td>
				<td colspan=3>
					<select name="accountGroup">
						<option value="-1"></option>
						<%
						set RS=conn.Execute ("SELECT * FROM accountGroups ORDER BY isPartner DESC, Name")
						Do while not RS.eof 
						
						%>
						<option value='<%=RS("ID")%>' class='<% if RS("isPartner")="True" then response.write "mar"%>'><%=RS("name")%></option>
						<%
							RS.movenext
						loop
						RS.close
						set RS=nothing
						%>
					</select>
				</td>
			</tr>
			<tr>
				<td>ﬁ«»· Å” :</td>
				<td></td>
				<td colspan=3>
					<input type="radio" name="isPostable" value="yes"><span>»«‘œ</span>
					<input type="radio" name="isPostable" value="no"><span>‰»«‘œ</span>
					<input type="radio" name="isPostable" value="all" checked><span>Â„Â</span>
				</td>
			</tr>
			
		</table>
	</div>
	</FORM>
	
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		document.all.search.focus();
	//-->
	</SCRIPT>
<%end if
conn.Close
%>
</font>
<script language="JavaScript">
<!--

function showAdvanced(){
document.getElementById("advancedSearch").style.visibility='visible';
document.getElementById("btnAdvanced").style.visibility='hidden';
document.getElementById("allData").style.visibility='hidden';
document.getElementById("btnSearch").value="Ã” ÃÊ ÅÌ‘—› Â";
}
function selectOperations(){
	var Arguments = new Array;
	notFound=true;
	for (i=0;i<document.getElementsByName("selectedCustomer").length;i++){
		if(document.getElementsByName("selectedCustomer")[i].checked){
			notFound=false;
		}
	}
	if (notFound)
		return false;
}

//-->
</script>

<!--#include file="tah.asp" -->
