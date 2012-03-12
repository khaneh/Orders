<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="„œÌ—Ì  »«‰ﬂ—Â«"
SubmenuItem=7
%>
<!--#include file="top.asp" -->
<BR><BR><BR><CENTER>
<%
function sqlSafeNoEnter (s)
  st=s
  st=replace(St,"'","`")
  st=replace(St,chr(34),"`")
  st=replace(St,vbCrLf," ")
  sqlSafeNoEnter=st
end function

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit New or Edit
'-----------------------------------------------------------------------------------------------------
if request("act")="submitBanker" then

	id = sqlSafeNoEnter(request.form("id"))
	bname = sqlSafeNoEnter(request.form("bname"))
	Responsible = sqlSafeNoEnter(request.form("Responsible"))

	if request.form("IsBankAccount")="on" then
		IsBankAccount=1
	else
		IsBankAccount=0
	end if

	if request.form("id") = "" then
		Conn.Execute ("INSERT INTO Bankers ([Name], IsBankAccount,  Responsible, LastCheckedDate, LastCheckedBy) VALUES ("&_
		"N'"& bname & "', '"& IsBankAccount & "','"& Responsible & "' , N'0', 0) ")
		'response.write "<B>«›“ÊœÂ ‘œ</B><BR>"
		set rsv=conn.execute ("select max(id) as maxID from Bankers where [Name]=N'" & bname & "'")
		id = rsv("maxID")
	else
		Conn.Execute ("UPDATE Bankers SET [name]=N'"& sqlSafeNoEnter(request.form("bname"))& "', [IsBankAccount]='"& IsBankAccount & "', [Responsible]='"& Responsible & "' WHERE ([ID]='"& id & "')")
		'response.write "<B>»Â —Ê“ ‘œ </B><BR>"
		Conn.Execute ("DELETE FROM BankerCheqStatusGLAccountRelation where (Banker = " & id & ") and GL = "  & OpenGL)
	end if

	for i=1 to request.form("GLaccounts").count 
		if not request.form("GLaccounts")(i) = "" then 
			Conn.Execute ("INSERT INTO BankerCheqStatusGLAccountRelation  (Banker, CheqStatus, GLAccount, GL) VALUES (" & id & ", " & request.form("GLaccountsID")(i) & "," & request.form("GLaccounts")(i) & "," & OpenGL & ")")	
		end if
	next

end if

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------- Add a new banker & Bankers List
'-----------------------------------------------------------------------------------------------------
EDITFLAG = false

if request("id") <> "" and isnumeric(request("id")) then 
	id = request("id")
	EDITFLAG = true
	set rsv=conn.execute ("SELECT Bankers.ID, Bankers.Name, Bankers.IsBankAccount, Bankers.CurrentBalance, Users.ID, Bankers.Responsible, Users.RealName FROM Bankers INNER JOIN Users ON Bankers.Responsible = Users.ID WHERE Bankers.ID=" & id ) 
	if rsv.eof then
		EDITFLAG = false
	else
		BankerName = rsv("Name")
		ID = rsv("ID")
		IsBankAccount = rsv("IsBankAccount") 
	end if
	rsv.close			
end if

%>
<TABLE width=80%>
<TR>
	<TD valign=top style="border:solid 1pt white">
		<TABLE border=0 width=100%>
		<form method=post action="?act=submitBanker">
		<TR bgcolor=white><TD align=center ><SPAN name="setGLsSpan" id="setGLsSpan">
		<% if EDITFLAG then%>
		ÊÌ—«Ì‘ «ÿ·«⁄«  
		<% else %>
		«›“Êœ‰   »«‰ﬂ— ÃœÌœ	
		<% end if %>
		</SPAN></TD></TR>
		<TR>
			<TD align=center>
					<BR>
					‰«„: <INPUT TYPE="hidden" name="id" value="<%=id%>"><input type="text" name="bname" value="<%=BankerName%>" size=35>&nbsp;&nbsp;&nbsp;
					„”∆Ê· ÅÌêÌ—Ì: 
								<select name="Responsible" class="custgeninput" style="width:90; font-family:tahoma; width:123;">
									<option value=""> «‰ Œ«» ﬂ‰Ìœ </option>
									<option value="">---------------------</option>
								<% set rsv=conn.execute ("SELECT * FROM Users WHERE Display=1 ORDER BY RealName") 
								do while not rsv.eof
								%>
									<option value="<%=rsv("ID")%>" <%
									if ID = rsv("ID") and not ID="" then
										response.write "selected" 
									end if 
									%>><%=rsv("realname")%></option>
								<%
								rsv.movenext
								loop
								rsv.close
								%>
								</select> &nbsp;&nbsp;&nbsp;
					Õ”«» »«‰ﬂÌ «” <INPUT TYPE="checkbox" NAME="IsBankAccount" onclick="changeGLsTD()" <%
					if IsBankAccount then response.write "checked" 
					%>>
					<BR><BR>
			</TD>
		</TR>
		<TR>
			<TD id="setGLsTD" name="setGLsTD">
			<% if EDITFLAG then %>
			<% if IsBankAccount then %>
				<table width=100%>
				<tr>
				<% set rsv=conn.execute ("select * from RcvPayChqStatus where IsBankAccount=1") 
				do while not rsv.eof 
					set rsw=conn.execute ("select * from BankerCheqStatusGLAccountRelation WHERE (CheqStatus = " & rsv("id") & ") AND (Banker = " & id & ") and GL=" & openGL) 
					%>
					<td align=center><%=rsv("Name")%>
					<br><INPUT TYPE='text' NAME='GLaccounts' size=7 value='<% if not rsw.eof then response.write rsw("GLAccount")%>'>
					<INPUT TYPE='hidden'  NAME='GLaccountsID'  value='<%=rsv("id")%>'><br>
					<% if rsv("IsRcvdChqStatus") then %> (œ—Ì«› ‰Ì) <% else %> (Å—œ«Œ ‰Ì)<% end if %>
					</td><%
					rsv.movenext
				loop
				rsv.close %>
				</tr></table>			
			<% else %>
				<table width=100%>
				<tr>
				<% set rsv=conn.execute ("select * from RcvPayChqStatus where IsBankAccount=0 and IsRcvdChqStatus = 0") 
				do while not rsv.eof 
					set rsw=conn.execute ("select * from BankerCheqStatusGLAccountRelation WHERE (CheqStatus = " & rsv("id") & ") AND (Banker = " & id & ") and GL=" & openGL) 
					%>
					<td align=center><%=rsv("Name")%><br><INPUT TYPE='text' NAME='GLaccounts' size=7 value='<% if not rsw.eof then response.write rsw("GLAccount")%>'>
					<INPUT TYPE='hidden'  NAME='GLaccountsID'  value='<%=rsv("id")%>'>
					</td><%
					rsv.movenext
				loop
				rsv.close %>
				</tr></table>			
			<% end if %>
			<% end if %>
			</TD>
		</TR>
		</TABLE>
		<BR>	
		<CENTER>
		<% if EDITFLAG then %>
		<INPUT TYPE="button" value="ÃœÌœ" onclick="window.location='BankerMng.asp'">
		<% end if %>
		<INPUT TYPE="submit" name="dokme" <% if EDITFLAG then %>value="ÊÌ—«Ì‘ «ÿ·«⁄«  "<% else %> value="«›“Êœ‰"<% end if %>></CENTER>
		<BR>	
		</TD>
		</form>
</TR>
<TR>
	<TD><BR><BR>
		<table id="result" width=100%>
		<tr style="background-color:white">
			<td>‰«„</td>
			<td align=center>„”∆Ê· ÅÌêÌ—Ì</td>
			<td align=center>»«·«‰” ›⁄·Ì</td>
			<td align=center>Õ”«» »«‰ﬂÌ «” ø</td>
		</tr>
		<%
		set rsv=conn.execute ("SELECT Bankers.ID as BID, Bankers.Name, Bankers.IsBankAccount,  Bankers.CurrentBalance, Users.ID, Bankers.Responsible, Users.RealName FROM Bankers INNER JOIN Users ON Bankers.Responsible = Users.ID ORDER BY Bankers.ID") 
		tmpColor2 = "#cccccc"
		tmpColor = ""
		do while not rsv.eof
		%>
		<a href="?id=<%=rsv("bid")%>">
		<tr style="cursor:hand" onMouseOver="this.style.backgroundColor='<%=tmpColor2%>'" onMouseOut="this.style.backgroundColor='<%=tmpColor%>'" onclick="copyInfo(this.rowIndex);">
			<td><INPUT TYPE="hidden" name="idList" value="<%=rsv("id")%>"><%=rsv("Name")%></td>
			<td><%=rsv("RealName")%> <INPUT TYPE="hidden" name=IDLIST value="<%=rsv("ID")%>"></td>
			<td><%=rsv("CurrentBalance")%></td>
			<td align=center><% if rsv("IsBankAccount") then %><INPUT TYPE="checkbox" NAME="isBank" disabled checked><% else %><INPUT TYPE="checkbox" NAME="isBank" disabled><% end if %></td>
		</tr>
		<%
		rsv.movenext
		loop
		rsv.close			
		%>
		</table>
		<BR>
	</TD>
</TR>
</TABLE>
</CENTER>
<SCRIPT LANGUAGE="JavaScript">
<!--

s1 = "<table width=100%><tr><% set rsv=conn.execute ("select * from RcvPayChqStatus where IsBankAccount=1") 
do while not rsv.eof
%><td align=center><%=rsv("Name")%><br><INPUT TYPE='text' NAME='GLaccounts' size=7><INPUT TYPE='hidden'  NAME='GLaccountsID'  value='<%=rsv("id")%>'><BR><% if rsv("IsRcvdChqStatus") then %> (œ—Ì«› ‰Ì) <% else %> (Å—œ«Œ ‰Ì)<% end if %></td><%
rsv.movenext
loop
rsv.close
%></tr></table>"

s2 = "<table width=100%><tr><% set rsv=conn.execute ("select * from RcvPayChqStatus where IsBankAccount=0 and IsRcvdChqStatus = 0") 
do while not rsv.eof
%><td align=center><%=rsv("Name")%><br><INPUT TYPE='text' NAME='GLaccounts' size=7><INPUT TYPE='hidden'  NAME='GLaccountsID'  value='<%=rsv("id")%>'></td><%
rsv.movenext
loop
rsv.close
%></tr></table>"



function copyInfo(index){
	var myObj=document.getElementsByTagName("table").item('result').getElementsByTagName("tr").item(index);
	document.all.bname.value				=myObj.getElementsByTagName("td").item(0).innerText;
	document.all.IsBankAccount.checked	=document.getElementsByName("isBank")[index-1].checked;
	document.all.id.value				=document.getElementsByName("idList")[index-1].value;
	document.all.Responsible.value				=document.getElementsByName("IDLIST")[index-1].value;
	document.all.dokme.value = "À»   €ÌÌ—« "
	document.all.setGLsSpan.innerText = "ÊÌ—«Ì‘ «ÿ·«⁄«  " + myObj.getElementsByTagName("td").item(0).innerText;
	if (document.getElementsByName("isBank")[index-1].checked)
		document.all.setGLsTD.innerHTML = s1;
	else
		document.all.setGLsTD.innerHTML = s2;
	document.all.bname.select();
	document.all.bname.focus();
}


function changeGLsTD(){
	if (document.all.IsBankAccount.checked)
		{
		s2 = document.all.setGLsTD.innerHTML;
		document.all.setGLsTD.innerHTML = s1;
		}
	else
		{
		s1 = document.all.setGLsTD.innerHTML;
		document.all.setGLsTD.innerHTML = s2;
		}
}


document.all.bname.focus();

<% if not EDITFLAG then %>
// by default, new bankers are BankAccount
document.all.setGLsTD.innerHTML = s1;
document.all.IsBankAccount.checked =  true ;
<% end if %>

//-->
</SCRIPT>
<!--#include file="tah.asp" -->