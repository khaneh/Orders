<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= " «ÌÌœ ”‰œ"
SubmenuItem=2
if not Auth(8 , 2) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->

<%


'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ List ALl GL Docs
'-----------------------------------------------------------------------------------------------------
if request("act")="" then
	response.write "<BR><BR><BLOCKQUOTE>"

	set RSS=Conn.Execute ("SELECT GLDocID, GLDocDate FROM GLDocs WHERE (deleted = 0) and (IsChecked = 0) GROUP BY GLDocID, GLDocDate, GL HAVING (GL = "& OpenGL & ")")
	
	Do while not RSS.eof
		response.write "<li class=alak2> <A HREF='Approve.asp?act=ShowDoc&id="& RSS("GLDocID") & "'>”‰œ ‘„«—Â " & RSS("GLDocID") & " »Â  «—ÌŒ <span dir=ltr>"  & RSS("GLDocDate") & "</span></A>"
	RSS.moveNext
	Loop
	response.write "</BLOCKQUOTE>"

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- List GL Super Groups
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submit" then
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	id=request("id")
	if id="" or not(isnumeric(id)) then
		%>
		<BR><BR><CENTER>
		Œÿ«! ‘„«—Â ”‰œ Ê«—œ ‰‘œÂ «” 
		</CENTER>
		<%
		response.end
	end if

	if request("submit")=" «ÌÌœ" then
		Conn.Execute ("UPDATE GLDocs SET IsChecked = 1 WHERE (GLDocID = "& id & ")  AND (GL = "& OpenGL & ")")
		response.write "<BR><BR><BR><CENTER><H2> «ÌÌœ ‘œ</H2></CENTER>"
	else
		response.redirect "GLMemoInput.asp?act=editDoc&id="& id 
		response.write "<BR><BR><BR><CENTER><H2>Under Construction</H2></CENTER>"
	end if

'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------- List GL Super Groups
'-----------------------------------------------------------------------------------------------------
elseif request("act")="ShowDoc" then
id=request("id")
if id="" or not(isnumeric(id)) then
	%>
	<BR><BR><CENTER>
	Œÿ«! ‘„«—Â ”‰œ Ê«—œ ‰‘œÂ «” 
	</CENTER>
	<%
	response.end
end if
%>
<br><br>
<%
	mySQL="SELECT * FROM GLDocs WHERE (GLDocID = "& id & ")  AND (GL = "& OpenGL & ") and (deleted = 0)"
	set RS1=conn.execute(mySQL)
	if RS1.eof then
		response.write "Œÿ«! ç‰Ì‰ ”‰œÌ ÊÃÊœ ‰œ«—œ"
		response.end
	end if 

	GLDoc = RS1("ID")
	GLDocDate = RS1("GLDocDate")
	GLDocID = RS1("GLDocID")

	mySQL="SELECT GLAccounts.Name AS accTitle, GLRows.Description, GLRows.Amount, GLRows.IsCredit, GLRows.Tafsil, GLRows.GLAccount FROM GLAccounts INNER JOIN GLRows ON GLAccounts.ID = GLRows.GLAccount WHERE (GLRows.GLDoc = "& GLDoc & ") and (GLRows.deleted = 0)  AND (GLAccounts.GL = "& OpenGL & ")"
	set RS2=conn.execute(mySQL)

	%>
<CENTER><H2>”‰œ ‘„«—Â <%=id%> </H2>
<H4> «—ÌŒ ”‰œ: <span dir=ltr><%=GLDocDate%></span></H4>
<br>

</CENTER>
<TABLE class="GLTable2" Cellspacing="0" Cellpadding="0" width=90% align=center  Dir="RTL">
<tr>
	<td style="width:26; border-right:none;"> # </td>
	<td style="width:50; ">Õ”«»</td>
	<td style="width:170;">‰«„ Õ”«»</td>
	<td style="width:300;">‘—Õ</td>
	<!--td style="width:100;"> «—ÌŒ À» </td-->
	<td style="width:80;">»œÂﬂ«—</td>
	<td style="width:80;">»” «‰ﬂ«—</td>
</tr>
</table>
<TABLE Border="0" width=90%  align=center Cellspacing="1" Cellpadding="0" Dir="RTL" bgcolor="#558855" class="GLTable3">
<Tbody id="GLrows">
	<%

	Do while not RS2.eof
		i = i + 1
		'CreatedDate = RS1("CreatedDate")
		GLAccount = RS2("GLAccount")
		accTitle = RS2("accTitle")
		theDescription = RS2("Description")
		Amount = RS2("Amount")
		IsCredit = RS2("IsCredit")

		credit = ""
		debit = ""
		if IsCredit  then 
			credit = Amount 
			totalCredit = totalCredit + clng(Amount)
		else
			debit = Amount 
			totalDebit = totalDebit + clng(Amount)
		end if

		%>
		<tr bgcolor='#F0F0F0' >
		<td style="width:25; border-right:none;"> <%=i%> </td>
		<td style="width:50;"><%=GLAccount%></td>
		<td style="width:170;"><%=accTitle%></td>
		<td style="width:300;"><%=theDescription%></td>
		<!--td style="width:100;" > <span dir=ltr><%=CreatedDate%></span></td-->
		<td style="width:80;"><%=debit%></td>
		<td style="width:80;"><%=credit%></td>
		</tr>
		<%
	RS2.movenext
	loop	
%>
</Tbody></TABLE>
<table  width=90% align="center" Cellspacing="0" Cellpadding="0"  align=center dir=rtl>
<tr style="border:none;">
	<td style="width:26;"></td>
	<td style="width:50;"></td>
	<td style="width:170;"></td>
	<td style="width:300;">Ã„⁄</td>
	<td style="width:100;"></td>
	<td style="width:80;"><%=totalDebit%></td>
	<td style="width:80;"><%=totalCredit%></td>
</tr>
</TABLE>	
<BR><BR>
<CENTER>
<FORM METHOD=POST ACTION="Approve.asp?act=submit">
<INPUT TYPE="hidden" name="id" value="<%=id%>">
<INPUT TYPE="submit" name="submit" value=" «ÌÌœ" class="GenButton" onclick="return confirm('ÕﬁÌﬁ « „Ì ŒÊ«ÂÌœ «Ì‰ ”‰œ —«  «ÌÌœ ﬂ‰Ìœøø')">
<INPUT TYPE="submit" name="submit" value="ÊÌ—«Ì‘" class="GenButton">
<INPUT TYPE="button" name="submit" value="»—ê‘ " class="GenButton" onclick="window.location='Approve.asp'">
</FORM>
</CENTER>
<BR><BR>
<%
end if
%>
<!--#include file="tah.asp" -->
