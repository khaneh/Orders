<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= "����� �����"
SubmenuItem=3
if not Auth(8 , 3) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<style>
	a {color:black;}

	.GLTable0 { font-family:tahoma; font-size: 9pt; background-color: #006666; border:2 solid #003333; padding:0; }
	.GLTable0 TH { background-color: #006666; border: 0; font-size:16pt;border-bottom: 2 solid #003333;}
	.GLTable0 TD { background-color: #DDDDDD; border: 0; font-size:10pt;}
	.GLTable0TR1 td{ background-color: #008888; border: 0; }
	.GLTable0TR2 td{ background-color: #006666; border: 0; }

	.GLTable1 { font-family:tahoma; font-size: 9pt; background-color: #3399CC; border:2 solid #330099; padding:0; }
	.GLTable1 TH { background-color: #3399CC; border: 0; font-size:16pt;border-bottom: 2 solid #330099;}
	.GLTable1 TD { background-color: #EEEEEE; border: 0; font-size:10pt;}
	.GLTable1TR1 td{ background-color: #88BBDD; border: 0; }
	.GLTable1TR2 td{ background-color: #3399CC; border: 0; }

	.GLTable2 { font-family:tahoma; font-size: 9pt; background-color: #CC3399; border:2 solid #993300; padding:0; }
	.GLTable2 TH { background-color: #CC3399; border: 0; font-size:16pt;border-bottom: 2 solid #993300;}
	.GLTable2 TD { background-color: #EEEEEE; border: 0; font-size:10pt;}
	.GLTable2TR1 td{ background-color: #DD88BB; border: 0; }
	.GLTable2TR2 td{ background-color: #CC3399; border: 0; }

	.GLTable3 { font-family:tahoma; font-size: 9pt; background-color: #AAAA33; border:2 solid #CCCC00; padding:0; }
	.GLTable3 TH { background-color: #AAAA33; border: 0; font-size:16pt;border-bottom: 2 solid #BBBB00;}
	.GLTable3 TD { background-color: #EEEEEE; border: 0; font-size:10pt;}
	.GLTable3TR1 td{ background-color: #CCCC11; border: 0; }
	.GLTable3TR2 td{ background-color: #AAAA33; border: 0; }

	.GLTable4 { font-family:tahoma; font-size: 9pt; background-color: #33AA99; border:2 solid #005533; padding:0; }
	.GLTable4 TH { background-color: #33AA99; border: 0; font-size:16pt;border-bottom: 2 solid #005533;}
	.GLTable4 TD { background-color: #EEEEEE; border: 0; font-size:10pt;}
	.GLTable4TR1 td{ background-color: #77BBAA; border: 0; }
	.GLTable4TR2 td{ background-color: #33AA99; border: 0; }

	.GLTable5 { font-family:tahoma; font-size: 9pt; background-color: #66CCCC; border:2 solid #225555; padding:0; }
	.GLTable5 TH { background-color: #66BBCC; border: 0; font-size:16pt;border-bottom: 2 solid #225555;}
	.GLTable5 TD { background-color: #EEEEEE; border: 0; font-size:10pt;}
	.GLTable5TR1 td{ background-color: #99DDDD; border: 0; }
	.GLTable5TR2 td{ background-color: #66CCCC; border: 0; }

	.GenRow1 TD { background-color: #EEEEEE; border: 0; font-size:10pt;}
	.GenRow2 TD { background-color: #DDDDDD; border: 0; font-size:10pt;}


</style>
<%
function ShowErrorMessage(theText)
%>
	<br>
	<TABLE width=70% align='center'>
	<TR>
		<TD align=center bgcolor=#FFBBBB style='border: dashed 2pt Red'><BR><b><%=theText%></b><BR><BR></TD>
	</TR>
	</TABLE>
<%
end function

ON ERROR RESUME NEXT
	DocID=	clng(request("DocID"))
	if Err.Number<>0 then Err.clear:DocID=0
	id=		clng(request("id"))
	if Err.Number<>0 then Err.clear:id=0
ON ERROR GOTO 0

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------- Changing status of a GL DOC
'-----------------------------------------------------------------------------------------------------
if request("act")="changeStatus" then
	
	'----- Check GL is closed
	if (session("IsClosed")="True") then 
		Conn.close
		response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("���! ��� ���� ���� ���� ��� � ��� ���� �� ����� �� �� ������.")
	end if 
	'----

	submit=request("submit")

	if DocID=0 or id=0 then
		%>
		<BR><BR><CENTER>
		���! ����� ��� ���� ���� ���
		</CENTER>
		<%
		response.end
	end if

	if submit="���" then
		mySQL="SELECT min(GLDocID) as MinGLDocID FROM GLDocs where (GL = "& OpenGL & ")"
		set RS1=conn.execute(mySQL)

		if RS1.eof then
			response.write "<BR><BR><BR><CENTER>����� ������� �� ����� ������ ���.</CENTER>"
			response.end
		else
			RemovedGLDocID = RS1("MinGLDocID") - 1
			if RemovedGLDocID=0 then RemovedGLDocID=-1
		end if 

		Conn.Execute ("UPDATE GLDocs SET IsRemoved=1,RemovedDate=N'"& shamsiToday() & "',RemovedBy='"& session("ID") & "' WHERE (ID = "& id & ") AND (GL = "& OpenGL & ")")
		Conn.Execute ("UPDATE GLRows SET deleted=1 WHERE (GLDoc = "& id & ")")
		Conn.Execute ("UPDATE GLDocs SET GLDocID="& RemovedGLDocID & ", OldGLDocID="& DocID & " WHERE (GLDocID = "& DocID & ") AND (GL = "& OpenGL & ")")
		DocID = RemovedGLDocID
	elseif submit="������ ��� ��������" then
		conn.close
		response.redirect "SubsysDocsEdit.asp?act=show&id=" & id
	elseif submit="������ ��� ����" then
		response.redirect "CompMemoInput.asp?act=editDoc&id="& DocID 
	elseif submit="������" then
		response.redirect "GLMemoInput.asp?act=editDoc&id="& DocID 
	elseif submit="��� ��� ����" then ' ------------------ADD BY SAM
		mySQL="SELECT min(GLDocID) as MinGLDocID FROM GLDocs where (GL = "& OpenGL & ")"
		set RS1=conn.execute(mySQL)

		if RS1.eof then
			response.write "<BR><BR><BR><CENTER>����� ������� �� ����� ������ ���.</CENTER>"
			response.end
		else
			RemovedGLDocID = RS1("MinGLDocID") - 1
			if RemovedGLDocID=0 then RemovedGLDocID=-1
		end if 
		set myRSS=Conn.Execute("SELECT * FROM EffectiveGLRows  WHERE GLDoc="&id &" AND SYS>''")
		do while not myRSS.eof
			link	= myRSS("Link")
			sys		= myRSS("SYS")
			isCredit= myRSS("isCredit")
			Amount	= myRSS("Amount")
			Account = myRSS("Tafsil")
			Conn.Execute ("UPDATE " & sys & "Items SET RemainedAmount=0, FullyApplied=0,Voided=1,GL_Update=1 WHERE ID="& link)
			conn.Execute ("UPDATE " & sys & "Memo SET Voided=1 WHERE ID in (SELECT Link FROM "&sys&"Items WHERE Type=3 AND ID="&link&")")
			if isCredit then
				mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance - '"& Amount & "' WHERE (ID='"& Account & "')"
			else
				mySQL="UPDATE Accounts SET "& sys & "Balance = "& sys & "Balance + '"& Amount & "' WHERE (ID='"& Account & "')"
			end if
			myRSS.moveNext
		loop
		myRSS.close
		Conn.Execute ("UPDATE GLDocs SET IsRemoved=1,RemovedDate=N'"& shamsiToday() & "',RemovedBy='"& session("ID") & "' WHERE (ID = "& id & ") AND (GL = "& OpenGL & ")")
		Conn.Execute ("UPDATE GLRows SET deleted=1 WHERE (GLDoc = "& id & ")")

		Conn.Execute ("UPDATE GLDocs SET GLDocID="& RemovedGLDocID & ", OldGLDocID="& DocID & " WHERE (GLDocID = "& DocID & ") AND (GL = "& OpenGL & ")")
		DocID = RemovedGLDocID
		'---------------------------------------------------------
	else
		select case submit
			case "����� �� �������"
				fieldUpdate = " IsTemporary=0 , IsChecked=0, IsFinalized=0"
			case "����� �� ����" 
				fieldUpdate = " IsTemporary=1 , IsChecked=0, IsFinalized=0, CreatedDate=N'" & shamsiToday() & "'"
			case "����� ��" 
				fieldUpdate = " IsTemporary=0 , IsChecked=1, IsFinalized=0, CheckedDate=N'" & shamsiToday() & "', CheckedBy="& session("id")
			case "���� ��" 
				fieldUpdate = " IsTemporary=0 , IsChecked=0, IsFinalized=1, FinalizedDate=N'" & shamsiToday() & "', FinalizedBy="& session("id")
			case else
				response.write "<BR><BR><H1>Very Bad Error!</H1><BR>System Has Been Colapsed!@"
				response.end
		end select
		Conn.Execute ("UPDATE GLDocs SET "& fieldUpdate & " WHERE (ID = "& id & ") AND (GL = "& OpenGL & ")")
	end if
	
end if

itemAmount=""

if DocID <> 0 then
	mySQL="SELECT * FROM GLDocs WHERE GLDocID="& DocID & " and (GL = "& OpenGL & ") order by id DESC"
	
	set RS1=conn.execute(mySQL)

	if RS1.eof then
		response.write "<BR><BR><BR><CENTER>���! ���� ���� �� ��� ��� ���� ���� �����</CENTER>"
		response.end
	else
		id = clng(RS1("id"))
	end if 
end if

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- GL Doc Search Form
'-----------------------------------------------------------------------------------------------------
if id=0 then
	showDescription=false
	itemDesc=sqlSafe(replace(request("itemDesc")," ",""))

	if request("act")="search" and itemDesc<>"" then

		SQL_ST = "SELECT TOP 20 GLDocs.ID, GLDocs.GLDocID, GLDocs.GLDocDate, GLDocs.IsTemporary, GLDocs.IsChecked, GLDocs.IsFinalized, GLDocs.IsRemoved, GLDocs.BySubSystem, GLDocs.IsCompound, GLRows.Description FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc WHERE (GL = "& OpenGL & ") AND (REPLACE(GLRows.Description,' ','') LIKE N'%" & itemDesc & "%') ORDER BY GLDocDate DESC"

		showDescription=true
	
	else

		fromDate =	request("fromDate") 
		toDate =	request("toDate") 
		Draft =		request("Draft") 
		Temporary=	request("Temporary") 
		Checked=	request("Checked") 
		Finalized=	request("Finalized") 
		Removed =	request("Removed")
		Order1=		request("Order1")

		itemAmount=	request("itemAmount")

		if request("fromSession") = "y" then
			toDate =	session("toDate")
			fromDate =	session("fromDate")
			Draft =		session("Draft") 
			Temporary=	session("Temporary") 
			Checked =	session("Checked") 
			Finalized=	session("Finalized") 
			Removed =	session("Removed") 
			Order1 =	session("Order1") 
		end if

		if not Draft="" then DraftSt = "checked"
		if not Temporary="" then	TemporarySt = "checked"
		if not Checked="" then CheckedSt = "checked"
		if not Finalized="" then FinalizedSt = "checked"
		if not Removed="" then RemovedSt = "checked"

		if ""&Order1="" then Order1 = "GLDocs.ID DESC"

		if toDate="" or fromDate="" then
			toDate =		shamsiToday()
			fromDate =		shamsiDate(Date()-7)
			DraftSt =		"checked"
			TemporarySt =	"checked"
			CheckedSt =		"checked"
		end if

		session("fromDate") =	fromDate 
		session("toDate") =		toDate
		session("Draft") =		Draft
		session("Temporary") =	Temporary
		session("Checked") =	Checked
		session("Finalized") =	Finalized
		session("Removed") =	Removed
		session("Order1") =		Order1
		
'--
	
		if itemAmount<>"" and isnumeric(itemAmount) then
			itemAmount = cdbl(itemAmount)

			SQL_ST = "SELECT TOP 50 GLDocs.ID, GLDocs.GLDocID, GLDocs.GLDocDate, GLDocs.IsTemporary, GLDocs.IsChecked, GLDocs.IsFinalized, GLDocs.IsRemoved, GLDocs.BySubSystem, GLDocs.IsCompound, GLRows.Description FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc WHERE (GL = "& OpenGL & ") AND (GLRows.Amount = '" & itemAmount & "') "

			showDescription=true
		else
			itemAmount = ""
			SQL_ST = "SELECT * FROM GLDocs WHERE (GL = "& OpenGL & ") and (GLDocDate >= N'"& fromDate & "') and (GLDocDate <= N'"& toDate & "')"
		end if

		SQL_ST = SQL_ST & " and ( 1=0 "

		if RemovedSt = "checked" then SQL_ST = SQL_ST & " or (IsRemoved=1) "
		if FinalizedSt = "checked" then	SQL_ST = SQL_ST & " or ( (IsFinalized=1) and (GLDocs.deleted=0) and (IsRemoved=0) )"
		if CheckedSt = "checked" then SQL_ST = SQL_ST & " or ( (IsChecked=1) and (GLDocs.deleted=0) and (IsRemoved=0) )"
		if TemporarySt = "checked" then	SQL_ST = SQL_ST & " or ( (IsTemporary=1) and (GLDocs.deleted=0) and (IsRemoved=0) )"
		if DraftSt = "checked" then SQL_ST = SQL_ST & " or ( (IsTemporary=0) and (IsChecked=0) and (IsFinalized=0) and (GLDocs.deleted=0) and (IsRemoved=0) )"

	'	SQL_ST = SQL_ST & " ) order by GLDocDate, GLDocID, id"
		SQL_ST = SQL_ST & " ) ORDER BY "& order1 '& ", id"
	end if

	%>
	<BR><BR><CENTER>
	<FORM METHOD=POST ACTION="GLMemoDocShow.asp?act=">
	����� ���: <INPUT TYPE="text" NAME="DocID"> &nbsp;&nbsp;<INPUT TYPE="submit" value="������">
	</FORM>
	<hr>
	<FORM METHOD=POST ACTION="GLMemoDocShow.asp?act=search">
	���: <INPUT TYPE="text" NAME="itemDesc" value="<%=itemDesc%>" > &nbsp;&nbsp;<INPUT TYPE="submit" value="�����">
	</FORM>
	<hr>
	<FORM METHOD=POST ACTION="GLMemoDocShow.asp">
	<TABLE border=0 align=center>
	<TR>
		<TD align=left>�� �����</TD>
		<TD align=right><INPUT TYPE="text" NAME="fromDate" value="<%=fromDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
		<TD align=left width=20></TD>
		<TD align=left>�� �����</TD>
		<TD align=right><INPUT TYPE="text" NAME="toDate" value="<%=toDate%>" dir=ltr onKeyPress="return maskDate(this);" onblur="acceptDate(this)" maxlength="10"></TD>
	</TR>
	<TR height=10>
		<TD colspan=5 align=center><hr></TD>
	</TR>
	<TR>
		<TD colspan=5>
		<INPUT TYPE="checkbox" NAME="Draft" <%=DraftSt%>> ������� 
		<INPUT TYPE="checkbox" NAME="Temporary" <%=TemporarySt%>> ����
<%	if Auth( 8 , 6 ) then%>
		<INPUT TYPE="checkbox" NAME="Checked" <%=CheckedSt%>> ����� ��� 
<%	end if
	if Auth( 8 , 5 ) then%>
		<INPUT TYPE="checkbox" NAME="Finalized" <%=FinalizedSt%>> ���� ��� 
<%	end if
	if Auth( 8 , 2 ) then%>
		<INPUT TYPE="checkbox" NAME="Removed" <%=RemovedSt%>> ��� ���
<%	end if%>
	</TR>
	<TR height=10>
		<TD colspan=5 align=center><hr></TD>
	</TR>
	<TR>
		<TD align=left>����:</TD>
		<TD align=right colspan=4>
			<INPUT TYPE="text" NAME="itemAmount" value="<%=itemAmount%>" >
		</TD>
	</TR>
	<TR height=10>
		<TD colspan=5 align=center><hr></TD>
	</TR>
	<TR>
		<TD colspan=4 align=center>���� ��� �� ����:
			<SELECT NAME="Order1" style="font-family:tahoma;font-size:9pt;width:120px;" onchange="document.all.submit.click();">
				<OPTION Value="GLDocDate" <%if Order1="GLDocDate" then response.write "selected"%>>����� ���</OPTION>
				<OPTION Value="GLDocID" <%if Order1="GLDocID" then response.write "selected"%>>����� ���</OPTION>
				<OPTION Value="GLDocs.ID" <%if Order1="GLDocs.ID" then response.write "selected"%>>����� ��� ���</OPTION>
				<OPTION Value="GLDocs.ID DESC" <%if Order1="GLDocs.ID DESC" then response.write "selected"%>>����� �������</OPTION>
			</SELECT>
		</TD>
		<TD align=center>
			<INPUT TYPE="submit" NAME="submit" class=inputBut value="�����">
		</TD>
	</TR>
	</TABLE>
	</FORM>
	</CENTER>
	<BR><BR>
	<div style='padding:10px;'>
	<%
'--
	set rsNote=Conn.Execute("select * from GLDocs where IsTemporary=0 and IsChecked=0 and IsFinalized=0 and IsRemoved=0 and deleted=0 and gl="&OpenGL)
	if not rsNote.eof then response.write ("<center><h2 style='color:yellow;'>���� ���� ���! ��� ����� ������� �����. ���� �э� ����� ����� ���� �� ���� ����</h2></center>")
	do while not rsNote.eof 
		statusString = " ������� "
		status =""
		if rsNote("BySubSystem") then
			statusString = statusString & " - �������� "
			status = "BySubSystem"
		elseif rsNote("IsCompound") then
			statusString = statusString & " - ��� ���� "
		end if
		response.write("<li class=alak2> <A style='font-size:8pt; color:red;' HREF='GLMemoDocShow.asp?id=" & rsNote("ID") & "' target='_blank'>")
		if status = "BySubSystem" then response.write("<b>")
		response.write(" ��� ����� "& rsNote("GLDocID")& " �� ����� <span dir=ltr>" & rsNote("GLDocDate") & "</span> &nbsp;&nbsp;" & statusString)
		if status = "BySubSystem" then response.write("</b>")
		response.write ("</a></li>")
		rsNote.moveNext
	loop
	response.write "<hr>"
	set RSS=Conn.Execute (SQL_ST)
	status =""
	Do while not RSS.eof
		memoNumberString=" ��� ����� " & RSS("GLDocID")
		if RSS("IsRemoved") then 
										statusString = " ��� ��� "
										memoNumberString= " <U>��� ��� ��� </U>"
		elseif RSS("IsFinalized") then 		
										statusString = " ���� ��� "
		elseif RSS("IsChecked") then	
										statusString = " ����� ��� "
		elseif RSS("IsTemporary") then	
										statusString = " ���� "
		else							
										statusString = " ������� "
		end if 

		if RSS("BySubSystem") then
			statusString = statusString & " - �������� "
			status = "BySubSystem"
		elseif RSS("IsCompound") then
			statusString = statusString & " - ��� ���� "
		end if 
%>
		<li class=alak2> <A style='font-size:8pt;' HREF='GLMemoDocShow.asp?id=<%=RSS("ID")%>' target='_blank'>
		<%=memoNumberString%> �� ����� <span dir=ltr><%=RSS("GLDocDate")%></span> &nbsp;&nbsp;(<%=statusString%>)
<%
		if showDescription then
%>			<span style='font-size:7pt;color:gray;'>[<%=RSS("Description")%>]</span>
<%		end if
%>		</A>
<%		
	RSS.moveNext
	Loop
%>
	</div>
<%
else
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Show a GL Doc
'-----------------------------------------------------------------------------------------------------

	'mySQL="SELECT GLDocs.*, GLRows.GLAccount, GLRows.Tafsil, GLRows.Amount, GLRows.Description, GLRows.SYS, GLRows.Link, GLRows.IsCredit, GLAccounts.Name FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc INNER JOIN GLAccounts ON GLRows.GLAccount = GLAccounts.ID WHERE (GLDocs.id="& id & ") AND (GLDocs.GL = "& OpenGL & ") AND (GLAccounts.GL = "& OpenGL & ") ORDER BY GLRows.ID"
	' Changed By kid 860728 Adding Accounts.AccountTitle
	mySQL="SELECT GLDocs.*, GLRows.GLAccount, GLRows.Tafsil, GLRows.Amount, GLRows.Description, GLRows.SYS, GLRows.Link, GLRows.IsCredit, GLAccounts.Name, Accounts.AccountTitle FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc INNER JOIN GLAccounts ON GLRows.GLAccount = GLAccounts.ID LEFT OUTER JOIN Accounts ON GLRows.Tafsil = Accounts.ID WHERE (GLDocs.id="& id & ") AND (GLDocs.GL = "& OpenGL & ") AND (GLAccounts.GL = "& OpenGL & ") ORDER BY GLRows.ID"

	set RS1=conn.execute(mySQL)

	if RS1.eof then
	response.write "<BR><BR><BR><CENTER>���! ���� ���� �� ��� ��� ���� ���� �����</CENTER>"
		response.end
	end if 

	GLDocID = RS1("GLDocID")
	Creator = RS1("CreatedBy")

	classType=1

	if RS1("IsRemoved") then 
		if Not Auth( 8 , 2 ) then
			ShowErrorMessage("��� ���� �� ���� ��� ��� ������.")		
			conn.close
			response.end
		end if	
		statusString = " ��� ��� "
		status = "Removed"
		classType=2

	elseif RS1("IsFinalized") then 
		if Not Auth( 8 , 5 ) then
			ShowErrorMessage("��� ���� �� ���� ��� ��� ������.")		
			conn.close
			response.end
		end if	
		statusString = " ���� ��� "
		status = "Finalized"

	elseif RS1("IsChecked") then 
		if Not Auth( 8 , 6 ) then
			ShowErrorMessage("��� ���� �� ���� ��� ��� ������.")		
			conn.close
			response.end
		end if	
		statusString = " ����� ��� "
		status = "Checked"

	elseif RS1("IsTemporary") then 
		if Not (session("ID") = Creator OR Auth( 8 , 7 ) )then
			ShowErrorMessage("��� ���� �� ���� ��� ��� ������.")		
			conn.close
			response.end
		end if	
		statusString = " ���� "
		status = "Temporary"
		classType=1

	else
		if Not (session("ID") = Creator OR Auth( 8 , 9 ) )then
			ShowErrorMessage("��� ���� �� ���� ��� ������� ������.")		
			conn.close
			response.end
		end if	
		statusString = " ������� "
		status = "Draft"
		classType=3
	end if 

	isBySubSystem = false

	if RS1("BySubSystem") then
		isBySubSystem = true
		if RS1("IsRemoved") then
			classType=2
		else
			classType=4
		end if
	end if

	isCompound = false

	if RS1("IsCompound") then
		isCompound = true
		if RS1("IsRemoved") then
			classType=2
		else
			classType=5
		end if
	end if 

	if isBySubSystem then
		stamp="<div style='width:220px; text-align:center; border:2 dashed green;padding:10px;font-size:12pt;font-weight:bold;'>��� �������� ��� ����� <br>("&statusString&")</div>"
	elseif isCompound then
		stamp="<div style='width:220px; text-align:center; border:2 dashed #9966CC;padding:10px;font-size:12pt;font-weight:bold;'>��� ����<br>("&statusString&")</div>"
	else
		stamp="<div style='width:120px; text-align:center; border:2 dashed red;padding: 10px;color:red;font-size:14pt;font-weight:bold;'>"&statusString&"</div>"
	end if 

	if RS1("deleted") then 
		statusString = statusString & " (�����) "
		stamp="<div style='border:2 dashed red;padding: 10px;color:red;font-size:14pt;font-weight:bold;'>�����</div>"
		status = "deleted"
		classType=0
	end if 

	%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function showAcc(acc,moeen){
		window.open('tafsili.asp?accountID='+acc+'&FromDate=<%=RS1("GLDocDate")%>&ToDate=<%=RS1("GLDocDate")%>&moeenFrom='+moeen+'&moeenTo='+moeen+'&act=Show');
	}
	function showGLAcc(num){
		window.open('moeen.asp?accountID='+num+'&FromDate=<%=RS1("GLDocDate")%>&ToDate=<%=RS1("GLDocDate")%>&act=Show');
	}
	//-->
	</SCRIPT>
	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr height="20">
			<td colspan=2></td>
		</tr>
		<tr height="10">
			<td width="250"></td>
			<td valign="top"><div style='position:absolute;'><%=stamp%></div></td>
		</tr>
		<tr height="20">
			<td colspan=2></td>
		</tr>
	</table>

	<table Border="0" align=center Cellspacing="1" Cellpadding="5" Dir="RTL" class="GLTable<%=classType%>">
	<tr>
		<th colspan=6>
<%		if RS1("GLDocID")>0 then 
			response.write "��� ����� " & RS1("GLDocID")
		else 
			response.write "��� ��� (����� ����: "& RS1("OldGLDocID") & ")"
		end if 
%>
		</th>
	</tr>
	<tr>
		<th colspan=7 style="padding:0">
		<TABLE width="100%" border=0 Cellspacing="1" Cellpadding="10" Dir="RTL">
		<TR class="GLTable<%=classType%>TR2">
			<TD>����� ���: <%= RS1("id")%></TD>
			<TD align=center>�����: <%=statusString%></TD>
			<TD align=left>����� ���: <span dir=ltr><%= RS1("GLDocDate")%></span></TD>
		</TR>
		</TABLE>
		</th>
	</tr>
	<tr class="GLTable<%=classType%>TR1">
		<td style="width:20;"> # </td>
		<td style="width:40;">������</td>
		<td style="width:30;">����</td>
		<td style="width:300;">���</td>
		<td style="width:70;">������</td>
		<td style="width:70;">��������</td>
	</tr>
<%
	i=0
	Do while not RS1.eof
		i = i + 1
		GLAccount = RS1("GLAccount")
		AccTitle = RS1("AccountTitle")
		GLAccTitle = RS1("name")
		theDescription = RS1("Description")
		Amount = Separate(RS1("Amount"))
		IsCredit = RS1("IsCredit")
		Tafsil = RS1("Tafsil")
		Sys =	RS1("Sys")
		ItemLink =	RS1("Link")

		if len(sys)>1 then  'if isBySubSystem then
			theLink = "ShowItem.asp?sys=" & sys & "&Item=" & ItemLink
		else
			theLink = "about:blank"
		end if

		credit = ""
		debit = ""
		if IsCredit then 
			credit = Amount 
			totalCredit = totalCredit + cdbl(Amount)
		else
			debit = Amount 
			totalDebit = totalDebit + cdbl(Amount)
		end if

		%>
		<tr <%if i mod 2 = 1 then%>class="GenRow1"<%else%>class="GenRow2"<%end if%>>
			<td style="width:20;"> <%=i%> </td>
			<td style="width:40;" title="<%=AccTitle%>"><A HREF="javascript:showAcc(<%=Tafsil%>,<%=GLAccount%>);" style="color:black;"><%=Tafsil%></A></td>
			<td style="width:30;" title="<%=GLAccTitle%>"><A HREF="javascript:showGLAcc(<%=GLAccount%>);"><%=GLAccount%></A></td>
			<td style="width:300; cursor:hand;" title="���� ����� ��� ������ ���� ����." onclick="window.open('<%=theLink%>');"><%=theDescription%></td>
			<td style="width:80;"><%=debit%></td>
			<td style="width:80;"><%=credit%></td>
		</tr>
		<%
	RS1.movenext
	loop	
%>
	<tr class="GLTable<%=classType%>TR1">
		<td></td>
		<td></td>
		<td></td>
		<td>���</td>
		<td><%=Separate(totalDebit)%></td>
		<td><%=Separate(totalCredit)%></td>
	</tr>
	</TABLE>
	<BR><BR>
	<CENTER>
	<FORM METHOD=POST ACTION="?act=changeStatus&msg=����� ����� ��� ����� ��">
		<INPUT TYPE="hidden" name="DocID" value="<%=GLDocID%>">
		<INPUT TYPE="hidden" name="id" value="<%=id%>">
	<% if status <> "deleted" then 
		if status = "Draft" then 
			if (((session("ID") = Creator) OR ( Auth( 8 , "H" ))) AND totalDebit=totalCredit) OR isBySubSystem then%>
				<INPUT TYPE="submit" name="submit" value="����� �� ����" class="GenButton" onclick="return confirm('����� ����� ��� ����� ��Ͽ')"><%
			end if

			if isBySubSystem AND Auth( 8 , 8 ) then %>
				<INPUT TYPE="submit" name="submit" value="������ ��� ��������" class="GenButton" style="border-color:red;">
		<%	elseif isCompound then
				if Auth( 8 , "F" ) then %>
					<INPUT TYPE="submit" name="submit" value="������ ��� ����" class="GenButton" style="border-color:red;">
					<INPUT TYPE="submit" name="submit" value="��� ��� ����" class="GenButton" style="border-color:red;" onclick="return confirm('��� ����� ����� �� �� ������ ��� ��� �� ��� ���Ͽ')">
		<%		end if
			else %>
				<INPUT TYPE="submit" name="submit" value="������" class="GenButton" style="border-color:red;">
				<INPUT TYPE="submit" name="submit" value="���" class="GenButton" style="border-color:red;" onclick="return confirm('��� ����� ����� �� �� ������ ��� ��� �� ��� ���Ͽ')"><%
			end if 

		elseif status = "Temporary" then 
			if session("ID") = Creator OR Auth( 8 , 8 ) then%>
				<INPUT TYPE="submit" name="submit" value="����� �� �������" class="GenButton" onclick="return confirm('����� ����� ��� ����� ��Ͽ')"><%
			end if
			if Auth( 8 , 8 ) then%>
				<INPUT TYPE="submit" name="submit" value="����� ��" class="GenButton" onclick="return confirm('����� ����� ��� ����� ��Ͽ')"><%

				if isBySubSystem AND Auth( 8 , 8 ) then %>
					<INPUT TYPE="submit" name="submit" value="������ ��� ��������" class="GenButton" style="border-color:red;">
			<%	elseif isCompound then
					if Auth( 8 , "F" ) then %>
						<INPUT TYPE="submit" name="submit" value="������ ��� ����" class="GenButton" style="border-color:red;">
						<INPUT TYPE="submit" name="submit" value="��� ��� ����" class="GenButton" style="border-color:red;" onclick="return confirm('��� ����� ����� �� �� ������ ��� ��� �� ��� ���Ͽ')">
			<%		end if
				else %>
					<INPUT TYPE="submit" name="submit" value="������" class="GenButton" style="border-color:red;">
					<INPUT TYPE="submit" name="submit" value="���" class="GenButton" style="border-color:red;" onclick="return confirm('��� ����� ����� �� �� ������ ��� ��� �� ��� ���Ͽ')"><%
				end if 

			end if
		elseif status = "Checked" then 	%>
			<INPUT TYPE="submit" name="submit" value="����� �� ����" class="GenButton" onclick="return confirm('����� ����� ��� ����� ��Ͽ')">
			<INPUT TYPE="submit" name="submit" value="���� ��" class="GenButton" onclick="return confirm('����� ����� ��� ����� ��Ͽ')">
	<%	end if 
	%>
	<br><br>
	<% if status <> "Draft" then %>
		<% 	ReportLogRow = PrepareReport ("GLDoc.rpt", "GLDoc_ID", id, "/beta/dialog_printManager.asp?act=Fin") %>
		<INPUT TYPE="button" value=" �ǁ " Class="GenButton" style="border:1 solid blue;" onclick="printThisReport(this,<%=ReportLogRow%>);">
		<!--BR><iframe name=f1 id=f1 src="/CRReports/?Id=<%=ReportLogRow%>" align=center style="width:750; height:210; border-style: none" border=0 FRAMEBORDER=0 scrollbars=no ></iframe-->
	<% end if %>
	<INPUT TYPE="button" name="submit" value="������" class="GenButton" onclick="window.location='GLMemoDocShow.asp?fromSession=y'">
	<% 
	end if %>
	</FORM>

	<br>
	<hr noshade color=white>
	<%
	mySQL="SELECT GLDocs.*, Users.RealName FROM GLDocs INNER JOIN Users ON GLDocs.createdBy = Users.ID where GLDocID="& GLDocID & " and GL = "& OpenGL & " ORDER BY GLDocs.ID"

	set RS1=conn.execute(mySQL)

	if not RS1.eof then
		response.write " ��� ��� �� ����� ��� ��� ������ ��� ���: <br><br>"
		do while not RS1.EOF
			response.write "<li class=alak2 "
			if trim(id) = trim(RS1("id")) then 
				response.write " disabled "
			end if
			response.write "><A HREF='GLMemoDocShow.asp?id="& RS1("id") & "'> <span dir=ltr>"& RS1("CreatedDate") & "</span> ���� "& RS1("RealName") & " </a> <br>"
			RS1.movenext
		loop
	end if 

end if
%>
	<hr noshade color=white>
	<BR>
</center>
<!--#include file="tah.asp" -->
