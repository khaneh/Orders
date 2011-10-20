<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= "”‰œ Â«Ì “Ì— ”Ì” „"
SubmenuItem=9
if not Auth(8 , "C") then NotAllowdToViewThisPage() ' this is the same as viewing subsystem memos
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.GLTable {font-family:tahoma; background-color:#330099;border:solid 1px Black;}
	.GLTable TR {background-color:#F0F0F0;}
	.HeaderTD {font-size:12pt;font-weight:bold;background-color:#FFCC66;}
	.RowsTD {font-size:9pt;padding-bottom:10px;}
	.GeneralInput {width:70px; font-family:tahoma; font-size:8pt; border:1pt solid gray; background:transparent; direction: LTR; }
</STYLE>
<%
if request("act")="submitEdit" then

	id=request("id")
	if id="" or not isnumeric(id) then
		call showAlert ("‘„«—Â ”‰œ €·ÿ «” .",CONST_MSG_ERROR) 
		Conn.close
		response.end
	end if
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	mySQL="SELECT * FROM GLDocs WHERE id="& id & " and  (GLDocs.GL = "& OpenGL & ")"
	set RS1=conn.execute(mySQL)
	errMessage=""
	if RS1.eof then
		errMessage="Œÿ«! ç‰Ì‰ ”‰œÌ  œ— «Ì‰ ”«· „«·Ì ÊÃÊœ ‰œ«—œ."
	else
		GLMemoNo=RS1("GLDocID")
		GLMemoDate=RS1("GLDocDate")
		if RS1("IsTemporary") then
			IsTemporary=1
		else
			IsTemporary=0
		end if

		if not RS1("BySubSystem") then
			errMessage="«Ì‰ Ìﬂ ”‰œ « Ê„« Ìﬂ ‰Ì” ."
		elseif RS1("deleted") OR RS1("IsRemoved") then
			errMessage="«Ì‰ ”‰œ Õ–› ‘œÂ Ê «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ."
		elseif RS1("IsFinalized") then
			errMessage="«Ì‰ ”‰œ ﬁÿ⁄Ì ‘œÂ Ê «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ."
		elseif RS1("IsChecked") then
			errMessage="«Ì‰ ”‰œ »——”Ì ‘œÂ Ê «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ."
		end if 
	end if

	if errMessage<>"" then
		response.write "<br><br>"
		call showAlert (errMessage,CONST_MSG_ERROR) 
		Conn.close
		response.end
	end if
	RS1.close


	GLDoc = id
	creationDate = shamsiToday()

	totalRows = request.form("GLDocRows").count
	'--- Checking For Errors:
	if totalRows > 0 then
		mySQL="SELECT DISTINCT LEFT(Description, 5) AS Des, SYS, Link FROM GLRows INNER JOIN GLDocs ON GLRows.GLDoc = GLDocs.ID WHERE (GLRows.ID IN ("
		for i=1 to totalRows - 1
			mySQL = mySQL & request.form("GLDocRows")(i) & ", "
		next
		mySQL = mySQL & request.form("GLDocRows")(i) & ")) AND (GLDocs.ID = '"& GLDoc & "') AND (GLDocs.GL = '"& OpenGL & "') AND (GLRows.Deleted = 0)"
	else
		response.write "<br>ÂÌç ŒÿÌ «‰ Œ«» ‰‘œÂ..."
		response.end
	end if

	set RS1=conn.execute(mySQL)

	errorFound = false
	Do While not RS1.eof AND not errorFound 
		if RS1("Des")<>"«»ÿ«·" then
			mySQL2="SELECT Voided, GL_Update FROM "& trim(RS1("SYS")) & "Items WHERE (ID = "& RS1("Link") & ")"
			set RS2=conn.execute(mySQL2)
			if RS2("Voided") AND not RS2("GL_Update") then
				errSYS=RS1("SYS")
				errLink=RS1("Link")
				errorFound = true
				RS2.close
				exit do
			end if
			RS2.close
		end if
		RS1.moveNext
	Loop
	RS1.close
	
	if errorFound then
		response.write "<br>Error Found!:<br>" 
		response.write "<br>SYS:" & errSYS
		response.write "<br>Link:" & errLink
		response.end
	end if
	'--- End of checking for errors.
	'--- 

	'--- Actual Deleting 
	set RS1=conn.execute(mySQL)
	Do While not RS1.eof 
		tempDes =	RS1("Des")
		tempSYS =	trim(RS1("SYS"))
		tempLink=	RS1("Link")
		if tempDes="«»ÿ«·" then
			' This is a reverse memo
			mySQL2="SELECT * FROM GLRows WHERE (SYS = '"& tempSYS & "') AND (Link = "& tempLink & ") AND (deleted = 0) AND (Description NOT Like N'«»ÿ«·%')"
			set RS2=conn.execute(mySQL2)
			if RS2.eof then
				' No Direct Memo, No need to Update GL
				tempGL_Update=0
			else
				' Has A Direct Memo. Must be added to GL later
				tempGL_Update=1
			end if
		else
			' This is a direct memo
			mySQL2="SELECT Voided, GL_Update FROM "& tempSYS & "Items WHERE (ID = "& tempLink & ")"
			set RS2=conn.execute(mySQL2)
			if RS2("Voided") then
				' Item is voided. No need to Update GL
				tempGL_Update=0
			else
				' Item is not voided. Must be added to GL later
				tempGL_Update=1
			end if
			RS2.close
		end if

		mySQL2="UPDATE "& tempSYS & "Items SET GL_Update="& tempGL_Update & " WHERE (ID = "& tempLink & ")"
		Conn.Execute (mySQL2)

		mySQL2="UPDATE GLRows SET deleted = 1 WHERE (GLDoc='"& GLDoc & "') AND (deleted = 0) AND (SYS = '"& tempSYS & "') AND (Link = "& tempLink & ")"
		Conn.Execute (mySQL2)

		RS1.moveNext
	Loop
	RS1.close
	'--- End of actual deleting 
	'--- 

	mySQL = "SELECT * FROM GLRows WHERE (GLDoc='"& GLDoc & "') AND (deleted = 0)"
	set RS1 = Conn.execute(mySQL)
	if RS1.eof then
		RS1.close

		mySQL="SELECT MIN(GLDocID) AS MinGLDocID FROM GLDocs WHERE (GL = "& OpenGL & ")"
		set RS1=conn.execute(mySQL)
		RemovedGLDocID = clng(RS1("MinGLDocID")) - 1
		RS1.close
		Conn.Execute("UPDATE GLDocs SET IsRemoved=1, RemovedDate=N'"& shamsiToday() & "', RemovedBy='"& session("ID") & "', OldGLDocID=GLDocID, GLDocID="& RemovedGLDocID & "  WHERE (ID = "& GLDoc & ")")

		response.redirect "GLMemoDocShow.asp?id=" & GLDoc & "&msg=" & Server.URLEncode("Õ–› ”‰œ « Ê„« Ìﬂ «‰Ã«„ ‘œ.")
	else
		RS1.close
		'---- Creating a new GLDoc 
		mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, createdBy, BySubSystem, IsTemporary) VALUES ("& openGL & " , "& GLMemoNo & ", N'"& GLMemoDate & "' , N'"& creationDate & "', "& session("ID") & ", 1, "& IsTemporary & ");SELECT @@Identity AS NewGLDoc"
		set RS1 = Conn.execute(mySQL).NextRecordSet
		NewGLDoc = RS1("NewGLDoc")
		RS1.close
		'---- 

		'---- Inserting new GLRows 
		mySQL="INSERT INTO GLRows ( GLDoc, GLAccount, Tafsil, Amount, Description, Ref1, Ref2, SYS, Link, IsCredit, deleted) SELECT '"& NewGLDoc & "' AS GLDoc, GLAccount, Tafsil, Amount, Description, Ref1, Ref2, SYS, Link, IsCredit, deleted FROM GLRows WHERE (GLDoc = "& GLDoc & ") AND (Deleted=0) "
		conn.Execute(mySQL)

		'---- Marking old GLDoc and its remaining GLRows as DELETED
		conn.Execute("UPDATE GLRows SET deleted = 1 WHERE (GLDoc = "& GLDoc & ")")
		conn.Execute("UPDATE GLDocs SET deleted = 1 WHERE (ID = "& GLDoc & ")")
		'---- 

		response.redirect "GLMemoDocShow.asp?id=" & NewGLDoc & "&msg=" & Server.URLEncode("ÊÌ—«Ì‘ ”‰œ « Ê„« Ìﬂ «‰Ã«„ ‘œ.")
	end if

	'---- 

'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Show a GL Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="show" then

	id=request("id")
	if id="" or not isnumeric(id) then
		call showAlert ("‘„«—Â ”‰œ €·ÿ «” .",CONST_MSG_ERROR) 
		Conn.close
		response.end
	end if

	mySQL="SELECT * FROM GLDocs WHERE id="& id & " and  (GLDocs.GL = "& OpenGL & ")"
	set RS1=conn.execute(mySQL)
	errMessage=""
	if RS1.eof then
		errMessage="Œÿ«! ç‰Ì‰ ”‰œÌ  œ— «Ì‰ ”«· „«·Ì ÊÃÊœ ‰œ«—œ."
	else
		GLMemoNo=RS1("GLDocID")
		GLMemoDate=RS1("GLDocDate")
		if RS1("IsTemporary") then
			IsTemporary=1
		else
			IsTemporary=0
		end if

		if not RS1("BySubSystem") then
			errMessage="«Ì‰ Ìﬂ ”‰œ « Ê„« Ìﬂ ‰Ì” ."
		elseif RS1("deleted") OR RS1("IsRemoved") then
			errMessage="«Ì‰ ”‰œ Õ–› ‘œÂ Ê «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ."
		elseif RS1("IsFinalized") then
			errMessage="«Ì‰ ”‰œ ﬁÿ⁄Ì ‘œÂ Ê «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ."
		elseif RS1("IsChecked") then
			errMessage="«Ì‰ ”‰œ »——”Ì ‘œÂ Ê «„ﬂ«‰ ÊÌ—«Ì‘ ¬‰ ÊÃÊœ ‰œ«—œ."
		end if 
	end if

	if errMessage<>"" then
		response.write "<br><br>"
		call showAlert (errMessage,CONST_MSG_ERROR) 
		Conn.close
		response.end
	end if
	RS1.close


	mySQL="SELECT GLDocs.GLDocID, GLDocs.ID, GLDocs.GLDocDate, GLDocs.CreatedBy, GLRows.ID AS RowID, GLRows.GLAccount, GLRows.Tafsil, GLRows.Amount, GLRows.Description, GLRows.SYS, GLRows.Link, GLRows.IsCredit, GLAccounts.Name FROM GLDocs INNER JOIN GLRows ON GLDocs.ID = GLRows.GLDoc INNER JOIN GLAccounts ON GLRows.GLAccount = GLAccounts.ID WHERE (GLDocs.id="& id & ") AND (GLRows.deleted = 0) AND (GLDocs.GL = "& OpenGL & ") AND (GLAccounts.GL = "& OpenGL & ") ORDER BY GLRows.ID"
	set RS1=conn.execute(mySQL)

	GLDocID = RS1("GLDocID")
	Creator = RS1("CreatedBy")
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	var selectedRowsBalance=0;
	function selectRow(srcRow){
		check = srcRow.getElementsByTagName("INPUT")[0];
		sysLinkName=srcRow.getElementsByTagName("INPUT")[1].name;

		if (check.checked){
			if (sysLinkName!='SysLink'){
				for (j=0;j<document.getElementsByName(sysLinkName).length;j++){
					tempRow = document.getElementsByName(sysLinkName)[j].parentNode.parentNode
					tempRow.getElementsByTagName("INPUT")[0].checked=true

					selectedRowsBalance = selectedRowsBalance + parseInt(tempRow.getElementsByTagName("INPUT")[2].value);

					for(i=0; i<tempRow.getElementsByTagName("TD").length; i++){
						tempRow.getElementsByTagName("TD")[i].setAttribute("bgColor","yellow")
					}
				}
			}else{
				selectedRowsBalance = selectedRowsBalance + parseInt(srcRow.getElementsByTagName("INPUT")[2].value);
				for(i=0; i<srcRow.getElementsByTagName("TD").length; i++){
					srcRow.getElementsByTagName("TD")[i].setAttribute("bgColor","yellow")
				}
			}
		}else{
			if (sysLinkName!='SysLink'){
				for (j=0;j<document.getElementsByName(sysLinkName).length;j++){
					tempRow = document.getElementsByName(sysLinkName)[j].parentNode.parentNode
					tempRow.getElementsByTagName("INPUT")[0].checked=false

					selectedRowsBalance = selectedRowsBalance - parseInt(tempRow.getElementsByTagName("INPUT")[2].value);

					for(i=0; i<tempRow.getElementsByTagName("TD").length; i++){
						tempRow.getElementsByTagName("TD")[i].setAttribute("bgColor","")
					}
				}
			}else{
				selectedRowsBalance = selectedRowsBalance - parseInt(srcRow.getElementsByTagName("INPUT")[2].value);
				for(i=0; i<srcRow.getElementsByTagName("TD").length; i++){
					srcRow.getElementsByTagName("TD")[i].setAttribute("bgColor","")
				}
			}
		}
	}

	function selectAll(src){
		if(src.checked){
			checks = document.getElementsByName("GLDocRows");
			selectedRowsBalance = 0;
			for(j=0; j< checks.length; j++){
				tempRow=checks[j].parentNode.parentNode
				tempRow.getElementsByTagName("INPUT")[0].checked=true;
				selectedRowsBalance = selectedRowsBalance + parseInt(tempRow.getElementsByTagName("INPUT")[2].value);
				for(i=0; i<tempRow.getElementsByTagName("TD").length; i++){
					tempRow.getElementsByTagName("TD")[i].setAttribute("bgColor","yellow")
				}
			}
		}
		else{
			selectedRowsBalance = 0;
			for(j=0; j< checks.length; j++){
				tempRow=checks[j].parentNode.parentNode
				tempRow.getElementsByTagName("INPUT")[0].checked=false;
				for(i=0; i<tempRow.getElementsByTagName("TD").length; i++){
					tempRow.getElementsByTagName("TD")[i].setAttribute("bgColor","")
				}
			}
		}
	}

	function checkAndSubmit(action){
		if (action=='Edit'){
			if (selectedRowsBalance != 0){
				alert("Œÿ Â«Ì «‰ Œ«» ‘œÂ  —«“ ‰Ì” ‰œ!");
				return false;
			}
		}else if (action=='Delete'){
			if (confirm('¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ”‰œ —« »Â ÿÊ— ﬂ«„· Õ–› ﬂ‰Ìœø')){
				checkBox=document.getElementsByName('checkAll')[0]
				checkBox.checked=true;
				selectAll(checkBox);
				checkAndSubmit('Edit');
			}
			return false;
		}
		document.all.act.value="submit"+action;
		document.forms[0].submit();
	}
	//-->
	</SCRIPT>
	<br><br>
	<FORM METHOD=POST ACTION="?">
		<INPUT TYPE="hidden" name="id" value="<%=id%>">

		<TABLE Border="0" width='90%' align=center Cellspacing="1" Cellpadding="5" class="GLTable">
		<TBODY Id="RowsTable">
		<tr style='height:70px;background-color:#FFFFDD;'>
			<td class='HeaderTD' colspan='7' align='center'> ÊÌ—«Ì‘ ”‰œ « Ê„« Ìﬂ “Ì— ”Ì” „ <br></td>
		</tr>
		<tr style='height:1px;background-color:black;'>
			<td colspan=7 ></td>
		</tr>
		<tr style='height:30px;background-color:#EEEEEE;'>
			<td colspan='3'>”‰œ ‘„«—Â <%=GLDocID%></td>
			<td colspan='2'>‘„«—Â ⁄ÿ›: <%=RS1("id")%></td>
			<td colspan='2'> «—ÌŒ ”‰œ: <span dir=ltr><%= RS1("GLDocDate")%></span></td>
		</tr>
		<tr style='height:1px;background-color:black;'>
			<td colspan=7 ></td>
		</tr>
		<tr style='height:40px;background-color:#DDDDDD;'>
			<td style="width:26; border-right:none;"> # </td>
			<td style="width:50; "> ›’Ì·Ì</td>
			<td style="width:50; ">„⁄Ì‰</td>
			<td style="width:170;">‰«„ Õ”«» „⁄Ì‰</td>
			<td style="width:500;">‘—Õ</td>
			<td style="width:80;">»œÂﬂ«—</td>
			<td style="width:80;">»” «‰ﬂ«—</td>
		</tr>
<%
		Do while not RS1.eof
			i = i + 1
			GLAccount = RS1("GLAccount")
			accTitle = RS1("name")
			theDescription = RS1("Description")
			Amount = RS1("Amount")
			IsCredit = RS1("IsCredit")
			Tafsil = RS1("Tafsil")

			credit = ""
			debit = ""
			if IsCredit  then 
				credit = Separate(Amount)
				totalCredit = totalCredit + cdbl(Amount)
			else
				debit = Separate(Amount)
				totalDebit = totalDebit + cdbl(Amount)
				Amount = -1 * cdbl(Amount)
			end if
%>
			<tr>
				<td class='RowsTD' align='center' style='font-size:9pt;color:gray;'><%=i%><br>
					<INPUT TYPE="checkbox" NAME="GLDocRows" Value='<%=RS1("RowID")%>' onclick='selectRow(this.parentNode.parentNode);'>
					<INPUT TYPE="hidden" Name='SysLink<%=RS1("Sys")&RS1("Link")%>'>
					<INPUT TYPE="hidden" Name='Amount' value='<%=Amount%>'>
				</td>
				<td class='RowsTD'><%=Tafsil%></td>
				<td class='RowsTD'><%=GLAccount%></td>
				<td class='RowsTD'><%=accTitle%></td>
				<td class='RowsTD'><%=theDescription%></td>
				<td class='RowsTD'><%=debit%></td>
				<td class='RowsTD'><%=credit%></td>
			</tr>
<%
		RS1.movenext
		loop	
%>
		<tr style='height:1px;background-color:black;'>
			<td colspan=7 ></td>
		</tr>
		<tr style="height:50px;">
			<td class='RowsTD' colspan='2' style='font-size:7pt;'>
				<INPUT TYPE="checkbox" NAME='checkAll' onclick='selectAll(this);'>&nbsp; «‰ Œ«» Â„Â</td>
			<td class='RowsTD' colspan='3' align='left'><B>Ã„⁄ :</B></td>
			<td class='RowsTD' style="width:80;"><%=Separate(totalDebit)%></td>
			<td class='RowsTD' style="width:80;"><%=Separate(totalCredit)%></td>
		</tr>
		<tr style='height:1px;background-color:black;'>
			<td colspan=7 ></td>
		</tr>
		<tr style="height:50px;">
			<td class='HeaderTD' colspan='7' align='center'>
				<INPUT TYPE="hidden" name="act" value="" >
				<TABLE width='100%'>
				<TR>
					<TD class='HeaderTD'><INPUT TYPE="button" value="Õ–› ”ÿ— Â«Ì «‰ Œ«» ‘œÂ" class="GenButton" onclick="checkAndSubmit('Edit');"></TD>
					<TD class='HeaderTD' align='left'><INPUT TYPE="button" value="«‰’—«›" class="GenButton" onclick="window.location='GLMemoDocShow.asp?id=<%=id%>';"></TD>
					<TD class='HeaderTD' align='left'><INPUT TYPE="button" value="Õ–› ﬂ«„· ”‰œ (Â„Â ”ÿ— Â«)" class="GenButton" Style='border-color:red' onclick="checkAndSubmit('Delete');"></TD>
				</TR>
				</TABLE>
			</td>
		</tr>
		</TBODY>
		</TABLE>
	</FORM>
	<BR>
<%
	if request("SysLink")<>"" then
%>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			sysLinkName='SysLink<%=request("SysLink")%>';
			a=document.getElementsByName(sysLinkName)[0].parentNode.getElementsByTagName("INPUT")[0];
			a.checked=true;
			selectRow(a.parentNode.parentNode);
			a.focus();
		//-->
		</SCRIPT>
<%
	end if
'-----------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------- Show a GL Doc
'-----------------------------------------------------------------------------------------------------
elseif request("act")="find" then
	
	sys=left(request("link"),2)
	lnk=right(request("link"),len(request("link"))-2)
	sys=sqlSafe(sys)
	lnk=clng(lnk)

	mySQL="SELECT DISTINCT GLDoc FROM EffectiveGLRows WHERE (GL = "& OpenGL & ") AND (Link = " & lnk & ") AND (SYS = '" & sys & "')"
	set RS1=conn.execute(mySQL)
	if RS1.eof then
		conn.close
		response.redirect "?errMsg="&Server.URLEncode("Œÿ«! ”‰œ ÅÌœ« ‰‘œ.")
	else
		GLDoc= RS1("GLDoc")
	end if
	RS1.close
	conn.close
	response.redirect "SubsysDocsEdit.asp?act=show&id=" & GLDoc & "&SysLink=" & sys & lnk
end if %>
<!--#include file="tah.asp" -->
