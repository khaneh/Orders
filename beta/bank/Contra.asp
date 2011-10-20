<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Bank (10 [=A])
PageTitle= "„€«Ì— "
SubmenuItem=9
if not Auth("A" , 9) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<style>
	.GenInput { font-family:tahoma; font-size: 9pt; border: 1px solid black; text-align:right; }
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
	.dateLine {background-Color:#003399;color:#FFFF00;}
	.row {}
	.row1 {background-Color:#99FFCC;}
	.row2 {background-Color:#66CC99;}
	.row3 {background-Color:#AADDFF;}
	.row9 {background-Color:#FFCC00;t:#66CC00;}
</style>

<BR><BR>
<%
Function showDate(inpStr)
'	if len(inpStr)=6 then
'		showDate = left(inpStr,2) & "." & mid(inpStr,3,2) & "." & right(inpStr,2)
'	else
'		showDate = "Error"
'	end if
	showDate = replace(inpStr,"/",".")
End Function

Function WriteBlankRows (inpRemainedRows) 

	For tmpCnt = 1 to inpRemainedRows
%>		<TR height=20>
			<TD colspan=5>&nbsp;</TD>
		</TR>
<%	Next

End Function

Function convertDate(inpStr)
	if len(inpStr)=10 then
		convertDate = mid(inpStr,3,2) & mid(inpStr,6,2) & right(inpStr,2)
	else
		convertDate = "Error"
	end if
End Function

Function lookupDescCode(inpBank, inpCode)
	result = inpBank & "," & inpCode ' = "&nbsp;"
	Select Case result
		' Tejarat
		Case "1,6":
			result = "Å—œ«Œ  ‰ﬁœÌ çﬂ"
		Case "1,7":
			result = "œ—Ì«›  ‰ﬁœÌ"
		Case "1,8":
			result = "Å—œ«Œ  ‰ﬁœÌ"
		Case "1,11":
			result = "Ê’Ê· çﬂ"
		Case "1,13":
			result = "«‰ ﬁ«·ÌÂ«"
		Case "1,22":
			result = "œ—Ì«›  ‰ﬁœÌ (‘⁄»)"
		Case "1,28":
			result = "çﬂ »«‰ﬂÌ"
		Case "1,33":
			result = "çﬂÂ«Ì ’«œ—Â"
		Case "1,38":
			result = "ÊÃÂ çﬂÂ«Ì Ê’Ê·Ì"
		Case "1,42":
			result = "Ê’Ê· »—« "
		Case "1,43":
			result = "Å”  Ê ﬂ«—„“œ »—« "
		Case "1,125":
			result = "œ—Ì«›  ‰ﬁœÌ"
		Case "1,251":
			result = "Å—œ«Œ  «‰ ﬁ«·Ì çﬂ"
		Case "1,122":
			result = "œ—Ì«›  ‰ﬁœÌ (»Ì‰ ‘⁄»)"
		Case "1,291":
			result = "»—ê‘  À»  «‘ »«Â"
		Case "1,162":
			result = "Å—œ«Œ  ‰ﬁœÌ ÿÌ ”‰œ"
		Case "1,262":
			result = "«‰ ﬁ«· »Â Õ”«»( ·›‰ »«‰ﬂ)"
		Case "1,223":
			result = "Ê«—Ì“ »«»  œ” Ê—Â«Ì À«"
		Case "1,253":
			result = "Å—œ«Œ  «‰ ﬁ«·Ì ‘»Â ÅÊ·"
		Case "1,222":
			result = "«‰ ﬁ«· «“ Õ”«»( ·›‰ »«‰ﬂ)"
		Case "1,273":
			result = "ﬂ«—„“œ çﬂ œÊ—‰ÊÌ”"
		Case "1,211":
			result = "Ê’Ê· çﬂ"
		Case "1,261":
			result = "«‰ ﬁ«· »Â Õ”«»"
		Case "1,161":
			result = "Å—œ«Œ  ‰ﬁœÌ çﬂ"
		Case "1,254":
			result = "Å—œ«Œ  «‰ ﬁ«·Ì çﬂÂ«Ì »«‰ﬂÌ"
		Case "1,":
			result = ""
		Case "1,":
			result = ""
		' Sepah
		Case "2,1":
			result = "çﬂ «‰ ﬁ«·Ì"
		Case "2,2":
			result = "›Ì‘ ‰ﬁœÌ"
		Case "2,3":
			result = "Ê’Ê·Ì"
		Case "2,4":
			result = "»—Ê« "
		Case "2,5":
			result = "ﬂ«—„“œ"
		Case "2,9":
			result = "ÕÊ«·Ã« "
		Case "2,11":
			result = "»œÂﬂ«—«‰"
		Case "2,12":
			result = "¬êÂÌ »œÂﬂ«—"
		Case "2,15":
			result = "¬êÂÌ Å—œ«Œ "
	End Select
	lookupDescCode = result
End Function
Function lookupBankType(inpGLAccount)
	Select Case inpGLAccount
		Case "12001":
			result = "1" ' Tejarat
		Case "12004":
			result = "1" ' Tejarat
		Case "12002":
			result = "2" ' Sepah
		Case else:
			result = "Error"
	End Select
	lookupBankType = result
End Function
Function lookupBankAccountNo(inpGLAccount)
	Select Case inpGLAccount
		Case "12001":
			result = "0033624050"
		Case "12004":
			result = "0033624816"
		Case "12002":
			result = "000434419"
		Case else:
			result = "Error"
	End Select
	lookupBankAccountNo = result
End Function

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------- Show Remained Un Related
'-----------------------------------------------------------------------------------------------------
if request("act")="showAll" then
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	StartDate =		request.QueryString("StartDate")
	EndDate =		request.QueryString("EndDate")
	BankGLAccount = cint(request.QueryString("BankGLAccount"))
	BankAccountNo = lookupBankAccountNo(BankGLAccount)
	BankType =		lookupBankType(BankGLAccount)

	GLRowsViewName = "VIEW_Contra_Remained_GLRows_" & openGL & "_" & BankGLAccount
	BankRowsViewName = "VIEW_Contra_Remained_BankRows_" & BankAccountNo

	'Creating Temporary Table
	mySQL="DELETE FROM BankContradiction_TempDates"
	Conn.Execute(mySQL)

'	mySQL="INSERT INTO BankContradiction_TempDates ([Date], MaxRows) SELECT LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) AS [date], CASE WHEN ISNULL(GLRowDates.CntGLRow, 0) > ISNULL(BankLogDates.CntBankLog, 0) THEN GLRowDates.CntGLRow ELSE BankLogDates.CntBankLog END AS MaxRows FROM (SELECT GLDocDate AS [Date], COUNT(*) AS CntGLRow FROM " & GLRowsViewName & " GROUP BY GLDocDate) GLRowDates FULL OUTER JOIN (SELECT [Date], COUNT(*) AS CntBankLog FROM " & BankRowsViewName & " GROUP BY [Date]) BankLogDates ON GLRowDates.[Date] = BankLogDates.[Date] WHERE (LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) >= '" & StartDate & "') AND (LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) <= '" & EndDate & "') ORDER BY GLRowDates.[Date]"
	mySQL="INSERT INTO BankContradiction_TempDates ([Date], MaxRows) SELECT LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) AS [date], CASE WHEN ISNULL(GLRowDates.CntGLRow, 0) > ISNULL(BankLogDates.CntBankLog, 0) THEN GLRowDates.CntGLRow ELSE BankLogDates.CntBankLog END AS MaxRows FROM (SELECT GLDocDate AS [Date], COUNT(*) AS CntGLRow FROM EffectiveGLRows WHERE (GL = " & openGL & ") AND (GLAccount = " & BankGLAccount & ") GROUP BY GLDocDate) GLRowDates FULL OUTER JOIN (SELECT [Date], COUNT(*) AS CntBankLog FROM BankLog WHERE (AccountNo = " & BankAccountNo & ") GROUP BY [Date]) BankLogDates ON GLRowDates.[Date] = BankLogDates.[Date] WHERE (LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) >= '" & StartDate & "') AND (LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) <= '" & EndDate & "') ORDER BY GLRowDates.[Date]"
	Conn.Execute(mySQL)
%>

	<table align=center>
	<tr>
		<td valign=top>
<!-- GL Rows -->
		<TABLE cellpadding=2 cellspacing=1 border=1 width=370>
		<TR>
			<TD colspan=5><B>„⁄Ì‰ <%=BankGLAccount%></B></TD>
		</TR>
<%

'		mySQL="SELECT BankContradiction_TempDates.[Date] AS BaseDate, BankContradiction_TempDates.MaxRows, BankingGLRows.*, ISNULL(BankContradictionRelations.Type,0) AS Type, BankContradictionRelations.Description AS RelationDesc, Users.RealName AS Creator, BankContradictionRelations.CreatedDate, BankContradictionRelations.CreatedTime FROM BankContradictionRelations INNER JOIN Users ON BankContradictionRelations.CreatedBy = Users.ID RIGHT OUTER JOIN (SELECT ID, isCredit, Amount, Description, GLDoc, GLDocID, GLDocDate FROM EffectiveGLRows WHERE (GL = " & openGL & ") AND (GLAccount = " & BankGLAccount & ")) BankingGLRows ON BankContradictionRelations.GLRowID = BankingGLRows.ID RIGHT OUTER JOIN BankContradiction_TempDates ON BankingGLRows.GLDocDate = BankContradiction_TempDates.[Date] ORDER BY BankContradiction_TempDates.[Date]"

		mySQL="SELECT BankContradiction_TempDates.[Date] AS BaseDate, BankContradiction_TempDates.MaxRows, ISNULL(BankContradictionRelations.Type, 0) AS Type, BankingGLRows.ID, BankingGLRows.isCredit, BankingGLRows.Amount, BankingGLRows.Description, BankingGLRows.GLDoc, BankingGLRows.GLDocID, BankingGLRows.GLDocDate FROM BankContradictionRelations RIGHT OUTER JOIN (SELECT ID, isCredit, Amount, Description, GLDoc, GLDocID, GLDocDate FROM EffectiveGLRows WHERE (GL = " & openGL & ") AND (GLAccount = " & BankGLAccount & ")) BankingGLRows ON BankContradictionRelations.GLRowID = BankingGLRows.ID RIGHT OUTER JOIN BankContradiction_TempDates ON BankingGLRows.GLDocDate = BankContradiction_TempDates.[Date] GROUP BY BankContradiction_TempDates.[Date], BankContradiction_TempDates.MaxRows, ISNULL(BankContradictionRelations.Type, 0), BankingGLRows.ID, BankingGLRows.Amount, BankingGLRows.Description, BankingGLRows.GLDoc, BankingGLRows.GLDocID, BankingGLRows.GLDocDate, BankingGLRows.isCredit ORDER BY BankContradiction_TempDates.[Date], BankingGLRows.isCredit DESC, BankingGLRows.Amount"

		Set RSDates = Conn.Execute(mySQL)
		if NOT RSDates.eof then
			RemainedRows = 0
			BaseDate = ""
			TempCount = 1
			Do While NOT RSDates.eof 
				RowsCount = clng(RSDates("MaxRows"))
				if BaseDate = RSDates("BaseDate") then
					RemainedRows = RemainedRows - 1
					TempCount = TempCount + 1
				else

					Call WriteBlankRows (RemainedRows)

%>
					<TR class="dateLine">
						<TD colspan=5><%=showDate(RSDates("BaseDate"))%></TD>
					</TR>
<%
					RemainedRows = RowsCount - 1 
					TempCount = 1
				end if

				if not isnull(RSDates("ID")) then
					if RSDates("isCredit") then
						Debit="&nbsp;"
						Credit=Separate(RSDates("Amount"))
					else
						Debit=Separate(RSDates("Amount"))
						Credit="&nbsp;"
					end if
					rowType = cint(RSDates("Type"))
					if rowType = 1 then
						RowTitle="« Ê„« Ìﬂ ‰Ê⁄ 1:" & vbCrLf & "„»·€°  «—ÌŒ° ‘„«—Â çﬂ"
					elseif rowType = 2 then
						RowTitle="« Ê„« Ìﬂ ‰Ê⁄ 2:" & vbCrLf & "„»·€ Ê ‘„«—Â çﬂ"
					elseif rowType = 3 then 
						RowTitle="« Ê„« Ìﬂ ‰Ê⁄ 3:" & vbCrLf & "„»·€ Ê  «—ÌŒ"
					elseif rowType = 9 then 
						RowTitle="œ” Ì" 
						'RowTitle="œ” Ì" & vbCrLf & " Ê”ÿ " & RSDates("Creator") & " " & RSDates("CreatedDate")
						'if RSDates("RelationDesc") <> "" then RowTitle = RowTitle & vbCrLf & " Ê÷ÌÕ: " & RSDates("RelationDesc")
					else
						RowTitle="‰« „‘Œ’"
					end if
%>

					<TR class="row<%=rowType%>" title="<%=RowTitle%>" height=20>
						<TD width=30><%=RSDates("GLDocID")%></TD>
						<TD title="<%=RSDates("Description")%>" style="cursor:col-resize;"><%=Left(RSDates("Description"),20) & ".."%></TD>
						<TD width=70><%=Debit%></TD>
						<TD width=70><%=Credit%></TD>
						<TD width=20 style="cursor:crosshair;" onclick="slctGLRow(this)">
							<input type="hidden" Value="<%=RSDates("ID")%>">
						</TD>
					</TR>
<%
				else

					Call WriteBlankRows (1)

				End if
				BaseDate = RSDates("BaseDate")
				
				RSDates.moveNext
			Loop
		end if
		RSDates.close

		Call WriteBlankRows (RemainedRows) 
%>
		</TABLE>
		</td>
		<td></td>
		<td valign=top>
<!-- Bank Rows -->
		<TABLE cellpadding=2 cellspacing=1 border=1 width=350>
		<TR>
			<TD colspan=5><B>ê“«—‘ »«‰ﬂ</B></TD>
		</TR>
<%

'		mySQL = "SELECT BankContradiction_TempDates.[Date] AS BaseDate, BankContradiction_TempDates.MaxRows, allBankRows.*, BankContradictionRelations.Description, ISNULL(BankContradictionRelations.Type,0) AS Type, BankContradictionRelations.Description AS RelationDesc, Users.RealName AS Creator, BankContradictionRelations.CreatedDate, BankContradictionRelations.CreatedTime FROM BankContradiction_TempDates LEFT OUTER JOIN Users INNER JOIN BankContradictionRelations ON Users.ID = BankContradictionRelations.CreatedBy RIGHT OUTER JOIN (SELECT * FROM BankLog WHERE AccountNo = " & BankAccountNo & ") allBankRows ON BankContradictionRelations.BankLogID = allBankRows.autoKey ON BankContradiction_TempDates.[Date] = allBankRows.[Date] ORDER BY BankContradiction_TempDates.[Date]"

		mySQL = "SELECT BankContradiction_TempDates.[Date] AS BaseDate, BankContradiction_TempDates.MaxRows, ISNULL(BankContradictionRelations.Type, 0) AS Type, allBankRows.AccountNo, allBankRows.CheqNo, allBankRows.[Date], allBankRows.isDebit, allBankRows.Amount, allBankRows.DescCode, allBankRows.BankType, allBankRows.autoKey FROM BankContradiction_TempDates LEFT OUTER JOIN BankContradictionRelations RIGHT OUTER JOIN (SELECT * FROM BankLog WHERE AccountNo = " & BankAccountNo & ") allBankRows ON BankContradictionRelations.BankLogID = allBankRows.autoKey ON BankContradiction_TempDates.[Date] = allBankRows.[Date] GROUP BY BankContradiction_TempDates.[Date], BankContradiction_TempDates.MaxRows, ISNULL(BankContradictionRelations.Type, 0), allBankRows.AccountNo, allBankRows.CheqNo, allBankRows.[Date], allBankRows.Amount, allBankRows.DescCode, allBankRows.BankType, allBankRows.autoKey, allBankRows.isDebit ORDER BY BankContradiction_TempDates.[Date], allBankRows.isDebit, allBankRows.Amount"

		Set RSDates = Conn.Execute(mySQL)
		if NOT RSDates.eof then
			RemainedRows = 0
			BaseDate = ""
			TempCount = 1
			Do While NOT RSDates.eof 
				RowsCount = clng(RSDates("MaxRows"))
				if BaseDate = RSDates("BaseDate") then
					RemainedRows = RemainedRows - 1
					TempCount = TempCount + 1
				else

					Call WriteBlankRows (RemainedRows)
%>
					<TR class="dateLine">
						<TD colspan=5><%=showDate(RSDates("BaseDate"))%></TD>
					</TR>
<%
					RemainedRows = RowsCount - 1 
					TempCount = 1
				end if

				if not isnull(RSDates("AutoKey")) then
				'crosshair
					if RSDates("isDebit") then
						Debit=Separate(cdbl(RSDates("Amount")))
						Credit="&nbsp;"
					else
						Debit="&nbsp;"
						Credit=Separate(cdbl(RSDates("Amount")))
					end if
					rowType = cint(RSDates("Type"))
					if rowType = 1 then
						RowTitle="« Ê„« Ìﬂ ‰Ê⁄ 1:" & vbCrLf & "„»·€°  «—ÌŒ° ‘„«—Â çﬂ"
					elseif rowType = 2 then
						RowTitle="« Ê„« Ìﬂ ‰Ê⁄ 2:" & vbCrLf & "„»·€ Ê ‘„«—Â çﬂ"
					elseif rowType = 3 then 
						RowTitle="« Ê„« Ìﬂ ‰Ê⁄ 3:" & vbCrLf & "„»·€ Ê  «—ÌŒ"
					elseif rowType = 9 then 
						RowTitle="œ” Ì" 
						'RowTitle="œ” Ì" & vbCrLf & " Ê”ÿ " & RSDates("Creator") & " " & RSDates("CreatedDate")
						'if RSDates("RelationDesc") <> "" then RowTitle = RowTitle & vbCrLf & " Ê÷ÌÕ: " & RSDates("RelationDesc")
					else
						RowTitle="‰« „‘Œ’"
					end if
%>

					<TR class="row<%=rowType%>" title="<%=RowTitle%>" height=20>
						<TD width=20 style="cursor:crosshair;" onclick="slctBnkRow(this)">
							<input type="hidden" Value="<%=RSDates("AutoKey")%>">
						</TD>
						<TD width=80 style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:col-resize;"><NoBr><%=lookupDescCode(BankType,RSDates("DescCode"))%></NoBr></TD>
						<TD width=70><%=Debit%></TD>
						<TD width=70><%=Credit%></TD>
						<TD width=30><%=RSDates("CheqNo")%></TD>
					</TR>
<%
				else

					Call WriteBlankRows (1)

				End if
				BaseDate = RSDates("BaseDate")
				
				RSDates.moveNext
			Loop
		end if
		RSDates.close

		Call WriteBlankRows (RemainedRows)
%>
		</TABLE>
		</td>
	</tr>
	<tr height="5">
		<td align='center' colspan="3">
		<table>
		<tr>
			<td>
				<FORM METHOD=POST ACTION="?act=delRelations&StartDate=<%=Server.URLEncode(StartDate)%>&EndDate=<%=Server.URLEncode(EndDate)%>&BankGLAccount=<%=Server.URLEncode(BankGLAccount)%>&NextAct=showAll">
					<span id="Relations"></span>
					<INPUT TYPE="submit" value="Õ–› —«»ÿÂ Â«" class="GenButton" onclick="this.disabled=true;submit();">
				</FORM>
			</td>
			<td>
				<FORM METHOD=POST ACTION="?act=autoRelate&StartDate=<%=Server.URLEncode(StartDate)%>&EndDate=<%=Server.URLEncode(EndDate)%>&BankGLAccount=<%=Server.URLEncode(BankGLAccount)%>&NextAct=showAll">
					<INPUT TYPE="submit" value="œÊŒ ‰ « Ê„« Ìﬂ" class="GenButton" onclick="this.disabled=true;submit();">
				</FORM>
			</td>
			<td>
				<FORM METHOD=POST ACTION="?act=showRemained&StartDate=<%=Server.URLEncode(StartDate)%>&EndDate=<%=Server.URLEncode(EndDate)%>&BankGLAccount=<%=Server.URLEncode(BankGLAccount)%>&NextAct=showAll">
					<INPUT TYPE="submit" value="‰„«Ì‘ „«‰œÂ Â«" class="GenButton" onclick="this.disabled=true;submit();">
				</FORM>
			</td>
		</tr>
		</table>
		</td>
	</tr>
	</table>

	<script language="JavaScript">
	<!--
	//---------------------------------------------
	var slctRowClr='#FFFF00'
	var rltdRowClr='#99FF88'
	var sRltdRowClr='#9988FF'
	var delRowClr='#FF0055'
	
	function slctGLRow(src){
		
		if (confirm("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« Õ–› ﬂ‰Ìœø\n")){
			src.onclick="";
			src.setAttribute("bgColor",delRowClr);
			src.title="Õ–› ‘œÂ"
			document.getElementById("Relations").innerHTML+="<input type='hidden' name='GLRow' value='"+src.getElementsByTagName("input")[0].value+"'>"
		}
	}
	function slctBnkRow(src){
		
		if (confirm("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« Õ–› ﬂ‰Ìœø\n")){
			src.onclick="";
			src.setAttribute("bgColor",delRowClr);
			src.title="Õ–› ‘œÂ"
			document.getElementById("Relations").innerHTML+="<input type='hidden' name='BnkRow' value='"+src.getElementsByTagName("input")[0].value+"'>"
		}
	}

	//-->
	</script>

<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------- Show Remained Un Related
'-----------------------------------------------------------------------------------------------------
elseif request("act")="showRemained" then
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	if request.QueryString("StartDate")<>"" then
	' Use QueryString Data
		StartDate =		request.QueryString("StartDate")
		EndDate =		request.QueryString("EndDate")
		BankGLAccount = cint(request.QueryString("BankGLAccount"))
		BankAccountNo = lookupBankAccountNo(BankGLAccount)
		BankType =		lookupBankType(BankGLAccount)

	else
	' Use Form Data
		set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
		mySmartUpload.Upload

		StartDate =		mySmartUpload.Form("StartDate") '"840101"
		EndDate =		mySmartUpload.Form("EndDate") '"841019"
		BankGLAccount = cint(mySmartUpload.Form("BankGLAccount")) '"841019"
		BankAccountNo = lookupBankAccountNo(BankGLAccount)
		BankType =		lookupBankType(BankGLAccount)

		If not mySmartUpload.Files(1).IsMissing Then
		' We have a FILE uploaded

			' Save The File
			SavedFileName = "./LogFiles/" & lookupBankAccountNo(BankGLAccount) & ".TXT"
			mySmartUpload.Files(1).SaveAs(SavedFileName)

			' Read and Insert
			Dim oFS 
			Dim oFile 
			Dim oStream 

			If BankGLAccount = 12001 OR BankGLAccount = 12004 then
			'Tejarat
				Set oFS = Server.CreateObject("Scripting.FileSystemObject") 
				Set oFile = oFS.GetFile(Server.MapPath(SavedFileName))
				Set oStream = oFile.OpenAsTextStream(1, -2) '1:ForReading , -2:TristateUseDefault

				Do While Not oStream.AtEndOfStream 

					theLine=oStream.ReadLine 

					if len(theLine) > 39 then
						BankAccountNo = mid(theLine,2,10) 'left(theLine,10)
						CheckNo = mid(theLine,12,6)
						mmoDate = "13" & mid(theLine,18,2) & "/" & mid(theLine,20,2) & "/" & mid(theLine,22,2)
						isDebit = mid(theLine,24,1)
						amount = mid(theLine,25,13)
						DescCode = mid(theLine,38,3)

						mySQL="INSERT INTO BankLog (AccountNo, CheqNo, [Date], isDebit, Amount, DescCode, BankType)" &_
						"VALUES ('" & BankAccountNo & "','" & CheckNo & "','" & mmoDate & "','" & isDebit & "','" & amount & "','" & DescCode & "','1')"
						Conn.Execute(mySQL)

					else
						exit do
					end if

				Loop 

				oStream.Close 

				Set oFS = Nothing
				Set oFile = Nothing
				Set oStream = Nothing

				'Removing Duplicate Rows
				mySQL="DELETE FROM BankLog WHERE (autoKey IN (SELECT MAX(autoKey) AS RepeatedLine FROM BankLog GROUP BY AccountNo, CheqNo, [Date], Amount, DescCode, BankType, isDebit HAVING (COUNT(*) > 1)))"
				Conn.Execute(mySQL)

			elseif BankGLAccount = 12002 then
			'Sepah

				Set oFS = Server.CreateObject("Scripting.FileSystemObject") 
				Set oFile = oFS.GetFile(Server.MapPath(SavedFileName))
				Set oStream = oFile.OpenAsTextStream(1, -2) '1:ForReading , -2:TristateUseDefault
				Do While Not oStream.AtEndOfStream 
					theLine=oStream.ReadLine 

					if theLine <> "" then
						BankAccNo = left(theLine,9)
						mmoDate = "13" & mid(theLine,10,2) & "/" & mid(theLine,12,2) & "/" & mid(theLine,14,2)
						CheckNo = mid(theLine,19,6)
						amount = mid(theLine,39,14)
						StupidMinusSign = mid(theLine,52,1)
						if not isnumeric(StupidMinusSign) then 
							amount = "-" & left(amount,13) & (asc(StupidMinusSign) - 112)
						end if

						isDebit = 2 - cint(mid(theLine,53,1))

						DescCode = mid(theLine,54,2)

						if isDebit <> 2 then
							mySQL="INSERT INTO BankLog (AccountNo, CheqNo, [Date], isDebit, Amount, DescCode, BankType)" &_
							"VALUES ('" & BankAccNo & "','" & CheckNo & "','" & mmoDate & "','" & isDebit & "','" & amount & "','" & DescCode & "','2')"
							Conn.Execute(mySQL)
						end if

					else
						exit do
					end if

				Loop 

				oStream.Close 

				Set oFS = Nothing
				Set oFile = Nothing
				Set oStream = Nothing

				'Removing Duplicate Rows
				mySQL="DELETE FROM BankLog WHERE (autoKey IN (SELECT MAX(autoKey) AS RepeatedLine FROM BankLog GROUP BY AccountNo, CheqNo, [Date], Amount, DescCode, BankType, isDebit HAVING (COUNT(*) > 1)))"
				Conn.Execute(mySQL)

			else
			'Error
			End If

		End If

	end if

'	CREATE VIEW VIEW_Contra_Remained_GLRows_84_12001
'	AS
'	SELECT     ID, GLDocID, GLDocDate COLLATE SQL_Latin1_General_CP1_CI_AS AS [Date], IsCredit, Amount, Description
'	FROM         EffectiveGLRows
'	WHERE     (GL = 84) AND (GLAccount = 12001) AND (ID NOT IN
'							  (SELECT     GLRowID
'								 FROM         BankContradictionRelations))
'
'	-	-	-	-	-	-	-	-	
'
'	CREATE VIEW dbo.VIEW_Contra_Remained_BankRows_0033624050
'	AS
'	SELECT     autoKey, CheqNo, [Date] COLLATE SQL_Latin1_General_CP1_CI_AS AS [Date], isDebit, Amount, DescCode, AccountNo
'	FROM         dbo.BankLog
'	WHERE     (AccountNo = 0033624050) AND (autoKey NOT IN
'							  (SELECT     BankLogID
'								 FROM         BankContradictionRelations))
'


	GLRowsViewName = "VIEW_Contra_Remained_GLRows_" & openGL & "_" & BankGLAccount
	BankRowsViewName = "VIEW_Contra_Remained_BankRows_" & BankAccountNo

	'Creating Temporary Table
	mySQL="DELETE FROM BankContradiction_TempDates"
	Conn.Execute(mySQL)

	mySQL="INSERT INTO BankContradiction_TempDates ([Date], MaxRows) SELECT LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) AS [date], CASE WHEN ISNULL(GLRowDates.CntGLRow, 0) > ISNULL(BankLogDates.CntBankLog, 0) THEN GLRowDates.CntGLRow ELSE BankLogDates.CntBankLog END AS MaxRows FROM (SELECT [Date], COUNT(*) AS CntGLRow FROM " & GLRowsViewName & " GROUP BY [Date]) GLRowDates FULL OUTER JOIN (SELECT [Date], COUNT(*) AS CntBankLog FROM " & BankRowsViewName & " GROUP BY [Date]) BankLogDates ON GLRowDates.[Date] = BankLogDates.[Date] WHERE (LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) >= '" & StartDate & "') AND (LEFT(ISNULL(GLRowDates.[Date], '') + ISNULL(BankLogDates.[Date], ''), 10) <= '" & EndDate & "') ORDER BY GLRowDates.[Date]"

	Conn.Execute(mySQL)

%>

	<table align=center>
	<tr>
		<td valign=top>
<!-- GL Rows -->
		<TABLE cellpadding=2 cellspacing=1 border=1 width=350>
		<TR>
			<TD colspan=5><B>„⁄Ì‰ <%=BankGLAccount%></B></TD>
		</TR>
<%

		mySQL="SELECT BankContradiction_TempDates.[Date] AS BaseDate, BankContradiction_TempDates.MaxRows, " & GLRowsViewName & ".* FROM BankContradiction_TempDates LEFT OUTER JOIN " & GLRowsViewName & " ON BankContradiction_TempDates.[Date] = " & GLRowsViewName & ".[Date] ORDER BY BankContradiction_TempDates.[Date], " & GLRowsViewName & ".IsCredit DESC, " & GLRowsViewName & ".Amount"

		Set RSDates = Conn.Execute(mySQL)
		if NOT RSDates.eof then
			RemainedRows = 0
			BaseDate = ""
			TempCount = 1
			Do While NOT RSDates.eof 
				RowsCount = clng(RSDates("MaxRows"))
				if BaseDate = RSDates("BaseDate") then
					RemainedRows = RemainedRows - 1
					TempCount = TempCount + 1
				else
					
					Call WriteBlankRows (RemainedRows)
%>
					<TR class="dateLine">
						<TD colspan=5><%=showDate(RSDates("BaseDate"))%></TD>
					</TR>
<%
					RemainedRows = RowsCount - 1 
					TempCount = 1
				end if

				if not isnull(RSDates("ID")) then
					if RSDates("isCredit") then
						Debit="&nbsp;"
						Credit=Separate(RSDates("Amount"))
					else
						Debit=Separate(RSDates("Amount"))
						Credit="&nbsp;"
					end if
%>

					<TR height="20">
						<TD width=30><%=RSDates("GLDocID")%></TD>
						<TD title="<%=RSDates("Description")%>" style="cursor:col-resize;"><%=Left(RSDates("Description"),20) & ".."%></TD>
						<TD width=70><%=Debit%></TD>
						<TD width=70><%=Credit%></TD>
						<TD width=20 style="cursor:crosshair;" onclick="slctGLRow(this)">
							<input type="hidden" Value="<%=RSDates("ID")%>">
						</TD>
					</TR>
<%

				else

					Call WriteBlankRows (1)

				End if
				BaseDate = RSDates("BaseDate")
				
				RSDates.moveNext
			Loop
		end if
		RSDates.close

		Call WriteBlankRows (RemainedRows)
%>
		</TABLE>
		</td>
		<td></td>
		<td valign=top>
<!-- Bank Rows -->
		<TABLE cellpadding=2 cellspacing=1 border=1 width=350>
		<TR>
			<TD colspan=5><B>ê“«—‘ »«‰ﬂ</B></TD>
		</TR>
<%
		mySQL = "SELECT BankContradiction_TempDates.[Date] AS BaseDate, BankContradiction_TempDates.MaxRows, " & BankRowsViewName & ".* FROM BankContradiction_TempDates LEFT OUTER JOIN " & BankRowsViewName & " ON BankContradiction_TempDates.[Date] = " & BankRowsViewName & ".[Date] ORDER BY BankContradiction_TempDates.[Date], " & BankRowsViewName & ".isDebit, " & BankRowsViewName & ".Amount"

		Set RSDates = Conn.Execute(mySQL)
		if NOT RSDates.eof then
			RemainedRows = 0
			BaseDate = ""
			TempCount = 1
			Do While NOT RSDates.eof 
				RowsCount = clng(RSDates("MaxRows"))
				if BaseDate = RSDates("BaseDate") then
					RemainedRows = RemainedRows - 1
					TempCount = TempCount + 1
				else

					Call WriteBlankRows (RemainedRows)
%>
					<TR class="dateLine">
						<TD colspan=5><%=showDate(RSDates("BaseDate"))%></TD>
					</TR>
<%
					RemainedRows = RowsCount - 1 
					TempCount = 1
				end if

				if not isnull(RSDates("AutoKey")) then
				'crosshair
					if RSDates("isDebit") then
						Debit=Separate(cdbl(RSDates("Amount")))
						Credit="&nbsp;"
					else
						Debit="&nbsp;"
						Credit=Separate(cdbl(RSDates("Amount")))
					end if
%>

					<TR height="20">
						<TD width=20 style="cursor:crosshair;" onclick="slctBnkRow(this)">
							<input type="hidden" Value="<%=RSDates("AutoKey")%>">
						</TD>
						<TD width=80 style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:col-resize;"><NoBr><%=lookupDescCode(BankType,RSDates("DescCode"))%></NoBr></TD>
						<TD width=70><%=Debit%></TD>
						<TD width=70><%=Credit%></TD>
						<TD width=30><%=RSDates("CheqNo")%></TD>
					</TR>
<%
				else

					Call WriteBlankRows (1)

				End if
				BaseDate = RSDates("BaseDate")
				
				RSDates.moveNext
			Loop
		end if
		RSDates.close

		Call WriteBlankRows (RemainedRows)
%>
		</TABLE>
		</td>
	</tr>
	<tr height="5">
		<td align='center' colspan="3">
		<table>
		<tr>
			<td>
				<FORM METHOD=POST ACTION="?act=submitRelations&StartDate=<%=Server.URLEncode(StartDate)%>&EndDate=<%=Server.URLEncode(EndDate)%>&BankGLAccount=<%=Server.URLEncode(BankGLAccount)%>&NextAct=showRemained">
					<span id="Relations"></span>
					<INPUT TYPE="submit" value="À»  —«»ÿÂ Â«" class="GenButton" onclick="this.disabled=true;submit();">
				</FORM>
			</td>
			<td>
				<FORM METHOD=POST ACTION="?act=autoRelate&StartDate=<%=Server.URLEncode(StartDate)%>&EndDate=<%=Server.URLEncode(EndDate)%>&BankGLAccount=<%=Server.URLEncode(BankGLAccount)%>&NextAct=showRemained">
					<INPUT TYPE="submit" value="œÊŒ ‰ « Ê„« Ìﬂ" class="GenButton" onclick="this.disabled=true;submit();">
				</FORM>
			</td>
			<td>
				<FORM METHOD=POST ACTION="?act=showAll&StartDate=<%=Server.URLEncode(StartDate)%>&EndDate=<%=Server.URLEncode(EndDate)%>&BankGLAccount=<%=Server.URLEncode(BankGLAccount)%>&NextAct=showRemained">
					<INPUT TYPE="submit" value="‰„«Ì‘ Â„Â" class="GenButton" onclick="this.disabled=true;submit();">
				</FORM>
			</td>
		</tr>
		</table>
		</td>
	</tr>
	</table>

	<script language="JavaScript">
	<!--
	//---------------------------------------------
	var slctRowClr='#FFFF00'
	var rltdRowClr='#99FF88'
	var sRltdRowClr='#9988FF'
	var slctdGLRow = null
	var slctdBnkRow = null
	var numRelations=1;

	//---------------------------------------------
	function DslctGLRow(src){
		if (src){
			src.parentNode.setAttribute("bgColor","");
		}
	}
	function slctGLRow(src){
		
		DslctGLRow(slctdGLRow)
		if (slctdGLRow == src){
			if (confirm("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« „” ﬁ·« Õ–› ﬂ‰Ìœø\n")){
				relate(true);
				return;
			}
			slctdGLRow = null
		}
		else{
			slctdGLRow = src
			src.parentNode.setAttribute("bgColor",slctRowClr)
			relate(false);
		}
	}
	//---------------------------------------------
	function DslctBnkRow(src){
		if (src){
			src.parentNode.setAttribute("bgColor","");
		}
	}
	function slctBnkRow(src){
		
		DslctBnkRow(slctdBnkRow)
		if (slctdBnkRow == src){
			if (confirm("¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ ”ÿ— —« „” ﬁ·« Õ–› ﬂ‰Ìœø\n")){
				relate(true);
				return;
			}
			slctdBnkRow = null
		}
		else{
			slctdBnkRow = src
			src.parentNode.setAttribute("bgColor",slctRowClr)
			relate(false);
		}
	}
	//---------------------------------------------
	function relate(singleRow){

		var BnkRow=0;
		var GLRow=0;
		var desc='';
		var tmpDlgTxt= new Object();
		tmpDlgTxt.value = ' Ê÷ÌÕÌ œ— «Ì‰ —«»ÿÂ: ';

		if (singleRow)
		{
			if (slctdBnkRow){
				BnkRow = slctdBnkRow.getElementsByTagName("input")[0].value
				slctdBnkRow.setAttribute("bgColor",sRltdRowClr)
				slctdBnkRow.innerHTML+=numRelations
			}
			else if (slctdGLRow){
				GLRow = slctdGLRow.getElementsByTagName("input")[0].value
				slctdGLRow.setAttribute("bgColor",sRltdRowClr)
				slctdGLRow.innerHTML+=numRelations
			}
		}
		else if (slctdBnkRow && slctdGLRow){

			BnkRow = slctdBnkRow.getElementsByTagName("input")[0].value
			GLRow = slctdGLRow.getElementsByTagName("input")[0].value

			slctdBnkRow.setAttribute("bgColor",rltdRowClr)
			slctdBnkRow.innerHTML+=numRelations
			slctdGLRow.setAttribute("bgColor",rltdRowClr)
			slctdGLRow.innerHTML+=numRelations
		}
		else{
			return;
		}

		dialogActive=true
		window.showModalDialog('../dialog_GenInput.asp',tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
		dialogActive=false
		desc=tmpDlgTxt.value

		numRelations++;


		document.getElementById("Relations").innerHTML+="<input type='hidden' name='BnkRow' value='"+BnkRow+"'>"
		document.getElementById("Relations").innerHTML+="<input type='hidden' name='GLRow' value='"+GLRow+"'>"
		document.getElementById("Relations").innerHTML+="<input type='hidden' name='Desc' value='"+escape(desc.replace("'","`"))+"'>"

		DslctBnkRow(slctdBnkRow)
		DslctBnkRow(slctdGLRow)

		slctdBnkRow = null
		slctdGLRow = null
	}

	//-->
	</script>

<%
elseif request("act")="autoRelate" then
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----
	StartDate =		request.QueryString("StartDate")
	EndDate =		request.QueryString("EndDate")
	BankGLAccount = cint(request.QueryString("BankGLAccount"))
	BankAccountNo = clng(lookupBankAccountNo(BankGLAccount))

	' Type 1:		Amount + Date + CheqNo
	mySQL="INSERT INTO BankContradictionRelations (BankLogID, GLRowID, Description, Type, CreatedBy, CreatedDate, CreatedTime) SELECT BankLog.autoKey AS BankLogID, RemainedGLRows.GLRowID, '' AS Description, 1 AS Type, '" & session ("ID") & "' AS CreatedBy, '" & shamsiToday() & "' AS CreatedDate, '" & currentTime10() & "' AS CreatedTime FROM BankLog INNER JOIN (SELECT ID AS GLRowID, CONVERT(bigint, Ref1) AS CheqNo, GLDocDate AS [Date], 1 - IsCredit AS IsDebit, Amount FROM dbo.EffectiveGLRows WHERE (GL = " & openGL & ") AND (GLAccount = " & BankGLAccount & ") AND (ID NOT IN (SELECT GLRowID FROM BankContradictionRelations))) RemainedGLRows ON BankLog.[Date] = RemainedGLRows.[Date] COLLATE Arabic_CI_AS AND BankLog.Amount = RemainedGLRows.Amount AND BankLog.isDebit = RemainedGLRows.IsDebit AND BankLog.CheqNo = RemainedGLRows.CheqNo WHERE (BankLog.AccountNo = " & BankAccountNo & ")"
	Conn.Execute(mySQL)
	response.write "Type 1: OK<BR><BR>"

	' Type 2:		Amount + Date + CheqNo
	mySQL="INSERT INTO BankContradictionRelations (BankLogID, GLRowID, Description, Type, CreatedBy, CreatedDate, CreatedTime) SELECT BankLog.autoKey AS BankLogID, RemainedGLRows.GLRowID, '' AS Description, 2 AS Type, '" & session ("ID") & "' AS CreatedBy, '" & shamsiToday() & "' AS CreatedDate, '" & currentTime10() & "' AS CreatedTime FROM BankLog INNER JOIN (SELECT ID AS GLRowID, CONVERT(bigint, Ref1) AS CheqNo, 1 - IsCredit AS IsDebit, Amount FROM dbo.EffectiveGLRows WHERE (GL = " & openGL & ") AND (GLAccount = " & BankGLAccount & ") AND (ID NOT IN (SELECT GLRowID FROM BankContradictionRelations))) RemainedGLRows ON BankLog.Amount = RemainedGLRows.Amount AND BankLog.isDebit = RemainedGLRows.IsDebit AND BankLog.CheqNo = RemainedGLRows.CheqNo WHERE (BankLog.AccountNo = " & BankAccountNo & ")"
	Conn.Execute(mySQL)
	response.write "Type 2: OK<BR><BR>"

	' Type 3:		Amount + Date 
	mySQL="INSERT INTO BankContradictionRelations (BankLogID, GLRowID, Description, Type, CreatedBy, CreatedDate, CreatedTime) SELECT BankLog.autoKey AS BankLogID, RemainedGLRows.GLRowID, '' AS Description, 3 AS Type, '" & session ("ID") & "' AS CreatedBy, '" & shamsiToday() & "' AS CreatedDate, '" & currentTime10() & "' AS CreatedTime FROM BankLog INNER JOIN (SELECT ID AS GLRowID, GLDocDate AS [Date], 1 - IsCredit AS IsDebit, Amount FROM dbo.EffectiveGLRows WHERE (GL = " & openGL & ") AND (GLAccount = " & BankGLAccount & ") AND (ID NOT IN (SELECT GLRowID FROM BankContradictionRelations))) RemainedGLRows ON BankLog.[Date] = RemainedGLRows.[Date] COLLATE Arabic_CI_AS AND BankLog.Amount = RemainedGLRows.Amount AND BankLog.isDebit = RemainedGLRows.IsDebit WHERE (BankLog.AccountNo = " & BankAccountNo & ")"
	Conn.Execute(mySQL)
	response.write "Type 3: OK<BR><BR>"

	response.redirect "?act=" & request("NextAct") & "&StartDate=" & Server.URLEncode(StartDate) & "&EndDate=" & Server.URLEncode(EndDate) & "&BankGLAccount=" & Server.URLEncode(BankGLAccount) & "&msg=" & Server.URLEncode("—Ê«»ÿ « Ê„« Ìﬂ «ÌÃ«œ ‘œ.")

elseif request("act")="submitRelations" then

	StartDate =		request.QueryString("StartDate")
	EndDate =		request.QueryString("EndDate")
	BankGLAccount = cint(request.QueryString("BankGLAccount"))

	for i=1 to request.form("BnkRow").count 
		BankRow = cdbl(request.form("BnkRow")(i))
		GLRow	= cdbl(request.form("GLRow")(i))
		Desc	= unescape(request.form("Desc")(i))

		' Type = 9 Means Manual
		mySQL="INSERT INTO BankContradictionRelations (BankLogID, GLRowID, Description, Type, CreatedBy, CreatedDate, CreatedTime) VALUES " &_ 
			 "('"& BankRow & "', '"& GLRow & "', N'"& sqlSafe(Desc) & "', '9', '"& session("ID") & "', '"& ShamsiToday() & "', '"& currentTime10() & "')"
		conn.Execute(mySQL)
	next

	response.redirect "?act=" & request("NextAct") & "&StartDate=" & Server.URLEncode(StartDate) & "&EndDate=" & Server.URLEncode(EndDate) & "&BankGLAccount=" & Server.URLEncode(BankGLAccount) & "&msg=" & Server.URLEncode("—Ê«»ÿ „Ê—œ ‰Ÿ— «ÌÃ«œ ‘œ.")

elseif request("act")="delRelations" then

	StartDate =		request.QueryString("StartDate")
	EndDate =		request.QueryString("EndDate")
	BankGLAccount = cint(request.QueryString("BankGLAccount"))

	for i=1 to request.form("BnkRow").count 
		BankRow = cdbl(request.form("BnkRow")(i))
		mySQL="DELETE FROM BankContradictionRelations WHERE (BankLogID='" & BankRow & "')"
		conn.Execute(mySQL)
	next

	for i=1 to request.form("GLRow").count 
		GLRow	= cdbl(request.form("GLRow")(i))
		mySQL="DELETE FROM BankContradictionRelations WHERE (GLRowID='" & GLRow & "')"
		conn.Execute(mySQL)
	next

	response.redirect "?act=" & request("NextAct") & "&StartDate=" & Server.URLEncode(StartDate) & "&EndDate=" & Server.URLEncode(EndDate) & "&BankGLAccount=" & Server.URLEncode(BankGLAccount) & "&msg=" & Server.URLEncode("—Ê«»ÿ ”ÿ— Â«Ì „Ê—œ ‰Ÿ— Õ–› ‘œ.")

'-----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------- Get parameters and file
'-----------------------------------------------------------------------------------------------------
elseif request("act")="" then

%>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<FORM METHOD=POST ACTION="?act=showRemained" onsubmit="return isEmpty()" ENCTYPE="multipart/form-data">
	<TABLE align=center style="border:1pt solid white">
	<TR>
		<TD>
			«“  «—ÌŒ: 
		</TD>
		<TD>
			<INPUT class="GenInput" style='direction:LTR;text-align:left;' NAME="StartDate" TYPE="text" maxlength="10" size="10" value="<%=Session("OpenGLStartDate")%>" onblur="acceptDate(this)">
		</TD>
	</TR>
	<TR>
		<TD>
			 «  «—ÌŒ: 
		</TD>
		<TD>
			<INPUT class="GenInput" style='direction:LTR;text-align:left;' NAME="EndDate" TYPE="text" maxlength="10" size="10" value="<%=shamsiToday()%>" onblur="acceptDate(this)">
		</TD>
	</TR>
	<TR>
		<TD>
			„⁄Ì‰ »«‰ﬂ :
		</TD>
		<TD>
			<select name="BankGLAccount" class=inputBut style="width:200; ">
				<option value="-1">«‰ Œ«» ﬂ‰Ìœ</option>
				<option value="12001">12001 - »«‰ﬂ  Ã«—  Ã«—Ì 24050 (»)</option>
				<option value="12004">12004 - »«‰ﬂ  Ã«—  Ã«—Ì 24816 («·›)</option>
				<option value="12002">12002 - »«‰ﬂ ”ÅÂ Ã«—Ì 4344/19 («·›)</option>
			</select>
		</TD>
	</TR>
	<TR>
		<TD>›«Ì· »«‰ﬂ</TD>
		<TD>
			<INPUT TYPE="FILE" NAME="LogFile" class=inputBut style="width:200; ">
		</TD>
	</TR>
	<TR>
		<TD colspan='2' height='2' ><hr></TD>
	</TR>
	<TR>
		<TD colspan='2' align='center'>
			<INPUT TYPE="submit" value=" «œ«„Â " class='GenButton' onclick="this.disabled=true;submit();"><br><br>
		</TD>
	</TR>
	</TABLE>
	</FORM>
	<BR><BR>

	<SCRIPT LANGUAGE="JavaScript">
	<!--

	var dialogActive=false;

	function isEmpty()
	{
		Bank = document.all.BankGLAccount.value
		if(Bank==-1){
			alert("Œÿ«! ›—„ ﬂ«„· Å— ‰‘œÂ «” ")
			return false
		}
		else{
			return true;
		}
	}
	//-->
	</SCRIPT>
<% 
end if
%>
<!--#include file="tah.asp" -->
