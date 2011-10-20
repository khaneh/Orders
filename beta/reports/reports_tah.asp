<%
	'mySQL = Request.Form("SQL") 
	'connector = " WHERE ("
	'Set Params = Request.Form
	'For Each p in Params
	'	Param = Params(p)
'		response.write "param : #" & p & " :: " & Param & "# <br>"
	'	If Param <> "" AND p <> "SQL" Then
	'		Do While Instr(Param,"||") <> 0
	'			mySQL = mySQL & connector & p & " " & Left(Param,Instr(Param,"||")-1)
	'			connector = " OR "
	'			Param = Right(Param,Len(Param) - Instr(Param,"||")-1)
	'		Loop
	'		mySQL = mySQL & connector & p & " " & Param & ")"
	'		connector = " AND ("
	'	End If
	'Next

'	------------------ 
'	-----  Show Result using:
'	-----  
'	-----  String:				mySQL	;	the SQL
'	-----  ADODB.Connection:	conn:	;	the connection string
'	-----  
'	------------------ 

'---------------------------------------------
'------------------------------------ Separate
'---------------------------------------------
function Separate(inputTxt)

if not isnumeric(inputTxt) or "" & inputTxt="" then 
	Separate=inputTxt
else
	myMinus=""
	input=inputTxt
	t=instr(input, ".")
	if t>0 then 
		expPart = mid(input, t+1, 2)
		input = left(input, t-1)
	end if
	if left(input,1)="-" then
		myMinus="-"
		input=right(input,len(input)-1)
	end if
	if len(input) > 3 then
		tmpr=right(input ,3)
		tmpl=left(input , len(input) - 3 )
		result = tmpr
		while len(tmpl) > 3
			tmpr=right(tmpl,3)
			result = tmpr & "," & result 
			tmpl=left(tmpl , len(tmpl) - 3 )
		wend
		if len(tmpl) > 0 then
			result = tmpl & "," & result
		end if 
	else
		result = input
	end if 
	if t>0 then 
		result = result & "." & expPart
	end if

	Separate=myMinus & result
end if
end function
'-------------------------------------------------------
'---------------------

	'Set conn = Server.CreateObject("ADODB.Connection")
	'conn.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("db1.mdb") & ";"

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open mySQL, conn, 3, 3
	
	'----------------------------------G e n   H T M L   T a b l e------------------------------------
	Dim myXls
	myXls = "<style> .resTable1 { Font-family:tahoma; Font-Size:9pt; Background-color:#336699; border: 1 solid #336699;} .resTable1 th {Background-color:#6699CC;} .resRow1 {Background-color:#F8F8FF;} .resRow0 {Background-color:#CCCCDD;} </style>"
	
	myXls = myXls & "<CENTER><H3>"&Server.HTMLEncode(reportTitle)&"</H3></CENTER>"
	myXls = myXls & "<table class='resTable1' Cellspacing=1 Cellpadding=2 align=center><tr>"
	
	For i = 0 To rs.Fields.Count - 1
		myXls = myXls & "<th>"&Server.HTMLEncode(rs.Fields(i).Name)&"</th>"
	
	Next
	myXls = myXls & "</tr>"
	
	if rs.eof then
		myXls = myXls & "<tr><td width='100%' class=resRow1 colspan="&rs.Fields.Count&" align='center'>جوابي يافت نشد</td></tr>"
	else
		rs.MoveFirst
		rowCounter=0
		do while Not rs.eof
			rowCounter = rowCounter + 1
			myXls = myXls &"<tr class='resRow"& (rowCounter mod 2) &"'>"
			For i = 0 To rs.Fields.Count - 1
				FieldValue = trim(rs.Fields(rs.Fields(i).Name).Value)
				'if isnumeric(FieldValue) then
				'FieldValue = 111
				'end if

				if FieldValue <> "" and not isnumeric(FieldValue) then
					FieldValue = Server.HTMLEncode(FieldValue)
				elseif FieldValue = "" then
					FieldValue = "&nbsp;"
				end if

				myXls = myXls & "<td >" & Separate(FieldValue) & "</td>"
			Next
		%>
		  </tr>
		<%
			rs.MoveNext
		loop
	End if
	rs.Close
	set rs = nothing
	conn.Close
	set conn = nothing
		
	myXls = myXls & "</table>"
	Response.Clear()
	Response.Buffer = True
	Response.AddHeader "Content-Disposition", "attachment;filename=" & reportTitle & ".xls"  
	Response.ContentType = "application/vnd.ms-excel"
	Response.Charset = "UTF-8"
	Response.Write myXls
	Response.End()
else
%>
<CENTER><H3><%=reportTitle%></H3></CENTER>
	<form METHOD="Post" ACTION="?act=show">
		<div align="left">
		<table align=center border="0" cellpadding="0" cellspacing="0" width="300">
		<tr align="center">
			<td width="142">date From</td>
			<td width="190"><input onblur="acceptDate(this)" NAME="dateFrom" size="20"></td>
		</tr>
		<tr align="center">
			<td width="142">date To</td>
			<td width="190"><input onblur="acceptDate(this)" NAME="dateTo" size="20"></td>
		</tr>
		<tr align="center">
			<td colspan=2>
				<br><input TYPE="Submit" VALUE="Run Query">
			</td>
		</tr>
		</table>
		</div>
	</form>

<%
end if
%>
</body>
</html>
