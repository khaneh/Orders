<!-- #INCLUDE FILE="../inc/mla_sql_config.asp" -->
<!-- #INCLUDE FILE="../inc/mla_sql_lang.asp" -->
<!-- #INCLUDE FILE="../inc/mla_sql_error.asp" -->
<!-- #INCLUDE FILE="../inc/mla_sql_start.asp" -->

<%
	' Global variables declaration
	Dim gObjDC	 ' connection object
	
	' ### Private functions to format and display content
	Dim mla_arr_num, mla_arr_date, mla_arr_bool, mla_arr_str, mla_arr_bin, mla_arr_text
	mla_arr_num = Array(16, 2, 3, 20, 17, 18, 19, 21, 4, 5, 6, 14, 131, 139)
	mla_arr_date= Array(7, 133, 134, 64, 137, 135)
	mla_arr_bool = Array(11)
	mla_arr_str = Array(8, 129, 130, 202, 200)
	mla_arr_bin= Array(128, 204, 205)
	mla_arr_text = Array(201, 203)

	Function getColAlign(pType)
		Select Case (True)
			Case (getOne(mla_arr_num, pType) >=0) : getColAlign = "RIGHT"
			Case (getOne(mla_arr_date, pType) >=0) : getColAlign = "CENTER"
			Case Else : getColAlign = "LEFT"
		End Select
	End Function

	Function getColSort(pType)
		Select Case (True)
			Case (getOne(mla_arr_bin, pType) >=0 Or getOne(mla_arr_bool, pType) >=0 Or getOne(mla_arr_text, pType) >=0) : getColSort = False
			Case Else : getColSort = True
		End Select
	End Function

	Function getColDisplay(pType)
		Select Case (True)
			Case (getOne(mla_arr_str, pType) >=0) : getColDisplay = 1
			Case (getOne(mla_arr_text, pType) >=0) : getColDisplay = 2
			Case (getOne(mla_arr_bin, pType) >=0) : getColDisplay = 3
			Case (getOne(mla_arr_date, pType) >=0) : getColDisplay = 4
			Case Else : getColDisplay = 0
		End Select
	End Function

	' ###  mlt_string : allows faster string concatenation than classical & operator
	Class mlt_string
		Dim gStrArray, gInitialSize, gCurrentSize
		Private Sub Class_Initialize
			strInit()
		End Sub
		Public Sub strInit()
			gInitialSize = 100
			gCurrentSize = 0
			ReDim gStrArray(gInitialSize)
		End Sub
		Public Sub strAppend(Byval pStr)
			If pStr <> "" And VarType(pStr) <> vbNull And Not IsNull(pStr)Then 
				If gCurrentSize >= gInitialSize Then Redim Preserve gStrArray(gInitialSize + gCurrentSize)
				gStrArray(gCurrentSize) = pStr
				gCurrentSize = gCurrentSize + 1
			End If
		End Sub
		Public Function getStr()
			getStr = Join(gStrArray, "")
		End Function
	End Class

	' ### mla_getConnStr : returns a connection string (DSN Less)
	Function mla_getConnStr(pProvider, pDataSource, pPortNumber, pNetworkLib, pIniCat, pTrusted, pUid, pPwd)
		Dim myConnStr
		If pProvider = "SQLOLEDB" Then		' ### OLE db connection string
			myConnStr  = "Provider=sqloledb;Data Source=" & pDataSource
			If pPortNumber <> "" Then myConnStr = myConnStr & "," & pPortNumber End If
			myConnStr = myConnStr & ";"
			If pNetworkLib <> "" Then myConnStr = myConnStr & "Network Library=" & pNetworkLib & ";" End If
			If pIniCat <> "" Then myConnStr = myConnStr & "Initial Catalog=" & pIniCat & ";" End If
			If pTrusted <> "" Then myConnStr = myConnStr & "Integrated Security=SSPI;" End If
			myConnStr = myConnStr & "User Id=" & pUid & ";"
			myConnStr = myConnStr & "Password=" & pPwd & ";"
		Else																									' ### ODBC connection string
			myConnStr  = "Driver={SQL Server};Server=" & pDataSource & ";"
			If pPortNumber <> "" Then myConnStr = myConnStr & "Address=" & pDataSource & "," & pPortNumber & ";" End If
			If pNetworkLib <> "" Then myConnStr = myConnStr & "Network	=" & pNetworkLib & ";" End If 
			If pIniCat <> "" Then myConnStr = myConnStr & "Database=" & pIniCat & ";" End If
			If pTrusted <> "" Then myConnStr = myConnStr & "Trusted_Connection=yes;" End If
			myConnStr = myConnStr & "Uid=" & pUid & ";"
			myConnStr = myConnStr & "Pwd=" & pPwd & ";"
		End If
		mla_getConnStr = myConnStr
	End Function

	' ### mla_getConnStr_dsn : returns a connection string (DSN)
	Function mla_getConnStr_dsn(pDSN, pUid, pPwd)
		Dim myConnStr
		myConnStr="dsn=" & pDSN & ";uid=" & pUid & ";pwd=" & pPwd
		mla_getConnStr_dsn = myConnStr
	End Function

	' ### remquote : replaces " by \" and vbCrlf by \n
	Function remdquote(pStr)
		If pStr = "" Or IsNull(pStr) Then
			remdquote = ""
		Else
			remdquote = Replace(Replace(Replace(pStr, "\", "\\"), """", "\"""), vbCrlf, "\n")
		End If
	End Function

	' ### remquote : replaces ' by ''
	Function remquote(pStr)
		If pStr = "" Or IsNull(pStr) Then
			remquote = ""
		Else
			remquote = Replace(pStr, "'", "''")
		End If
	End Function

	' ### rembracket : replaces ] by ]]
	Function rembracket(pStr)
		If pStr = "" Or IsNull(pStr) Then
			rembracket = ""
		Else
			rembracket = Replace(pStr, "]", "]]")
		End If
	End Function

	' ### resbracket : replaces ]] by ]
	Function resbracket(pStr)
		If pStr = "" Or IsNull(pStr) Then
			resbracket = ""
		Else
			resbracket = Replace(pStr, "]]", "]")
		End If
	End Function

	' ### txt2html : replaces vbCrlf by <BR> and vbTab by &nbsp;&nbsp;&nbsp;
	Function txt2html(pStr)
		If pStr = "" Or IsNull(pStr) Then
			txt2html = ""
		Else
			txt2html = Replace(Replace(Replace(Server.HTMLEncode(pStr), vbCrlf, "<BR>"), vbTab, "&nbsp;&nbsp;&nbsp;"), "  ", "&nbsp;&nbsp;")
		End If
	End Function

	' ### getStrBegin : returns an array with the X first characters of the string and a boolean to know if the string has been cut
	Function getStrBegin(pStr, pLength)
		Dim myC
		If pStr = "" Or IsNull(pStr) Then
			getStrBegin = Array("", false)
		ElseIf Len(pStr) <= pLength Then
			getStrBegin = Array(pStr, false)
		Else
			myC = InStr(pLength, pStr, " ")
			If myC > 0 Then getStrBegin = Array(Left(pStr, myC), true) Else getStrBegin = Array(pStr, false) End If
		End If
	End Function

	' ### bin2hex : returns hexadecimal representation of a binary string
	Function bin2hex(pBin, pLen)
		Dim i, myL, myStr, myFlag
		myStr = "0x"
		If LenB(pBin) < pLen Then
			myL = LenB(pBin)
			myFlag = false
		Else
			myL = pLen
			myFlag = true
		End If
		For i = 1 To myL
			myStr = myStr & Hex(AscB(MidB(pBin, i, 1))) 
		Next
		bin2hex = Array(myStr, myFlag)
	End Function

	' ### i2a : returns number with a specified length
	Function i2a(pInt, pLength)
		i2a = Right(String(pLength, "0") & CStr(pInt) , pLength)
	End Function

	' ### str2date : returns date in ISO format
	Function str2date(pStr)
	Dim myDate, myDD, myMM, myYYYY, myHH, myMin, mySec, mySep, myArray, myTime, myVBDllBug
		mySep = ""
		myArray = Split(Trim(pStr), " ", 2)
		If UBound(myArray) > 0 Then
			myDate = Trim(myArray(0))
			myTime = Trim(myArray(1))
			If UBound(Split(myDate, "/")) = 1 AND UBound(Split(myTime, ":")) = 2 Then
				myVBDllBug = True
			End If
		End If
		If myVBDllBug = True Then
			str2date = NULL
		ElseIf Not IsDate(pStr) Then
			str2date = NULL
		Else
			myDate = CDate(pStr)
			myDD = i2a(Day(myDate), 2)
			myMM = i2a(Month(myDate), 2)
			myYYYY = Year(myDate)
			myHH = i2a(Hour(myDate), 2)
			myMin = i2a(Minute(myDate), 2)
			mySec = i2a(Second(myDate), 2)
			str2date = myYYYY & mySep & myMM & mySep & myDD & " " & myHH & ":" & myMin & ":" & mySec
		End If
	End Function

	' ### returns max value between 2 int
	Function Max(pInt1, pInt2)
		If pInt1 > pInt2 Then
			Max = pInt1
		Else
			Max = pInt2
		End If
	End Function

	' ### returns min value between 2 int
	Function Min(pInt1, pInt2)
		If pInt1 < pInt2 Then
			Min = pInt1
		Else
			Min = pInt2
		End If
	End Function
	
	' ### returns item position inside an array
	Function getOne(pArray, pStr)
	Dim i
		If IsArray(pArray) Then
			getOne = -13
			For i = 0 To UBound(pArray)
				If pStr = pArray(i) Then
					getOne = i
					Exit For
				End If
			Next
		Else
			getOne = -1
		End If
	End Function

	' ### openConnection : opens a connection to the dabase with the connection string stored in the Session Variable
	Sub openConnection
		If Session("isConnected") Then
			On Error Resume Next
			Set gObjDC = Server.CreateObject("ADODB.Connection")
			gObjDC.ConnectionTimeOut = 15
			gObjDC.CommandTimeOut = 30
			gObjDC.Open Session("ConnStr")
'			If gObjDC.Errors.Count > 0 Then 
'				Session.Abandon()
'				Call mla_displayError(Err, Session("ConnStr"))
'			End If
			If Err <> 0 Then 
				Session.Abandon()
				Call mla_displayError(Err, Session("ConnStr"))
			End If
		Else
			Response.Redirect "../conn/expired.asp"
			Response.End
		End If
	End Sub

	' ### closeConnection : closes the connection opened by openConnection
	Sub closeConnection
		If gObjDC.State > 0 Then gObjDC.Close
		Set gObjDC = Nothing
	End Sub

	' ### getTreeStr : returns tree HTML sentence
	Function getTreeStr(pType, pArray)
		Dim myStrHTML
		myStrHTML = "<A HREF=""../hlp/default.asp"">myLittleAdmin</A>"
		Select Case Left(pType, 1)
			Case "1" :	' ### Connections
				myStrHTML = myStrHTML & " \ <A HREF=""../conn/default.asp"">" & myTObj.getTerm(2) & "</A>"
				If Mid(pType, 3, 1) <> "" Then
					Select Case Mid(pType, 3, 1)
						Case "1" : myStrHTML = myStrHTML & " \ <A HREF=""../conn/dsnless.asp"">" & myTObj.getTerm(3) & "</A>"	' ### DSN Less
						Case "2" : myStrHTML = myStrHTML & " \ <A HREF=""../conn/dsn.asp"">" & myTObj.getTerm(4) & "</A>"	' ### DSN
					End Select
				End If
			Case "2" :	' ### Server
				myStrHTML = myStrHTML & " \ <A HREF=""../srv/srvinfo.asp"">" & getSrvName() & "</A>"
				Select Case Mid(pType, 3, 1)
					Case "1" :	' ### Databases
						myStrHTML = myStrHTML & " \ <A HREF=""../db/default.asp"">" & myTObj.getTerm(5) & "</A>"
						If Mid(pType, 5, 1) <> "" Then
							myStrHTML = myStrHTML & " \ <A HREF=""../db/dbinfo.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & pArray(0) & "</A>"	' ### Database
							Select Case Mid(pType, 7, 1)
								Case "1" : 	' ### Tables
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbtbl.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(7) & "</A>"
									Select Case Mid(pType, 9, 1) 
										Case "1" : myStrHTML = myStrHTML & " \ <A HREF=""../db/content.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&obj=" & Server.URLEncode(pArray(1)) & "&type=tbl"">" & pArray(1) & "</A>"	' ### Table Content
										Case "2" : myStrHTML = myStrHTML & " \ <A HREF=""../db/tblstruct.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&tbl=" & Server.URLEncode(pArray(1)) & """>" & pArray(1) & "</A>"	' ### Table Struct
										Case "3" : myStrHTML = myStrHTML & " \ <A HREF=""../db/search.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&obj=" & Server.URLEncode(pArray(1)) & "&type=tbl"">" & pArray(1) & "</A>"	' ### Table Search
									End Select
									If Mid(pType, 11, 1) <> "" Then myStrHTML = myStrHTML & " \ <A HREF=""../db/trstruct.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&tbl=" & Server.URLEncode(pArray(1)) & "&tr=" & Server.URLEncode(pArray(2)) & """>" & pArray(2) & "</A>"	' ### Trigger Struct
								Case "2" : 	' ### Views
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbview.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(8) & "</A>"
									Select Case Mid(pType, 9, 1) 
										Case "1" : myStrHTML = myStrHTML & " \ <A HREF=""../db/content.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&obj=" & Server.URLEncode(pArray(1)) & "&type=view"">" & pArray(1) & "</A>"	' ### View Content
										Case "2" : myStrHTML = myStrHTML & " \ <A HREF=""../db/viewstruct.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&view=" & Server.URLEncode(pArray(1)) & """>" & pArray(1) & "</A>"	' ### View Struct
										Case "3" : myStrHTML = myStrHTML & " \ <A HREF=""../db/search.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&obj=" & Server.URLEncode(pArray(1)) & "&type=view"">" & pArray(1) & "</A>"	' ### Table Search
									End Select
								Case "3" : 	' ### SP
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbsp.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(9) & "</A>"
									If Mid(pType, 9, 1) <> "" Then myStrHTML = myStrHTML & " \ <A HREF=""../db/spstruct.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&sp=" & Server.URLEncode(pArray(1)) & """>" & pArray(1) & "</A>"	' ### SP Struct
								Case "4" :	' ### USR
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbuser.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(10) & "</A>"
									If Mid(pType, 9, 1) <> "" Then myStrHTML = myStrHTML & " \ <A HREF=""../db/edituser.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&usr=" & Server.URLEncode(pArray(1)) & """>" & pArray(1) & "</A>"
								Case "5" :	' ### ROLE
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbrole.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(11) & "</A>"
									If Mid(pType, 9, 1) <> "" Then myStrHTML = myStrHTML & " \ <A HREF=""../db/editrole.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&role=" & Server.URLEncode(pArray(1)) & """>" & pArray(1) & "</A>"
								Case "6" :	' ### RULE
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbrule.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(12) & "</A>"
									If Mid(pType, 9, 1) <> "" Then myStrHTML = myStrHTML & " \ " & pArray(1)
								Case "7" : ' ### DEFAULT
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbdefault.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(13) & "</A>"
									If Mid(pType, 9, 1) <> "" Then myStrHTML = myStrHTML & " \ " & pArray(1)
								Case "8" :	' ### UDT
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbudt.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(14) & "</A>"
									If Mid(pType, 9, 1) <> "" Then myStrHTML = myStrHTML & " \ " & pArray(1)
								Case "9" : 	' ### UDF
									myStrHTML = myStrHTML & " \ <A HREF=""../db/dbudf.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & """>" & myTObj.getTerm(16) & "</A>"
									If Mid(pType, 9, 1) <> "" Then myStrHTML = myStrHTML & " \ <A HREF=""../db/udfstruct.asp?nid=M2_1_" & Mid(pType, 5, 1) & "&db=" & Server.URLEncode(pArray(0)) & "&udf=" & Server.URLEncode(pArray(1)) & """>" & pArray(1) & "</A>"	' ### UDF Struct
							End Select
						End If
					Case "2" :	' ### Management
						myStrHTML = myStrHTML & " \ <A HREF=""../act/default.asp"">" & myTObj.getTerm(15) & "</A>"
						If Mid(pType, 5, 1) <> "" Then
							myStrHTML = myStrHTML & " \ <A HREF=""../act/default.asp"">" & myTObj.getTerm(26) & "</A>"	' ### Current Activity
							Select Case Mid(pType, 7, 1)
								Case "1" : myStrHTML = myStrHTML & " \ <A HREF=""../act/actspid.asp"">" & myTObj.getTerm(27) & "</A>"	' ### Processus info
								Case "2" : myStrHTML = myStrHTML & " \ <A HREF=""../act/actlockid.asp"">" & myTObj.getTerm(28) & "</A>"	' ### Lock / Processus ID
								Case "3" : myStrHTML = myStrHTML & " \ <A HREF=""../act/actlockobj.asp"">" & myTObj.getTerm(29) & "</A>"	' ### Lock / Objects
							End Select
						End If
					Case "3" :	' ### Security
						myStrHTML = myStrHTML & " \ <A HREF=""../srv/default.asp"">" & myTObj.getTerm(19) & "</A>"
						Select Case Mid(pType, 5, 1)
							Case "1" : myStrHTML = myStrHTML & " \ <A HREF=""../srv/srvcon.asp"">" & myTObj.getTerm(20) & "</A>"	' ### Connections
							Case "2" : myStrHTML = myStrHTML & " \ <A HREF=""../srv/srvrole.asp"">" & myTObj.getTerm(21) & "</A>"	' ### Server roles
						End Select
					Case "4" :	' ### Tools
						myStrHTML = myStrHTML & " \ <A HREF=""../tools/default.asp"">" & myTObj.getTerm(30) & "</A>"
						Select Case Mid(pType, 5, 1)
							Case "1" : myStrHTML = myStrHTML & " \ <A HREF=""../tools/query.asp"">" & myTObj.getTerm(31) & "</A>"	' ### Query Analyser
						End Select
				End Select
			Case "3" :	' ### Preferences
				myStrHTML = myStrHTML & " \ <A HREF=""../pref/default.asp"">" & myTObj.getTerm(22) & "</A>"
				If Mid(pType, 3, 1) <> "" Then
					Select Case Mid(pType, 3, 1)
						Case "1" : myStrHTML = myStrHTML & " \ <A HREF=""../pref/theme.asp"">" & myTObj.getTerm(23) & "</A>"	' ### Theme
						Case "2" : myStrHTML = myStrHTML & " \ <A HREF=""../pref/language.asp"">" & myTObj.getTerm(24) & "</A>"	' ### Language
						Case "3" : myStrHTML = myStrHTML & " \ <A HREF=""../pref/display.asp"">" & myTObj.getTerm(25) & "</A>"	' ### Display
					End Select
				End If
		End Select
		getTreeStr = myStrHTML
	End Function

	' ### getListBox : returns the HTML code for a list box with the result of the request
	Function getListBox(pName, pArrValue, pValueIndex, pDisplayIndex, pFirstLine, pSelectedValue, pClass, pJavaScript)
		Dim myHTMLstr, i, myLen
		myHTMLstr = "<SELECT NAME=""" & pName & """"
		If pClass <> "" Then myHTMLstr = myHTMLstr & " CLASS=""" & pClass & """"
		If pJavaScript <> "" Then myHTMLstr = myHTMLstr & " onchange=""" & pJavaScript & """"
		myHTMLstr = myHTMLstr & ">" & vbCrLf
		If pFirstLine <> "" Then
			myHTMLstr = myHTMLstr & vbTab & "<OPTION VALUE=-1"
			If pSelectedValue = -1 Then myHTMLstr = myHTMLstr & " SELECTED"
			myHTMLstr = myHTMLstr & ">" & pFirstLine & "</OPTION>" & vbCrLf
		End If
		If isArray(pArrValue) Then
			myLen = UBound(pArrValue, 2)
			For i = 0 To myLen
				myHTMLstr = myHTMLstr & "<OPTION"
				If pValueIndex <> pDisplayIndex Then myHTMLstr = myHTMLstr & " VALUE=""" & pArrValue(pValueIndex, i) & """"
				If CStr(pSelectedValue) = CStr(pArrValue(pValueIndex, i)) Then myHTMLstr = myHTMLstr & " SELECTED"
				myHTMLstr = myHTMLstr & ">" & pArrValue(pDisplayIndex, i) & "</OPTION>" & vbCrLf
			Next
		End If
		myHTMLstr = myHTMLstr & "</SELECT>"  & vbCrlf
		getListBox = myHTMLstr
	End Function

	' ### getArrFromSQL : returns array from a SQL request using the getRows method
	Function getArrFromSQL(pStr)
		Dim myObjRS, myArrRet
		On Error Resume Next
		Set myObjRS = Server.CreateObject("ADODB.RecordSet")
		myObjRS.ActiveConnection = gObjDC
		myObjRS.CursorType = 0
		myObjRS.Open pStr
		If Err <> 0 Then
			Set myObjRS = Nothing
			Call mla_displayError(Err, pStr)
		End If
		If Not myObjRS.EOF Then
			myArrRet = myObjRS.getRows()
		Else
			myArrRet = Empty
		End If
		myObjRS.Close
		Set myObjRS = Nothing
		getArrFromSQL = myArrRet
	End Function

	' ### mla_execute : executre a SQL statement ###
	Function mla_execute(pDbName, pStr)
	Dim myLng
		On Error Resume Next
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		If Err <> 0 Then Call mla_displayError(Err, "")
		gObjDC.execute pStr, myLng
		If Err <> 0 Then Call mla_displayError(Err, pStr)
		mla_execute = myLng
	End Function

	' ### getSQLSrvDate() : returns date from SQL Server
	Function getSQLSrvDate()
		Dim myStrSQL
		myStrSQL = "SELECT GETDATE();"
		getSQLSrvDate = getArrFromSQL(myStrSQL)(0, 0)
	End Function

	' ### getSrvName() : returns server name
	Function getSrvName()
		Dim myStrSQL
		myStrSQL = "SELECT @@SERVERNAME;"
		getSrvName = getArrFromSQL(myStrSQL)(0, 0)
	End Function

	' ### getSrvInfo() : returns server info
	Function getSrvInfo()
		Dim myStrSQL
		myStrSQL = "SELECT @@SERVERNAME, @@VERSION, @@LANGUAGE;"
		getSrvInfo = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbList : returns database list
	Function getDbList
		Dim myStrSQL
		If mla_cfg_showsysdatabases Then
			myStrSQL = "SELECT name FROM master.dbo.sysdatabases WHERE has_dbaccess(name) = 1 ORDER BY name;"
		Else
			myStrSQL = "SELECT name FROM master.dbo.sysdatabases WHERE has_dbaccess(name) = 1 AND name NOT IN ('master', 'tempdb', 'msdb', 'model') ORDER BY name;"
		End If
		getDbList = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbFullList : returns complete database list
	Function getDbFullList
		Dim myStrSQL
		myStrSQL = "SELECT name FROM master.dbo.sysdatabases ORDER BY name;"
		getDbFullList = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbInfo : returns database information
	Function getDbInfo(pDbName)
		Dim myStrSQL
		myStrSQL = "SELECT t1.owner, t1.crdate, t1.size, t2.DBBupDate, t3.DifBupDate, t4.JournalBupDate FROM (SELECT d.name, suser_sname(d.sid) AS owner, d.crdate, (SELECT STR(SUM(CONVERT(DEC(15), f.size)) * (SELECT v.low FROM master.dbo.spt_values v WHERE v.type = 'E' AND v.number = 1) / 1048576, 10, 2) + 'MB' FROM [" & remquote(pDbName) & "].dbo.sysfiles f) AS size FROM master.dbo.sysdatabases d WHERE d.name = '" & remquote(pDbName) & "') AS t1 LEFT JOIN (SELECT '" & remquote(pDbName) & "' AS name, MAX(backup_finish_date) AS DBBupDate FROM msdb.dbo.backupset WHERE type = 'D' AND database_name = '" & remquote(pDbName) & "') AS t2 ON t1.name = t2.name LEFT JOIN (SELECT '" & remquote(pDbName) & "' AS name, MAX(backup_finish_date) AS DifBupDate FROM msdb.dbo.backupset WHERE type = 'I' AND database_name = '" & remquote(pDbName) & "') AS t3 ON t1.name = t3.name LEFT JOIN (SELECT '" & remquote(pDbName) & "' AS name, MAX(backup_finish_date) AS JournalBupDate FROM msdb.dbo.backupset WHERE type = 'L' AND database_name = '" & remquote(pDbName) & "') AS t4 ON t1.name = t4.name;"
		getDbInfo = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbFileInfo : returns database files information
	Function getDbFileInfo(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = " EXEC sp_helpfile;"
		getDbFileInfo = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbFileGroup : returns database file groups
	Function getDbFileGroup(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT groupname FROM sysfilegroups"
		getDbFileGroup = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbTbl : returns table list
	Function getDbTbl(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		If mla_cfg_showsystables Then
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, g.groupname, i.rows, o.type FROM [" & remquote(pDBName) & "].dbo.sysobjects o INNER JOIN [" & remquote(pDBName) & "].dbo.sysindexes i ON o.id = i.id INNER JOIN [" & remquote(pDBName) & "].dbo.sysfilegroups g ON i.groupid = g.groupid WHERE o.type IN ('U', 'S') AND i.indid < 2 AND (PERMISSIONS(o.id) & 1 = 1 OR PERMISSIONS(o.id) & 4096 = 4096) ORDER BY o.uid, o.name;"
		Else
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, g.groupname, i.rows, o.type FROM [" & remquote(pDBName) & "].dbo.sysobjects o INNER JOIN [" & remquote(pDBName) & "].dbo.sysindexes i ON o.id = i.id INNER JOIN [" & remquote(pDBName) & "].dbo.sysfilegroups g ON i.groupid = g.groupid WHERE o.type = 'U' AND o.name <> 'dtproperties' AND i.indid < 2 AND (PERMISSIONS(o.id) & 1 = 1 OR PERMISSIONS(o.id) & 4096 = 4096) ORDER BY o.uid, o.name;"
		End If
		getDbTbl = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbView : returns view list
	Function getDbView(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		If mla_cfg_showsysviews Then
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, o.category FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE o.type = 'V' AND (PERMISSIONS(o.id) & 1 = 1 OR PERMISSIONS(o.id) & 4096 = 4096) ORDER BY o.uid, o.name;"
		Else
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, o.category FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE o.type = 'V' AND o.category = 0 AND (PERMISSIONS(o.id) & 1 = 1 OR PERMISSIONS(o.id) & 4096 = 4096) ORDER BY o.uid, o.name;"
		End If
		getDbView = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbSp : returns sp list
	Function getDbSp(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		If mla_cfg_showsysprocedures Then
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, o.category FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE o.type = 'P' AND PERMISSIONS(o.id) & 32 = 32 ORDER BY o.uid, o.name;"
		Else
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, o.category FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE o.type = 'P' AND o.category = 0 AND PERMISSIONS(o.id) & 32 = 32 ORDER BY o.uid, o.name;"
		End If
		getDbSp = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbFunction : returns function list
	Function getDbFunction(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		If mla_cfg_showsysfunctions Then
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, o.category FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE (o.type = 'IF' OR o.type = 'FN') AND PERMISSIONS(o.id) & 32 = 32 ORDER BY o.uid, o.name;"
		Else
			myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, o.category FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE (o.type = 'IF' OR o.type = 'FN') AND o.category = 0 AND PERMISSIONS(o.id) & 32 = 32 ORDER BY o.uid, o.name;"
		End If
		getDbFunction = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbRule : returns rule list
	Function getDbRule(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, CASE WHEN CHARINDEX('AS', c.text) > 0 THEN RIGHT(c.text, LEN(c.text) - CHARINDEX('AS', c.text) - 1) ELSE c.text END FROM [" & remquote(pDBName) & "].dbo.sysobjects o INNER JOIN [" & remquote(pDBName) & "].dbo.syscomments c ON o.id = c.id WHERE o.type = 'R' AND o.category = 0 ORDER BY o.uid, o.name;"
		getDbRule = getArrFromSQL(myStrSQL)
	End Function

	' ### getSimpleDbRule : returns rule list
	Function getSimpleDbRule(pDbName)
		Dim myStrSQL
		myStrSQL = "SELECT o.id, o.name FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE o.type = 'R' AND o.category = 0 ORDER BY o.name;"
		getSimpleDbRule = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbDefault : returns default list
	Function getDbDefault(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT o.name, user_name(o.uid) AS owner, o.crdate, CASE WHEN CHARINDEX('AS', c.text) > 0 THEN RIGHT(c.text, LEN(c.text) - CHARINDEX('AS', c.text)-1) ELSE c.text END FROM [" & remquote(pDBName) & "].dbo.sysobjects o INNER JOIN [" & remquote(pDBName) & "].dbo.syscomments c ON o.id = c.id WHERE o.type = 'D' AND o.category = 0 ORDER BY o.uid, o.name;"
		getDbDefault = getArrFromSQL(myStrSQL)
	End Function

	' ### getSimpleDbDefault : returns default list
	Function getSimpleDbDefault(pDbName)
		Dim myStrSQL
		myStrSQL = "SELECT o.id, o.name FROM [" & remquote(pDBName) & "].dbo.sysobjects o WHERE o.type = 'D' AND o.category = 0 ORDER BY o.name;"
		getSimpleDbDefault = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbUdt : returns user defined type list
	Function getDbUdt(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT t.name, user_name(t.uid) AS owner, t2.name, t.length, CONVERT(TINYINT, t.allownulls), user_name(o2.uid) + '.' + o2.name, user_name(o.uid) + '.' + o.name FROM [" & remquote(pDBName) & "].dbo.systypes t INNER JOIN [" & remquote(pDBName) & "].dbo.systypes t2 ON t.xtype = t2.xusertype LEFT JOIN [" & remquote(pDBName) & "].dbo.sysobjects o ON t.domain = o.id LEFT JOIN [" & remquote(pDBName) & "].dbo.sysobjects o2 ON t.tdefault = o2.id WHERE t.xusertype >= 256 AND t.name <> 'sysname' ORDER BY t.uid, t.name;"
		getDbUdt = getArrFromSQL(myStrSQL)
	End Function

	' ### getUdtInfo : returns info for a specified user defined type
	Function getUdtInfo(pDbName, pUdtName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT t.name, user_name(t.uid) AS owner, t2.name, t.length, CONVERT(TINYINT, t.allownulls), ISNULL(o2.name, ''), ISNULL(o.name, '') FROM [" & remquote(pDBName) & "].dbo.systypes t INNER JOIN [" & remquote(pDBName) & "].dbo.systypes t2 ON t.xtype = t2.xusertype LEFT JOIN [" & remquote(pDBName) & "].dbo.sysobjects o ON t.domain = o.id LEFT JOIN [" & remquote(pDBName) & "].dbo.sysobjects o2 ON t.tdefault = o2.id WHERE t.name = '" & remquote(pUdtName) & "';"
		getUdtInfo = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbUser : returns user list
	Function getDbUser(pDbName)
		Dim myStrSQL
		myStrSQL = "SELECT u.name, l.loginname FROM [" & remquote(pDBName) & "].dbo.sysusers u LEFT JOIN master.dbo.syslogins l ON u.sid = l.sid WHERE u.islogin = 1 AND u.isaliased = 0 AND u.hasdbaccess = 1 ORDER BY u.name;"
		getDbUser = getArrFromSQL(myStrSQL)
	End Function

	' ### getDbRole : returns role list
	Function getDbRole(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "EXEC sp_helprole;"
		getDbRole = getArrFromSQL(myStrSQL)
	End Function

	' ### getSrvCon : returns connection list
	Function getSrvCon()
		Dim myStrSQL
		myStrSQL = "SELECT loginname, CASE WHEN isntname = 1 AND isntgroup = 1 THEN 1 WHEN isntname = 1 AND isntuser = 1 THEN 2 ELSE 0 END AS type, dbname, la.alias, createdate FROM master.dbo.syslogins lo INNER JOIN master.dbo.syslanguages la ON lo.language = la.name ORDER BY loginname;"
		getSrvCon = getArrFromSQL(myStrSQL)
	End Function

	' ### : getSrvRole : returns server role list
	Function getSrvRole()
		Dim myStrSQL
		myStrSQL = "EXEC sp_helpsrvrole;"
		getSrvRole = getArrFromSQL(myStrSQL)
	End Function

	' ### getTblCol : returns column list
	Function getTblCol(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT c.name, type_name(c.xusertype) AS type, c.length AS length, CASE WHEN CHARINDEX(TYPE_NAME(c.xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney') > 0 THEN COLUMNPROPERTY(c.id, c.name, 'precision') ELSE NULL END AS mPrecision, CASE WHEN CHARINDEX(TYPE_NAME(c.xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney') > 0 THEN OdbcScale(c.xtype,c.xscale) ELSE NULL END AS mScale, CONVERT(BIT, c.isnullable), CASE WHEN o.parent_obj > 0 THEN co.text ELSE USER_NAME(o.uid)+ '.' + o.name END AS mDefault, CASE WHEN o2.parent_obj > 0 THEN co2.text ELSE USER_NAME(o2.uid)+ '.' + o2.name END AS mRule, CONVERT(BIT, COLUMNPROPERTY(OBJECT_ID('" & remquote(pTblName) & "'), c.name, 'IsIdentity')) AS isIdentity, CASE WHEN COLUMNPROPERTY(OBJECT_ID('" & remquote(pTblName) & "'), c.name, 'IsIdentity') = 1 THEN IDENT_INCR('" & remquote(pTblName) & "') ELSE NULL END AS mStart, CASE WHEN COLUMNPROPERTY(OBJECT_ID('" & remquote(pTblName) & "'), c.name, 'IsIdentity') = 1 THEN IDENT_SEED('" & remquote(pTblName) & "') ELSE NULL END AS mSeed, CONVERT(BIT, COLUMNPROPERTY(OBJECT_ID('" & remquote(pTblName) & "'), c.name, 'IsRowGuidCol')) AS isRowGuidCol, CASE WHEN c.colid IN (SELECT k.colid FROM sysindexes i RIGHT JOIN sysindexkeys k ON i.id = k.id AND i.indid = k.indid WHERE status & 2048 = 2048 AND k.id = c.id) THEN 1 ELSE 0 END FROM syscolumns c LEFT JOIN (sysobjects o INNER JOIN syscomments co ON o.id = co.id) ON c.cdefault = o.id LEFT JOIN (sysobjects o2 INNER JOIN syscomments co2 ON o2.id = co2.id) ON c.domain = o2.id WHERE c.id = OBJECT_ID('" & remquote(pTblName) & "') ORDER BY c.colid;"
		getTblCol = getArrFromSQL(myStrSQL)
	End Function

	' ### getSimpleTblCol : returns column name list and misc. prop. for display
	Function getSimpleTblCol(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT name, CASE WHEN CHARINDEX(TYPE_NAME(xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney, bit') > 0 THEN 'RIGHT' WHEN CHARINDEX(TYPE_NAME(xtype), N'datetime,smalldatetime,timestamp') > 0 THEN 'CENTER' ELSE 'LEFT' END AS mAlign, CASE WHEN CHARINDEX(TYPE_NAME(xtype), N'text, ntext, binary, varbinary, image, bit') > 0 THEN CONVERT(BIT, 0) ELSE CONVERT(BIT, 1) END AS isSortable, CASE WHEN CHARINDEX(TYPE_NAME(xtype), N'char, nchar, varchar, nvarchar') > 0 THEN 1 WHEN CHARINDEX(TYPE_NAME(xtype), N'text, ntext') > 0 THEN 2 WHEN CHARINDEX(TYPE_NAME(xtype), N'binary, varbinary, image,timestamp') > 0 THEN 3 WHEN CHARINDEX(TYPE_NAME(xtype), N'datetime,smalldatetime') > 0 THEN 4 WHEN CHARINDEX(TYPE_NAME(xtype), N'bit') > 0 THEN 5 WHEN CHARINDEX(TYPE_NAME(xtype), N'uniqueidentifier') > 0 THEN 6 ELSE 0 END AS displayType, COLUMNPROPERTY(OBJECT_ID('" & remquote(pTblName) & "'), name, 'IsIdentity') AS isIdentity, COLUMNPROPERTY(OBJECT_ID('" & remquote(pTblName) & "'), name, 'AllowsNull') AS isIdentity FROM syscolumns WHERE id = OBJECT_ID('" & remquote(pTblName) & "') ORDER BY colid;"
		getSimpleTblCol = getArrFromSQL(myStrSQL)
	End Function

	'### getTblRowCount : returns row count
	Function getTblRowCount(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT rows FROM sysindexes WHERE id = OBJECT_ID('" & remquote(pTblName) & "') AND indid < 2;"
		getTblRowCount = getArrFromSQL(myStrSQL)(0, 0)
	End Function

	' ### getTblIx : returns table index list
	Function getTblIx(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT i.indid, i.name, c.name, CONVERT(BIT, i.status & 2048) AS isPrimary, CONVERT(BIT, i.status & 2) AS isUnique, CONVERT(BIT, i.status & 1) AS IgnoreDuplicateKeys, CONVERT(BIT, i.status & 4096) AS isUniqueConstraint, CONVERT(BIT, i.status & 16) AS isClustered, CONVERT(BIT, i.status & 16777216) AS startnorecp, i.OrigFillFactor, g.groupname FROM sysindexes i RIGHT JOIN sysindexkeys k ON i.id = k.id AND i.indid = k.indid INNER JOIN syscolumns c ON k.colid = c.colid AND k.id = c.id INNER JOIN [" & remquote(pDBName) & "].dbo.sysfilegroups g ON i.groupid = g.groupid WHERE i.id = OBJECT_ID('" & remquote(pTblName) & "') AND i.indid > 0 AND i.indid < 255 AND i.status & 64 = 0 ORDER BY i.indid, k.keyno;"
		getTblIx = getArrFromSQL(myStrSQL)
	End Function

	' ### getTblPrimaryKey : returns primary key columns for the specified table
	Function getTblPrimaryKey(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT c.name, CASE WHEN CHARINDEX(TYPE_NAME(c.xtype), N'char, nchar, varchar, nvarchar') > 0 THEN 1 WHEN CHARINDEX(TYPE_NAME(c.xtype), N'text, ntext') > 0 THEN 2 WHEN CHARINDEX(TYPE_NAME(c.xtype), N'binary, varbinary, image,timestamp') > 0 THEN 3 WHEN CHARINDEX(TYPE_NAME(c.xtype), N'datetime,smalldatetime') > 0 THEN 4 WHEN CHARINDEX(TYPE_NAME(c.xtype), N'bit') > 0 THEN 5  WHEN CHARINDEX(TYPE_NAME(c.xtype), N'uniqueidentifier') > 0 THEN 6 ELSE 0 END AS displayType FROM sysindexes i RIGHT JOIN sysindexkeys k ON i.id = k.id AND i.indid = k.indid INNER JOIN syscolumns c ON k.colid = c.colid AND k.id = c.id WHERE i.id = OBJECT_ID('" & remquote(pTblName) & "') AND CONVERT(BIT, i.status & 2048) = 1 ORDER BY i.indid, k.keyno"
		getTblPrimaryKey = getArrFromSQL(myStrSQL)
	End Function

	' ### getTblForeignKey : returns foreign key list
	Function getTblForeignKey(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT constid, OBJECT_NAME(constid) AS 'RelationShip Name', OBJECT_NAME(rkeyid) AS 'Master Table', OBJECT_NAME(fkeyid) AS 'Foreign Table', COL_NAME(rkeyid, rkey) AS 'Primary Key Col', COL_NAME(fkeyid, fkey) AS 'Foreign Key Col', CASE WHEN fkeyid = OBJECT_ID('" & remquote(pTblName) & "') THEN 1 ELSE 0 END AS FFlag FROM sysforeignkeys WHERE rkeyid = OBJECT_ID('" & remquote(pTblName) & "') OR  fkeyid = OBJECT_ID('" & remquote(pTblName) & "') ORDER BY 2;"
		getTblForeignKey = getArrFromSQL(myStrSQL)
	End Function

	' ### getTblTrigger : returns trigger list
	Function getTblTrigger(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT name, USER_NAME(uid), crdate FROM sysobjects WHERE type = 'TR' AND parent_obj = OBJECT_ID('" & remquote(pTblName) & "');"
		getTblTrigger = getArrFromSQL(myStrSQL)
	End Function

	' ### getTblConstraint : returns constraint list
	Function getTblConstraint(pDbName, pTblName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT o.name, USER_NAME(o.uid), o.crdate, c.text FROM sysobjects o INNER JOIN syscomments c ON o.id = c.id WHERE type = 'C' AND parent_obj = OBJECT_ID('" & remquote(pTblName) & "');"
		getTblConstraint = getArrFromSQL(myStrSQL)
	End Function

	' ### getViewCol : returns column list
	Function getViewCol(pDbName, pViewName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT c.name, type_name(c.xusertype) AS type, c.length AS length, CASE WHEN CHARINDEX(TYPE_NAME(c.xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney') > 0 THEN COLUMNPROPERTY(c.id, c.name, 'precision') ELSE NULL END AS mPrecision, CASE WHEN CHARINDEX(TYPE_NAME(c.xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney') > 0 THEN OdbcScale(c.xtype,c.xscale) ELSE NULL END AS mScale, CONVERT(BIT, c.isnullable), CONVERT(BIT, COLUMNPROPERTY(OBJECT_ID('" & remquote(pViewName) & "'), c.name, 'IsIdentity')) AS isIdentity, CASE WHEN COLUMNPROPERTY(OBJECT_ID('" & remquote(pViewName) & "'), c.name, 'IsIdentity') = 1 THEN IDENT_INCR('" & remquote(pViewName) & "') ELSE NULL END AS mStart, CASE WHEN COLUMNPROPERTY(OBJECT_ID('" & remquote(pViewName) & "'), c.name, 'IsIdentity') = 1 THEN IDENT_SEED('" & remquote(pViewName) & "') ELSE NULL END AS mSeed, CONVERT(BIT, COLUMNPROPERTY(OBJECT_ID('" & remquote(pViewName) & "'), c.name, 'IsRowGuidCol')) AS isRowGuidCol, CASE WHEN c.colid IN (SELECT k.colid FROM sysindexes i RIGHT JOIN sysindexkeys k ON i.id = k.id AND i.indid = k.indid WHERE status & 2048 = 2048 AND k.id = c.id) THEN 1 ELSE 0 END FROM syscolumns c LEFT JOIN (sysobjects o INNER JOIN syscomments co ON o.id = co.id) ON c.cdefault = o.id LEFT JOIN (sysobjects o2 INNER JOIN syscomments co2 ON o2.id = co2.id) ON c.domain = o2.id WHERE c.id = OBJECT_ID('" & remquote(pViewName) & "') ORDER BY c.colid;"
		getViewCol = getArrFromSQL(myStrSQL)
	End Function

	' ### getSimpleViewCol : returns column name list and misc. prop. for display
	Function getSimpleViewCol(pDbName, pViewName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT name, CASE WHEN CHARINDEX(TYPE_NAME(xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney, bit') > 0 THEN 'RIGHT' WHEN CHARINDEX(TYPE_NAME(xtype), N'datetime,smalldatetime,timestamp') > 0 THEN 'CENTER' ELSE 'LEFT' END AS mAlign, CASE WHEN CHARINDEX(TYPE_NAME(xtype), N'text, ntext, binary, varbinary, image, bit') > 0 THEN CONVERT(BIT, 0) ELSE CONVERT(BIT, 1) END AS isSortable, CASE WHEN CHARINDEX(TYPE_NAME(xtype), N'char, nchar, varchar, nvarchar') > 0 THEN 1 WHEN CHARINDEX(TYPE_NAME(xtype), N'text, ntext') > 0 THEN 2 WHEN CHARINDEX(TYPE_NAME(xtype), N'binary, varbinary, image') > 0 THEN 3 ELSE 0 END AS displayType FROM syscolumns WHERE id = OBJECT_ID('" & remquote(pViewName) & "') ORDER BY colid;"
		getSimpleViewCol = getArrFromSQL(myStrSQL)
	End Function

	'### getViewRowCount : returns row count
	Function getViewRowCount(pDbName, pViewName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT COUNT(*) FROM " & pViewName & ";"
		getViewRowCount = getArrFromSQL(myStrSQL)(0, 0)
	End Function

	' ### getSpParam : returns parameter list
	Function getSpParam(pDbName, pSpName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT c.name, type_name(c.xusertype) AS type, c.length AS length, CASE WHEN CHARINDEX(TYPE_NAME(c.xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney') > 0 THEN COLUMNPROPERTY(c.id, c.name, 'precision') ELSE NULL END AS mPrecision, CASE WHEN CHARINDEX(TYPE_NAME(c.xtype), N'tinyint,smallint,decimal,int,real,money,float,numeric,smallmoney') > 0 THEN OdbcScale(c.xtype,c.xscale) ELSE NULL END AS mScale, CONVERT(BIT, c.isnullable), CONVERT(BIT, c.isoutparam) FROM syscolumns c LEFT JOIN (sysobjects o INNER JOIN syscomments co ON o.id = co.id) ON c.cdefault = o.id WHERE c.id = OBJECT_ID('" & pSpName & "') ORDER BY c.colid;"
		getSpParam = getArrFromSQL(myStrSQL)
	End Function

	' ### getObjTxt : returns object text
	Function getObjTxt(pDbName, pObjName)
		Dim myStrSQL, myArr, myRC, i, myTxt
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT c.text FROM syscomments c WHERE c.id = OBJECT_ID('" & (remquote(pObjName)) & "');"
		myArr = getArrFromSQL(myStrSQL)
		If isArray(myArr) Then myRC = UBound(myArr, 2) Else myRC = -1 End If
		myTxt = ""
		For i = 0 To myRC
			myTxt = myTxt & myArr(0, i)
		Next
		getObjTxt = myTxt
	End Function

	' ### getTypeList : returns allowed type list
	Function getTypeList(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT name, length, prec, scale, allownulls, tdefault, domain, variable FROM systypes WHERE name <> 'sysname' ORDER BY name;"
		getTypeList = getArrFromSQL(myStrSQL)
	End Function

	' ### getSimpleTypeList : returns allowed type list (but no udt)
	Function getSimpleTypeList(pDbName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT name, length, allownulls FROM systypes WHERE xusertype = xtype ORDER BY name;"
		getSimpleTypeList = getArrFromSQL(myStrSQL)
	End Function

	' ### returns object shortname
	Function getObjShortName(pDbName, pObjName)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT OBJECT_NAME(OBJECT_ID('" & remquote(pObjName) & "'))"
		getObjShortName = getArrFromSQL(myStrSQL)(0, 0)
	End Function

	' ### getSQLRecordCount : returns record count from a request
	Function getSQLRecordCount(pDbName, pObjName, pWhere)
		Dim myStrSQL
		gObjDC.execute "USE [" & rembracket(pDbName) & "];"
		myStrSQL = "SELECT COUNT(*) AS rc FROM " & remquote(pObjName)
		If pWhere <> "" Then myStrSQL = myStrSQL & " WHERE " & pWhere
		getSQLRecordCount = getArrFromSQL(myStrSQL)(0, 0)
	End Function

	' ### getSQLstr : returns an array of sql to get table rows _ 1st item is the query with a TOP statement to get a faster display, the 2nd item is the complete query
	Function getSQLstr(pObjName, pField, pWhere, pOrderColName, pOrderWay, pPageSize, pPageNumber)
		Dim myStrSQL, myStrSQL2
		myStrSQL = "SELECT"
		myStrSQL2 = "SELECT"
		If pPageNumber > 0 Then myStrSQL = myStrSQL & " TOP " & (pPageSize * pPageNumber)
		If pField = "" Then
			myStrSQL = myStrSQL & " *"
			myStrSQL2 = myStrSQL2 & " *"
		Else
			myStrSQL = myStrSQL & " " & pField
			myStrSQL2 = myStrSQL2 & " " & pField
		End If
		myStrSQL = myStrSQL & " FROM " & remquote(pObjName)
		myStrSQL2 = myStrSQL2 & " FROM " & remquote(pObjName)
		If pWhere <> "" Then
			myStrSQL = myStrSQL & " WHERE " & pWhere
			myStrSQL2 = myStrSQL2 & " WHERE " & pWhere
		End If
		If pOrderColName <> "" Then 
			myStrSQL = myStrSQL & " ORDER BY [" & remquote(rembracket(pOrderColName)) & "]"
			myStrSQL2 = myStrSQL2 & " ORDER BY [" & remquote(rembracket(pOrderColName)) & "]"
			If pOrderWay <> "" Then 
				myStrSQL = myStrSQL & " " & pOrderWay
				myStrSQL2 = myStrSQL2 & " " & pOrderWay
			End If
		End If
		getSQLstr = Array(myStrSQL, myStrSQL2)
	End Function

	' ### No Cache Method
	Sub setNoCacheHeader
		Response.Expires = 0
		Response.ExpiresAbsolute = #September 23, 1991 21:00:00#
		Response.AddHeader "Last-Modified ", CStr(now())
		Response.AddHeader "Cache-Control", "no-cache, must-revalidate"
		Response.AddHeader "Pragma", "no-cache"
	End Sub

	' ### Authorization Function
	Function mla_auth(pObjectIndex, pActionIndex)
		mla_auth = True
	End Function
%>
