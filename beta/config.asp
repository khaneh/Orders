<%
' Remember to Check "Enable Parent Paths" in:
'		IIS > Home Directory > Configuration > Options

Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
orderFolder="C://myfile/"
'---------------------------------------------
'------------------------------- Check Session 
'---------------------------------------------
if (session("ID")="") then
	response.cookies("OldURL") = Request.ServerVariables("URL")
	'response.cookies("OldForm") = request.form
	session.abandon
	response.redirect "/login.asp?err="& server.URLEncode("»—«Ì œÌœ‰ «Ì‰ ’›ÕÂ »«Ìœ Ê«—œ ”Ì” „ ‘ÊÌœ")
end if

'---------------------------------------------
'------------------------------- DB Connection
'---------------------------------------------

'***** NOTE: If conStr changed, you should change it in Home/errHandler.asp too
'conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=jame;Data Source=.\sqlexpress;PWD=5tgb;"
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 180
conn.open conStr
conn.CommandTimeout = 120 

'---------------------------------------------
'---------------------------- Check GL Routine
'---------------------------------------------
dim OpenGL

if Session("OpenGL")="" then
	response.write "Œÿ«! ”«· „«·Ì „‘Œ’ ‰Ì” ."
	response.end

else
	OpenGL = Session("OpenGL")
	OpenGLName = Session("OpenGLName")
	FiscalYear = Session("FiscalYear")
end if


'---------------------------------------------
'--------------------------- Check Permissions
'---------------------------------------------
function Auth(menuID, subMenuID)
	pr = session("Permission")

	if (pr="") then
		response.cookies("OldURL") = Request.ServerVariables("URL")
		session.abandon
		response.redirect "/login.asp?err="& server.URLEncode("»—«Ì œÌœ‰ «Ì‰ ’›ÕÂ »«Ìœ Ê«—œ ”Ì” „ ‘ÊÌœ")
	end if

	st = inStr(pr,"#"&menuID)
	
	if subMenuID >= "A" then 
		subMenuID_int = cint(asc(subMenuID)-55)
	else
		subMenuID_int = cint(subMenuID)
	end if

	Auth = false
	if st > 0 then
		en = inStr(st+1, pr, "#")
		sm = inStr(st+2, pr, subMenuID)
		if subMenuID_int = 0 or ((sm <> 0) and (en > sm or en = 0 )) then
			Auth = true
		end if		
	end if
end function

'---------------------------------------------
'--------------------------------- Access Deny
'---------------------------------------------
function NotAllowdToViewThisPage()
	response.redirect "top.asp?errmsg=" & Server.URLEncode("‘„« „Ã«“ »Â œÌœ‰ «Ì‰ ’›ÕÂ ‰Ì” Ìœ.")
end function

'---------------------------------------------
'-------------------------------------- Darsad
'---------------------------------------------
function Pourcent(input1,input2)
	if input2=0 then
		result="-"
	else
		result=round(input1/input2 * 100)
	end if
	Pourcent=result
end function

'---------------------------------------------
'------------------------------------ SQL Safe
'---------------------------------------------
function sqlSafe (inpStr)
  tmpStr=inpStr
  tmpStr=replace(tmpStr,"'","`")
  tmpStr=replace(tmpStr,"""","`")
  sqlsafe=tmpStr
end function


'---------------------------------------------
'---------------------------------- text2value
'---------------------------------------------
function text2value(input)
	result=replace(input,",","")
	if result="" then
		text2value=0
	else
		text2value=replace(input,",","")
	end if
	if input="True" then text2value=1
	if input="False" then text2value=0
end function


const CONST_MSG_NONE	= -1
const CONST_MSG_INFORM	= 0
const CONST_MSG_ALERT	= 1
const CONST_MSG_ERROR	= 2

'---------------------------------------------
'-------------------------- SHOW ALERT MESSAGE
'---------------------------------------------
sub showAlert( alertMsg , msgStatus )
	select case msgStatus
	case -1:
		response.write "<TABLE width=45% align='center' bgcolor=ffffff style='border: dashed 1px Black'>"
	case 0:
		response.write "<TABLE width=45% align='center' bgcolor=ffffff style='border: dashed 1px Blue'>"
	case 1:
		response.write "<TABLE width=45% align='center' bgcolor=ffffdd style='border: dashed 1px Red'>"
	case 2:
		response.write "<TABLE width=45% align='center' bgcolor=ffeeee style='border: dashed 1px Red'>"
	case else:
		response.write "<TABLE width=45% align='center' bgcolor=ffffff style='border: dashed 1px Black'>"
	end select
	response.write "<TR>"
	response.write "<td align=center width='10%'>"
	response.write "<IMG SRC='../images/alertIcon"& trim(msgStatus) & ".gif' >"
	response.write "</td>"
	response.write "<TD align=center><BR>"
	response.write alertMsg
	response.write "<BR><BR></TD>"
	response.write "</TR>"
	response.write "</TABLE>"
end sub

'---------------------------------------------
'-------------- Un-Approve an invoice when its 
'-------------- related order has been changed
'---------------------------------------------
sub UnApproveInvoice(invoiceID,ApprovedBy)
	MsgTo			=	ApprovedBy
	msgTitle		=	"Invoice Unapproved"
	msgBody			=	"›«? Ê— ›Êﬁ »Â œ·Ì·  €ÌÌ— ”›«—‘ „—»ÊÿÂ Ì« »Â œ·Ì· ÊÌ—«Ì‘°  Ê”ÿ "& session("CSRName") & " «“   «ÌÌœ Œ«—Ã ‘œ."
	RelatedTable	=	"invoices"
	relatedID		=	invoiceID
	replyTo			=	0
	IsReply			=	0
	urgent			=	1
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("UPDATE Invoices SET Approved=0 WHERE (ID='"& invoiceID & "')")
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")
end sub

'---------------------------------------------
'-------------- Inform CSR that Order Is Ready
'---------------------------------------------
sub InformCSRorderIsReady(orderID,CSR)
	MsgTo			=	CSR
	msgTitle		=	"Order Is Ready"
	msgBody			=	"”›«—‘ ¬„«œÂ  ÕÊÌ· «” ."
	RelatedTable	=	"orders"
	relatedID		=	orderID
	replyTo			=	0
	IsReply			=	0
	urgent			=	3
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")
end sub

'---------------------------------------------
'----------------- Inventory Item Request Deny
'---------------------------------------------
sub InformRequestDenied(itemName,CSR)
	MsgTo			=	CSR
	msgTitle		=	"Request Denied"
	msgBody			=	"œ—ŒÊ«”  ‘„« »—«Ì "& itemName & "  —œ „Ì ‘Êœ. "
	RelatedTable	=	"NaN"
	relatedID		=	0
	replyTo			=	0
	IsReply			=	0
	urgent			=	0
	MsgFrom			=	session("ID")
	MsgDate			=	shamsiToday()
	MsgTime			=	currentTime10()
	Conn.Execute ("INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable, urgent) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "', "& urgent & ")")
end sub

'---------------------------------------------
'------------------------------------ Separate
'---------------------------------------------
function Separate(inputTxt)

if (not isnumeric(trim(inputTxt))) or ("" & inputTxt="") then 
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


'------------------------------------------------------------START
' Convert Digits to Letter ( e.g. 1 --> Yek )
' Ver 2.0  -    1381-11-7    - Alix
'---------------------------------------FUNC_1
Function ConvertIT(inputDigits)
digits=inputDigits
if not IsNumeric(digits) then
	ConvertIT=digits
else
	dim radeh(11)
	radeh(0)=""				'non
	radeh(1)="Â“«—"			'hezar
	radeh(2)="„Ì·ÌÊ‰"		'milion
	radeh(3)="„Ì·Ì«—œ"		'miliard
	radeh(4)="»Ì·ÌÊ‰"		'bilion 
	radeh(5)=" —Ì·ÌÊ‰"		'trilion 
	radeh(6)="ﬂÊ«—œÌ·ÌÊ‰"	'quadrilion 
	radeh(7)="ﬂÊÌ‰ Ì·ÌÊ‰"	'quintilion 
	radeh(8)="”Ìﬂ” Ì·ÌÊ‰"	'sextilion 
	radeh(9)="”Å Ì·ÌÊ‰"		'septilion
	radeh(10)="«Êﬂ Ì·ÌÊ‰"	'oktilian 
	
	dim nroundn, roundn
	digits = replace(replace(digits,chr(13),""),chr(10),"")
	nroundn=(len(digits))/3
	roundn=fix(nroundn)
	if not roundn=nroundn then
		roundn = roundn + 1
	end if
	ConvertIT=""
	for i=0 to roundn-1
		pr = trim(Convert3dig(right(digits,3)))
			if trim(pr)<>"" then
				pr = pr &" "& radeh(i) 
				if i<>roundn-1 then
				pr = " Ê " & pr 
				end if
			else
				pr = " " & pr &" "
			end if
		ConvertIT= trim(pr) & " " & trim(ConvertIT)
		t=len(digits)-3
		if t<=0 then
			t = len(digits)
		end if
		digits = left(digits, t)
	next
	ConvertIT = trim(ConvertIT)
end if
End Function
'---------------------------------------FUNC_2
Function Convert3dig(digits3)
	dim yekan(20)
	yekan(0)=""			'0
	yekan(1)="Ìﬂ"		'1
	yekan(2)="œÊ"		'2
	yekan(3)="”Â"		'3
	yekan(4)="çÂ«—"		'4
	yekan(5)="Å‰Ã"		'5
	yekan(6)="‘‘"		'6
	yekan(7)="Â› "		'7
	yekan(8)="Â‘ "		'8
	yekan(9)="‰Â"		'9
	yekan(10)="œÂ"		'10
	yekan(11)="Ì«“œÂ"	'11
	yekan(12)="œÊ«“œÂ"	'12
	yekan(13)="”Ì“œÂ"	'13
	yekan(14)="çÂ«—œÂ"	'14
	yekan(15)="Å«‰“œÂ"	'15
	yekan(16)="‘«‰“œÂ"	'16
	yekan(17)="Â›œÂ"	'17
	yekan(18)="ÂÃœÂ"	'18
	yekan(19)="‰Ê“œÂ"	'19

	dim dahgan(10)
	dahgan(0)=""		'0
	dahgan(1)="œÂ"		'10
	dahgan(2)="»Ì” "	'20
	dahgan(3)="”Ì"		'30
	dahgan(4)="çÂ·"		'40
	dahgan(5)="Å‰Ã«Â"	'50
	dahgan(6)="‘’ "		'60
	dahgan(7)="Â› «œ"	'70
	dahgan(8)="Â‘ «œ"	'80
	dahgan(9)="‰Êœ"		'90

	dim sadgan(10)
	sadgan(0)=""		'0
	sadgan(1)="’œ"		'100
	sadgan(2)="œÊÌ” "	'200
	sadgan(3)="”Ì’œ"	'300
	sadgan(4)="çÂ«—’œ"	'400
	sadgan(5)="Å«‰’œ"	'500
	sadgan(6)="‘‘’œ"	'600
	sadgan(7)="Â› ’œ"	'700
	sadgan(8)="Â‘ ’œ"	'800
	sadgan(9)="‰Â’œ"	'900

	n=len(digits3)

	Convert3dig=""
	a1 = digits3/100
	if a1>0 then 
		Convert3dig=Convert3dig & sadgan(int(a1)) & " "
	end if

	digits3_2=digits3-int(a1)*100
	if not (digits3_2=0 or a1<1) then 
		Convert3dig=Convert3dig &  " Ê "
	end if
	
	if digits3_2>19 then
		a2 = digits3_2/10
		if a2>0 then 
			Convert3dig=Convert3dig & dahgan(int(a2)) & " "
		end if

		digits3_3=digits3_2-int(a2)*10
		if not (digits3_3=0 or a2<1) then 
			Convert3dig=Convert3dig &  " Ê "
		end if
		Convert3dig=Convert3dig & yekan(int(digits3_3)) & " "
	else
		a2 = digits3_2
		if a2>0 then 
			Convert3dig=Convert3dig & yekan(int(a2)) & " "
		end if
	end if
	
End Function

function MapURL(path)

     dim rootPath, url

     'Convert a physical file path to a URL for hypertext links.

     rootPath = Server.MapPath("/")
     url = Right(path, Len(path) - Len(rootPath)+2)
     MapURL = Replace(url, "\", "/")

end function 

sub ListFolderContents(path)

 dim fs, folder, file, item, url

 set fs = CreateObject("Scripting.FileSystemObject")
 set folder = fs.GetFolder(path)

'Display the target folder and info.

 Response.Write("<li><b>" & folder.Name & "</b> - " _
   & folder.Files.Count & " files, ")
 if folder.SubFolders.Count > 0 then
   Response.Write(folder.SubFolders.Count & " directories, ")
 end if
 Response.Write(Round(folder.Size / 1024) & " KB total." _
   & "</li>" & vbCrLf)

 Response.Write("<ul style='margin:0px 0px 0px 10px;'>" & vbCrLf)

 'Display a list of sub folders.

 for each item in folder.SubFolders
   ListFolderContents(item.Path)
 next

 'Display a list of files.

 for each item in folder.Files
   url = MapURL(item.path)
   Response.Write("<li><a href=""myfile/" & url & """>" _
     & item.Name & "</a> - " _
     & item.Size & " bytes, " _
     & "last modified on " & item.DateLastModified & "." _
     & "</li>" & vbCrLf)
 next

 Response.Write("</ul>" & vbCrLf)

end sub   
'------------------------------------------------------------ END


'<SCRIPT LANGUAGE="JavaScript" src="/beta/jquery-1.5.2.min.js"></SCRIPT>
%>
