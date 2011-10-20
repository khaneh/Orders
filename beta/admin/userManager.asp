<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="ãÏíÑíÊ ßÇÑÈÑÇä"
SubmenuItem=1
%>
<!--#include file="top.asp" -->
<%
NumberOfCategories = 14
dim pagesPermissions(20,25)
dim pages(20,25)
'-------------------------------------
 pages ( 0 , 0 ) = "ÕİÍå Çæá"         
	pages ( 0 , 1 ) = "ÎæÇäÏä íÇã"   
	pages ( 0 , 2 ) = "äæÔÊä íÇã"    
	pages ( 0 , 3 ) = "ÏíÏä ÏİÊÑ Êáİä"     
	pages ( 0 , 4 ) = "ÏÑÎæÇÓÊ ãÑÎÕí" 
	pages ( 0 , 5 ) = "ÏÑÎæÇÓÊ ßÇáÇ"  
	pages ( 0 , 6 ) = "æíÑÇíÔ ÏİÊÑ Êáİä"     
	pages ( 0 , 7 ) = "ÇÑÓÇá íÇã Èå åãå"     
	pages ( 0 , 8 ) = "ßÊÇÈÎÇäå"
'-------------------------------------
 pages ( 1 , 0 ) = "İåÑÓÊ äÇãåÇ"
	pages ( 1 , 1 ) = "äãÇíÔ ÇØáÇÚÇÊ"
	pages ( 1 , 2 ) = "ÌÏíÏ"
	pages ( 1 , 3 ) = "ÇÕáÇÍ"
	pages ( 1 , 4 ) = "ÊÛííÑ ãÓÆæá æ ÓŞİ ÇÚÊÈÇÑ"
	pages ( 1 , 5 ) = "ÊÛííÑ æÖÚíÊ ÍÓÇÈ"
	pages (1, 6)	= "ÒÇÑÔåÇ"
	pages (1, 7)	= "ÇÕáÇÍ ÚäæÇä ÍÓÇÈ"
	'pages (1, 8)	= "ÊÛííÑ íÔ İÑÖ Çáİ"
'-------------------------------------
 pages ( 2 , 0 ) = "ÓİÇÑÔÇÊ"
	pages ( 2 , 1 ) = "æÑæÏ ÓİÇÑÔ"
	pages ( 2 , 2 ) = "ÇÕáÇÍ ÓİÇÑÔ"
	pages ( 2 , 3 ) = "ííÑí ÓİÇÑÔ"
	pages ( 2 , 4 ) = "<s>ÊÎãíä ŞíãÊ ÏíÌíÊÇá</s> deprecated"
	pages ( 2 , 5 ) = "ÌÓÊÌæí  íÔÑİÊå"
	pages ( 2 , 6 ) = "ÒÇÑÔ ßáí"
	pages ( 2 , 7 ) = "ÒÇÑÔ ãÏíÑíÊ"
	pages ( 2 , 8 ) = "ÒÇÑÔ ãÏíÑíÊ - ŞÇÈáíÊ ÊÛííÑ ÔÑÇíØ"
	pages ( 2 , 9 ) = "ÇÓÊÚáÇã"
'-------------------------------------
 pages ( 3 , 0 ) = "ÊæáíÏ"
	pages ( 3 , 1 ) = "æÑæÏ ãÑÇÍá ÊæáíÏ"
	pages ( 3 , 2 ) = "ÈÑäÇãå ÑíÒí ÊæáíÏ"
	pages ( 3 , 3 ) = "ííÑí ÓİÇÑÔ"
	pages ( 3 , 4 ) = "ÌÓÊÌæí  íÔÑİÊå"
	pages ( 3 , 5 ) = "ÇíÌÇÏ ÏÑÎæÇÓÊ ßÇáÇ íÇ ÓÑæíÓ ÈÑÇí ÓİÇÑÔí ßå İÇßÊæÑ Âä ŞÈáÇ ÕÇÏÑ ÔÏå ÇÓÊ"
	pages ( 3 , 6 ) = "ÌÏæá ÊæáíÏ"
'-------------------------------------
 pages ( 4 , 0 ) = "ÎÑíÏ"
	pages ( 4 , 1 ) = "ÏÑÎæÇÓÊ ÎÑíÏ ßÇáÇí ÇäÈÇÑ"
	pages ( 4 , 2 ) = "ÏÑÎæÇÓÊ ÎÑíÏ ÓÇíÑßÇáÇ åÇ æ ÎÏãÇÊ"
	pages ( 4 , 3 ) = "ÓİÇÑÔ ÎÑíÏ"
	pages ( 4 , 4 ) = "ííÑí ÎÑíÏ"
	pages ( 4 , 5 ) = "ÏæÎÊä"
	pages ( 4 , 6 ) = "ÊÛííÑ ÓİÇÑÔ ãÑÈæØå"
'-------------------------------------
 pages ( 5 , 0 ) = "ÇäÈÇÑ"
	pages ( 5 , 1 ) = "ÕÏæÑ ÍæÇáå ÇäÈÇÑ ÛíÑ ãÑÈæØ Èå ÓİÇÑÔ åÇ æ íÑÇãæä"
	pages ( 5 , 2 ) = "ÎÑæÌ ßÇáÇ"
	pages ( 5 , 3 ) = "æÑæÏ ßÇáÇ"
	pages ( 5 , 4 ) = "ÊÚÑíİ ßÇáÇí ÌÏíÏ"
	pages ( 5 , 5 ) = "ÒÇÑÔ"
	pages ( 5 , 6 ) = "ÇÕáÇÍ ßÇáÇ"
	pages ( 5 , 7 ) = "ÊäÙíã ÏÓÊå Çí ãæÌæÏí "
	pages ( 5 , 8 ) = "ÒÇÑÔ ãæÌæÏí"
	pages ( 5 , 9 ) = "ÕÏæÑ ÍæÇáå ÇäÈÇÑ ãÑÈæØ Èå ÓİÇÑÔ åÇ"
	pages ( 5 , 10 ) = "ËÈÊ ÏÑÎæÇÓÊ ßÇáÇ"					'A
	pages ( 5 , 11 ) = "ÇäÊŞÇá ãÇáßíÊ"						'B
	pages ( 5 , 12 ) = "ËÈÊ æÑæÏ/ÎÑæÌ ÏÑ ÊÇÑíÎ ÏáÎæÇå"		'C
	pages ( 5 , 13 ) = "ËÈÊ ŞíãÊ ßÇáÇ ÏÑ íß ÈÇÒå ÒãÇäí"		'D
'-------------------------------------
 pages ( 6 , 0 ) = "ÍÓÇÈÏÇÑí İÑæÔ"
	pages ( 6 , 1 ) = "æÑæÏ İÇßÊæÑ "
	pages ( 6 , 2 ) = "ÇÚáÇãíå ÈÏåßÇÑ/ÈÓÊÇäßÇÑ"
	pages ( 6 , 3 ) = "ÇÕáÇÍ İÇßÊæÑ"
	pages ( 6 , 4 ) = "æÑæÏ İÇßÊæÑ ÈÑÔÊ"
	pages ( 6 , 5 ) = "ÇÕáÇÍ İÇßÊæÑ ÈÑÔÊ"
	pages ( 6 , 6 ) = "ÒÇÑÔ ÍÓÇÈ"
	pages ( 6 , 7 ) = "ÏæÎÊä"
	pages ( 6 , 8 ) = "ÒÇÑÔ İÑæÔ ÎæÏÔ"
	pages ( 6 , 9 ) = "ÒÇÑÔ İÑæÔ åãå"
	pages ( 6 , 10 ) = "ÊÛííÑÇÊ ÌÑíí İÇßÊæÑ ÈÚÏ ÇÒ ÕÏæÑ"	'A
	pages ( 6 , 11 ) = "<s>ÏíÏä ÓÇíÑ ÒÇÑÔåÇ</s> unused"	'B
	pages ( 6 , 12 ) = "ÊÇííÏ İÇßÊæÑ"						'C
	pages ( 6 , 13 ) = "ÕÏæÑ İÇßÊæÑ"						'D
	pages ( 6 , 14 ) = "Ç İÇßÊæÑ"							'E
	pages ( 6 , 15 ) = "ÇÈØÇá İÇßÊæÑ"						'F
	pages ( 6 , 16 ) = "ÍĞİ íÔ äæíÓ"						'G
	pages ( 6 , 17 ) = "ÇÈØÇá ÇÚáÇãíå İÑæÔ"					'H
	pages ( 6 , 18 ) = "ÕÏæÑ İÇßÊæÑ ÏÑ ÊÇÑíÎ ÏáÎæÇå"		'I
	pages ( 6 , 19 ) = "Ç ÒÇÑÔ ÍÓÇÈ"						'J
	pages ( 6 , 20 ) = "æÑæÏ (İÇßÊæÑ / ÈÑÔÊ ) ÈÏæä ÓİÇÑÔ"	'K
	pages ( 6 , 21 ) = "ßí íÔ äæíÓ"						'L
	pages ( 6 , 22 ) = "ÇÈØÇá İÇßÊæÑí ßå åäæÒ ÓäÏ ÍÓÇÈÏÇÑí äÎæÑÏå"						'M
	pages (6, 23)	 = "ÊÛííÑÇÊ ßáí İÇßÊæÑ ÈÚÏ ÇÒ ÕÏæÑ"					'N
'-------------------------------------
 pages ( 7 , 0 ) = "ÍÓÇÈÏÇÑí ÎÑíÏ"
	pages ( 7 , 1 ) = "æÑæÏ İÇßÊæÑ ÎÑíÏ"
	pages ( 7 , 2 ) = "ÊÇííÏ İÇßÊæÑ ÎÑíÏ"
	pages ( 7 , 3 ) = "ÒÇÑÔ"
	pages ( 7 , 4 ) = "ÇÚáÇãíå ÈÏåßÇÑ/ÈÓÊÇäßÇÑ"
	pages ( 7 , 5 ) = "ÒÇÑÔ ÍÓÇÈ"
	pages ( 7 , 6 ) = "ÏæÎÊä"
	pages ( 7 , 7 ) = "ÒÇÑÔ ÎÑíÏ"
	pages ( 7 , 8 ) = "ÇÈØÇá ÇÚáÇãíå ÎÑíÏ"
	pages ( 7 , 9 ) = "ÇÈØÇá İÇßÊæÑ ÎÑíÏ"
	pages ( 7 , 10 ) = "Ç ÒÇÑÔ ÍÓÇÈ"						'A

'-------------------------------------
 pages ( 8 , 0 ) = "ÍÓÇÈÏÇÑí "
	pages ( 8 , 1 ) = "æÑæÏ  ÓäÏ"
	pages ( 8 , 2 ) = "ÏíÏä ÓäÏ åÇí ÍĞİ ÔÏå"
	pages ( 8 , 3 ) = "[ÕİÍå] ÂÑÔíæ ÇÓäÇÏ"
	pages ( 8 , 4 ) = "ÏİÊÑ ßá"
	pages ( 8 , 5 ) = "ÏíÏä ÓäÏ åÇí ŞØÚí ÔÏå"
	pages ( 8 , 6 ) = "ÏíÏä ÓäÏ åÇí ÈÑÑÓí ÔÏå æ ÊÛííÑ æÖÚíÊ ÂäåÇ"
	pages ( 8 , 7 ) = "ÏíÏä åãå ÓäÏ åÇí ãæŞÊ"
	pages ( 8 , 8 ) = "ÇÕáÇÍ íÇ ÊÛííÑ æÖÚíÊ åãå ÓäÏ åÇí ãæŞÊ"
	pages ( 8 , 9 ) = "ÏíÏä åãå íÇÏÏÇÔÊ åÇ"
	pages ( 8 , 10 ) = "ÒÇÑÔ ãÚíä"							'A
	pages ( 8 , 11 ) = "ÒÇÑÔ ÊİÕíáí"						'B
	pages ( 8 , 12 ) = "ÓäÏåÇí ÒíÑ ÓíÓÊã"					'C
	pages ( 8 , 13 ) = "ÊÛííÑ ÏİÊÑ ßá íÔ İÑÖ ÓíÓÊã"		'D
	pages ( 8 , 14 ) = "ÏíÏä ÓÇíÑ ÒÇÑÔ åÇ"					'E
	pages ( 8 , 15 ) = "ÓäÏ ãÑßÈ (æÑæÏ æ æíÑÇíÔ) "			'F
	pages ( 8 , 16 ) = "ÒÇÑÔ ÏÔÈæÑÏ ÎÒÇäåÏÇÑí"			'G
	'pages ( 8 , 17 ) = "íÇÏÏÇÔÊåÇ"						'H
'-------------------------------------
 pages ( 9 , 0 ) = "ÕäÏæŞ"
	pages ( 9 , 1 ) = "ÏÑíÇİÊ"
	pages ( 9 , 2 ) = "ÑÏÇÎÊ äŞÏ"
	pages ( 9 , 3 ) = "ÒÇÑÔ ÕäÏæŞ"
	pages ( 9 , 4 ) = "ÈÓÊä"
	pages ( 9 , 5 ) = "ÇíÌÇÏ"
	pages ( 9 , 6 ) = "ÒÇÑÔ ÓÑÑÓÊ ÕäÏæŞ"
	pages ( 9 , 9 ) = "ÒÇÑÔ ÓÑÑÓÊ ÕäÏæŞ (æíå)"
	pages ( 9 , 7 ) = "ÇÈØÇá Èí ŞíÏ æ ÔÑØ ÏÑíÇİÊ/ÑÏÇÎÊ"
	pages ( 9 , 8 ) = "Ç İÑã İÇßÊæÑ"
'-------------------------------------
 pages ( 10 , 0 ) = "ÈÇäß"									'A
	pages ( 10 , 1 ) = "ÕÏæÑ ß"
	pages ( 10 , 2 ) = "ííÑí ß ÏÑíÇİÊí"
	pages ( 10 , 3 ) = "ííÑí ß ÑÏÇÎÊí"
	pages ( 10 , 4 ) = "ËÈÊ ÍæÇáå"
	pages ( 10 , 5 ) = "ÑÏÇÎÊ İÇßÊæÑ"
	pages ( 10 , 6 ) = "ÏÑíÇİÊ ß ãÚíä"
	pages ( 10 , 7 ) = "ÕÏæÑ ß ãÚíä"
	pages ( 10 , 8 ) = "ÏİÊÑ ß"
	pages ( 10 , 9 ) = "ãÛÇíÑÊ"
'-------------------------------------
 pages ( 11 , 0 ) = "ÍÓÇÈÏÇÑí ÓÇíÑ"							'B
	pages ( 11 , 1 ) = "ÒÇÑÔ ÍÓÇÈ"
	pages ( 11 , 2 ) = "ÓÇíÑ ÇÚáÇãíå åÇ"
	pages ( 11 , 3 ) = "ÇäÊŞÇá"
	pages ( 11 , 4 ) = "ÏæÎÊä"
	pages ( 11 , 5 ) = "ÒÇÑÔ ÓÇíÑ"
	pages ( 11 , 6 ) = "ÇÈØÇá ÇÚáÇãíå ÓÇíÑ"
	pages ( 11 , 7 ) = "Ç ÒÇÑÔ ÍÓÇÈ"						
	pages ( 11 , 8 ) = "ÍŞæŞ"						

'-------------------------------------
 pages ( 12 , 0 ) = "ÓÇíÑ ÒÇÑÔ åÇí ÈÎÔ ÍÓÇÈÏÇÑí İÑæÔ"		'C
	pages ( 12 , 1 ) = "ÒÇÑÔ İÑæÔ ÑæÒÇäå"
	pages ( 12 , 2 ) = "ÒÇÑÔ İÑæÔ ÑæÒÇäå ÈÇ ÊÇÑíÎ ÏáÎæÇå"
	pages ( 12 , 3 ) = "ÒÇÑÔ ÂíÊã åÇí İÑæÔ "
	pages ( 12 , 4 ) = "ÒÇÑÔ İÇßÊæÑ åÇí Çáİ"
	pages ( 12 , 5 ) = "ÒÇÑÔ ÑÓíÏ åÇí Çáİ"
	pages ( 12 , 6 ) = "ÒÇÑÔ ÂíÊã åÇí İÑæÔ ÈÑÍÓÈ İÑæÔäÏå"
	pages ( 12 , 7 ) = "ÒÇÑÔ ŞíãÊ ÊãÇã ÔÏå"
	pages ( 12 , 8 ) = "ÒÇÑÔ ÊÑÇÒ åÇÑ ÓÊæäí"
	pages ( 12 , 9 ) = "ÒÇÑÔ ãÔÊÑíÇäí ßå ÇÚÊÈÇÑ ÂäåÇ ÕİÑ äíÓÊ"

'-------------------------------------
 pages ( 13 , 0 ) = "ãíÒ ßÇÑ"								'D
	pages ( 13 , 1 ) = "ÓİÇÑÔ åÇí ÈÇÒ"
	pages ( 13 , 2 ) = "İÇßÊæÑ åÇí ÊÇííÏ äÔÏå"
	pages ( 13 , 3 ) = "İÇßÊæÑ åÇí ÊÇííÏ ÔÏå (ÕÇÏÑ äÔÏå)"
	pages ( 13 , 4 ) = "İÇßÊæÑ åÇí ÕÇÏÑ ÔÏå (ÊÓæíå äÔÏå)"
	pages ( 13 , 5 ) = "ÍÓÇÈ åÇí ãÇäÏå ÏÇÑ"
	pages ( 13 , 6 ) = "ß åÇí ÈÑÔÊí"
	pages ( 13 , 7 ) = "ÒÇÑÔ ßáí ãÓææá ÍÓÇÈ"

'-------------------------------------
 pages ( 14 , 0 ) = "ãÏíÑíÊ"								'E
	pages ( 14 , 1 ) = "Reserved - æÑæÏ Èå ŞÓãÊ admin"
	pages ( 14 , 2 ) = "Reserved - ÊÛííÑ ãÌæÒ ßÇÑÈÑÇä"
	pages ( 14 , 3 ) = "Reserved"
	pages ( 14 , 4 ) = "Reserved"
	pages ( 14 , 5 ) = "ÏíÏä ãíÒ ßÇÑ ÓÇíÑ ßÇÑÈÑÇä"
	pages ( 14 , 6 ) = "ÏíÏä ÒÇÑÔåÇí ÇßÓáí"


function Auth(menuID, subMenuID, permission)
	pr = permission

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


%>
<style>
	.TABLE1 {font-family: tahoma; font-size: 8pt; Background-Color:navy; border:2 solid navy; cursor:pointer;}
	.TD_Rowspan {Background-Color:navy;}
	.textbox1 {border:none; width:100px; font-family: tahoma; font-size: 8pt; Background-Color:transparent; cursor:pointer;} 
	.TABLE1 TR {Background-Color:white; height:20px; }
	.TD10 {Background-Color:#CCCCCC;height:10px;}
	.TD00 {Background-Color:#BBBBFF;height:10px;}
	.TD11 {Background-Color:#FFFFFF;height:10px;}
	.TD01 {Background-Color:#DDDDFF;height:10px;}
</style>
<SCRIPT LANGUAGE="JavaScript">
<!--
function showUser(user){
	window.location="?act=edit&userID="+user;
}
//-->
</SCRIPT>
<%

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------ Users List
'-----------------------------------------------------------------------------------------------------
if request("act")="" then
	set RSc=Conn.Execute ("SELECT count(*) as c FROM Users WHERE ID<>0")
%>
	<BR>

	<TABLE class="table1" border=0 cellspacing=1 cellpadding=1 align=right>
	<TR>
		<TD rowspan="3" height=100>ßÇÑÈÑ</TD>
	</tr>
	<TR>
		<%
		for i=0 to NumberOfCategories
			j=0
			if not pagesPermissions(i,j)="ok" then 
				disableChecks="disabled"
				rowBGColor="#F0F0F0"
				groupChecked=""
			else
				disableChecks=""
				rowBGColor="#33AACC"
				groupChecked="checked"
			end if

			Do While pages(i,j+1)<>"" 
				j=j+1
			Loop

			%>
			<td colspan=<%=j%> align=center ><%=pages(i,0)%></td>
 			<td rowspan=<%=RSc("c")+3%> class="TD_Rowspan" ></td>
<%
		next
%>
	</tr>
	<TR>
<%
		for i=0 to NumberOfCategories
			j=0
			while pages(i,j+1)<>"" 
				j=j+1

				if j <10 then
					jj = j
				else
					jj = chr(55+j)
				end if
				
				%>
				<td title="<%=pages(i,j)%>"><%response.write jj'pages(i,j)%></td>
				<%
			wend
		next

%>
	</TR>
<% 
	set RSV=Conn.Execute ("SELECT * FROM Users WHERE ID<>0 ORDER BY RealName") 
	tmpRowCounter=0
	Do while not RSV.eof
		Permission = RSV("Permission")
		RealName = RSV("RealName")
		tmpRowCounter = tmpRowCounter + 1

%>
		<TR height='10px' onclick="showUser(<%=RSV("ID")%>);">
			<TD height='10px' class="TD<%=tmpRowCounter Mod 2%>1"><INPUT TYPE="text" class="textbox1" Value="<%=RealName%>"><br></TD>
<%
			tmpColCounter=0
			for i=0 to NumberOfCategories
				j=0
				while pages(i,j+1)<>"" 
					tmpColCounter = tmpColCounter + 1
					j=j+1

					if i <10 then
						ii = i
					else
						ii = chr(55+i)
					end if

					if j <10 then
						jj = j
					else
						jj = chr(55+j)
					end if
					
					%>
					<td height='10px' class="TD<%=tmpRowCounter Mod 2%><%=tmpColCounter Mod 2%>" title="<%=RealName & ": " & vbCrLf & pages(i,j) %>"><%if Auth (ii,jj,Permission) then %><B>x</B><%else%>&nbsp;<%end if%></td>
					<%
				wend
			next

%>

		</TR>
<%
		RSV.moveNext

	Loop
	RSV.close
%>
		<TR height='10px' onclick="showUser('');">
			<TD height='10px' class="TD<%=(tmpRowCounter+1) Mod 2%>1"><INPUT TYPE="text" class="textbox1" Value="*ÌÏíÏ*"><br></TD>
			
			
			
		</TR>
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Show Permissions
'-----------------------------------------------------------------------------------------------------
elseif request("act")="edit" then
%>
	<BR><BR>
<%
	userID = request("userID")
	if userID="" then 
		Password	= ""
%>
		<center>
		<H3>ßÇÑÈÑ ÌÏíÏ</H3>
		</center>
<%
	else
		userID = cint(request("userID"))
		set RSM = conn.Execute ("SELECT * from Users where ID="& userID & "")

		UserName	= RSM("UserName")
		Password	= "$aMe@sB4" 'Same as Before  instead of : RSM("Password") 
		if left(RSM("Password"),1)="'" then
			DisabledChecked="checked"
		else
			DisabledChecked=""
		end if 
		RealName	= RSM("RealName") 
		Account		= RSM("Account") 
		Permission	= RSM("Permission")
		Display		= RSM("Display")
		if Display then
			DisplayChecked="checked"
		else
			DisplayChecked=""
		end if

		sp = split (Permission,"#")
		for u=1 to ubound(sp)
			x = left(sp(u),1)
			if x >= "A" then 
				i=cint(asc(x)-55)
			else
				i=cint(x)
			end if

			pagesPermissions(i,0)="ok"
			for v = 2 to len(sp(u))
				tmpp=mid(sp(u),v,1)

				if tmpp >= "A" then 
					j=cint(asc(tmpp)-55)
				else
					j=cint(tmpp)
				end if
				'j=cint(mid(sp(u),v,1))
				pagesPermissions(i,j)="ok"
			next
		next
%>
		<center>
		<H3>æíÑÇíÔ ßÇÑÈÑ: "<%=RealName%>"</H3>
		</center>
<%
	end if
%>

	<FORM METHOD=POST ACTION="?act=submit">

	<TABLE style="font-family:tahoma;font-size:9pt; " Cellspacing="0" Cellpadding="5" width=100%>
	<TR>
		<TD valign=top>

		ãÔÎÕÇÊ:<br><br>
		<TABLE border=1>
		<TR>
			<TD>ßÏ ßÇÑÈÑ:</TD>
			<TD dir=LTR><%=userID%>&nbsp;<INPUT TYPE="hidden" Name="userID" Value="<%=userID%>"></TD>
		</TR>
		<TR>
			<TD>äÇã ßÇÑÈÑí:</TD>
			<TD><INPUT TYPE="text" NAME="UserName" Value="<%=UserName%>" dir=LTR></TD>
		</TR>
		<TR>
			<TD>ÑãÒ ÚÈæÑ</TD>
			<TD><INPUT TYPE="Password" NAME="Password" Value="<%=Password%>" dir=LTR></TD>
		</TR>
		<TR>
			<TD>äÇã äãÇíÔ ÏÇÏå ÔÏå:</TD>
			<TD><INPUT TYPE="text" NAME="RealName" Value="<%=RealName%>"></TD>
		</TR>
		<TR>
			<TD>ÔãÇÑå ÍÓÇÈ ÏÑ ÓíÓÊã:</TD>
			<TD><INPUT TYPE="text" NAME="Account" Value="<%=Account%>"></TD>
		</TR>
		<TR>
			<TD>ÏÑ áíÓÊ äãÇíÔ ÏÇÏå ÔæÏ</TD>
			<TD><INPUT TYPE="checkbox" NAME="Display" <%=DisplayChecked%> ></TD>
		</TR>
		<TR>
			<TD>äãí ÊæÇäÏ Èå ÓíÓÊã æÇÑÏ ÔæÏ</TD>
			<TD><INPUT TYPE="checkbox" NAME="disable" <%=DisabledChecked%>></TD>
		</TR>
		<TR>
			<TD colspan=2 align=center><INPUT TYPE="submit" value="ĞÎíÑå"></TD>
		</TR>
		</TABLE>

		</TD>
		<TD Width="350" >
		ãÌæÒ åÇ:<br><br>
		<table style="font-family:tahoma;font-size:9pt; border:1 dashed #888888; direction:RTL;"  Cellspacing="0" Cellpadding="5">
		<tbody id="PermissionsTable">
		<% 
			for i=0 to NumberOfCategories
				j=0
				if not pagesPermissions(i,j)="ok" then 
					disableChecks="disabled"
					rowBGColor="#F0F0F0"
					groupChecked=""
				else
					disableChecks=""
					rowBGColor="#33AACC"
					groupChecked="checked"
				end if
				%>
				<tr bgcolor='<%=rowBGColor%>'>
					<td width=20> <INPUT TYPE="checkbox" NAME="P<%=i%>" value="<%=j%>" onclick="activeGroup(this)" <%=groupChecked%>> </td>
					<td align=right><B><%=pages(i,0)%></B></td>
				</tr>
				<tr>
					<td width=20> </td>
					<td align=right> 
				<%
				while pages(i,j+1)<>"" 
					j=j+1

					if j <10 then
						jj = j
					else
						jj = chr(55+j)
					end if

					%>
					<INPUT TYPE="checkbox" <%=disableChecks%> NAME="P<%=i%>" value="<%=jj%>" <%if pagesPermissions(i,j)="ok" then %>checked<% end if %>> <%=pages(i,j)%> <BR>
					<%
				wend
			next
			%>
					</td>
				</tr>
		</table>
		</TD>
	</TR>
	</TABLE>
	</FORM>

<%
'-----------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------- Submit Changes
'-----------------------------------------------------------------------------------------------------
elseif request("act")="submit" then
	userID		= request("userID")

	RealName	= request("RealName") 
	RealName	= "N'" & RealName & "'"

	Account		= request("Account") 

	UserName	= request("UserName")
	UserName	= "'" & UserName & "'"

	Password	= request("Password") 
	if Password="$aMe@sB4" then
		Password= "Password"
	else
		Password= "'" & Password & "'"
	end if

	if request("disable")="on" then 
		Password = "'''' + REPLACE(" & Password & ",'''','')"
	else
		Password = "REPLACE(" & Password & ",'''','')"
	end if

	Permission = ""
	for i=0 to NumberOfCategories 
		if request.form("P"& i ) <> "" then
			if i <10 then
				ii = i
			else

				ii = chr(55+i)
			end if
			alll = replace(request.form("P"& i ),", ","")
			Permission = Permission	& "#" & ii &  right( alll, len(alll)-1)
		end if
	next
	Permission	= "'" & Permission & "'"

	Display		= request("Display")
	if Display="on" then
		Display="1"
	else
		Display="0"
	end if


	if userID="" then 
		' Add New User
		mySQL="SELECT MAX(ID)+1 AS NewID FROM Users"
		set RS=Conn.Execute (mySQL) 
		userID = RS("NewID")
		RS.close
		mySQL="INSERT INTO Users (ID, UserName, Password, RealName, Account, Permission, Display) VALUES (" & userID & ", " & userName & ", " & Password & ", " & RealName & ", " & Account & ", " & Permission & "," & Display & ")"
		msg="ßÇÑÈÑ ÌÏíÏ ÇíÌÇÏ ÔÏ."
	else
		' Update User Info
		userID = cint(request("userID"))

		mySQL="UPDATE Users SET UserName=" & UserName & ", Password=" & Password & ", RealName=" & RealName & ", Account=" & Account& ", Permission= " & Permission & ", Display=" & Display & " WHERE ID="& userID & ""
		msg="ÊÛííÑÇÊ ÈÑ Ñæí ÇíÔÇä ÇäÌÇã ÑÏíÏ"
	end if

	conn.Execute (mySQL)
	conn.close
	response.redirect "?act=edit&userID=" & userID & "&msg=" & Server.URLEncode(msg)
end if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--

function activeGroup(src){
	rowNo=src.parentNode.parentNode.rowIndex;
	invTable=document.getElementById("PermissionsTable");
	theRowPr=invTable.getElementsByTagName("tr")[rowNo];
	theRow=invTable.getElementsByTagName("tr")[rowNo+1];
	boxCount=theRow.getElementsByTagName("INPUT").length;
	if (src.checked){
		theRowPr.bgColor= '#33AACC';
		for (i=0;i<boxCount;i++){
			theRow.getElementsByTagName("INPUT")[i].disabled=false;
		}
//		theRow.disabled=false;
	}
	else{
		for (i=0;i<boxCount;i++){
			theRow.getElementsByTagName("INPUT")[i].disabled=true;
		}
		theRowPr.bgColor= '#F0F0F0';
//		theRow.disabled=true;
	}

}
//-->
</SCRIPT>
<!--#include file="tah.asp" -->
