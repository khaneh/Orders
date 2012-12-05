<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="„œÌ—Ì  ﬂ«—»—«‰"
SubmenuItem=1
%>
<!--#include file="top.asp" -->
<%
NumberOfCategories = 14
dim pagesPermissions(20,25)
dim pages(20,25)
'-------------------------------------
 pages ( 0 , 0 ) = "’›ÕÂ «Ê·"         
	pages ( 0 , 1 ) = "ŒÊ«‰œ‰ ÅÌ«„"   
	pages ( 0 , 2 ) = "‰Ê‘ ‰ ÅÌ«„"    
	pages ( 0 , 3 ) = "œÌœ‰ œ› —  ·›‰"     
	pages ( 0 , 4 ) = "œ—ŒÊ«”  „—Œ’Ì" 
	pages ( 0 , 5 ) = "œ—ŒÊ«”  ﬂ«·«"  
	pages ( 0 , 6 ) = "ÊÌ—«Ì‘ œ› —  ·›‰"     
	pages ( 0 , 7 ) = "«—”«· ÅÌ«„ »Â Â„Â"     
	pages ( 0 , 8 ) = "ﬂ «»Œ«‰Â"
	pages (0,9)		= "›⁄«·Ì  —Ê“«‰Â"
'-------------------------------------
 pages ( 1 , 0 ) = "›Â—”  ‰«„ùÂ«"
	pages ( 1 , 1 ) = "‰„«Ì‘ «ÿ·«⁄« "
	pages ( 1 , 2 ) = "ÃœÌœ"
	pages ( 1 , 3 ) = "«’·«Õ"
	pages ( 1 , 4 ) = " €ÌÌ— „”∆Ê· Ê ”ﬁ› «⁄ »«—"
	pages ( 1 , 5 ) = " €ÌÌ— Ê÷⁄Ì  Õ”«»"
	pages (1, 6)	= "ê“«—‘ùÂ«"
	pages (1, 7)	= "«’·«Õ ⁄‰Ê«‰ Õ”«»"
	pages (1, 8)	= "œ‘»Ê—œ „‘ —Ì«‰"
'-------------------------------------
 pages ( 2 , 0 ) = "”›«—‘« "
	pages ( 2 , 1 ) = " »œÌ· »Â ”›«—‘"
	pages ( 2 , 2 ) = "«’·«Õ ”›«—‘"
	pages ( 2 , 3 ) = "ÅÌêÌ—Ì ”›«—‘"
	pages ( 2 , 4 ) = " €ÌÌ— Õ”«» ”›«—‘ Ê ’Ê— Õ”«»"
	pages ( 2 , 5 ) = "Ã” ÃÊÌ  ÅÌ‘—› Â"
	pages ( 2 , 6 ) = "ê“«—‘ ﬂ·Ì"
	pages ( 2 , 7 ) = "ê“«—‘ „œÌ—Ì "
	pages ( 2 , 8 ) = "ê“«—‘ „œÌ—Ì  - ﬁ«»·Ì   €ÌÌ— ‘—«Ìÿ"
	pages ( 2 , 9 ) = "«” ⁄·«„"
	pages ( 2 , 10 ) = " «ÌÌœ"						'A
	pages ( 2 , 11 ) = "ç«Å"						'B
	pages ( 2 , 12 ) = "ﬂ‰”· ”›«—‘"					'C
	pages ( 2 , 13 ) = "«⁄·«„ Œÿ« œ— ”›«—‘"			'D
	pages ( 2 , 14 ) = "ç«Å œ” Ê—  Ê·Ìœ"				'E
	pages ( 2 , 15 ) = " Œ›Ì›  « ”ﬁ›"				'F
	pages ( 2 , 16 ) = " Œ›Ì› »œÊ‰ „ÕœÊœÌ "			'G
	pages ( 2 , 17 ) = " €ÌÌ—  —„ùÂ«Ì “„«‰Ì ”›«—‘"			'H
'-------------------------------------
 pages ( 3 , 0 ) = " Ê·Ìœ"
	pages ( 3 , 1 ) = "Ê—Êœ „—«Õ·  Ê·Ìœ"
	pages ( 3 , 2 ) = "»—‰«„Â —Ì“Ì  Ê·Ìœ"
	pages ( 3 , 3 ) = "ÅÌêÌ—Ì ”›«—‘"
	pages ( 3 , 4 ) = "Ã” ÃÊÌ  ÅÌ‘—› Â"
	pages ( 3 , 5 ) = "«ÌÃ«œ œ—ŒÊ«”  ﬂ«·« Ì« ”—ÊÌ” »—«Ì ”›«—‘Ì ﬂÂ ›«ﬂ Ê— ¬‰ ﬁ»·« ’«œ— ‘œÂ «” "
	pages ( 3 , 6 ) = "ÃœÊ·  Ê·Ìœ"
	pages (3,7)		= "⁄„·ﬂ—œ"
	pages (3,9)		= "ÊÌ—«Ì‘ ⁄„·ﬂ—œ"
	pages (3,10)	= "ê“«—‘ ⁄„·ﬂ—œ"					'A
	pages (3,8)		= "„ÊÃÊœÌ ﬂ«€–"
	pages (3,11)	= " €ÌÌ— »Â«Ì ﬂ«€–"				'B
'-------------------------------------
 pages ( 4 , 0 ) = "Œ—Ìœ"
	pages ( 4 , 1 ) = "œ—ŒÊ«”  Œ—Ìœ ﬂ«·«Ì «‰»«—"
	pages ( 4 , 2 ) = "œ—ŒÊ«”  Œ—Ìœ ”«Ì—ﬂ«·« Â« Ê Œœ„« "
	pages ( 4 , 3 ) = "”›«—‘ Œ—Ìœ"
	pages ( 4 , 4 ) = "ÅÌêÌ—Ì Œ—Ìœ"
	pages ( 4 , 5 ) = "œÊŒ ‰"
	pages ( 4 , 6 ) = " €ÌÌ— ”›«—‘ „—»ÊÿÂ"
	pages (4,7)		= " €ÌÌ— ﬁÌ„   Ê«›ﬁ ‘œÂ° œ— ÅÌêÌ—Ì Œ—Ìœ"
	pages (4,8)		= " €ÌÌ— ›—Ê‘‰œÂ"
'-------------------------------------
 pages ( 5 , 0 ) = "«‰»«—"
	pages ( 5 , 1 ) = "’œÊ— ÕÊ«·Â «‰»«— €Ì— „—»Êÿ »Â ”›«—‘ Â« Ê ÅÌ—«„Ê‰"
	pages ( 5 , 2 ) = "Œ—ÊÃ ﬂ«·«"
	pages ( 5 , 3 ) = "Ê—Êœ ﬂ«·«"
	pages ( 5 , 4 ) = " ⁄—Ì› ﬂ«·«Ì ÃœÌœ"
	pages ( 5 , 5 ) = "ê“«—‘"
	pages ( 5 , 6 ) = "«’·«Õ ﬂ«·«"
	pages ( 5 , 7 ) = " ‰ŸÌ„ œ” Â «Ì „ÊÃÊœÌ "
	pages ( 5 , 8 ) = "ê“«—‘ „ÊÃÊœÌ"
	pages ( 5 , 9 ) = "’œÊ— ÕÊ«·Â «‰»«— „—»Êÿ »Â ”›«—‘ Â«"
	pages ( 5 , 10 ) = "À»  œ—ŒÊ«”  ﬂ«·«"					'A
	pages ( 5 , 11 ) = "«‰ ﬁ«· „«·ﬂÌ "					'B
	pages ( 5 , 12 ) = "À»  Ê—Êœ/Œ—ÊÃ œ—  «—ÌŒ œ·ŒÊ«Â"			'C
	pages ( 5 , 13 ) = "À»  ﬁÌ„  ﬂ«·« œ— Ìﬂ »«“Â “„«‰Ì"			'D
	pages (5,14) = "À»  Ê—Êœ ﬂ«·«Ì «—”«·Ì „‘ —Ì"				'E
	pages (5,15) = "À»  Ê—Êœ «“  Ê·Ìœ"						'F
	pages (5,16) = "À»  ﬂ«·«Ì „—ÃÊ⁄Ì"						'G
	pages (5,17) = "Ê—Êœ «“ ”«Ì— «‰»«—ùÂ«"						'H
	pages (5,18) = "Ê—Êœ «“ ÷«Ì⁄« " 						'I
	pages (5,19) = "Ê—Êœ „«“«œ  Ê·Ìœ" 						'J
	pages (5,20) = "œ—ŒÊ«”  ﬂ«·« »Â œ·Ì· Œ—«»Ì" 				'K
'-------------------------------------
 pages ( 6 , 0 ) = "Õ”«»œ«—Ì ›—Ê‘"
	pages ( 6 , 1 ) = "Ê—Êœ ›«ﬂ Ê— "
	pages ( 6 , 2 ) = "«⁄·«„ÌÂ »œÂﬂ«—/»” «‰ﬂ«—"
	pages ( 6 , 3 ) = "«’·«Õ ›«ﬂ Ê—"
	pages ( 6 , 4 ) = "Ê—Êœ ›«ﬂ Ê— »—ê‘ "
	pages ( 6 , 5 ) = "«’·«Õ ›«ﬂ Ê— »—ê‘ "
	pages ( 6 , 6 ) = "ê“«—‘ Õ”«»"
	pages ( 6 , 7 ) = "œÊŒ ‰"
	pages ( 6 , 8 ) = "ê“«—‘ ›—Ê‘ ŒÊœ‘"
	pages ( 6 , 9 ) = "ê“«—‘ ›—Ê‘ Â„Â"
	pages ( 6 , 10 ) = " €ÌÌ—«  Ã—ÌÌ ›«ﬂ Ê— »⁄œ «“ ’œÊ—"			'A
	pages ( 6 , 11 ) = "unused"						'B
	pages ( 6 , 12 ) = " «ÌÌœ ›«ﬂ Ê—"						'C
	pages ( 6 , 13 ) = "’œÊ— ›«ﬂ Ê—"						'D
	pages ( 6 , 14 ) = "ç«Å ÅÌ‘‰ÊÌ” ›«ﬂ Ê—"					'E
	pages ( 6 , 15 ) = "Œ—ÊÃ ›«ﬂ Ê— «“ ’œÊ—"				'F
	pages ( 6 , 16 ) = "Œ—ÊÃ ›«ﬂ Ê— «“  «ÌÌœ"				'G
	pages ( 6 , 17 ) = "«»ÿ«· «⁄·«„ÌÂ ›—Ê‘"					'H
	pages ( 6 , 18 ) = "’œÊ— ›«ﬂ Ê— œ—  «—ÌŒ œ·ŒÊ«Â"			'I
	pages ( 6 , 19 ) = "ç«Å ê“«—‘ Õ”«»"					'J
	pages ( 6 , 20 ) = "Ê—Êœ (›«ﬂ Ê— / »—ê‘  ) »œÊ‰ ”›«—‘"		'K
	pages ( 6 , 21 ) = "ﬂÅÌ ÅÌ‘ ‰ÊÌ”"					'L
	pages ( 6 , 22 ) = " €ÌÌ—«  «·›/» ﬁ»· «“ ’œÊ—"			'M
	pages (6, 23)	 = " €ÌÌ—«  «·›/» »⁄œ «“ ’œÊ—"			'N
	pages (6,24)	 = "ç«Å Œÿ—‰«ﬂ"						'O

'-------------------------------------
 pages ( 7 , 0 ) = "Õ”«»œ«—Ì Œ—Ìœ"
	pages ( 7 , 1 ) = "Ê—Êœ ›«ﬂ Ê— Œ—Ìœ"
	pages ( 7 , 2 ) = " «ÌÌœ ›«ﬂ Ê— Œ—Ìœ"
	pages ( 7 , 3 ) = "ê“«—‘"
	pages ( 7 , 4 ) = "«⁄·«„ÌÂ »œÂﬂ«—/»” «‰ﬂ«—"
	pages ( 7 , 5 ) = "ê“«—‘ Õ”«»"
	pages ( 7 , 6 ) = "œÊŒ ‰"
	pages ( 7 , 7 ) = "ê“«—‘ Œ—Ìœ"
	pages ( 7 , 8 ) = "«»ÿ«· «⁄·«„ÌÂ Œ—Ìœ"
	pages ( 7 , 9 ) = "«»ÿ«· ›«ﬂ Ê— Œ—Ìœ"
	pages ( 7 , 10 ) = "ç«Å ê“«—‘ Õ”«»"						'A
	pages ( 7 , 11 ) = "œ‘»Ê—œ Œ—Ìœ"							'B
'-------------------------------------
 pages ( 8 , 0 ) = "Õ”«»œ«—Ì "
	pages ( 8 , 1 ) = "Ê—Êœ  ”‰œ"
	pages ( 8 , 2 ) = "œÌœ‰ ”‰œ Â«Ì Õ–› ‘œÂ"
	pages ( 8 , 3 ) = "[’›ÕÂ] ¬—‘ÌÊ «”‰«œ"
	pages ( 8 , 4 ) = "œ› — ﬂ·"
	pages ( 8 , 5 ) = "œÌœ‰ ”‰œ Â«Ì ﬁÿ⁄Ì ‘œÂ"
	pages ( 8 , 6 ) = "œÌœ‰ ”‰œ Â«Ì »——”Ì ‘œÂ Ê  €ÌÌ— Ê÷⁄Ì  ¬‰Â«"
	pages ( 8 , 7 ) = "œÌœ‰ Â„Â ”‰œ Â«Ì „Êﬁ "
	pages ( 8 , 8 ) = "«’·«Õ Ì«  €ÌÌ— Ê÷⁄Ì  Â„Â ”‰œ Â«Ì „Êﬁ "
	pages ( 8 , 9 ) = "œÌœ‰ Â„Â Ì«œœ«‘  Â«"
	pages ( 8 , 10 ) = "ê“«—‘ „⁄Ì‰"							'A
	pages ( 8 , 11 ) = "ê“«—‘  ›’Ì·Ì"						'B
	pages ( 8 , 12 ) = "”‰œÂ«Ì “Ì— ”Ì” „"					'C
	pages ( 8 , 13 ) = " €ÌÌ— œ› — ﬂ· ÅÌ‘ ›—÷ ”Ì” „"		'D
	pages ( 8 , 14 ) = "œÌœ‰ ”«Ì— ê“«—‘ Â«"					'E
	pages ( 8 , 15 ) = "”‰œ „—ﬂ» (Ê—Êœ Ê ÊÌ—«Ì‘) "			'F
	pages ( 8 , 16 ) = "ê“«—‘ œ‘»Ê—œ Œ“«‰Âùœ«—Ì"			'G
	pages ( 8 , 17 ) = " „«„ Ì«œœ«‘  ‘œÂùÂ« —« » Ê«‰œ „Ê›ﬁ  ﬂ‰œ"	'H
'-------------------------------------
 pages ( 9 , 0 ) = "’‰œÊﬁ"
	pages ( 9 , 1 ) = "œ—Ì«› "
	pages ( 9 , 2 ) = "Å—œ«Œ  ‰ﬁœ"
	pages ( 9 , 3 ) = "ê“«—‘ ’‰œÊﬁ"
	pages ( 9 , 4 ) = "»” ‰"
	pages ( 9 , 5 ) = "«ÌÃ«œ"
	pages ( 9 , 6 ) = "ê“«—‘ ”—Å—”  ’‰œÊﬁ"
	pages ( 9 , 9 ) = "ê“«—‘ ”—Å—”  ’‰œÊﬁ (ÊÌéÂ)"
	pages ( 9 , 7 ) = "«»ÿ«· »Ì ﬁÌœ Ê ‘—ÿ œ—Ì«› /Å—œ«Œ "
	pages ( 9 , 8 ) = "ç«Å ›—„ ›«ﬂ Ê—"
'-------------------------------------
 pages ( 10 , 0 ) = "»«‰ﬂ"									'A
	pages ( 10 , 1 ) = "’œÊ— çﬂ"
	pages ( 10 , 2 ) = "ÅÌêÌ—Ì çﬂ œ—Ì«› Ì"
	pages ( 10 , 3 ) = "ÅÌêÌ—Ì çﬂ Å—œ«Œ Ì"
	pages ( 10 , 4 ) = "À»  ÕÊ«·Â"
	pages ( 10 , 5 ) = "Å—œ«Œ  ›«ﬂ Ê—"
	pages ( 10 , 6 ) = "œ—Ì«›  çﬂ „⁄Ì‰"
	pages ( 10 , 7 ) = "’œÊ— çﬂ „⁄Ì‰"
	pages ( 10 , 8 ) = "œ› — çﬂ"
	pages ( 10 , 9 ) = "„€«Ì— "
'-------------------------------------
 pages ( 11 , 0 ) = "Õ”«»œ«—Ì ”«Ì—"							'B
	pages ( 11 , 1 ) = "ê“«—‘ Õ”«»"
	pages ( 11 , 2 ) = "”«Ì— «⁄·«„ÌÂ Â«"
	pages ( 11 , 3 ) = "«‰ ﬁ«·"
	pages ( 11 , 4 ) = "œÊŒ ‰"
	pages ( 11 , 5 ) = "ê“«—‘ ”«Ì—"
	pages ( 11 , 6 ) = "«»ÿ«· «⁄·«„ÌÂ ”«Ì—"
	pages ( 11 , 7 ) = "ç«Å ê“«—‘ Õ”«»"						
	pages ( 11 , 8 ) = "ÕﬁÊﬁ"						

'-------------------------------------
 pages ( 12 , 0 ) = "”«Ì— ê“«—‘ Â«Ì »Œ‘ Õ”«»œ«—Ì ›—Ê‘"		'C
	pages ( 12 , 1 ) = "ê“«—‘ ›—Ê‘ —Ê“«‰Â"
	pages ( 12 , 2 ) = "ê“«—‘ ›—Ê‘ —Ê“«‰Â »«  «—ÌŒ œ·ŒÊ«Â"
	pages ( 12 , 3 ) = "ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘ "
	pages ( 12 , 4 ) = "ê“«—‘ ›«ﬂ Ê— Â«Ì «·›"
	pages ( 12 , 5 ) = "ê“«—‘ —”Ìœ Â«Ì «·›"
	pages ( 12 , 6 ) = "ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘ »—Õ”» ›—Ê‘‰œÂ"
	pages ( 12 , 7 ) = "ê“«—‘ ﬁÌ„   „«„ ‘œÂ"
	pages ( 12 , 8 ) = "ê“«—‘  —«“ çÂ«— ” Ê‰Ì"
	pages ( 12 , 9 ) = "ê“«—‘ „‘ —Ì«‰Ì ﬂÂ «⁄ »«— ¬‰Â« ’›— ‰Ì” "
	pages (12,10) = "ê“«—‘ »—ê‘  «“ ›—Ê‘"			'A
	pages ( 12 , 11 ) = "ê“«—‘ ¬Ì „ Â«Ì ›—Ê‘ ÿ—«ÕÌ"	'B
'-------------------------------------
 pages ( 13 , 0 ) = "„Ì“ ﬂ«—"								'D
	pages ( 13 , 1 ) = "”›«—‘ Â«Ì »«“"
	pages ( 13 , 2 ) = "›«ﬂ Ê— Â«Ì  «ÌÌœ ‰‘œÂ"
	pages ( 13 , 3 ) = "›«ﬂ Ê— Â«Ì  «ÌÌœ ‘œÂ (’«œ— ‰‘œÂ)"
	pages ( 13 , 4 ) = "›«ﬂ Ê— Â«Ì ’«œ— ‘œÂ ( ”ÊÌÂ ‰‘œÂ)"
	pages ( 13 , 5 ) = "Õ”«» Â«Ì „«‰œÂ œ«—"
	pages ( 13 , 6 ) = "çﬂ Â«Ì »—ê‘ Ì"
	pages ( 13 , 7 ) = "ê“«—‘ ﬂ·Ì „”ÊÊ· Õ”«»"

'-------------------------------------
 pages ( 14 , 0 ) = "„œÌ—Ì "								'E
	pages ( 14 , 1 ) = "Reserved - Ê—Êœ »Â ﬁ”„  admin"
	pages ( 14 , 2 ) = "Reserved -  €ÌÌ— „ÃÊ“ ﬂ«—»—«‰"
	pages ( 14 , 3 ) = "Reserved"
	pages ( 14 , 4 ) = "Reserved"
	pages ( 14 , 5 ) = "œÌœ‰ „Ì“ ﬂ«— ”«Ì— ﬂ«—»—«‰"
	pages ( 14 , 6 ) = "œÌœ‰ ê“«—‘ùÂ«Ì «ﬂ”·Ì"


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
		<TD rowspan="3" height=100>ﬂ«—»—</TD>
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
			<TD height='10px' class="TD<%=(tmpRowCounter+1) Mod 2%>1"><INPUT TYPE="text" class="textbox1" Value="*ÃœÌœ*"><br></TD>
			
			
			
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
		<H3>ﬂ«—»— ÃœÌœ</H3>
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
		'costCenterString= RSM("costCenter")
		RSM.close
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
		<H3>ÊÌ—«Ì‘ ﬂ«—»—: "<%=RealName%>"</H3>
		</center>
<%
	end if
%>

	<FORM METHOD=POST ACTION="?act=submit">

	<TABLE style="font-family:tahoma;font-size:9pt; " Cellspacing="0" Cellpadding="5" width=100%>
	<TR>
		<TD valign=top>

		„‘Œ’« :<br><br>
		<TABLE border=1>
		<TR>
			<TD>ﬂœ ﬂ«—»—:</TD>
			<TD dir=LTR><%=userID%>&nbsp;<INPUT TYPE="hidden" Name="userID" Value="<%=userID%>"></TD>
		</TR>
		<TR>
			<TD>‰«„ ﬂ«—»—Ì:</TD>
			<TD><INPUT TYPE="text" NAME="UserName" Value="<%=UserName%>" dir=LTR></TD>
		</TR>
		<TR>
			<TD>—„“ ⁄»Ê—</TD>
			<TD><INPUT TYPE="Password" NAME="Password" Value="<%=Password%>" dir=LTR></TD>
		</TR>
		<TR>
			<TD>‰«„ ‰„«Ì‘ œ«œÂ ‘œÂ:</TD>
			<TD><INPUT TYPE="text" NAME="RealName" Value="<%=RealName%>"></TD>
		</TR>
		<TR>
			<TD>‘„«—Â Õ”«» œ— ”Ì” „:</TD>
			<TD><INPUT TYPE="text" NAME="Account" Value="<%=Account%>"></TD>
		</TR>
		<TR>
			<TD>œ— ·Ì”  ‰„«Ì‘ œ«œÂ ‘Êœ</TD>
			<TD><INPUT TYPE="checkbox" NAME="Display" <%=DisplayChecked%> ></TD>
		</TR>
		<TR>
			<TD>‰„Ì  Ê«‰œ »Â ”Ì” „ Ê«—œ ‘Êœ</TD>
			<TD><INPUT TYPE="checkbox" NAME="disable" <%=DisabledChecked%>></TD>
		</TR>
		<tr>
			<td>„—«ﬂ“ Â“Ì‰Â</td>
			<td>
				<table width="100%">
					<%
					'----------------------------- COST CENTER ---------------------------------
					if userID<>"" then
						mySQL="SELECT cost_centers.name as costCenterName, cost_drivers.*,isnull(cost_user_relations.driver_id,-1) as driver_id FROM cost_centers inner join cost_drivers on cost_centers.id=cost_drivers.cost_center_id left outer join cost_user_relations on cost_drivers.id=cost_user_relations.driver_id and cost_user_relations.user_id=" & userID
					else
						mySQL="SELECT cost_centers.name as costCenterName, cost_drivers.*,isnull(cost_user_relations.driver_id,-1) as driver_id FROM cost_centers inner join cost_drivers on cost_centers.id=cost_drivers.cost_center_id left outer join cost_user_relations on cost_drivers.id=cost_user_relations.driver_id"
					end if
					set rrs=Conn.Execute(mySQL)
					oldCostCenter=-1
					while not rrs.eof
						theTitle=""
						set oprs=Conn.Execute("select * from cost_operation_type where driver_id=" & rrs("id"))
						'response.write ("select * from cost_operation_type where driver_id=" & rrs("id"))
						while not oprs.eof 
							theTitle= theTitle & oprs("name") & "° "
							oprs.moveNext
						wend
						oprs.close
						if oldCostCenter=cint(rrs("cost_center_id")) then
							
							%>
							<tr>
								<td title="<%=theTitle%>"><%=rrs("name")%></td>
								<td><input type="checkbox" name="costDriver-<%=rrs("id")%>" <%if cint(rrs("driver_id"))>0 then response.write("checked='checked'")%>></td>
							</tr>
							<%
						else
							%>
							<tr bgcolor="#33AACC">
								<td colspan="2" align="center"><b><%=rrs("costCenterName")%></b></td>
							</tr>
							<tr>
								<td title="<%=theTitle%>"><%=rrs("name")%></td>
								<td><input type="checkbox" name="costDriver-<%=rrs("id")%>" <%if cint(rrs("driver_id"))>0 then response.write("checked='checked'")%>></td>
							</tr>
							<%
						end if
							%>
						<%
						oldCostCenter=cint(rrs("cost_center_id"))
						rrs.MoveNext
					wend
					rrs.close
					'--------------------------------------------------------------------------------
					%>
				</table>
			</td>
		</tr>
		<TR>
			<TD colspan=2 align=center><INPUT TYPE="submit" value="–ŒÌ—Â"></TD>
		</TR>
		</TABLE>

		</TD>
		<TD Width="350" >
		„ÃÊ“ Â«:<br><br>
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
	'------------------------ COST CENTER -----------------------------------
	if userID<>"" then
		conn.Execute("delete cost_user_relations where user_id=" & userID)
	end if
	set rrs = Conn.execute("select id from cost_drivers")
	while not rrs.eof
		if request("costDriver-"&rrs("id"))="on" then 
			conn.Execute("insert into cost_user_relations (user_id,driver_id) values ("& userID & "," & rrs("id") & ")")
		end if
		rrs.MoveNext
	wend
	rrs.close
	if userID="" then 
		' Add New User
		mySQL="SELECT MAX(ID)+1 AS NewID FROM Users"
		set RS=Conn.Execute (mySQL) 
		userID = RS("NewID")
		RS.close
		mySQL="INSERT INTO Users (ID, UserName, Password, RealName, Account, Permission, Display) VALUES (" & userID & ", " & userName & ", " & Password & ", " & RealName & ", " & Account & ", " & Permission & "," & Display & ")"
		msg="ﬂ«—»— ÃœÌœ «ÌÃ«œ ‘œ."
	else
		' Update User Info
		userID = cint(request("userID"))

		mySQL="UPDATE Users SET UserName=" & UserName & ", Password=" & Password & ", RealName=" & RealName & ", Account=" & Account& ", Permission= " & Permission & ", Display=" & Display & " WHERE ID="& userID & ""
		msg=" €ÌÌ—«  »— —ÊÌ «Ì‘«‰ «‰Ã«„ ê—œÌœ"
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
