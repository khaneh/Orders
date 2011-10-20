<%
'menuItem = request ("menuItem")

'Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
if (session("AID")="") then
'	session.abandon
	response.redirect "login.asp?err=»—«Ì œÌœ‰ «Ì‰ ’›ÕÂ »«Ìœ Ê«—œ ”Ì” „ ‘ÊÌœ"
end if
'conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=jame;Data Source=.\sqlexpress;PWD=5tgb;"

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 6000
conn.open conStr

AppBgColor = "#BBBBBB"  'Bala:"#99AACC"
AppFgColor =  "#C3DBEB" '"#DEEBD9"
SelectedMenuColor = "#0E5499"
UnSelectedMenuColor = "#309261"
SelectedSubMenuColor = "#DEEBD9" '"#C3DBEB" 
UnSelectedSubMenuColor = "#0E5499" ' "#609250"
TabWidth=66
ImgTabSelected="/images/tab-1.gif"
ImgTabNotSelected="/images/tab-2.gif"

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
function sqlSafe (inpStr)
  tmpStr=inpStr
  tmpStr=replace(tmpStr,"'","`")
  tmpStr=replace(tmpStr,"""","`")
  sqlsafe=tmpStr
end function
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
'Session.Timeout = 300
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<TITLE> <%=PageTitle%> </TITLE>
<style>
	body { font-family: tahoma; font-size: 8pt;}
	Input { font-family: tahoma; font-size: 9pt;}
	td { font-family: tahoma; font-size: 8pt;}
	.tt { font-family: tahoma; font-size: 8pt; color:yellow;}
	.tt2 { font-family: tahoma; font-size: 8pt; color:yellow;}
	.inputBut { font-family: tahoma; font-size: 8pt; richness: 10}
	.t7pt { font-size: 8pt;}
	.t8pt { font-size: 10pt;}
	.alak a { color: #cccccc; text-decoration: none;  font-size: 8pt;}
	.alak2 a { color: black; text-decoration: none;  font-size: 8pt; }
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
</HEAD>
<%

CSRName = session("CSRName")

%>
<BODY bgcolor=<%=AppBgColor%>  topmargin=0 leftmargin=0><!<% if onunload<>"" then %> onunload="<%=onunload%>"<% end if %> >
<BR>
<TABLE cellspacing=0 cellpadding=0 width=760 border=0 dir=rtl align=center>
<TR >
	<TD>
	<TABLE cellspacing=0 cellpadding=0>
	<TR height=30 class="alak">
	<%if SubmenuItem=0 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'> ‰ŸÌ„«  ﬂ·Ì</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='default.asp'> ‰ŸÌ„«  ﬂ·Ì</A></TD>
	<%end if %>

	<%if SubmenuItem=1 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'><A HREF='userManager.asp' style="color:yellow;">„œÌ—Ì  <BR>ﬂ«—»—«‰</A></TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='userManager.asp'>„œÌ—Ì  <BR>ﬂ«—»—«‰</A></TD>
	<%end if %>
		<!--
	<%if SubmenuItem=2 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>Õ”«»Â«Ì <br>ÅÌ‘ ›—÷</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='#'>Õ”«»Â«Ì <br>ÅÌ‘ ›—÷</A></TD>
	<%end if %>

	<%if SubmenuItem=3 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>„—«Õ· <BR>”›«—‘</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A HREF='#'>„—«Õ· <BR>”›«—‘</A></TD>
	<%end if %>

	<%if SubmenuItem=4 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>Ê÷⁄Ì  ”›«—‘</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='#'>Ê÷⁄Ì  ”›«—‘</A></TD>
	<%end if %>
	-->
	<%if SubmenuItem=5 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>ﬁ›· ”Ì” „</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='sysLock.asp'>ﬁ›· ”Ì” „</A></TD>
	<%end if %>

	<%if SubmenuItem=6 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>¬Ì „ Â«Ì ›—Ê‘</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='InvItemMng.asp'>¬Ì „ Â«Ì ›—Ê‘</A></TD>
	<%end if %>

	<%if SubmenuItem=7 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>„œÌ—Ì  »«‰ﬂ—Â«</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='BankerMng.asp'>„œÌ—Ì  »«‰ﬂ—Â«</A></TD>
	<%end if %>

	<%if SubmenuItem=8 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>ê“«—‘ Â«Ì Œ›‰</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='Reports.asp'>ê“«—‘ Â«Ì Œ›‰</A></TD>
	<%end if %>

	<%if SubmenuItem=9 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>„œÌ—Ì  œ› — ﬂ·</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='AccountInfo.asp'>„œÌ—Ì  œ› — ﬂ·</A></TD>
	<%end if %>

	<%if SubmenuItem=10 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>—”ÌœêÌ »Â Œÿ«Â«</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='errorHandler.asp'>—”ÌœêÌ »Â Œÿ«Â«</A></TD>
	<%end if %>
	<%if SubmenuItem=11 then %> 
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabSelected%>'>„—«ﬂ“ Â“Ì‰Â</TD>
	<%else %>  
		<TD class=tt2 width="<%=TabWidth%>" align=center background='<%=ImgTabNotSelected%>' ><A style='font-size:8pt;' HREF='costs.asp'>„—«ﬂ“ Â“Ì‰Â</A></TD>
	<%end if %>

	</TR>
	</TABLE>
	</TD>
</TR>
<TR BGCOLOR="<%=SelectedMenuColor%>">
	<TD height=20 align=left  class=alak>
<%
set tmpRS=Conn.Execute ("SELECT Users.RealName FROM Messages INNER JOIN Users ON Messages.MsgFrom = Users.ID WHERE     (Messages.MsgTo = "& session("AID") & ") AND (Messages.Urgent = 2) AND (Messages.IsSmall = 0) AND (Messages.IsRead = 0)")	
if not tmpRS.EOF then%>
	<span style='font-size:12pt;font-weight:bold;align:center;color:yellow;'><marquee dir=rtl width=100%>“Ì‰Â«— ! ‘„« Ìﬂ ÅÌ«„ ŒÌ·Ì ›Ê—Ì «“ <%=tmpRS("RealName")%> œ«—Ìœ! </marquee></span>
<%else%>
	&nbsp;
<%end if%>

<A HREF="logout.asp">Œ—ÊÃ </A>&nbsp;&nbsp;&nbsp;
</TD>
</TR>
<TR >
	<TD >
<TABLE cellspacing=0 cellpadding=0 width=100% height=450 style='border:4px solid <%=SelectedMenuColor%>;' dir=rtl align=center>
<TR>
	<TD colspan=2 valign=top bgcolor=<%=AppFgColor%>>
	<%
	if request.queryString("errmsg")<>"" then
		response.write "<br>" 
		call showAlert (request.queryString("errmsg"),CONST_MSG_ERROR) 
		response.write "<br>" 
	end if
	if request.queryString("msg")<>"" then
		response.write "<br>" 
		call showAlert (request.queryString("msg"),CONST_MSG_INFORM) 
		response.write "<br>" 
	end if
	%>
	
