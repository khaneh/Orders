<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
'Home (0)
PageTitle= " ������� �����"
SubmenuItem=4

if not Auth(0 , 4) then NotAllowdToViewThisPage()

sendTo = session("id")
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<%

function sqlSafe (s)
  st=s
  st=replace(St,"'","`")
  st=replace(St,chr(34),"`")
  sqlsafe=st
end function

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ Send New Message
'-----------------------------------------------------------------------------------------------------
if request.form("Submit")="��� �������" then

'	MsgTo			=	"4"		'=====>   4 = Mr. Kavakeb
'	Changed by kid 840727
	MsgTo			=	"6"		'=====>   1 = Mr. Vazehi
	msgTitle		=	"request for furlough"
	msgDate			=	sqlSafe(request.form("msgDate"))
	msgBody			=	sqlSafe(left(request.form("msgBody"),1999))
	fromTime		=	sqlSafe(request.form("fromTime"))
	toTime			=	sqlSafe(request.form("toTime"))
	RelatedTable	=	"NaN"
	relatedID		=	"0"
	replyTo			=	"0"
	IsReply			=	"0"
	MsgFrom			=	session("ID")
	MsgTime			=	currentTime10()
	MsgDate			=	shamsiToday()

	'msgBody2 = "������� ����� ���� ����� <span dir=ltr>" & msgDate & "</span> �� ���� " & fromTime & "�� ���� " & toTime & "<hr noshade style=''height:1''>" & msgBody
	msgBody2 = "������� ����� ���� ����� " & msgDate & " �� ���� " & fromTime & "�� ���� " & toTime & "----------- �������: " & msgBody

	
	MySQL = "INSERT INTO Messages (MsgFrom, MsgTo, MsgTime, MsgDate, IsRead, MsgTitle, MsgBody, replyTo, IsReply, relatedID, RelatedTable) VALUES ( "& MsgFrom & ", "& MsgTo & ", N'"& MsgTime & "', N'"& MsgDate & "', 0, N'"& MsgTitle & "', N'"& MsgBody2 & "', "& replyTo & ", "& IsReply & ", "& relatedID & ", '"& RelatedTable & "')"

	conn.Execute MySQL 
	response.write "<center><br><br>������� ��� ���� ������ ����� ��</center>"

end if

'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------ New Message Form
'-----------------------------------------------------------------------------------------------------
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function hideIT()
{
//alert(document.all.RelatedTable.value)
if(document.all.RelatedTable.value!="no") 
	{
		document.all.relatedIDSpan.style.visibility= 'visible'
		document.all.relatedID.value = "0"
		document.all.relatedID.focus()
	}
	else
	{
		document.all.relatedIDSpan.style.visibility= 'hidden'
		document.all.relatedID.value = "0"
	}
}
//-->
</SCRIPT>

<center><BR><BR><BR>
<FORM METHOD=POST ACTION="furlough.asp">
<INPUT TYPE="hidden" name="replyTo" value="0">
<INPUT TYPE="hidden" name="IsReply" value="0">
<TABLE>
<TR>
	<TD colspan=2 align=center><H3>������� �����</H3></TD>
</TR>
<!--TR>
	<TD align=left>���: </TD>
	<TD align=right>
		<select name="kind" class=inputBut  >
		<option value="0">�� ��� �� ����</option>
		<option value="1">��� ����</option>
		</select> <BR>
	</TD>
</TR-->
<TR>
	<TD align=left>�����</TD>
	<TD align=right>
		<INPUT TYPE="text" NAME="msgDate" value="<%=shamsiDate(date()+1)%>" onKeyPress="return maskDate(this);" onblur="acceptDate(this)" dir=ltr class=inputBut size=31>
	</TD>
</TR>
<TR>
	<TD align=left>�� ����</TD>
	<TD align=right>
		<INPUT TYPE="text" NAME="fromTime"  class=inputBut size=8 onKeyPress="return maskTime(this);"  dir=ltr>&nbsp;
		������: 
		<INPUT TYPE="text" NAME="toTime"  class=inputBut size=9 onKeyPress="return maskTime(this);" dir=ltr>
	</TD>
</TR>
<TR>
	<TD align=left>�����</TD>
	<TD align=right>
		<TEXTAREA NAME="msgBody" ROWS="4"  class=inputBut COLS="32"></TEXTAREA>
	</TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=center><br><INPUT TYPE="submit" name="submit" value="��� �������"></TD>
</TR>
<TR>
	<TD align=left></TD>
	<TD align=right>
	</TD>
</TR>
</TABLE>
</FORM>

<!--#include file="tah.asp" -->