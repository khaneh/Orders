<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="��� �����"
SubmenuItem=5
%>
<!--#include file="top.asp" -->
<BR><BR><BR><BR><BR><CENTER>
<%
'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit new AO MEMO
'-----------------------------------------------------------------------------------------------------
if request("act")="open" then
	application("syslock") = ""
	application("syslockerName") = ""

'-----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------- Submit new AO MEMO
'-----------------------------------------------------------------------------------------------------
elseif request("act")="close" then
	application("syslock") = session("id")
	application("syslockerName") = session("CSRName")

end if

if application("syslock") <> ""  then 
	response.write "<H3>����� ����  "& application("syslockerName") & " ��� ��� ���</H3>"
	%>
	<FORM METHOD=POST ACTION="?act=open">
		<INPUT TYPE="submit" value="��� ���� ���">
	</FORM>
	<%
else
	%>
	<FORM METHOD=POST ACTION="?act=close">
		<INPUT TYPE="submit" value="���� ���">
	</FORM>
	<%
end if 

%>

</CENTER>

<!--#include file="tah.asp" -->