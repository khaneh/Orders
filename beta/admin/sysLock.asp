<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
' Admin
PageTitle="ﬁ›· ”Ì” „"
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
	response.write "<H3>”Ì” „  Ê”ÿ  "& application("syslockerName") & " ﬁ›· ‘œÂ «” </H3>"
	%>
	<FORM METHOD=POST ACTION="?act=open">
		<INPUT TYPE="submit" value="»«“ ﬂ—œ‰ ﬁ›·">
	</FORM>
	<%
else
	%>
	<FORM METHOD=POST ACTION="?act=close">
		<INPUT TYPE="submit" value="»” ‰ ﬁ›·">
	</FORM>
	<%
end if 

%>

</CENTER>

<!--#include file="tah.asp" -->