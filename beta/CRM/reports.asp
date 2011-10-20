<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'AR (6)
PageTitle="ê“«—‘ùÂ«"
SubmenuItem=6
if not Auth(1 , 6) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<!--#include File="../include_UtilFunctions.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {border:1pt solid white;vertical-align:top;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTable2 th {font-size:9pt; padding:5px; background-color:#0080C0;height:25px;}
	.RepTable2 td {height:25px;}
	.RepTable2 input {font-family:tahoma; font-size:9pt; border:1 solid black;}
	.RepTable2 select {font-family:tahoma; font-size:9pt; border:1 solid black;}
</STYLE>
<BR>
<TABLE class="RepTable" align=center cellspacing=10 >
	<TR height=110>
		<TD style="border:1pt solid white;width:130pt" align=center>
		<FORM METHOD=POST ACTION="Rep_Credit.asp">
			<table class="RepTable2" id="AInvoices">
			<tr>
				<th colspan="2">ê“«—‘ ”ﬁ› «⁄ »«—</td>
			</tr>
			<tr>
				<td align=left>»Ì‘ — «“</td>
				<td><INPUT TYPE="text" NAME="limit" Value="0" style="width:75px;direction:LTR;" maxlength=10></td>
			</tr>
			
			</table>
			<INPUT Class="GenButton" style="border:1 solid black;" TYPE="submit" value=" ‰„«Ì‘ ">&nbsp;
			
			</FORM>
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</td>
	</tr>
	<TR height=110>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</td>
	</tr>
	<TR height=110>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</TD>
		<TD style="border:1pt solid white;width:130pt" align=center>
		</td>
	</tr>
<%
'-----------------------------------------------------------------------------------------------------

'---------------------------------------------------------------

%>

<!--#include file="tah.asp" -->
