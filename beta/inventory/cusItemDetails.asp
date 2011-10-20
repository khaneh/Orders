<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Inventory (5)
'PageTitle= " اصلاح كالا"
'SubmenuItem=6
'if not Auth(5 , 6) then NotAllowdToViewThisPage()
%>
<!--#include File="../config.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<TITLE> <%=PageTitle%> </TITLE>
<style>
	body { font-family: tahoma; font-size: 8pt;}
	body A { Text-Decoration : none ;}
	Input { font-family: tahoma; font-size: 9pt;}
	td { font-family: tahoma; font-size: 8pt;}
	.tt { font-family: tahoma; font-size: 10pt; color:yellow;}
	.tt2 { font-family: tahoma; font-size: 8pt; color:yellow;}
	.inputBut { font-family: tahoma; font-size: 8pt; richness: 10}
	.t7pt { font-size: 8pt;}
	.t8pt { font-size: 10pt;}
	.GenButton { font-family:tahoma; font-size: 9pt; border: 1px solid black; }
</style>
</HEAD>
<BODY bgcolor="#C3DBEB" topmargin=0 leftmargin=0>
<%
'-----------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------ Update Item Properties
'-----------------------------------------------------------------------------------------------------
if isNumeric(request("id")) then
itemId = cint(request("id"))
set RSS=Conn.Execute ("SELECT SUM(DERIVEDTBL.sumQtty) AS sumQttys, DERIVEDTBL.AccountID, dbo.Accounts.AccountTitle FROM (SELECT SUM((CONVERT(tinyint, dbo.InventoryLog.IsInput) - .5) * 2 * dbo.InventoryLog.Qtty) AS sumQtty,  dbo.InventoryLog.owner AS AccountID FROM dbo.InventoryLog WHERE (dbo.InventoryLog.ItemID = " & itemId & " and dbo.InventoryLog.voided=0) GROUP BY dbo.InventoryLog.owner) DERIVEDTBL INNER JOIN dbo.Accounts ON DERIVEDTBL.AccountID = dbo.Accounts.ID GROUP BY DERIVEDTBL.AccountID, dbo.Accounts.AccountTitle having SUM(DERIVEDTBL.sumQtty)<>0")	
%><BR>
<CENTER><B><%=request("name")%> ارسالي ديگران</B></CENTER><BR>
<TABLE dir=rtl align=center width=100%>
<TR bgcolor="eeeeee">
	<TD align=center><!A HREF="default.asp?s=1"><SMALL>شماره حساب</SMALL></A></TD>
	<TD><!A HREF="default.asp?s=2"><SMALL>عنوان حساب</SMALL></A></TD>
	<TD align=center><!A HREF="default.asp?s=3"><SMALL>تعداد</SMALL></A></TD>
</TR>
<%
tmpCounter=0
Do while not RSS.eof
	tmpCounter = tmpCounter + 1
	if tmpCounter mod 2 = 1 then
		tmpColor="#FFFFFF"
		tmpColor2="#FFFFBB"
	Else
		tmpColor="#FFFFFF"
		tmpColor2="#FFFFBB"
		'tmpColor="#DDDDDD"
		'tmpColor2="#EEEEBB"
	End if 

%>
<TR bgcolor="<%=tmpColor%>" height=25>
	<TD align=center dir=ltr><%=RSS("AccountID")%></TD>
	<TD><%=RSS("AccountTitle")%></TD>
	<TD align=center dir=ltr><%=RSS("sumQttys")%></TD>
</TR>
	  
<% 
RSS.moveNext
Loop
%>
</TABLE><br>
</FORM>
<%
end if
%>
<!--#include file="tah.asp" -->