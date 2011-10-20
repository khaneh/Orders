<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	Table { font-size: 9pt;}
	Input { font-family:tahoma; font-size: 9pt;}
</style>
<TITLE>كالاهاي انبار</TITLE>
<script language="JavaScript">
<!--
var Arguments = new Array(2)

function documentKeyDown() {
	var theKey = event.keyCode;
	if (theKey == 27) { 
		window.close();
	}
}

document.onkeydown = documentKeyDown;


//-->
</script>
</HEAD>

<BODY>
<font face="tahoma">
<%

	mySQL="SELECT * From InventoryItems WHERE (REPLACE([Name], ' ', '') LIKE REPLACE(N'%"& sqlSafe(request("name")) & "%', ' ', '')) ORDER BY Name"

	Set RS1 = conn.Execute(mySQL)

	if (RS1.eof) then %>
		<br><br><br><hr>
		<table align='center'>
			<tr><td bgcolor='#FFCCCC' dir='rtl' align='center'> آيتمي با اين مشخصات پيدا نشد</td>
			</tr>
		</table>
		<hr>
<%	else 
%>
	<TABLE ID="InvoiceItems" border="1" width="100%" cellspacing="1" cellpadding="5" dir="RTL">
		<tr bgcolor='#C3C3FF'>
			<td align='center' width="5%"> <input type="radio" disabled> </td>
			<td align='center' width="15%"> كد آيتم </td>
			<td align='center'> عنوان </td>
		</tr>
<%		tempCounter=0
		while Not (RS1.EOF)
			tempCounter=tempCounter+1
			if (tempCounter Mod 2 = 1)then
				tempColor="#FFFFFF"
			else
				tempColor="#DDDDEE"
			end if
%>				<tr bgcolor='<%=tempColor%>'>
					<td align='center'><input type="radio" name="selectedItem" value="<%=RS1("ID")%>" onclick="selectItem(this)">&nbsp;</td>
					<td dir='ltr'><%=RS1("OldItemID") %>&nbsp;</td>
					<td><%=RS1("Name")%>&nbsp;</td>
				</tr>
<%			RS1.movenext
		wend
%>
		<tr bgcolor='#C3C3FF'>
			<td align='center' colspan="3"><input type="submit" style="font-family:Tahoma" value="انتخاب" onclick="selectAndClose();">&nbsp;</td>
		</tr>
	</TABLE>
	<BR>
	<HR>
	<BR>
<%	end if    
conn.Close
%>
</font>
</BODY>
</HTML>
<script language="JavaScript">
<!--
function selectItem(src){
	rowNo=src.parentNode.parentNode.rowIndex;
//	thePrice=document.getElementsByName("selectedItem")[rowNo-1].value
	myRow=document.getElementById("InvoiceItems").getElementsByTagName("tr")[rowNo];
	Arguments[0]=myRow.getElementsByTagName("td")[1].innerText;
	Arguments[1]=myRow.getElementsByTagName("td")[2].innerText;
}
function selectAndClose(){
myString=Arguments.join("#");
window.dialogArguments.value=myString
window.close();
}
//-->
</script>