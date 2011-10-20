<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#Include file='../config.asp' -->
<%
function Link2Trace(OrderNo)
	Link2Trace="<A HREF='../order/orderEdit.asp?e=n&radif="& OrderNo & "' target='_balnk'>"& OrderNo & "</A>"
end function

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	body { font-family: tahoma; font-size: 8pt; background-color:#DDDDDD;}
	Input { font-family: tahoma; font-size: 9pt;}
	td { font-family: tahoma; font-size: 8pt;}
</style>
<TITLE>حساب ها</TITLE>
<script language="JavaScript">
<!--
var Arguments = new Array(2)

function documentKeyDown() {
	var theKey = event.keyCode;
	if (theKey == 27) { 
		window.close();
	}
	else if (theKey == 13) {
		selectOperations()
	}
}

document.onkeydown = documentKeyDown;

//-->
</script>
</HEAD>
<BODY>
<%
if request("act")="select" then
	if request("search") <> "" then
		SA_TitleOrName=request("search")
		SA_Action="return selectOperations();"
		SA_SearchAgainURL="InvoiceInput.asp"
		SA_StepText="" '"گام دوم : انتخاب حساب"
		SA_ActName			= "select"	
		SA_SearchBox	="search"		
%>
		<FORM METHOD=POST ACTION="">
		<!--#include File="../ar/include_SelectAccount.asp"-->
		</FORM>
<%
	end if
end if
conn.Close
%>
</BODY>
</HTML>
<script language="JavaScript">
<!--
function selectOperations(){
	var Arguments = new Array;
	notFound=true;
	for (i=0;i<document.getElementsByName("selectedCustomer").length;i++){
		if(document.getElementsByName("selectedCustomer")[i].checked){
			notFound=false;
			Arguments[0]=document.getElementsByName("selectedCustomer")[i].value;
			Arguments[1]=document.getElementById("AccountsTable").getElementsByTagName("tr")[i+1].getElementsByTagName("td")[2].innerText;
		}
	}
	if (notFound)
		return false;
	myString=Arguments.join("#");
	window.dialogArguments.value=myString
	window.close();
	return false;
}
//-->
</script>