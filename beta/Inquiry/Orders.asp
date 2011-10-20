<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><% 
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Server.ScriptTimeout = 600
Response.Expires= -1
if (session("ID")="") then
	session.abandon
	response.redirect "/default.asp?err=session expired"
end if

'conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"

Set conn = Server.CreateObject("ADODB.Connection")
conn.open conStr

%>
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
	Table { font-size: 9pt;}
	Input { font-family:tahoma; font-size: 9pt;}
</style>
<TITLE>سفارش ها</TITLE>
</HEAD>

<BODY>
<font face="tahoma">
<%
if request("act")="select" then
	if request("customer") <> "" then
		SO_Customer=request("customer")
		SO_Action="return selectOperations();"
		SO_StepText="" '"گام سوم : انتخاب سفارش هاي مربوطه"
%>
		<FORM METHOD=POST ACTION="">
		<!--#include File="include_SelectOrder.asp"-->
		</FORM>
<%
	end if
end if
conn.Close
%>
</font>
</BODY>
</HTML>
<script language="JavaScript">
<!--
function documentKeyDown() {
	var theKey = event.keyCode;
	if (theKey == 27) { // Esc
//		event.keyCode=0
		window.dialogArguments.value="[Esc]"
		window.close();
	}
}

document.onkeydown = documentKeyDown;

function selectOperations(){
	var Arguments = new Array;
	argCounter=0
	for (i=0;i<document.getElementsByName("selectedOrders").length;i++){
		if(document.getElementsByName("selectedOrders")[i].checked){
			argCounter++;
			Arguments[argCounter]=document.getElementsByName("selectedOrders")[i].value;
		}
	}
	Arguments[0]=argCounter;
	myString=Arguments.join("#");
//	alert(myString)
	window.dialogArguments.value=myString
	window.close();
	return false;
}
//-->
</script>