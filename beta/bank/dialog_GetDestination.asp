<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256">
<meta http-equiv="Content-Language" content="fa">
<style>
	Table { font-size: 9pt;}
	Input { font-family:tahoma; font-size: 9pt;}
</style>
<TITLE>وارد كنيد</TITLE>
<script language="JavaScript">
<!--
var Arguments = new Array(2)

function documentKeyDown() {
	var theKey = event.keyCode;
	var GLA = 0
	var Acc = 0
	if (theKey == 27) { 
		window.close();
	}
	else if (theKey == 13){
		check(document.all.GLAccount);
		if ((document.all.GLAccountName.value != "خالي است " ) && (document.all.GLAccountName.value != "خطا! " ) && (document.all.GLAccountName.value != "چنين حسابي وجود ندارد" ))
			GLA=document.all.GLAccount.value
		check(document.all.Account);
		if ((document.all.AccountName.value != "خالي است " ) && (document.all.AccountName.value != "خطا! " ) && (document.all.AccountName.value != "چنين حسابي وجود ندارد" ))
			Acc=document.all.Account.value
		if (GLA != 0){
			window.dialogArguments.value=GLA+"#"+Acc;
			window.close();
		}
	}
	else if (theKey==9){
		event.keyCode=0;
		if (document.all.GLAccount.value=="")
			document.all.GLAccount.focus();
		else if (document.all.Account.value=="")
			document.all.Account.focus();
		else
			document.all.GLAccount.focus();
	}
}

document.onkeydown = documentKeyDown;

function onMyLoad(){
	if (window.dialogArguments){
		document.getElementById('dlgDispTxt').innerHTML=window.dialogArguments.value;
		window.dialogArguments.value=""
		window.dialogArguments.value=""
	}
	document.all.GLAccount.focus();
}

//-->
</script>
<!--#include file="../config.asp" -->
</HEAD>
<BODY leftmargin=0 topmargin=0 bgcolor='#DDDDFF' onload="onMyLoad()" style="font-family:tahoma;" DIR=RTL>
<BR>
<Center>
<TABLE>
<TR>
	<TD colspan='2'><span id='dlgDispTxt' style="width:350px;">هيچ سوالي نيست</span></TD>
</TR>
<TR>
	<TD align=center>تفصيلي</TD>
	<TD align=center>معين</TD>
</TR>
<TR>
	<TD><INPUT dir="LTR" TYPE="text" NAME="Account" maxlength="6" value="<%=Account%>" onKeyPress='return mask(this);' onBlur='check(this);' style="width:250px;border:solid 1pt black"></TD>
	<TD><INPUT dir="LTR" TYPE="text" NAME="GLAccount" maxlength="5" value="<%=GLAccount%>" onKeyPress='return mask(this);' onBlur='check(this);' style="width:100px;border:solid 1pt black"></TD>
</TR>
<TR>
	<TD><TextArea NAME="AccountName" id="AccountName" readonly style="width:250px;height:40px;font-family:tahoma;font-size:9pt;background:transparent; border:solid 1px white"><%=AccountName%></TextArea></TD>
	<TD><TextArea NAME="GLAccountName" id="GLAccountName" readonly  style="width:100px;height:40px;font-family:tahoma;font-size:9pt;background:transparent; border:solid 1px white;"><%=GLAccountName%></TextArea></TD>
</TR>
</TABLE>
<input type="hidden" Name='tmpDlgArg' value=''>
<input type="hidden" Name='tmpDlgTxt' value=''>
<SCRIPT LANGUAGE="JavaScript">
<!--
var dialogActive=false;

	function mask(src){ 
		var theKey=event.keyCode;
		if (theKey==32){
			event.keyCode=9
			document.all.tmpDlgArg.value="#"
			if (src.name=='GLAccount'){
				document.all.tmpDlgTxt.value="جستجو در نام حساب هاي معين:"
				dialogActive=true
				window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				dialogActive=false
				if (document.all.tmpDlgTxt.value !="") {
					dialogActive=true
					window.showModalDialog('../accounting/dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					dialogActive=false
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						document.all.GLAccountName.value=Arguments[1];
					}
				}
			}
			else if (src.name=='Account') {
				document.all.tmpDlgTxt.value="جستجو در نام حساب هاي تفصيلي:"
				dialogActive=true
				window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
				dialogActive=false
				if (document.all.tmpDlgTxt.value !="") {
					dialogActive=true
					window.showModalDialog('../AR/dialog_SelectAccount.asp?act=select&search='+escape(document.all.tmpDlgTxt.value), document.all.tmpDlgArg, 'dialogWidth:780px; dialogHeight:500px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
					dialogActive=true
					if (document.all.tmpDlgArg.value!="#"){
						Arguments=document.all.tmpDlgArg.value.split("#")
						src.value=Arguments[0];
						document.all.AccountName.value=Arguments[1];
					}
				}
			}
		}
		else if ((theKey >= 48 && theKey <= 57) || theKey==13 ) // [0]-[9] and [Enter] are acceptible
			return true;
		else
			return false;
	}

function check(src){ 
	if (!dialogActive){
		badCode = false;
		if (window.XMLHttpRequest) {
				var objHTTP=new XMLHttpRequest();
			} else if (window.ActiveXObject) {
				var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
			}
		if (src.name=='GLAccount') {
			objHTTP.open('GET','../accounting/xml_GLAccount.asp?id='+src.value,false)
			objHTTP.send()
			tmpStr = unescape(objHTTP.responseText)
			document.all.GLAccountName.value=tmpStr;
		}
		else if (src.name=='Account') {
			objHTTP.open('GET','../accounting/xml_CustomerAccount.asp?id='+src.value,false)
			objHTTP.send()
			tmpStr = unescape(objHTTP.responseText)
			document.all.AccountName.value=tmpStr;
		}
	}
}
//-->
</SCRIPT></BODY>
</HTML>
<% 
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Server.ScriptTimeout = 600
Response.Expires= -1
if (session("ID")="") then
	session.abandon
	response.redirect "../login.asp?err=براي ديدن اين صفحه بايد وارد سيستم شويد"
end if

'conStr="DRIVER={SQL Server};SERVER=(local);DATABASE=sefareshat;UID=sefadmin; PWD=5tgb;"
'conStr = "Provider=SQLNCLI10.1;Persist Security Info=False;User ID=sefadmin;Initial Catalog=sefareshat;Data Source=(local);PWD=5tgb;"

'Set conn = Server.CreateObject("ADODB.Connection")
'conn.open conStr
%>
