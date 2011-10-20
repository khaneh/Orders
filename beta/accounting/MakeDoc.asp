<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Accounting (8)
PageTitle= "”‰œ Â«Ì “Ì— ”Ì” „"
SubmenuItem=9
if not Auth(8 , "C") then NotAllowdToViewThisPage()

if request.querystring("act")="showItems" then
	response.buffer = false
end if
%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<STYLE>
	.RepTable {font-family:tahoma; font-size:9pt; direction: RTL; }
	.RepTable td {padding:5;border:1pt solid gray;}
	.RepTable a {text-decoration:none; color:#222288;}
	.RepTable a:hover {text-decoration:underline;}
	.RepTableTitle {background-color: #CCCCFF; text-align: center; height:50;}
	.RepTableHeader {background-color: #BBBBBB; text-align: center; font-weight:bold;}
	.RepTableFooter {background-color: #BBBBBB; direction: LTR; }
	.RepTD1 {width:50px;}
	.RepTR1 {background-color: #DDDDDD;}
	.RepTR2 {background-color: #FFFFFF;}
	.RepGeneralInput {width:70px; font-family:tahoma; font-size:8pt; border:1pt solid gray; background:transparent; direction: LTR; }
	.RepGLInput {width:35px; font-family:tahoma; font-size:8pt; border:none; background:#FFCCCC;d:transparent;}
	.RepGLInput2 {width:35px; font-family:tahoma; font-size:8pt; border:none; background:transparent;}
	.RepAccountInput {width:50px; font-family:tahoma; font-size:8pt; text-align:left; border:none; background:#FFCCCC;d:transparent;}
	.RepAccountInput2 {width:50px; font-family:tahoma; font-size:8pt; text-align:left; border:none; background:transparent;}
</STYLE>
<%

Dim Descriptions(10) 
' type 10 is defined 'Ê«—Ì“ ÊÃÂ ‰ﬁœ'
Descriptions(10)="Ê«—Ì“ ÊÃÂ ‰ﬁœ"

if request.querystring("act")="SubmitDoc" then
'	response.write request("DocDate")
	DocDate = sqlSafe(request.form("DocDate"))
	'---- Checking wether EffectiveDate is valid in current open GL
	if (DocDate < session("OpenGLStartDate")) OR (DocDate > session("OpenGLEndDate")) then
		Conn.close
		response.redirect "?errMsg=" & Server.URLEncode("Œÿ«!  «—ÌŒ Ê«—œ ‘œÂ „ ⁄·ﬁ »Â ”«· „«·Ì Ã«—Ì ‰Ì” .")
	end if 
	'----
	'----- Check GL is closed
	if (session("IsClosed")="True") then
		Conn.close
		response.redirect "AccountInfo.asp?errMsg=" & Server.URLEncode("Œÿ«! ”«· „«·Ì Ã«—Ì »” Â ‘œÂ Ê ‘„« ﬁ«œ— »Â  €ÌÌ— œ— ¬‰ ‰Ì” Ìœ.")
	end if 
	'----

	'session.Timeout=20
	if request.form("Items").count > 0 then

	  '--- Checking if these lines are already updated in GL
		ARItems="0"
		APItems="0"
		AOItems="0"
		for i=1 to request.form("Items").count
			n=request.form("Items")(i)
			sys=request.form("Sys"&n)
			Item=request.form("Item"&n)

			select case sys
			case "AR":
				ARItems = ARItems & ", " & Item 
			case "AP":
				APItems = APItems & ", " & Item 
			case "AO":
				AOItems = AOItems & ", " & Item 
			end select
		next

		errMsg=Server.URLEncode("Œÿ«! <br> Õœ «ﬁ· ÌﬂÌ «“ ”ÿ— Â«Ì «‰ Œ«» ‘œÂ ﬁ»·« ”‰œ ŒÊ—œÂ «” . <br><br> ÂÌç ”‰œÌ À»  ‰‘œ.")

		Set RS1=Conn.Execute ("SELECT * FROM ARItems WHERE ID IN("& ARItems & ") AND (GL_Update=0)")
		if not RS1.eof then
			Conn.close
			response.redirect "?errMsg=" & errMsg
		end if
		RS1.Close

		Set RS1=Conn.Execute ("SELECT * FROM APItems WHERE ID IN("& APItems & ") AND (GL_Update=0)")
		if not RS1.eof then
			Conn.close
			response.redirect "?errMsg=" & errMsg
		end if
		RS1.Close

		Set RS1=Conn.Execute ("SELECT * FROM AOItems WHERE ID IN("& AOItems & ") AND (GL_Update=0)")
		if not RS1.eof then
			Conn.close
			response.redirect "?errMsg=" & errMsg
		end if
		RS1.Close

	  '--- End of Checking if the lines are already updated in GL

		DocNo=request.form("DocNo")
		creationDate=ShamsiToday()
		WarningMsg=""
		Set RS1=Conn.Execute ("SELECT GLDocID FROM GLDocs WHERE (GLDocID='"& DocNo & "') AND (GL='"& request.form("GL") & "')")
		if not RS1.eof then
			Set RS1= Conn.Execute("SELECT Max(GLDocID) AS MaxGLDocID FROM GLDocs WHERE (GL='"& request.form("GL") & "')")
			DocNo=RS1("MaxGLDocID")+1
			WarningMsg="‘„«—Â ”‰œ  ﬂ—«—Ì »Êœ.<br>”‰œ »« ‘„«—Â <B>"& DocNo & "</B> À»  ‘œ."
		end if
		RS1.Close
		mySQL="INSERT INTO GLDocs (GL, GLDocID, GLDocDate, CreatedDate, CreatedBy, IsTemporary, BySubSystem) VALUES ("&_
		request.form("GL") & ", " & DocNo & ", N'" & DocDate & "', N'" & creationDate & "', " & session("ID") & ", 1, 1)"
		
		Conn.Execute(mySQL)

		Set RS1 = Conn.Execute("SELECT MAX(ID) AS MaxID FROM GLDocs WHERE (GL="& request.form("GL") & ") AND (GLDocID="& DocNo & ") AND (CreatedDate=N'"& creationDate & "')")
		GLDocID=RS1("MaxID")

		for i=1 to request.form("Items").count
			n=request.form("Items")(i)
			sys=request.form("Sys"&n)
			Item=request.form("Item"&n)
			Lines=request.form("Lines"&n)

			for j=1 to Lines
				Account=request.form("Account"&n)(j)
				GLAccount=request.form("GLAccount"&n)(j)
				Description=request.form("Description"&n)(j) 
				Ref1=request.form("Ref1"&n)(j) 
				Ref2=request.form("Ref2"&n)(j) 
				Amount=request.form("Amount"&n)(j)
				IsCredit=request.form("IsCredit"&n)(j) 
				if IsCredit="True" then
					IsCredit=1
				else
					IsCredit=0
				end if

				if Account="" or Account="0" then
					mySQL="INSERT INTO GLRows (GLDoc, GLAccount, Amount, Description, Ref1, Ref2, SYS, Link, IsCredit) VALUES ("&_
					GLDocID & ", " & GLAccount & ", " & Amount & ", N'" & Description & "', N'" & Ref1 & "', N'" & Ref2 & "', '" & sys & "', " & Item & ", " & IsCredit & ")"
				else
					mySQL="INSERT INTO GLRows (GLDoc, Tafsil, GLAccount, Amount, Description, Ref1, Ref2, SYS, Link, IsCredit) VALUES ("&_
					GLDocID & ", " & Account & ", " & GLAccount & ", " & Amount & ", N'" & Description & "', N'" & Ref1 & "', N'" & Ref2 & "', '" & sys & "', " & Item & ", " & IsCredit & ")"
				end if
				Conn.Execute(mySQL)
			next

			Conn.Execute("UPDATE "& sys & "Items SET GL_Update =0 WHERE ("& sys & "Items.ID='"& Item & "')")
		next
		Conn.Close
		response.redirect "GLMemoDocShow.asp?id="& GLDocID &"&msg="& Server.URLEncode("”‰œ »« „Ê›ﬁÌ  «ÌÃ«œ ‘œ.") &"&errmsg="& Server.URLEncode(WarningMsg)
	end if
elseif request("act")="showItems" then

'	session.Timeout=120

	Set RS1= Conn.Execute("SELECT Max(GLDocID) AS MaxGLDocID FROM GLDocs WHERE (GL='"& OpenGL & "')")
	DocNo=RS1("MaxGLDocID")+1
	RS1.Close
	Set RS1= Nothing
%>
	<BR>
	<input type="hidden" Name='tmpDlgArg' value=''>
	<input type="hidden" Name='tmpDlgTxt' value=''>
	<FORM METHOD=POST ACTION="?act=SubmitDoc" onsubmit="return checkValidation();" >
	<TABLE class="RepTable" width='90%' align='center'>
	<TR>
		<TD dir='rtl' align='center'>
		<table width='100%'>
		<tr>
			<td style="border:none;width:60px;text-align:left;">‘„«—Â ”‰œ:</td>

			<td style="border:none;width:60px;">
				<INPUT NAME="GL" TYPE="Hidden" Value="<%=OpenGL%>">
				<INPUT NAME="DocNo" TYPE="text" Class="RepGeneralInput" Value="<%=DocNo%>"></td>
			<td style="border:none;width:*;"></td>
			<td style="border:none;width:60px;text-align:left;"> «—ÌŒ:</td>
			<td style="border:none;width:60px;">
				<INPUT NAME="DocDate" TYPE="text" Class="RepGeneralInput" Value="" onBlur="acceptDate(this);"></td>
		</tr>
		<tr>
			<td style="border:none;text-align:left;">Ã” ÃÊ:</td>
			<td colspan='4' style="border:none;width:60px;">
				<INPUT NAME="SearchBox" TYPE="Text" style="border:1 solid black;width:150px;" Value="" onKeyPress="return handleSearch();"></td>
		</tr>
		</table>
		</TD>
	</TR>
	</TABLE>

	<div align='center' id='errors'></div>

	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function documentKeyDown() {
		var theKey = window.event.keyCode;
		var obj = window.event.srcElement;
		if (theKey == 114) { 
			if (obj.name=="SearchBox"){document.all.SearchBox.select();};
			window.event.keyCode=0;
			document.all.SearchBox.focus();
			return false;
		}
	}

	document.onkeydown = documentKeyDown;

	function checkValidation(){
	  try{
		foundErr=false;
		totalItems=document.getElementsByName("Items").length
		var checkedCounter = 0;
		for (i=0;i<totalItems;i++){
			tmpRow=document.getElementsByName("Items")[i].parentNode.parentNode;
			if (tmpRow.getElementsByTagName("Input")[0].checked) {
				checkedCounter ++;
				gla=txt2val(tmpRow.getElementsByTagName("TD")[3].getElementsByTagName("INPUT")[0].value);
				if (gla==0){
					foundErr=true;
					errObj=tmpRow.getElementsByTagName("TD")[3].getElementsByTagName("INPUT")[0];
					errMsg="‘„«—Â Õ”«» €·ÿ «” "
					break;
				}

				rows=tmpRow.getElementsByTagName("TD")[0].rowSpan
				for (t=1;t<rows;t++){
					tmpRow=tmpRow.nextSibling;
					box=tmpRow.getElementsByTagName("TD")[1].getElementsByTagName("INPUT")[0];
					check(box);
					gla=txt2val(box.value);
					if (gla==0){
						foundErr=true;
						errObj=tmpRow.getElementsByTagName("TD")[1].getElementsByTagName("INPUT")[0];
						errMsg="‘„«—Â Õ”«» €·ÿ «” "
						break;
					}
				}
				if (foundErr) break;
			}
		}
		if (checkedCounter==0){
			foundErr=true;
			errObj=document.all.SearchBox;
			errMsg="ÂÌç ”ÿ—Ì «‰ Œ«» ‰‘œÂ «” ."
		}
		if (foundErr) {
			tmpCol=errObj.style.backgroundColor;
			errObj.style.backgroundColor="red";
			errObj.focus();
			alert("!Œÿ« \n\n"+errMsg);
			errObj.style.backgroundColor=tmpCol;
			return false;
		}
		docDate=document.all.DocDate;
		if (docDate.value!='' && docDate.value!=null){
			if (!acceptDate(docDate))
				return false;
		}
		else{
			docDate.focus();
			return false;
		}

	  }catch(e){
			//alert("Œÿ«Ì €Ì— „‰ Ÿ—Â");
			alert(e);
			return false;
	  }

	}
	function selectAll(src){
		totalItems=document.getElementsByName("Items").length
		checked=src.checked
		for (i=0;i<totalItems;i++)
			document.getElementsByName("Items")[i].checked=checked;
	}
	var lastFund = 0;
	var lastSrch = "";
	function handleSearch(){
		var theKey=event.keyCode;
		if (theKey==13){
			event.keyCode=0;
			srch=document.all.SearchBox.value;
			if (srch == '') {
				return;
			}
			if (srch!=lastSrch){
				lastFund = 0;
				lastSrch=srch;
			}
			var found = false;
			var text = document.body.createTextRange();
			found=text.findText(srch)
			for (var i=0; i<=lastFund && found ; i++) {
				found=text.findText(srch)
				text.moveStart("character", 1);
				text.moveEnd("textedit");
			}
			if (found) {
				text.moveStart("character", -1);
				text.findText(srch);
				text.select();
				lastFund++;
				theRow=text.parentElement();
				while(theRow.nodeName!='TR'){
					theRow=theRow.parentNode;
				}
//				alert(theRow.innerHTML);
				if(theRow.getElementsByTagName("input")[0]){
					if(theRow.getElementsByTagName("input")[0].type=='checkbox'){
						theRow.getElementsByTagName("input")[0].checked=true;
						selectRow(theRow);
					}
				}
				theRow.scrollIntoView();

			}
			else{
				if (lastFund == '0'){
					alert('⁄»«—  "' + srch +'" œ— «Ì‰ ’›ÕÂ ÅÌœ« ‰‘œ.');
				}
				else{
					alert('ÂÌç Ã«Ì œÌê—Ì ⁄»«—  "' + srch +'" ÅÌœ« ‰‘œ.');
				}
				lastFund=0;
				lastSrch="";
			}
		}
	}
	function selectRow(src){
		if (src.getElementsByTagName("Input")[0].checked) {
			src.getElementsByTagName("TD")[0].setAttribute("bgColor","yellow")
			src.getElementsByTagName("TD")[1].setAttribute("bgColor","yellow")
			rows=src.getElementsByTagName("TD")[0].rowSpan
			tmpRow=src;
			for (i=1;i<rows;i++){
				tmpRow=tmpRow.nextSibling;
				tmpRow.getElementsByTagName("TD")[1].setAttribute("bgColor","yellow");
				tmpRow.getElementsByTagName("TD")[2].setAttribute("bgColor","yellow");
			}
		}else{
			src.getElementsByTagName("TD")[0].setAttribute("bgColor","")
			src.getElementsByTagName("TD")[1].setAttribute("bgColor","")
			rows=src.getElementsByTagName("TD")[0].rowSpan
			tmpRow=src;
			for (i=1;i<rows;i++){
				tmpRow=tmpRow.nextSibling;
				tmpRow.getElementsByTagName("TD")[1].setAttribute("bgColor","");
				tmpRow.getElementsByTagName("TD")[2].setAttribute("bgColor","");
			}
		}
	}
	function check(src,writeResponse){
		badCode = false;
		if (window.XMLHttpRequest) {
			var objHTTP=new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			var objHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		}
		objHTTP.open('GET','xml_GLAccount.asp?id='+src.value,false)
		objHTTP.send()
		tmpStr = unescape(objHTTP.responseText)
		if (tmpStr == 'ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ')
			src.value='';
		if (writeResponse)
			src.parentNode.nextSibling.innerText=tmpStr;
	}
	function mask(src,writeResponse){

		var theKey=event.keyCode;
		if (theKey==32){
			event.keyCode=0;
			dialogActive=true;
			document.all.tmpDlgArg.value="#"
			document.all.tmpDlgTxt.value="Ã” ÃÊ œ— ‰«„ Õ”«» Â«Ì œ› — ﬂ·:"
			var myTinyWindow = window.showModalDialog('../dialog_GenInput.asp',document.all.tmpDlgTxt,'dialogHeight:200px; dialogWidth:440px; dialogTop:; dialogLeft:; edge:None; center:Yes; help:No; resizable:No; status:No;');
			dialogActive=false;
			if (document.all.tmpDlgTxt.value !="") {
				var myTinyWindow = window.showModalDialog('dialog_selectGL.asp?act=select&name='+escape(document.all.tmpDlgTxt.value),document.all.tmpDlgArg,'dialogHeight:500px; dialogWidth:380px; dialogTop:; dialogLeft:; edge:Raised; center:Yes; help:No; resizable:Yes; status:No;');
				if (document.all.tmpDlgArg.value!="#"){
					Arguments=document.all.tmpDlgArg.value.split("#")

					src.value=Arguments[0];
//					src.title=Arguments[1];
					if (Arguments[1] == 'ç‰Ì‰ Õ”«»Ì ÊÃÊœ ‰œ«—œ')
						src.value='';
					if (writeResponse)
						src.parentNode.nextSibling.innerText=Arguments[1];

				}
			}
		}
		else if (theKey >= 48 && theKey <= 57 ) // [0]-[9] are acceptible
			return true;
		else
			return false;

	}
	//-->
	</SCRIPT>
<%
	tempCounter = 0

	ItemLinks=sqlSafe(trim(replace(request.form("ItemLinks"),vbCrLf," ")))
	Do While right(ItemLinks ,1)=","
		ItemLinks=left(ItemLinks,len(ItemLinks)-1)
	Loop
	ItemLinks=replace(ItemLinks,"  "," ")
	ItemLinks=replace(ItemLinks,"  "," ")
	ItemLinks=replace(ItemLinks,", ",",")
	ItemLinks=replace(ItemLinks," ,",",")
	ItemLinks=replace(ItemLinks,",,",",")
	ItemLinks=replace(ItemLinks,",,",",")

	ItemLinksArray=split(ItemLinks,",")


	sysChequeGLAccountB = 	"17001"
	sysChequeGLAccountBName = "«”‰«œ œ—Ì«› ‰Ì (» )"
	sysChequeGLAccountA = 	"17002"
	sysChequeGLAccountAName = "«”‰«œœ—Ì«› ‰Ì («·› )"

	sysCashGLAccountB	=	"11005"
	sysCashGLAccountBName =	"’‰œÊﬁ »"
	sysCashGLAccountA = 	"11007"
	sysCashGLAccountAName = "’‰œÊﬁ «·›"
	'sysCashGLAccountA	=	"11005"
	'sysCashGLAccountAName =	"’‰œÊﬁ »"

	sysVatAccount = "49010"
	sysVatAccountName = "„«·Ì«  »— «—“‘ «›“ÊœÂ"

	sys="AR"
	sysName="›—Ê‘"
	sysDefaultGLAccount = "13003"%>
<!--#include file="include_MakeDocItemRows.asp" -->

<%	sys="AP"
	sysName="Œ—Ìœ"
	sysDefaultGLAccount = "41001"%>
<!--#include file="include_MakeDocItemRows.asp" -->

<%	sys="AO"
	sysName="”«Ì—"
	sysDefaultGLAccount = "18001"%>
<!--#include file="include_MakeDocItemRows.asp" -->

	</FORM>
<%
	someItemsRemained=False
	errorMessageText=""
	for i=0 to ubound(ItemLinksArray)
		if ItemLinksArray(i)<>0 then
			someItemsRemained=True
			errorMessageText=errorMessageText & "‘„«—Â "& ItemLinksArray(i) & " œ— ·Ì”  ‰Ì«„œÂ «” .<br>"
		end if
	next

	if someItemsRemained then
		%>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			document.all.errors.innerHTML='<br><div style="width:60%;font-size:10pt;background-color=yellow;border:2 solid black;"><b> ÊÃÂ!</b><br><%=errorMessageText%><br></div><br>';
		//-->
		</SCRIPT>
		<%
	end if
else
%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		function commaNumberMask(src){
			var theKey=event.keyCode;
			if (theKey==13 || theKey==9 || theKey==32 ){ // [Enter] or [Tab] or [Space] 
				return true;
			}
			else if (theKey >= 48 && theKey <= 57) { // [0-9]
				return true;
			}
			else if (theKey == 44 || theKey == 1608) { // comma [,] or [Ê] (the same key) 
				event.keyCode=44;
				return true;
			}
			return false;
		}
	//-->
	</SCRIPT>
	<BR><BR>
	<FORM METHOD=POST ACTION="?act=showItems">
	<TABLE width='500' align='center' style='border:1 solid black;'>
	<TR height='25'>
		<TD style="width:200px;text-align:left;"> ‰„«Ì‘ ¬Ì „ Â«Ì ”‰œ ‰ŒÊ—œÂ «“  «—ÌŒ:</TD>
		<TD style="width:60px;">
			<INPUT NAME="FromDate" TYPE="text" Class="RepGeneralInput" Value="" onBlur="acceptDate(this);"></TD>
		<TD style="width:40px;text-align:left;"> «  «—ÌŒ:</TD>
		<TD style="width:60px;">
			<INPUT NAME="ToDate" TYPE="text" Class="RepGeneralInput" Value="" onBlur="acceptDate(this);"></TD>
	</TR>
	<TR>
		<TD colspan='4' style="text-align:center;">
			<table ><tr>
			<% Set rs=Conn.Execute("SELECT * FROM AXItemTypes")
			Do While not rs.eof
%>
				<td><INPUT TYPE="checkbox" NAME="ItemTypes" Value="<%=rs("ID")%>"><%=rs("Name")%>
				</td>
<%
				rs.MoveNext
			Loop%>
			</tr></table>
		</TD>
	</TR>
	<TR>
		<TD colspan='4' style="text-align:center;">
		<span id='ItemNumbers'>
			<TEXTAREA style='width:100%;direction:LTR' NAME="ItemLinks" ROWS="3" onChange="document.getElementsByName('ItemTypes')[0].checked=true;" onkeypress="return commaNumberMask(this)"></TEXTAREA><br>
			<span style="FONT-SIZE:7pt">* ‘„«—Â Â« —« »« ﬂ«„« «“ Â„ Ãœ« ﬂ‰Ìœ („À«·: [ 81001,81002,81003 ] )</span>
		</span>
		</TD>
	</TR>
	<TR height='25'>
		<td colspan='4' style="border:none;text-align:center;">
			<INPUT TYPE="Submit" style="font-family:tahoma;border:1 solid black; width:50px;" Value="«œ«„Â"></td>
	</TR>
	</TABLE>
<%
end if
conn.Close
%>
<!--#include file="tah.asp" -->
