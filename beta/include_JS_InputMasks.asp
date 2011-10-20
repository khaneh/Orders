<%' This Include File Uses shamsiToday() ASP function 
  ' and it is needed to be included after *include_farsiDateHandling.asp*
  '
  ' to use this, add **onKeyPress='return maskDate(this);'** to the 
  '    releavant <input> tag (this is an example for Date Mask ####/##/##) 
  '	
  ' to use this, add **onBlur='acceptDate(this);'** to the 
  '    releavant <input> tag (this is an example for Date Acceptance YYYY/MM/DD) 
  '
  '	Version 2.0   1382-05-19
  ' The Masking Javascript fuctions provided are : acceptDate, maskDate , maskNumber , maskTime
  '
  '	Version 1.5   1381-09-10
  ' The Masking Javascript fuctions provided are : maskDate , maskNumber , maskTime
  '
  '	Version 1.0   1381-09-10
  ' The Masking Javascript fuctions provided are : maskDate , maskNumber
%>
<SCRIPT LANGUAGE="JavaScript">
<!-- 

function acceptDate(objCtl){ 
	
	var str = objCtl.value;
	if (str=="") return true;
	else if (str=="//") str = "<%=shamsitoday()%>";

	var myRegExp = new RegExp("^(13)?[7-9][0-9]/[0-1]?[0-9]/[0-3]?[0-9]$")
	if( myRegExp.test(str) ) {
		var SP = str.split("/");
		if (SP[0].length == 2) SP[0] = "13" + SP[0] ;
		if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
		if (SP[2].length == 1) SP[2] = "0"  + SP[2] ;
		str = SP.join("/")
	}
	if(!myRegExp.test(str)||( SP[0]<1376 || SP[1]>12 || SP[2]>31 )) {
		alert("YYYY/MM/DD : Œÿ«!  «—ÌŒ »«Ìœ »Â «Ì‰ ›—„  »«‘œ");
		objCtl.select();
		objCtl.focus();
		return false;
	};
	objCtl.value = str;
	return true;
}


function maskDate(objCtl){ 
	return true;
	if(document.selection.createRange().text != ''){
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
	};
    var len = objCtl.value.length; 
    var soutput = objCtl.value;
	var theKey=event.keyCode;
	var tempString;
	if (theKey < 47 || theKey > 57 || len > 9) { // 0-9 and / are acceptible
		return false;
	}
	else{
		if (theKey==47){
			if (len<4){
				soutput ='<%=Left(shamsiToday(),4)%>'
			}
			else if (len==5){
					soutput = soutput + '<%=Mid(shamsiToday(),6,2)%>'
			}
			else if (len==6){
					tempString=soutput.charAt(len-1)
					soutput = soutput.substring(0,len-1) + '0' + tempString
			}
			else if (len==8){
					soutput = soutput + '<%=Right(shamsiToday(),2)%>'
					event.keyCode=0
			}
			else if (len==9){
					tempString=soutput.charAt(len-1)
					soutput = soutput.substring(0,len-1) + '0' + tempString
					event.keyCode=0
			}
			else if (len!=4 && len!=7){
				event.keyCode=0
			}
		}
		else{
			if (len == 3){ 
				soutput = soutput + String.fromCharCode(theKey) ; 
				event.keyCode=47
			}
			else if (len == 6){ 
				soutput = soutput + String.fromCharCode(theKey) ; 
				event.keyCode=47
			}
		}
	}
objCtl.value = soutput 
objCtl.focus(); 
}
function maskTime(objCtl){ 
	if(document.selection.createRange().text != ''){
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
		document.execCommand('Delete');
	};
    var len = objCtl.value.length; 
    var soutput = objCtl.value;
	var theKey=event.keyCode;
	var tempString;
	if (theKey < 48 || theKey > 58 || len >= 5) { // 0-9 and / are acceptible
		return false;
	}
	else{
		if (theKey==58){
			if (len==0){
				soutput = '00'
				event.keyCode=58
			}
			else if (len==1){
				tempString=soutput.charAt(len-1)
				soutput = soutput.substring(0,len-1) + '0' + tempString
				event.keyCode=58
			}
			else if (len!=2){
				event.keyCode=0
			}
		}
		else{
			if (len == 1){ 
				soutput = soutput + String.fromCharCode(theKey) ; 
				event.keyCode=58
			}
		}
	}
objCtl.value = soutput 
objCtl.focus(); 
}
function maskNumber(objCtl){ 
	var theKey=event.keyCode;
	if (theKey==13){
		return true;
	}
	else if ((theKey < 48 || theKey > 57) && theKey!=46) { // 0-9 and [.] are acceptible
		return false;
	}
}

function val2txt(inp){
	inp=''+inp;
	s=''; // s is sign
	t = inp.search("[.]")
	if (t>0) 
		{
		 if ( inp.length < 2 ) 
			t2 = inp.length 
		 else	
			t2 = 2
		 expPart = inp.substr(t+1,t2)
		 inp = inp.substr(0,t)
		}
	if (inp.charAt(0)=='-'){
		s='-';
		inp=inp.substr(1,inp.length -1);
	}
	while(inp.charAt(0)=='0')
		inp=inp.substr(1,inp.length -1);
	if (inp=='')
		return 0;
	tmp=inp.substr(inp.length -3 ,3);
	inp=inp.substr(0,inp.length -3);
	result = tmp;
	while(inp.length > 0){
		tmp=inp.substr(inp.length -3 ,3)
		inp=inp.substr(0,inp.length -3)
		result = tmp +","+ result
	}
	if (t>0) 
	{
	 result = result + "." + expPart
	}
	return (s+result);
}

function txt2val(inp){
	inp=''+inp
	while(inp.charAt(0)=='0')
		inp=inp.substr(1,inp.length -1);
	if (inp=='' || inp=='null')
		return 0;
	result=inp.replace(',','');
	while (result.search(",")>0){
		result=result.replace(',','')
	}
	result=parseFloat(result);
	if(''+result=="NaN")
		result=0;
	return result;
}

//-->
</SCRIPT>