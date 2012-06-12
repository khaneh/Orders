function disGroup (e){
	var groupName=$(e).parent(".mySection").attr("groupname");
	if (e.checked) {
		$(e).parent(".mySection").children('[name^="'+groupName+'"]').prop("disabled", false);
	} else {
		$(e).parent(".mySection").children('[name^="'+groupName+'"]').prop("disabled", true);
		$(e).parent(".mySection").children('[name$="disBtn"]').prop("disabled", false);
	}
}

$(document).ready(function () {
	$('[name$=-addValue]').hide();
	$('input[name$=-price]').change(function() {
		calc_total();
	});
});
function getNum(n){
	return parseInt(n.replace(/,/gi,''));
}
function echoNum(str){
	var regex = /(-?[0-9]+)([0-9]{3})/;
    str += '';
    while (regex.test(str)) {
        str = str.replace(regex, '$1,$2');
    }
    //str += ' kr';
    return str;
}
function calc_total(){
	var totalPrice = 0;
	$('input[name$=-price]').each(function(i){
		if ($.isNumeric(getNum($(this).val())))
			totalPrice += getNum($(this).val());
	});
	$('#totalPrice').val(echoNum(totalPrice));
}
function cloneRow(key){
	var maxID = $("#" + key.replace(/\//gi,"-") + "-maxID").val();
	var newRow = $("#" + key.replace(/\//gi,"-") + "-"+maxID).clone().attr('id', key.replace(/\//gi,"-") + "-" + (parseInt(maxID)+1));
	$('input:checkbox',newRow).each(function (){
		if ($(this).val().substr(0,3)=='on-')
			$(this).val('on-'+(parseInt(maxID)+1));
	});
	$("input:checkbox[name$=-disBtn]",newRow).each(function (){
		$(this).val(parseInt($(this).val()) + 1);
	});
	newRow.appendTo("#extreArea" + key.replace(/\//gi,"-") );
	$("#" + key.replace(/\//gi,"-") + "-maxID").val(parseInt(maxID)+1);
}
function removeRow(key){
	var maxID = parseInt($("#" + key.replace(/\//gi,"-") + "-maxID").val());
	if (maxID>0) {
		$("#" + key.replace(/\//gi,"-") + "-"+maxID).remove();
	}
	$("#" + key.replace(/\//gi,"-") + "-maxID").val(maxID-1);
}
function checkOther(e){
	if ($(e).val()==-1 || $(e).val().substr(0,6)=="other:") {
		if ($(e).find("option:selected").text()=="ساير") {
			$(e).next().val("مقدار را وارد كنيد");
		} else {
			$(e).next().val(
				$(e).find("option:selected").text());
		}
		$(e).next().show();
		$(e).next().focus();
	} else {
		$(e).next().hide();
	}
}
function addOther(e){
	if ($(e).val()!="" && $(e).val()!="مقدار را وارد كنيد" && $(e).val()!="نمي‌شه كه خالي باشه!"){
		$(e).val($(e).val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
		$(e).prev().find("option:selected").text($(e).val());
		$(e).prev().find("option:selected").val('other:'+$(e).val());
		$(e).hide();
		if ($(e).parent().attr("groupname")=="paper") calc_paper(e);
		if ($(e).parent().attr("groupname")=="selefon") calc_selefon(e);
	} else {
		$(e).val("نمي‌شه كه خالي باشه!");
		$(e).focus();
	}
}
function calc_plate(e){
	if ($(e).attr("name")!="plate-price" || ($(e).val()=="" && $(e).attr("name")=="plate-price")){
		var myGroup = $(e).parent();
		var color = parseInt(myGroup.children("input[name=plate-color-count]:first").val());
		var qtty = getNum(myGroup.children("input[name=plate-qtty]:first").val());
		var price = parseInt(myGroup.children("select[name=plate-type]:first").children(":selected").attr("price"));
		myGroup.children("input[name=plate-qtty]:first").val(echoNum(qtty));
		myGroup.children("input[name=plate-price]").val(echoNum(color*qtty*price));
		//------------- Related Event -------------
		myGroup.parent().children("div[groupname=proof]").children("[name=proof-dammy]").click();
		if ($(e).attr("name")=="plate-qtty"){
			myGroup.parent().children("div[groupname=print]").children("input[name=machin-form]:first").val(qtty);
			myGroup.parent().children("div[groupname=verni]").children("input[name=verni-form]:first").val(qtty);
			myGroup.parent().children("div[groupname=uv]").children("input[name=uv-form]:first").val(qtty);
		} else if ($(e).attr("name")=="plate-type"){
			var type = parseInt(myGroup.children("select[name=plate-type]:first").children(":selected").val());
			var size="35X50";
			if (type==1)
				size = "35X50";
			else if (type==2)
				size = "35X50";
			else if (type==3)
				size = "50X70";
			else if (type==5)
				size = "100X70";
			myGroup.parent().children("div[groupname=paper]").children("input[name=after-cut_size]:first").val(size);
			if (type!=4){
				if (type==3)
					type=4;
				else if (type==5)
					type=3;
				myGroup.parent().children("div[groupname=print]").children("select[name=machin]:first").val(type);
				//console.log(type);
			}
		}
	} else{
		$(e).val(echoNum($(e).val()));
	}
	
	calc_addedPaper(e);
	calc_total();
}
function calc_print(e){
	if ($(e).attr("name")!="machin-price" || ($(e).val()=="" && $(e).attr("name")=="machin-price")){
		var myGroup = $(e).parent();			
		var color = parseInt(myGroup.parent().children("div[groupname=plate]:first").children("input[name=plate-color-count]:first").val());
		var spcolor = parseInt(myGroup.children("input[name=machin-sp-color]:first").val());
		var form = parseInt(myGroup.children("input[name=machin-form]:first").val());
		var qtty = getNum(myGroup.children("input[name=machin-qtty]:first").val());
		var price = parseInt(myGroup.children("select[name=machin]:first").children(":selected").attr("price"));
		var dup = parseInt(myGroup.children("select[name=machin-d-type]:first").val());
		var ac_qtty = form * qtty;
		//console.log("print!");
		//console.log(dup+ " ac: "+ac_qtty);
		if (dup==2)
			ac_qtty = ac_qtty / 2;
		myGroup.children("input[name=machin-qtty]:first").val(echoNum(qtty));
		myGroup.parent().children("div[groupname=paper]:first").children("input[name=after-cut_qtty]").val(ac_qtty);
		if (qtty<5000) 
			qtty=5000;
		myGroup.children("input[name=machin-price]").val(echoNum(color*form*qtty*price + 1.5*spcolor*form*qtty*price));
	} else {
		$(e).val(echoNum($(e).val()));
	}
	//------------- Related Event -------------
	if ($(e).attr("name")=="machin-form" || $(e).attr("name")=="machin-qtty"){
		if (myGroup.parent().children("div[groupname=verni]").children("[name=verni-disBtn]").is(":checked"))
			myGroup.parent().children("div[groupname=verni]").children("[name=verni-disBtn]").click();
		if (myGroup.parent().children("div[groupname=selefon]").children("[name=selefon-disBtn]").is(":checked"))
			myGroup.parent().children("div[groupname=selefon]").children("[name=selefon-disBtn]").click();
		if (myGroup.parent().children("div[groupname=uv]").children("[name=uv-disBtn]").is(":checked"))
			myGroup.parent().children("div[groupname=uv]").children("[name=uv-disBtn]").click();
		
	}
	calc_addedPaper(e);
	calc_total();
}
function calc_binding(e){
	if ($(e).attr("name")!="binding-price" || ($(e).val()=="" && $(e).attr("name")=="binding-price")){
		var myGroup = $(e).parent();
		//console.log(myGroup.children("input[name=binding-qtty]:first").val());
		var qtty = getNum(myGroup.children("input[name=binding-qtty]:first").val());
		var form = parseInt(myGroup.children("input[name=binding-form]:first").val());
		var type = parseInt(myGroup.children("select[name=binding-type]:first").children(":selected").val());
		var price = parseInt(myGroup.children("select[name=binding-size]:first").children(":selected").attr("price").split(",")[type-1]);
		var orient = parseInt(myGroup.children("input[name=binding-orient]:checked").val());
		myGroup.children("input[name=binding-qtty]:first").val(echoNum(qtty));
		if (type==1){
			if (form < 10) 
				form = 10;
			if (qtty < 2000) 
				qtty = 2000;
		} else if (type==2){
			if (form < 3) 
				form = 3;
			if (qtty < 2000)
				qtty = 2000;
		}
		var result = price*form*qtty;
		if (orient==2 || orient==3)
			result += .2*result;
		if (myGroup.children("[name=binding-disBtn]").is(":checked"))
			myGroup.children("input[name=binding-price]").val(echoNum(result));
		else 
			myGroup.children("input[name=binding-price]").val("");
	} else {
		$(e).val(echoNum($(e).val()));
	}
	//calc_addedPaper(e);
	calc_total();
}
function calc_paper(e){
	var myGroup = $(e).parent();
	var afterCut = myGroup.children("[name=after-cut_size]");
	var beforCut = myGroup.children("[name=paper-size]:first").children(":selected");
	afterCut.val(afterCut.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
	var aw=parseFloat(afterCut.val().split("X")[0]);
	var ah=parseFloat(afterCut.val().split("X")[1]);
	var bw=parseFloat(beforCut.text().split("X")[0]);
	var bh=parseFloat(beforCut.text().split("X")[1]);
	if ((bh*bw)<(ah*aw)){
		rate = 1;
		afterCut.val(beforCut.text());
	} else {
		var rate = parseInt(bh * bw / (ah * aw));
	}
	//console.log(bw +" X " + bh);
	myGroup.children("input[name=paper-in-form]").val(rate);
	if ($(e).attr("name")!="paper-price" || ($(e).val()=="" && $(e).attr("name")=="paper-price")){
		
		var weight = parseInt(myGroup.children("[name=paper-weight]").children(":selected").text());
		var price = parseInt(myGroup.children("[name=paper-type]").children(":selected").attr("price"));
		var qtty = getNum(myGroup.children("[name=paper-qtty]").val());
		var result = bh*bw/10000*weight/1000*price*qtty;
		myGroup.children("[name=paper-qtty]").val(echoNum(qtty));
		if (myGroup.children("[name=paper-customer]").is(":checked"))
			myGroup.children("input[name=paper-price]").val("0");
		else
			myGroup.children("input[name=paper-price]").val(echoNum(Math.round(result)));
	}
	//------------- Related Event -------------
	if ($(e).attr("name")=="after-cut_size"){
		if (myGroup.parent().children("div[groupname=selefon]").children("[name=selefon-disBtn]").is(":checked"))
			myGroup.parent().children("div[groupname=selefon]").children("[name=selefon-disBtn]").click();
	} 
	calc_total();
}
function calc_snap(e){
	if ($(e).attr("name")=="snap-size"){
		$(e).val($(e).val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
	} else if ($(e).attr("name")=="snap-price"){
		$(e).val(echoNum($(e).val()))
	}
	calc_addedPaper(e);
	calc_total();
}
function calc_addedPaper(e){
	var thisRow = $(e).parent().parent();//.children(".exteraArea");
	//console.log(thisRow);
	var myPrint = thisRow.children("div[groupname=print]");
	var qtty = getNum(myPrint.children("input[name=machin-qtty]:first").val());
	var form = parseInt(myPrint.children("input[name=machin-form]:first").val());
	var color = parseInt(thisRow.children("div[groupname=plate]:first").children("input[name=plate-color-count]").val());
	var spColor= parseInt(myPrint.children("input[name=machin-sp-color]:first").val());
	var step = color + spColor + thisRow.children("div[groupname=verni]").children("[name=verni-disBtn]").is(':checked') +
		thisRow.children("div[groupname=uv]").children("[name=uv-disBtn]").is(':checked') +
		thisRow.children("div[groupname=selefon]").children("[name=selefon-disBtn]").is(':checked') +
		thisRow.children("div[groupname=fold]").children("[name=fold-disBtn]").is(':checked') +
		thisRow.children("div[groupname=snap]").children("[name=snap-disBtn]").is(':checked');
	step += $("input[name=binding-disBtn]:checked").size();
	//console.log("s: " + step+" q: "+qtty+" f: "+form);
	var step5 = false;
	var step10= false;
	var result= 0;
	
	for (remQtty = qtty; remQtty>0;){
		//console.log("rem: " + remQtty);
		if (!step5){
			step5 = true;
			if (remQtty>5000){
				remQtty -= 5000;
				result = 5 / 1000 * step * 5000 * form;
			} else {
				if (remQtty<5000) 
					remQtty = 5000;
				result = 5 / 1000 * step * remQtty * form;
				remQtty = 0;
			}
		} else if (!step10){
			step10 = true;
			if (remQtty>5000){
				remQtty -= 5000;
				result += 3 / 1000 * step * 5000 * form;
			} else {
			
				result += 3 / 1000 * step * qtty * form;
			}
		} else {
			result += 1 / 1000 * step * remQtty * form;
			remQtty = 0;
		}
	}
	result = Math.ceil(result); 
	thisRow.children("div[groupname=paper]:first").children("input[name=wastepaper]:first").val(echoNum(result));
	var rate = parseInt(thisRow.children("div[groupname=paper]:first").children("input[name=paper-in-form]:first").val());
	var paperAfter = parseInt(thisRow.children("div[groupname=paper]:first").children("input[name=after-cut_qtty]:first").val());
	//console.log("r: "+rate + " after: "+ paperAfter);
	paperQtty = Math.ceil((result + paperAfter)/rate);
	thisRow.children("div[groupname=paper]:first").children("input[name=paper-qtty]:first").val(echoNum(paperQtty));
	//calc_paper(thisRow.children("div[groupname=paper]:first").children("input[name=paper-in-form]:first"));
	thisRow.children("div[groupname=paper]").children("[name=paper-in-form]").blur();
}
/*

محاسبه ميزان باطله

تا 5000 نسخه
تيراژ هر فرم *تعداد مرحله * 5 / 1000
از 5000 تا 10000 نسخه
تيراژ- 5000) هر فرم *تعداد مرحله * 3   / 1000
از10000 نسخه به بالا
تيراژ- 10000) هر فرم *تعداد مرحله * 1  / 1000

*/		

function calc_verni(e){
	if ($(e).attr("name")!="verni-price" || ($(e).val()=="" && $(e).attr("name")=="verni-price")){
		var myGroup = $(e).parent();
		var myPrint = myGroup.parent().children("div[groupname=print]");
		var qtty = getNum(myPrint.children("input[name=machin-qtty]:first").val());
		var form = parseInt(myGroup.children("input[name=verni-form]:first").val());
		var verni_wat = parseInt(myGroup.children("select[name=verni-wat]:first").children(":selected").val());
		var price = parseInt(myGroup.children("select[name=verni-mat]:first").children(":selected").attr("price").split(",")[verni_wat-1]);
		//console.log("qtty: "+qtty+" form: "+form+" price: "+price);
		if (myGroup.children("[name=verni-disBtn]").is(":checked")){
			myGroup.children("input[name=verni-price]").val(echoNum(qtty*form*price));
		} else {
			myGroup.children("input[name=verni-price]").val("");
		}
	} else {
		$(e).val(echoNum($(e).val()));
	}
	calc_addedPaper(e);
	calc_total();
}
function calc_selefon(e){
	if ($(e).attr("name")!="selefon-price" || ($(e).val()=="" && $(e).attr("name")=="selefon-price")) {
		var myGroup = $(e).parent();
		var paperGroup = myGroup.parent().children("div[groupname=paper]");
		//var afterCut = paperGroup.children("[name=after-cut_size]");
		//afterCut.val(afterCut.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
		var w=parseFloat(myGroup.children("[name=selefon-size]:first").children(":selected").text().split("X")[0]);
		var h=parseFloat(myGroup.children("[name=selefon-size]:first").children(":selected").text().split("X")[1]);
		if (h>w){
			var t=w;
			w=h;
			h=t;
		}
		if (h<35 || w<50){
			h=35;
			w=50;
		}
		var face = parseInt(myGroup.children("[name=selefon-face]:first").children(":selected").val());
		var mat = parseInt(myGroup.children("[name=selefon-mat]:first").children(":selected").val());
		var price = parseFloat(myGroup.children("[name=selefon-hot]:first").children(":selected").attr("price").split(",")[mat-1]);
		var myPrint = myGroup.parent().children("div[groupname=print]");
		var qtty = getNum(myPrint.children("input[name=machin-qtty]:first").val());
		var form = parseInt(myPrint.children("input[name=machin-form]:first").val());
		if (myPrint.children("select[name=machin-d-type]:first").children(":selected").val()==2)
			form = form / 2;
		//console.log("w: "+w+" h: "+h+" price: "+price+" mat: "+mat+" face: "+face+" from "+form+" qtty "+qtty);
		if (myGroup.children("[name=selefon-disBtn]").is(":checked")){
				myGroup.children("input[name=selefon-price]").val(echoNum(face*price*h*w*form*qtty));
			} else {
				myGroup.children("input[name=selefon-price]").val("");
			}
	} else {
		$(e).val(echoNum($(e).val()));
	}
	calc_total();
	calc_addedPaper(e);
}
function calc_uv(e){
	/*
if ($(e).attr("name")=="uv-size")
		$(e).val($(e).val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
*/
	if ($(e).attr("name")=="uv-price"){
		$(e).val(echoNum($(e).val()));
	} else {
		var myGroup = $(e).parent();
		var myPrint = myGroup.parent().children("div[groupname=print]");
		var qtty = getNum(myPrint.children("input[name=machin-qtty]:first").val());
		var form = parseInt(myGroup.children("input[name=uv-form]:first").val());
		var mat = parseInt(myGroup.children("[name=uv-mat]:first").children(":selected").val()) - 1;
		var spot = parseInt(myGroup.children("[name=uv-spot]:first").children(":selected").val()) - 1;
		var price = parseInt(myGroup.children("[name=uv-size]").children(":selected").attr("price").split(",")[spot].split("|")[mat]);
		if (qtty<1000)
			qtty = 1000;
		//console.log("qtty: "+qtty+" form: "+form+" price: "+price);
		myGroup.children("input[name=uv-price]:first").val(echoNum(qtty * form * price));
	}
	calc_addedPaper(e);
	calc_total();
}
function calc_fold(e){
	if ($(e).attr("name")=="fold-price"){
		$(e).val(echoNum($(e).val()));
		calc_total();
	}
	calc_addedPaper(e);
}
function calc_packing(e){
	if ($(e).attr("name")=="packing-price"){
		$(e).val(echoNum($(e).val()));
		calc_total();
	}
}
function calc_service(e){
	if ($(e).attr("name")=="service-price"){
		$(e).val(echoNum($(e).val()));
		calc_total();
	}
}
function calc_proof(e){
	if ($(e).attr("name")!="proof-price" || ($(e).val()=="" && $(e).attr("name")=="proof-price")){
		var myGroup = $(e).parent();
		var paperGroup = myGroup.parent().children("div[groupname=paper]");
		var afterCut = paperGroup.children("[name=after-cut_size]");
		afterCut.val(afterCut.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
		var w=parseFloat(afterCut.val().split("X")[0]);
		var h=parseFloat(afterCut.val().split("X")[1]);
		var qtty = parseInt(myGroup.parent().children("div[groupname=plate]").children("input[name=plate-qtty]").val());
		var result = 0;
		if (myGroup.children("[name=proof-dammy]").is(":checked"))
			result += Math.round(parseInt(myGroup.children("[name=proof-dammy]").attr("price")) * w * h * qtty);
		if (myGroup.children("[name=proof_digital]").is(":checked"))
			result += parseInt(myGroup.children("[name=proof_digital]").attr("price")) * qtty;
		if (myGroup.children("[name=proof_iso]").is(":checked"))
			result += Math.round(parseInt(myGroup.children("[name=proof_iso]").attr("price")) * w * h * qtty);
		
		myGroup.children("input[name=proof-price]:first").val(echoNum(result));
	} else {
		$(e).val(echoNum($(e).val()));
	}
	if ($(e).attr("name")=="proof-date"){
		if ($(e).val()=="//") {
			var today = new Date();
			$(e).val($.format.date(today,"yyyy/MM/dd"));
		} else {
			var rege=/^(13)?[7-9][0-9]\/[0-1]?[0-9]\/[0-3]?[0-9]$/;
			if( rege.test($(e).val()) ) {
				var SP = $(e).val().split("/");
				if (SP[0].length == 2) SP[0] = "13" + SP[0] ;
				if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
				if (SP[2].length == 1) SP[2] = "0"  + SP[2] ;
				$(e).val(SP.join("/"));	
			}
			if(!rege.test($(e).val())||( SP[0]<'1376' || SP[1]>'12' || SP[2]>'31' )) {
				$(e).val("خطا!");
				$(e).focus();
			}
		}
	}
	calc_total();
}
function calc_delivery(e){
	if ($(e).attr("name")=="delivery-price"){
		$(e).val(echoNum($(e).val()));
		calc_total();
	}
}
function calc_cutting(e){
	if ($(e).attr("name")=="cutting-price"){
		$(e).val(echoNum($(e).val()));
		calc_total();
	}
}
