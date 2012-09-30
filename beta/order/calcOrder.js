$(document).ready(function(){
	$.ajaxSetup({
		cache: false
	});
});

function makeOutXML(){
	if ($(".forceErr:not(:disabled)").size()>0){
		$("#errMsg").html($(".forceErr:not(:disabled)").size() + " ›Ì·œ «Ã»«—Ì Å— ‰‘œÂù«‰œ°ù «» œ« ¬‰Â« —« Å— ﬂ‰Ìœ");
		return false;
	}
	else {
		try {
			var out=$("<data></data>");
			var id=0
			$(".myRow").each(function(i,myRow){
				id=0;
				$(myRow).find(".exteraArea").each(function (i,rowArea){
					
					var rowName=$(rowArea).attr("type");
					var thisRow = $("<row name='" + rowName + "' id='" + id + "'></row>");
					$(rowArea).find("[out=yes]").each(function (i,key){
						if ($(key).val()!="" && !$(key).attr("disabled") && ($(key).closest("div.group").find("[name$=-disBtn]").is(":checked") || isNaN($(key).closest("div.group").find("[name$=-disBtn]").val()) ))
							switch ($(key).attr("type")) {
								case 'checkbox':
									if ($(key).is(":checked"))
										thisRow.append("<key name='" + $(key).attr("name") + "'>" + $(key).val() + "</key>");
									break;	
								case 'radio':	
									if ($(key).is(":checked")){
										var thisName = $(key).attr("name").split("*");
										thisRow.append("<key name='" + thisName[0] + "'>" + $(key).val() + "</key>");
									}
									break;
								default:
									thisRow.append("<key name='" + $(key).attr("name") + "'>" + $(key).val() + "</key>");
							}
					});
					out.append(thisRow);
					id++;
				});
			});
			$("#outXML").val('<data>'+out.html()+'</data>');
/* 			console.log('<data>'+out.html()+'</data>'); */
			$("#returnDateTime").val($.jalaliCalendar.jalaliToGregorianStr($("input[name=ReturnDate]").val()) + " " + $("input[name=ReturnTime]").val());
			return true;
		} 
		catch (e){
			alert("Œÿ«Ì ⁄ÃÌ»!");
			 return false;
		}
	}

}
function acceptPaper(o){
 var obj=$(o);
 obj.val(obj.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
}
function acceptDate(o){
	$("#errMsg").html("");
	var obj=$(o); 
	if (obj.val()=="") {
		$("#errMsg").html("·ÿ›«  «—ÌŒ —« Ê«—œ ﬂ‰Ìœ");
		obj.focus();
	}
	else if (obj.val()=="//") {
		var today = new Date();
		obj.val($.format.date(today,"yyyy/MM/dd"));
	} else {
		var rege=/^(13)?[7-9][0-9]\/[0-1]?[0-9]\/[0-3]?[0-9]$/;
		if( rege.test(obj.val()) ) {
			var SP = obj.val().split("/");
			if (SP[0].length == 2) SP[0] = "13" + SP[0] ;
			if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
			if (SP[2].length == 1) SP[2] = "0"  + SP[2] ;
			obj.val(SP.join("/"));	
		}
		if(!rege.test(obj.val())||( SP[0]<'1376' || SP[1]>'12' || SP[2]>'31' )) {
			$("#errMsg").html("›—„   «—ÌŒ »«Ìœ YYYY/MM/DD »«‘œ.");
			obj.focus();
		};
	}
	
}
function acceptTime(o) {
	$("#errMsg").html("");
	var obj=$(o);
	if (obj.val()=="") {
		$("#errMsg").html("·ÿ›« ”«⁄  —« Ê«—œ ﬂ‰Ìœ.");
		obj.focus();
	} else if (obj.val()==":"){
		var now = new Date();
		obj.val(now.getHours()+':'+now.getMinutes());
		var rege = /^[0-2]?[0-9]:[0-5]?[0-9]$/;
		if( rege.test(obj.val()) ) {
			var SP = obj.val().split(":");
			if (SP[0].length == 1) SP[0] = "0" + SP[0] ;
			if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
			obj.val(SP.join(":"));
		}
	} else {
		var rege = /^[0-2]?[0-9]:[0-5]?[0-9]$/;
		if( rege.test(obj.val()) ) {
			var SP = obj.val().split(":");
			if (SP[0].length == 1) SP[0] = "0" + SP[0] ;
			if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
			obj.val(SP.join(":"));
		}
		if(!rege.test(obj.val())||( SP[0]>23 || SP[1]>59)) {
			$("#errMsg").html("›—„  ”«⁄  »«Ìœ HH:MM »«‘œ.");
			obj.focus();
		};
	}
}
function disGroup (e){
	if ($(e).hasClass("btn")) {
		// Group is visible
		var groupName = $(e).attr("name");
		var myRow = $(e).closest(".exteraArea");
		var myGroup = myRow.find('div.group[groupname=' + groupName + ']')
		myGroup.find('[name^="'+groupName+'"]').prop("disabled", false);
		myGroup.find('input.disBtn').prop("checked",true);
		myGroup.css("display","block");
		$(e).remove();
	} else {
		// Group is invisible
		var myGroup = $(e).closest("div.group");
		var myRow = myGroup.closest(".exteraArea");
		var groupName = myGroup.attr("groupname");			
		myGroup.find('[name^="'+groupName+'"]').prop("disabled", true);
		myGroup.find('[name$="disBtn"]').prop("disabled", false);
		myGroup.css("display","none");
		myRow.find(".unusedGroup").append('<label onclick="disGroup(this);" class="btn btn-inverse offset0" name="' + groupName + '">' + myGroup.find(".groupName").html() + '</label>');
	}
}

function readyForm() {
	$('[name$=-disBtn]:not(:checked)').each(function(i){
		disGroup($(this));
	});
	$('[name$=-addValue]').hide();
	$("[name$=-price]").addClass("price");
	$("[name$=-dis]").addClass("dis");
	$("[name$=-over]").addClass("over");
	$('input[name$=-price]').change(function() {
		calc_total();
	});
	$("[force=yes]:not(:disabled)").each(function(i){
		if ($(this).val()!="") 
			$(this).removeClass("forceErr");
		else
			$(this).addClass("forceErr");
	});
	$.getJSON("/service/json_getService.asp",
		{act:"getList"},
		function (json){
			var thisItem;
			var myAppend;
			$("select[name=service-item]").each(function(j){
				thisItem = $(this);
				$.each(json, function(i,s){
					myAppend = "<option ";
					if (s.id == thisItem.attr("thisValue"))
						myAppend += "selected ";
					myAppend +="value="+s.id+">"+s.name+"</option>"
					thisItem.append(myAppend);
				});
			});
	});
	$('.funcBtn p').tooltip({
      selector: "a[rel=tooltip]"
    });
	$('div.key-label').tooltip({
      selector: $(this)
    });
	$('div.priceGroupValue').tooltip({
      selector: "input[rel=tooltip]"
    });
    
    
}
function checkForce(e){
	var myGroup = $(e).closest("div.group");
//	console.log('force checking');
	$(myGroup).find("[force=yes]").each(function (i,obj){
		if ($(obj).val()=="" && !$(obj).attr("disabled")) {
			$(obj).addClass("forceErr");
		} else {
			$(obj).removeClass("forceErr");
		}
	});
}
function clearThis(e){
	//console.log('clear');
	$(e).val("");
}
function getNum(n){
	var out="0";
	//if (!isNaN(n))
		out = parseInt(n.replace(/,/gi,''));
	return out;
}
function echoNum(str){
	var regex = /(-?[0-9]+)([0-9]{3})/;
	str = Math.floor(str);
    str += '';
    while (regex.test(str)) {
        str = str.replace(regex, '$1,$2');
    }
    //str += ' kr';
    return str;
}

function getPer(n){
	var n = parseInt(n.replace(/%/gi,''));
	if (n>100) n=100;
	if (n<0) n=0;
	if (isNaN(n)) n=0;
	return n;
}
function echoPer(str){
	var n = parseInt(str);
	return n + '%';
}
function calc_total(){
	var totalPrice = 0;
	var totalVatedPrice = 0;
	$('input[name$=-price]:not(:disabled)').each(function(i){
		if ($.isNumeric(getNum($(this).val())))
			totalPrice += getNum($(this).val());
			if ($(this).closest("div.group").attr("groupname")!="paper")
				totalVatedPrice += getNum($(this).val()) * (100 + parseFloat($("#vatRate").val())) / 100;
	});
	$('#totalPrice').val(echoNum(totalPrice));
	$('#totalVatedPrice').val(echoNum(totalVatedPrice));
}
function cloneRow(e){
	var myRow = $(e).closest(".exteraArea");
	
	var newRow = myRow.clone();
/*
	$('input:checkbox',newRow).each(function (){
		if ($(this).val().substr(0,3)=='on-')
			$(this).val('on-'+(parseInt(maxID)+1));
	});
	$("input:checkbox[name$=-disBtn]",newRow).each(function (){
		$(this).val(parseInt($(this).val()) + 1);
	});
*/
	newRow.appendTo(myRow.closest(".myRow"));
/* 	$("#" + key + "-maxID").val(parseInt(maxID)+1); */
}
function removeRow(e){
	if (confirm('„ÿ„∆‰ Â” Ìœ ﬂÂ «Ì‰ ”ÿ— Õ–› ‘Êœø'))
	var myRow = $(e).closest(".exteraArea");
	var parentRow = myRow.closest(".myRow");
	if (parentRow.find(".exteraArea").size()!=1){
		myRow.remove();
	}
/*
	var maxID = parseInt($("#" + key + "-maxID").val());
	if (maxID>0) {
		$("#" + key + "-"+maxID).remove();
	}
	$("#" + key + "-maxID").val(maxID-1);
*/
}
function checkOther(e){
	if ($(e).val()==-1 || $(e).val().substr(0,6)=="other:") {
		if ($(e).find("option:selected").text()=="”«Ì—") {
			$(e).next().val("„ﬁœ«— —« Ê«—œ ﬂ‰Ìœ");
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
	if ($(e).val()!="" && $(e).val()!="„ﬁœ«— —« Ê«—œ ﬂ‰Ìœ" && $(e).val()!="‰„Ìù‘Â ﬂÂ Œ«·Ì »«‘Â!"){
		var err=false;
		switch ($(e).prev().attr("checkOther")){
			case "num":
				var v = parseFloat($(e).val());
				if (isNaN(v))
					err = true;
				else
					$(e).val(v);
				break;
			case "size":
				$(e).val($(e).val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
				var w = parseFloat($(e).val().split("X")[0]);
				var h = parseFloat($(e).val().split("X")[1]);
				if (isNaN(h) || isNaN(w))
					err = true;
				else
					$(e).val(w + 'X' + h);
				break;
		}
		if (!err) {
			$(e).prev().find("option:selected").text($(e).val());
			$(e).prev().find("option:selected").val('other:'+$(e).val());
			$(e).hide();
			if ($(e).parent().attr("groupname")=="paper") calc_paper(e);
			if ($(e).parent().attr("groupname")=="selefon") calc_selefon(e);
		} else {
			$(e).val("„ﬁœ«— —« Ê«—œ ﬂ‰Ìœ");
			$(e).focus();
		}
	} else {
		$(e).val("‰„Ìù‘Â ﬂÂ Œ«·Ì »«‘Â!");
		$(e).focus();
	}
}

function calc_desc(e){
	checkForce(e);
}

function calc_plate(e){
	var myGroup = $(e).closest("div.group");
	var type = parseInt(myGroup.find("select[name=plate-type]:first option:selected").val());
	var size="35X50";
	if (type==1) 
		size = "35X50";
	else if (type==2)
		size = "35X50";
	else if (type==3)
		size = "50X70";
	else if (type==5)
		size = "100X70";
	myGroup.parent().find("input[name=paper-after-cut_size]:first").val(size);
	if (type!=4){
		if (type==3)
			type=4;
		else if (type==5)
			type=3;
		myGroup.parent().find("select[name=print-type]:first").val(type);
		//console.log(type);
	}
	switch ($(e).attr("name")){
		case 'plate-qtty':
			var qtty = getNum(myGroup.find("input[name=plate-qtty]:first").val());
			myGroup.parent().find("input[name=print-form]:first").val(qtty);
			myGroup.parent().find("input[name=verni-form]:first").val(qtty);
			myGroup.parent().find("input[name=uv-form]:first").val(qtty);
		case 'plate-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		case 'plate-over':
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
		default:
			var customer = 1;
			var color = parseInt(myGroup.find("input[name=plate-color-count]:first").val());
			var qtty = getNum(myGroup.find("input[name=plate-qtty]:first").val());
			var price = parseInt(myGroup.find("select[name=plate-type]:first").children(":selected").attr("price"));
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			if (myGroup.find("input[name=plate-customer]:first").is(":checked")){
				customer = 0;
				over = 0;
				dis = 0;
			}
			var stockName = myGroup.find("span.groupName").html();
			stockName += " ";
			stockName += myGroup.find("select[name=plate-type] option:selected").text();
			if (size.indexOf('X')>0){
				myGroup.find("input[name$=-l]").val(size.split("X")[0]);
				myGroup.find("input[name$=-w]").val(size.split("X")[1]);
			}
			myGroup.find("input[name$=-stockName]").val(stockName);
			myGroup.find("input[name$=-stockDesc]").val(myGroup.find("input[name$=-desc]").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			myGroup.find("input[name=plate-qtty]:first").val(echoNum(qtty));
			myGroup.find("input[name=plate-price]").val(echoNum(customer*color*qtty*price*(1+over/100-dis/100)));
			//------------- Related Event -------------
			calc_proof(myGroup.parent().find("[name=proof-dis]"));
			calc_print(myGroup.parent().find("[name=print-type]:first"));
	}
	checkForce(e);
	calc_addedPaper(e);
	calc_total();
}
function calc_print(e){
	var myGroup = $(e).closest("div.group");
	var color = parseInt(myGroup.parent().find("input[name=plate-color-count]:first").val());
	var spcolor = parseInt(myGroup.find("input[name=print-sp-color]:first").val());
	var form = parseInt(myGroup.find("input[name=print-form]:first").val());
	var qtty = getNum(myGroup.find("input[name=print-qtty]:first").val());
	var price = parseInt(myGroup.find("select[name=print-type]:first").children(":selected").attr("price"));
	var dup = parseInt(myGroup.find("select[name=print-d-type]:first").val());
	if (dup==2 && form % 2 !=0){
		form += 1;
		 myGroup.find("input[name=print-form]:first").val(form);
	}
	var ac_qtty = form * qtty;
	if (dup!=1)
			ac_qtty = ac_qtty / 2;
	myGroup.find("input[name=print-qtty]:first").val(echoNum(qtty));
	myGroup.parent().find("input[name=paper-after-cut_qtty]").val(ac_qtty);
	if (qtty<5000) 
		qtty=5000;
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	myGroup.find("input[name=print-price]").val(echoNum((color*form*qtty*price + 1.5*spcolor*form*qtty*price)*(1+over/100-dis/100)));
	//------------- Related Event -------------
	if ($(e).attr("name")=="print-form" || $(e).attr("name")=="print-qtty"){
		if (myGroup.parent().find("[name=verni-disBtn]").is(":checked"))
			calc_verni(myGroup.parent().find("[name=verni-disBtn]"));
		if (myGroup.parent().find("[name=selefon-disBtn]").is(":checked"))
			calc_selefon(myGroup.parent().find("[name=selefon-disBtn]"));
		if (myGroup.parent().find("[name=uv-disBtn]").is(":checked"))
			calc_uv(myGroup.parent().find("[name=uv-disBtn]"));
		
	}
	checkForce(e);
	calc_addedPaper(e);
	calc_paper(myGroup.parent().find("[name=paper-weight]"));
	calc_verni(myGroup.parent().find("[name=verni-mat]"));
	calc_total();
}
function calc_binding(e){
	var myGroup = $(e).closest("div.group");
	var qtty = getNum(myGroup.find("input[name=binding-qtty]:first").val());
	var form = parseInt(myGroup.find("input[name=binding-form]:first").val());
	var type = parseInt(myGroup.find("select[name=binding-type]:first").children(":selected").val());
	var price = parseInt(myGroup.find("select[name=binding-size]:first").children(":selected").attr("price").split(",")[type-1]);
	var orient = parseInt(myGroup.find("input[name=binding-orient]:checked").val());
	if (qtty < 2000) 
		qtty = 2000;
	myGroup.find("input[name=binding-qtty]:first").val(echoNum(qtty));
	var other = false;
	var result = 0;
	switch (type){
		case 1:
			if (form < 10) form = 10;
			break;
		case 2:
			if (form < 3) form = 3;
			break;
		case 3:
			if (form < 15) form = 15;
			break;
		case 4:
			if (form < 15) form = 15;
			break;
		default:
			other = true;
			result = getNum(myGroup.find("[name$=-price]").val());
			myGroup.find("[name$=-price]").prop("readonly",false);
	}
	myGroup.find("input[name=binding-form]:first").val(form);
//	console.log("form:"+form+", qtty:"+qtty+", price:"+price)
	if (!other){
		result = price*form*qtty;
		if (orient==2 || orient==3)
			result += .2*result;
	}
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	if (myGroup.find("[name=binding-disBtn]").is(":checked"))
		myGroup.find("input[name=binding-price]").val(echoNum(result*(1+over/100-dis/100)));
	else 
		myGroup.find("input[name=binding-price]").val("0");
	checkForce(e);
	$("[name=print-type]").each(function(i){
		calc_addedPaper($(this));
	});
	calc_total();
}
function calc_paper(e){
	var myGroup = $(e).closest("div.group");
	var afterCut = myGroup.find("[name=paper-after-cut_size]");
	var beforCut = myGroup.find("[name=paper-size]:first").children(":selected");
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
	myGroup.find("input[name$=-l]").val(bh);
	myGroup.find("input[name$=-w]").val(bw);
	myGroup.find("input[name=paper-in-form]").val(rate);
	if ($(e).attr("name")!="paper-price" || ($(e).val()=="" && $(e).attr("name")=="paper-price")){
		
		var weight = parseInt(myGroup.find("[name=paper-weight]").children(":selected").text());
		var price = parseInt(myGroup.find("[name=paper-type]").children(":selected").attr("price"));
		var qtty = getNum(myGroup.find("[name=paper-qtty]").val());
		var result = bh*bw/10000*weight/1000*price*qtty;
		myGroup.find("[name=paper-qtty]").val(echoNum(qtty));
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		if (myGroup.find("[name=paper-customer]").is(":checked"))
			myGroup.find("input[name=paper-price]").val("0");
		else
			myGroup.find("input[name=paper-price]").val(echoNum(Math.round(result*(1+over/100-dis/100))));
	}
	//------------- Related Event -------------
	if ($(e).attr("name")=="paper-after-cut_size"){
		if (myGroup.parent().find("[name=selefon-disBtn]").is(":checked"))
			calc_selefon(myGroup.parent().find("[name=selefon-disBtn]"));
	} 
	//----------- CHECK STOCK
	if (myGroup.find("[name=paper-customer]").is(":checked")) {
		myGroup.find("input[name$=-stockName]").val("");
		myGroup.find("input[name$=-stockDesc]").val("");
	} else {
		var stockName = myGroup.find("span.groupName").html();
		stockName += " ";
		stockName += myGroup.find("select[name=paper-type] option:selected").text();
		stockName += " ";
		stockName += myGroup.find("select[name=paper-size] option:selected").text();
		stockName += " ";
		stockName += myGroup.find("select[name=paper-weight] option:selected").text()+"g";
		myGroup.find("input[name$=-stockName]").val(stockName);
		myGroup.find("input[name$=-stockDesc]").val(myGroup.find("input[name$=-desc]").val());
	} 
	
	checkForce(e);
	calc_total();
}
function calc_snap(e){
	var myGroup = $(e).closest("[groupname=snap]");
	switch ($(e).attr("name")){
		case 'snap-size':
			$(e).val($(e).val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
			myGroup.find("input[name$=-l]").val($(e).val().split("X")[0]);
			myGroup.find("input[name$=-w]").val($(e).val().split("X")[1]);
			break;
		case 'snap-price':
			$(e).val(echoNum(getNum($(e).val())));
			break;
		case 'snap-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			var price=getNum(myGroup.find("input[name=snap-price]:first").val());
			myGroup.find("input[name=snap-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
		case 'snap-over':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			var price=getNum(myGroup.find("input[name=snap-price]:first").val());
			myGroup.find("input[name=snap-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
	}
	if ($(e).attr("name")==""){
		
	} else if ($(e).attr("name")==""){
		
	}
	checkForce(e);
	calc_addedPaper(e);
	calc_total();
}
function calc_addedPaper(e){
	var thisRow = $(e).closest(".exteraArea");
	var qtty = getNum(thisRow.find("input[name=print-qtty]:first").val());
	var form = parseInt(thisRow.find("input[name=print-form]:first").val());
	var color = parseInt(thisRow.find("input[name=plate-color-count]").val());
	var spColor= parseInt(thisRow.find("input[name=print-sp-color]:first").val());
	var step = color + spColor + thisRow.find("[name=verni-disBtn]").is(':checked') +
		thisRow.find("[name=uv-disBtn]").is(':checked') +
		thisRow.find("[name=selefon-disBtn]").is(':checked') +
		thisRow.find("[name=fold-disBtn]").is(':checked') +
		thisRow.find("[name=snap-disBtn]").is(':checked');
	step += $("input[name=binding-disBtn]:checked").size();
	step += $("select[name=binding-type]:not(:disabled) option:selected[value=4]").size();
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
	thisRow.find("input[name=paper-wastepaper]:first").val(echoNum(result));
	var rate = parseInt(thisRow.find("input[name=paper-in-form]:first").val());
	var paperAfter = parseInt(thisRow.find("input[name=paper-after-cut_qtty]:first").val());
	//console.log("r: "+rate + " after: "+ paperAfter);
	paperQtty = Math.ceil((result + paperAfter)/rate);
	thisRow.find("input[name=paper-qtty]:first").val(echoNum(paperQtty));
	//calc_paper(thisRow.children("div[groupname=paper]:first").children("input[name=paper-in-form]:first"));
// temporary I disable it !
//	thisRow.find("[name=paper-in-form]").blur();
}
/*

„Õ«”»Â „Ì“«‰ »«ÿ·Â

 « 5000 ‰”ŒÂ
 Ì—«é Â— ›—„ * ⁄œ«œ „—Õ·Â * 5 / 1000
«“ 5000  « 10000 ‰”ŒÂ
 Ì—«é- 5000) Â— ›—„ * ⁄œ«œ „—Õ·Â * 3   / 1000
«“10000 ‰”ŒÂ »Â »«·«
 Ì—«é- 10000) Â— ›—„ * ⁄œ«œ „—Õ·Â * 1  / 1000

*/		

function calc_verni(e){
	var myGroup = $(e).closest("div.group");
	var myPrint = myGroup.parent().children("div[groupname=print]");
	var qtty = getNum(myPrint.find("input[name=print-qtty]:first").val());
	if (qtty<5000)
		qtty = 5000;
	var form = parseInt(myGroup.find("input[name=verni-form]:first").val());
	var verni_wat = parseInt(myGroup.find("select[name=verni-wat]:first").children(":selected").val());
	var printType = parseInt(myPrint.find("select[name=print-type]:first").children(":selected").val());
	var printPrice = parseInt(myPrint.find("select[name=print-type]:first").children(":selected").attr("price"));
	var rate = 30; 
	if (printType==3) {
		myGroup.children("select[name=verni-wat]:first").val(1);
		if (parseInt(myGroup.find("select[name=verni-mat] option:selected").val())==1)
			rate = 50;
			// mat
		else
			rate = 40;
			// barragh
	} else {
		myGroup.children("select[name=verni-wat]:first").val(2);
	}
	
	//var price = parseInt(myGroup.find("select[name=verni-mat]:first").children(":selected").attr("price").split(",")[verni_wat-1]);
/* 	console.log("rate:" + rate+", form:"+form+", qtty:"+qtty+", price:"+machinPrice); */
	var price = printPrice * (100 + rate)/100;
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	if (myGroup.find("[name=verni-disBtn]").is(":checked")){
		myGroup.find("input[name=verni-price]").val(echoNum(qtty*form*price*(1+over/100-dis/100)));
	} else {
		myGroup.find("input[name=verni-price]").val("0");
	}
	checkForce(e);
	calc_addedPaper(e);
	calc_total();
}
function calc_selefon(e){
	var myGroup = $(e).closest("div.group");
	var paperGroup = myGroup.parent().children("div[groupname=paper]");
	var w=parseFloat(myGroup.find("[name=selefon-size]:first").children(":selected").text().split("X")[0]);
	var h=parseFloat(myGroup.find("[name=selefon-size]:first").children(":selected").text().split("X")[1]);
	if (h>w){
		var t=w;
		w=h;
		h=t;
	}
	if (h<35 || w<50){
		h=35;
		w=50;
	}
	myGroup.find("input[name$=-l]").val(h);
	myGroup.find("input[name$=-w]").val(w);
	var face = parseInt(myGroup.find("[name=selefon-face]:first").children(":selected").val());
	var mat = parseInt(myGroup.find("[name=selefon-mat]:first").children(":selected").val());
	var price = parseFloat(myGroup.find("[name=selefon-hot]:first").children(":selected").attr("price").split(",")[mat-1]);
	var myPrint = myGroup.parent().children("div[groupname=print]");
	var qtty = getNum(myPrint.find("input[name=print-qtty]:first").val());
	if (qtty<1000)
		qtty = 1000;
	var form = parseInt(myPrint.find("input[name=print-form]:first").val());
	if (myPrint.find("select[name=print-d-type]:first").children(":selected").val()==2)
		form = form / 2;
	//console.log("w: "+w+" h: "+h+" price: "+price+" mat: "+mat+" face: "+face+" from "+form+" qtty "+qtty);
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	if (myGroup.find("[name=selefon-disBtn]").is(":checked")){
		myGroup.find("input[name=selefon-price]").val(echoNum(face*price*h*w*form*qtty*(1+over/100-dis/100)));
	} else {
		myGroup.find("input[name=selefon-price]").val("0");
	}
	checkForce(e);
	calc_total();
	calc_addedPaper(e);
}
function calc_uv(e){
	/*
if ($(e).attr("name")=="uv-size")
		$(e).val($(e).val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
*/
	var myGroup = $(e).closest("div.group");
	var myPrint = myGroup.parent().children("div[groupname=print]");
	var qtty = getNum(myPrint.find("input[name=print-qtty]:first").val());
	var form = parseInt(myGroup.find("input[name=uv-form]:first").val());
	var mat = parseInt(myGroup.find("[name=uv-mat]:first").children(":selected").val()) - 1;
	var spot = parseInt(myGroup.find("[name=uv-spot]:first").children(":selected").val()) - 1;
	var price = parseInt(myGroup.find("[name=uv-size]").children(":selected").attr("price").split(",")[spot].split("|")[mat]);
	if (qtty<1000)
		qtty = 1000;
	//console.log("qtty: "+qtty+" form: "+form+" price: "+price);
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	if (myGroup.find("[name=uv-disBtn]").is(":checked")){
		myGroup.find("input[name=uv-price]:first").val(echoNum(qtty * form * price*(1+over/100-dis/100)));
	} else {
		myGroup.find("input[name=uv-price]:first").val("0");
	}
	checkForce(e);
	calc_addedPaper(e);
	calc_total();
}
function calc_fold(e){
	var myGroup = $(e).closest("div.group");
	switch ($(e).attr("name")){
		case 'fold-price':
			$(e).val(echoNum(getNum($(e).val())));
			break;
		case 'fold-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
		case 'fold-over':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
	}
	checkForce(e);
	calc_total();
	calc_addedPaper(e);
}
function calc_packing(e){
	var myGroup = $(e).closest("div.group");
//	console.log($(e));
	switch ($(e).attr("name")){
		case 'packing-price':			
			$(e).val(echoNum(getNum($(e).val())));
			break;
		case 'packing-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
		case 'packing-over':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
	}
	checkForce(e);
	calc_total();
}
function calc_service(e){
	var myGroup = $(e).closest("div.group");
	switch ($(e).attr("name")){
		case 'service-price':
			$(e).val(echoNum(getNum($(e).val())));
			break;
		case 'service-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
		case 'service-over':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
	}
	myGroup.find("input[name$=-purchaseName]").val(myGroup.find("[name=service-item] option:selected").text());
	myGroup.find("input[name$=-purchaseDesc]").val((myGroup.find("[name=service-description]")[0]).value);
	myGroup.find("input[name$=-purchaseTypeID]").val(myGroup.find("[name=service-item] option:selected").val());
	checkForce(e);
	calc_total();
}
function calc_proof(e){
	var myGroup = $(e).closest("div.group");
	switch ($(e).attr("name")){
		case 'proof-date':
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
					$(e).val("Œÿ«!");
					$(e).focus();
				}
			}
			break;
		case 'proof-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		case 'proof-over':
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=$-over]:first").val(echoPer(over));
		default:
			var paperGroup = myGroup.parent().find("div[groupname=paper]");
			var afterCut = paperGroup.find("[name=paper-after-cut_size]");
			afterCut.val(afterCut.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
			var w=parseFloat(afterCut.val().split("X")[0]);
			var h=parseFloat(afterCut.val().split("X")[1]);
			var qtty = parseInt(myGroup.parent().find("input[name=plate-qtty]").val());
			var result = 0;
			if (myGroup.find("[name=proof-dammy]").is(":checked"))
				result += Math.round(parseInt(myGroup.find("[name=proof-dammy]").attr("price")) * w * h * qtty);
			if (myGroup.find("[name=proof_digital]").is(":checked"))
				result += parseInt(myGroup.find("[name=proof_digital]").attr("price")) * qtty;
			if (myGroup.children("[name=proof_iso]").is(":checked"))
				result += Math.round(parseInt(myGroup.find("[name=proof_iso]").attr("price")) * w * h * qtty);
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			myGroup.find("input[name=proof-price]:first").val(echoNum(result*(1+over/100-dis/100)));
	}
	checkForce(e);
	calc_total();
}
function calc_delivery(e){
	var myGroup = $(e).closest("div.group");
	switch ($(e).attr("name")){
		case 'delivery-price':
			$(e).val(echoNum(getNum($(e).val())));
			break;
		case 'delivery-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
		case 'delivery-over':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
		case 'delivery-point':
			myGroup.find("[name=delivery-address]").val($(e).children(":selected").attr("addr"));
			if ($(e).children(":selected").val()==3){
				$.getJSON("../CRM/json_getAccount.asp",
					{act:"getAddr",id:$("input[name=customerID]:first").val()},
					function(json){
						if (json.addr1!="")
							myGroup.find("[name=delivery-address]").val(json.addr1);
						else
							myGroup.find("[name=delivery-address]").val(json.addr2);
					});
			}
			break;
	}
		
	checkForce(e);
	calc_total();

}
function calc_cutting(e){
	var myGroup = $(e).closest("div.group");
	console.log(myGroup.find("input[name$=-price]:first"));
	switch ($(e).attr("name")){
		case 'cutting-price':
			$(e).val(echoNum(getNum($(e).val())));
			myGroup.find("input[name$=-dis]:first").val("0%");
			myGroup.find("input[name$=-over]:first").val("0%");
			break;
		case 'cutting-dis':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			console.log("d:" + dis + " o: "+ over + " p: "+price);
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
		case 'cutting-over':
			var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
			var over = getPer(myGroup.find("input[name$=-over]:first").val());
			myGroup.find("input[name$=-over]:first").val(echoPer(over));
			var price=getNum(myGroup.find("input[name$=-price]:first").val());
			myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
			break;
	}
	checkForce(e);
	calc_total();
}
function calc_photobookDesc(e){
	var myGroup = $(e).closest("div.group");
	var qtty = parseInt(myGroup.find("[name=photobook-qtty]").val());
	if (qtty<1)
		myGroup.find("[name=photobook-qtty]").val(1);
	calc_photobookBlock(myGroup.parent().children("div[groupname=photobookBlock]").find("[name=block-qtty-luster]:first"));
	calc_photobookCover(myGroup.parent().children("div[groupname=photobookCover]").find("[name=cover-type]:first"));
	checkForce(e);
	calc_total();
}
function calc_photobookBlock(e){
	var myGroup = $(e).closest("div.group");
	var descGroup = myGroup.parent().children("div[groupname=photobookDesc]");
	var qtty = parseInt(descGroup.find("[name=photobook-qtty]").val());
	var thickness = parseInt(myGroup.find("[name=thickness]").children(":selected").val());
	var siz = parseInt(descGroup.find("[name=album-size]:first").children(":selected").val());
	var lasterQtty = parseInt(myGroup.find("[name=block-qtty-luster]:first").val());
	var silkQtty = parseInt(myGroup.find("[name=block-qtty-silk]:first").val());
	var metalicQtty = parseInt(myGroup.find("[name=block-qtty-metalic]:first").val());
	if (siz==-1 || isNaN(siz)){
		var w = parseFloat(descGroup.find("[name=album-size]:first").children(":selected").text().split("X")[0]);
		var h = parseFloat(descGroup.find("[name=album-size]:first").children(":selected").text().split("X")[1]);
		var priceLaster = parseFloat(myGroup.find("[name=thickness]:checked").attr("price").split(",")[0].split("|")[0]);
		var priceSilk = parseFloat(myGroup.find("[name=thickness]:checked").attr("price").split(",")[0].split("|")[1]);
		var priceMetalic = parseFloat(myGroup.find("[name=thickness]:checked").attr("price").split(",")[0].split("|")[2]);
		priceLaster *= lasterQtty * w * h;
		priceSilk *= silkQtty * w * h;
		priceMetalic *= metalicQtty * w * h;
	} else {
		var priceLaster = parseFloat(myGroup.find("[name=thickness]:checked").attr("price").split(",")[siz].split("|")[0]);
		var priceSilk = parseFloat(myGroup.find("[name=thickness]:checked").attr("price").split(",")[siz].split("|")[1]);
		var priceMetalic = parseFloat(myGroup.find("[name=thickness]:checked").attr("price").split(",")[siz].split("|")[2]);
		if (priceLaster==0) 
			myGroup.find("[name=block-qtty-luster]:first").val(0);
		if (priceSilk==0) 
			myGroup.find("[name=block-qtty-silk]:first").val(0);
		if (priceMetalic==0) 
			myGroup.find("[name=block-qtty-metalic]:first").val(0);
		priceLaster *= lasterQtty;
		priceSilk *= silkQtty;
		priceMetalic *= metalicQtty;
	}
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	myGroup.find("input[name=photobookBlock-price]").val(echoNum((priceLaster + priceMetalic + priceSilk) * qtty * (1+over/100-dis/100)));
	
	checkForce(e);
	calc_total();
}
function calc_photobookCover(e){
	var myGroup = $(e).closest("div.group");
	var descGroup = myGroup.parent().children("div[groupname=photobookDesc]");
	var qtty = parseInt(descGroup.find("[name=photobook-qtty]").val());
	var astar = parseInt(myGroup.find("[name=cover-astar]:checked").val());
	var coverType = parseInt(myGroup.find("[name=cover-type]:first").children(":selected").val());
	var siz = parseInt(descGroup.find("[name=album-size]:first").children(":selected").val());
	var price = 0;
	var coverPrice = 0;
	var astarPrice = 0;
	if (siz==-1 || isNaN(siz)){
		var w = parseFloat(descGroup.find("[name=album-size]:first").children(":selected").text().split("X")[0]);
		var h = parseFloat(descGroup.find("[name=album-size]:first").children(":selected").text().split("X")[1]);
		coverPrice = parseFloat(myGroup.find("[name=cover-type]:first").children(":selected").attr("price").split(",")[0]) * h * w;
		astarPrice = parseFloat(myGroup.find("[name=cover-astar]:checked").attr("price").split(",")[0]) * h * w;
	} else {
		if (coverType!=0){
				coverPrice = parseFloat(myGroup.find("[name=cover-type]:first").children(":selected").attr("price").split(",")[siz]);
			if (coverPrice == 0)
				myGroup.find("[name=cover-type]:first").children(":first").prop("selected",true);
		}
		if (astar!=0){
			astarPrice = parseFloat(myGroup.find("[name=cover-astar]:checked").attr("price").split(",")[siz]);
			console.log(astarPrice);
			if (astarPrice == 0)
				myGroup.find("[name=cover-astar]:first").prop("checked",true);
		}
	}
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	myGroup.find("input[name=photobookCover-price]").val(echoNum((coverPrice + astarPrice) * qtty * (1+over/100-dis/100)));
	
	checkForce(e);
	calc_total();
}

function calc_design(e){
	var myGroup = $(e).closest("div.group");
	var qtty = parseInt(myGroup.find("[name$=-qtty]").val());
	var price = parseFloat(myGroup.find("[name$=-type] option:selected").attr("price"));
	var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
	var over = getPer(myGroup.find("input[name$=-over]:first").val());
	myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
	myGroup.find("input[name$=-over]:first").val(echoPer(over));
	myGroup.find("[name$=-price]:first").val(echoNum(price*qtty*(1+over/100-dis/100)));
	checkForce(e);
	calc_total();
}
function calc_leaflet(e){
	var myGroup = $(e).closest("div.group");
	var qtty = parseInt(myGroup.find("[name$=-qtty]").val());
	if (qtty<6)
		myGroup.find("[name$=-qtty]").val(6);
	calc_design(e);
}
function calc_scan(e){
	calc_design(e);
}
function calc_tract(e){
	calc_design(e);
}
function calc_ads(e){
	calc_design(e);
}
function calc_folder(e){
	calc_design(e);
}
function calc_adminPaper(e){
	calc_design(e);
}
function calc_label(e){
	calc_design(e);
}
function calc_dsigenPacking(e){
	calc_design(e);
}
function calc_logo(e){
	calc_design(e);
}

function calc_konica(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var qtty = parseInt(myGroup.find("[name=konica-qtty]").val());
		var form = parseInt(myGroup.find("[name=konica-form]").val());
		var size = parseInt(myGroup.find("[name^=konica-size]:checked").val());
		var dup = parseInt(myGroup.find("[name^=konica-dup]:checked").val());
		var paper = parseInt(myGroup.find("[name=konica-paper] option:selected").val());
		if (paper==6 && dup==2){
			myGroup.find("[name^=konica-dup][value=1]").prop("checked", true);
			dup=1;
		}
		var price = parseInt((myGroup.find("[name=konica-paper] option:selected").attr("price").split(",")[size-1]).split("|")[dup-1]);
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*qtty*form*(1+over/100-dis/100)));
		//console.log("price:"+price+", qtty:"+qtty+", form:"+form);
		calc_digitalCut(myGroup.parent().find("[name=digitalCut-befor]"));
		calc_digitalLaminate(myGroup.parent().find("[name=digitalLaminate-type]"));
		calc_digitalBinding(myGroup.parent().find("[name=digitalBinding-type]"));
		myGroup.parent().find("[name=envelope-qtty]").val(qtty * form);
	}
	checkForce(e);
	calc_total();
}
function calc_plot(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var qtty = parseInt(myGroup.find("[name=plot-qtty]").val());
		var size = myGroup.find("[name=plot-size]");
		size.val(size.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
		var w = parseFloat(size.val().split("X")[0]);
		var h = parseFloat(size.val().split("X")[1]);
		myGroup.parent().find("[name=foam-size]").val(w + "X" + h);
		myGroup.parent().find("[name=springFrame-size]").val(w + "X" + h);
		if (w*h<5000){
			w = 50;
			h = 100;
		}
		var price = parseInt(myGroup.find("[name=plot-paper] option:selected").attr("price"));
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*qtty*h*w*(1+over/100-dis/100)));
		calc_digitalLaminate(myGroup.parent().find("[name=digitalLaminate-type]"));
		calc_foam(myGroup.parent().find("[name=foam-size]"));
	}
	checkForce(e);
	calc_total();
}
function calc_digitalCut(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var konicaGroup = myGroup.parent().children("div[groupname=konica]");
		var size = myGroup.find("[name=digitalCut-befor]");
		if (size.val().indexOf(" ")>-1 || size.val().indexOf("-")>-1 || size.val().indexOf("*")>-1 || size.val().indexOf("x")>-1)
			size.val(size.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
		var qtty = parseInt(konicaGroup.find("[name=konica-qtty]").val());
		var form = parseInt(konicaGroup.find("[name=konica-form]").val());
		var price=0;
		if (qtty*form<=20)
			price = parseInt(myGroup.find("[name=digitalCut-befor]").attr("price").split(",")[0]);
		else
			price = parseInt(myGroup.find("[name=digitalCut-befor]").attr("price").split(",")[1]);
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
	}
	checkForce(e);
	calc_total();
}
function calc_digitalLaminate(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var konicaGroup = myGroup.parent().children("div[groupname=konica]");
		var plotGroup = myGroup.parent().children("div[groupname=plot]");
		var result = 0;
		
		if (konicaGroup.find("[name$=-disBtn]").is(":checked")){
			var qtty = parseInt(konicaGroup.find("[name=konica-qtty]").val());
			var form = parseInt(konicaGroup.find("[name=konica-form]").val()); 
			var price = parseInt(myGroup.find("[name=digitalLaminate-type] option:selected").attr("price").split(",")[0]);
			result += price * qtty * form; 
		}
		if (plotGroup.find("[name$=-disBtn]").is(":checked")){
			var qtty = parseInt(plotGroup.find("[name=plot-qtty]").val());
			var size = plotGroup.find("[name=plot-size]");
			var w = parseFloat(size.val().split("X")[0]);
			var h = parseFloat(size.val().split("X")[1]);
			var price = parseInt(myGroup.find("[name=digitalLaminate-type] option:selected").attr("price").split(",")[1]);
			result += price * qtty * w * h; 
		}
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(result*(1+over/100-dis/100)));
	}
	checkForce(e);
	calc_total();
}
function calc_digitalBinding(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var konicaGroup = myGroup.parent().children("div[groupname=konica]");
		var qtty = parseInt(konicaGroup.find("[name=konica-qtty]").val());
		var form = parseInt(konicaGroup.find("[name=konica-form]").val());
		var price = 0;
		if (qtty*form<=20){
			price = parseInt(myGroup.find("[name=digitalBinding-type] option:selected").attr("price").split(",")[0]);
		} else {
			if (parseInt(myGroup.find("[name=digitalBinding-type] option:selected").val())>2){
				price = parseInt(myGroup.find("[name=digitalBinding-type] option:selected").attr("price").split(",")[1]) * form * qtty;
			} else {
				price = parseInt(myGroup.find("[name=digitalBinding-type] option:selected").attr("price").split(",")[1]);
			}
		}
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
	}
	checkForce(e);
	calc_total();
}
function calc_envelope(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var qtty = parseInt(myGroup.find("[name=envelope-qtty]").val());
		var price = parseInt(myGroup.find("[name=envelope-qtty]").attr("price"));
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*qtty*(1+over/100-dis/100)));
	}
	checkForce(e);
	calc_total();
}
function calc_estand(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var qtty = parseInt(myGroup.find("[name=estand-qtty]").val());
		var price = parseInt(myGroup.find("[name=estand-type] option:selected").attr("price"));
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*qtty*(1+over/100-dis/100)));
	}
	checkForce(e);
	calc_total();
}
function calc_foam(e){
	var myGroup = $(e).closest("div.group");
	var plotGroup = myGroup.parent().children("div[groupname=plot]");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var qtty = parseInt(myGroup.find("[name=foam-qtty]").val());
		var price = parseInt(myGroup.find("[name=foam-size]").attr("price"));
		var size = myGroup.find("[name=foam-size]");
		size.val(size.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
		var w = parseFloat(size.val().split("X")[0]);
		var h = parseFloat(size.val().split("X")[1]);
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*qtty*w*h*(1+over/100-dis/100)));
	}
	checkForce(e);
	calc_total();
}
function calc_sampling(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var price = parseInt(myGroup.find("[name=sampling-desc]").attr("price"));
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum(price*(1+over/100-dis/100)));
		
	}
	checkForce(e);
	calc_total();
}
function calc_springFrame(e){
	var myGroup = $(e).closest("div.group");
	if (myGroup.find("[name$=-disBtn]").is(":checked")){
		var qtty = parseInt(myGroup.find("[name=springFrame-qtty]").val());
		var priceFrame = parseInt(myGroup.find("[name=springFrame-size]").attr("price"));
		var priceCorner = parseInt(myGroup.find("[name=springFrame-qtty]").attr("price"));
		var size = myGroup.find("[name=springFrame-size]");
		size.val(size.val().replace("x","X").replace("-","X").replace("*","X").replace(" ","X"));
		var w = parseFloat(size.val().split("X")[0]);
		var h = parseFloat(size.val().split("X")[1]);
		var dis = getPer(myGroup.find("input[name$=-dis]:first").val());
		var over = getPer(myGroup.find("input[name$=-over]:first").val());
		myGroup.find("input[name$=-dis]:first").val(echoPer(dis));
		myGroup.find("input[name$=-over]:first").val(echoPer(over));
		myGroup.find("input[name$=-price]:first").val(echoNum((2 * (w+h) * priceFrame + priceCorner) * qtty *(1+over/100-dis/100)));
		console.log("w: "+w+", h: "+h+",frame: "+priceFrame+",corner: " + priceCorner + ",qtty: "+ qtty);
	}
	checkForce(e);
	calc_total();
}
