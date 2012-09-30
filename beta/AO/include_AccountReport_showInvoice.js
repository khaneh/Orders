$(document).ready(function(){
	$.ajaxSetup({
		cache: false
	});
	$("#descDialog").dialog({ 
		autoOpen: false,
		buttons: {" «ÌÌœ":function(){
			var lineID=$("#descDialog input[name=id]").val();
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"updateLineDesc",id:lineID,desc:$("#descDialog input[name=newDesc]").val()},
				dataType:"json"
			}).done(function (data){
				$(".invoiceItem[lineID=" + lineID +"] td.chgDesc span.desc").html(data.desc);
			});
			$(this).dialog("close");
		}},
		title: " €ÌÌ— ‘—Õ"
	});
	$("#isaDialog").dialog({ 
		autoOpen: false,
		buttons: {" «ÌÌœ":function(){
			var invID=$("input[name=InvoiceID]").val();
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"updateIsA",id:invID,isa:$("#isaDialog input[name=newIsA]:checked").val()},
				dataType:"json"
			}).done(function (data){
				if (data.errMsg=="")
					location.reload();
				else
					location.href=$(location).attr('href')+"&errMsg="+escape(data.errMsg);
			});
			$(this).dialog("close");
		}},
		title: " €ÌÌ— «·›/»"
	});
	$("#noDialog").dialog({ 
		autoOpen: false,
		buttons: {" «ÌÌœ":function(){
			var invID=$("input[name=InvoiceID]").val();
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"updateNo",id:invID,no:$("#noDialog input[name=newNo]").val()},
				dataType:"json"
			}).done(function (data){
				$("input[name=InvoiceNo]").val(data.no);
			});
			$(this).dialog("close");
		}},
		title: " €ÌÌ— ‘„«—Â ›«ﬂ Ê— —”„Ì"
	});
	$("#removeIssuedDialog").dialog({ 
		autoOpen: false,
		buttons: {" «ÌÌœ":function(){
			var invID=$("input[name=InvoiceID]").val();
			$("#removeIssue").addClass("disabled");
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"removeIssue",id:invID,comment:$("#removeIssuedDialog input[name=comment]").val()},
				dataType:"json"
			}).done(function (data){
				if (data.err==1)
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&errmsg=" +data.msg;
				else
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&msg=" +data.msg;
			});
			$(this).dialog("close");
			$("#removeIssue").addClass("disabled");
		}},
		title: "Œ—ÊÃ «“ Õ«·  ’œÊ—",
		beforeClose: function(event, ui){console.log(event); $("#removeIssue").removeClass("disabled");}
	});
	$(".chgDesc").mouseover(function(event){
		$(this).find(".chgDescBtn").css("display","block");
	});
	$(".chgDesc").mouseout(function(event){
		$(this).find(".chgDescBtn").css("display","none");
	});
	$(".chgNo").mouseover(function(event){
		$(this).find(".chgNoBtn").css("display","block");
	});
	$(".chgNo").mouseout(function(event){
		$(this).find(".chgNoBtn").css("display","none");
	});
	$(".chgIsA").mouseover(function(event){
		$(this).find(".chgIsaBtn").css("display","block");
	});
	$(".chgIsA").mouseout(function(event){
		$(this).find(".chgIsaBtn").css("display","none");
	});
	$(".chgDescBtn").click(function(){
		$("#descDialog input[name=newDesc]").val($(this).prev(".desc").html());
		$("#descDialog input[name=id]").val($(this).closest("tr").attr("lineID"));
		$("#descDialog").dialog('open');
	});
	$(".chgIsaBtn").click(function(){
		//$("#isaDialog input[name=newDesc]").val($(this).prev(".desc").html());
		$("#isaDialog").dialog('open');
	});
	$(".chgNoBtn").click(function(){
		$("#noDialog input[name=newNo]").val($("input[name=InvoiceNo]").val());
		$("#noDialog").dialog('open');
	});
	$("#ApproveInvoice").click(function(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« ' «ÌÌœ' ﬂ‰Ìœø\n\n")){
			var invID=$("input[name=InvoiceID]").val();
			$("#ApproveInvoice").addClass("disabled");
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"approveInvoice",id:invID},
				dataType:"json"
			}).done(function (data){
				if (data.err==1)
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&errmsg=" +data.msg;
				else
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&msg=" +data.msg;
			});
		}
	});
	$("#removeApprove").click(function(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« «“ ' «ÌÌœ' Œ«—Ã ﬂ‰Ìœø\n\n")){
			var invID=$("input[name=InvoiceID]").val();
			$("#removeApprove").addClass("disabled");
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"removeApprove",id:invID},
				dataType:"json"
			}).done(function (data){
				if (data.err==1)
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&errmsg=" +data.msg;
				else
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&msg=" +data.msg;
			});
		}
	});
	$("#removeIssue").click(function(){
		$("#removeIssue").addClass("disabled");
		$("#removeIssuedDialog").dialog("open");
	});
	$("#IssueInvoice").click(function(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '’«œ—' ﬂ‰Ìœø\n\n")){
			var invID=$("input[name=InvoiceID]").val();
			$("#IssueInvoice").addClass("disabled");
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"issueInvoice",id:invID},
				dataType:"json"
			}).done(function (data){
				if (data.err==1)
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&errmsg=" +data.msg;
				else
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&msg=" +data.msg;
			});
		}
	});
	$("#IssueInvoiceInDate").click(function(){
		if (confirm("¬Ì« „ÿ„∆‰ Â” Ìœ ﬂÂ „Ì ŒÊ«ÂÌœ «Ì‰ ›«ﬂ Ê— —« '’«œ—' ﬂ‰Ìœø\n\n")){
			var invID=$("input[name=InvoiceID]").val();
			$("#IssueInvoiceInDate").addClass("disabled");
			$.ajax({
				type:"POST",
				url:"/service/json_invoice.asp",
				data:{act:"issueInvoice",id:invID,issueDate:$("input[name=IssueDate]").val()},
				dataType:"json"
			}).done(function (data){
				if (data.err==1)
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&errmsg=" +data.msg;
				else
					location.href="AccountReport.asp?act=showInvoice&invoice="+invID+"&msg=" +data.msg;
			});
		}
	});
	

});
