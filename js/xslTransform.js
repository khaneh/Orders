function loadXMLDoc(fname, retfunc) { 
    $.ajax({
        url: fname,
        cache: true,
        dataType: 'xml',
        success: function(data) { if (retfunc) retfunc(data); }
    }); 
}
function TransformXml(xml,xslUrl,retfunc) {
	loadXMLDoc(xslUrl, function (xsl) {
	if (window.ActiveXObject) {
	    retfunc(xml.transformNode(xsl));
	}
	else if (document.implementation && document.implementation.createDocument) {
	    var xsltProcessor = new XSLTProcessor();
	    xsltProcessor.importStylesheet(xsl);
	    var resultDocument = xsltProcessor.transformToFragment(xml, document);
	    retfunc(resultDocument);
	    
	} 
	}); 
}
function TransformXmlURL(xmlUrl,xslUrl,retfunc) {
	loadXMLDoc(xmlUrl,function(xml){
		loadXMLDoc(xslUrl, function (xsl) {
		if (window.ActiveXObject) {
		    retfunc(xml.transformNode(xsl));
		}
		else if (document.implementation && document.implementation.createDocument) {
		    var xsltProcessor = new XSLTProcessor();
		    xsltProcessor.importStylesheet(xsl);
		    var resultDocument = xsltProcessor.transformToFragment(xml, document);
		    retfunc(resultDocument);
		    
		} 
		});
	}); 
}