<%

	dim xml, text
	set xml = server.CreateObject("MSXML2.ServerXMLHTTP")
	xml.open "GET", "https://192.168.0.9/test/getvmi.php?exten=119", false
	xml.SetOption(2) = 13056
	xml.send
	text = xml.ResponseText
	Response.write(text)
%>