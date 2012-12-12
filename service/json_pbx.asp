<% 
    url = "http://192.168.10.10/json_campain.php?exten=" & Request("exten") 
    set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
    xmlhttp.open "GET", url, false 
    xmlhttp.send "" 
    Response.write xmlhttp.responseText 
    set xmlhttp = nothing 
%>