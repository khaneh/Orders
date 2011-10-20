<% 
	Session("isConnected") = False
	Session("ConnStr") = ""	
	Session.Abandon()
	Response.Redirect "default.asp"
	Response.End
%>
