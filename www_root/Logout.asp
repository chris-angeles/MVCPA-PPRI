<%
Dim Con
Set Con = server.CreateObject("ADODB.connection")
Con.Open(Application("ConnectionString"))
Con.Execute("UPDATE System.LoginLog SET LogoutTime=getdate() WHERE SessionID=" & Session.SessionID)
Session.Abandon
Response.Redirect("Default.asp?Message=Your session has ended")
Response.Cookies("SystemID")=0 
%>