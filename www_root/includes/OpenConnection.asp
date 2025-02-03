<%
Dim con, rs, sql
Set con = server.CreateObject("ADODB.connection")
Con.Open(Application("ConnectionString"))
%>