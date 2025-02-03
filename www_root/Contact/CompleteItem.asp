<%@ language=VBScript %>
<% Option Explicit %>
<!--#include file="../includes/EnsureLogin.asp"--> 
<!--#include file="../includes/adovbs.asp"--> 
<!--#include file="../includes/OpenConnection.asp"--> 
<%
Dim ContactPhoneCallID
ContactPhoneCallID = Request.QueryString("ContactPhoneCallID")

If IsNumeric(ContactPhoneCallID) = True Then
	sql = "UPDATE Contact.PhoneCalls " & _
		"SET DateComplete='" & date() & "', UpdateID=" & UserSystemID & _
		", UpdateTimeStamp=getdate() WHERE ContactPhoneCallID=" & ContactPhoneCallID
	Con.Execute(sql)
End If
%>
<html lang="en-us">
<head>
	<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
	<title>Complete Contact Item</title>
	<link rel=stylesheet href="../styles/main.css" type="text/css">
</head>
<body bgColor=#e5e5e5>
<!-- Main Table to hold all of page contents. Body backgound color will serve as border to this.-->
<table border=0 align=center bgcolor="FFFFFF" width=760><tr><td align=center>
<!-- End Main Table Code-->

<script langage=JavaScript>
window.close();
</script>


<!-- End of outer Table: start tags-->
</td></tr></table>
<!--End of outer table: end tags-->
</body>
</html>

