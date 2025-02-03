<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, LastName
debug = false
If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLF)
	For each i in Request.Form
		Response.Write("Request.Form(""" & i & """)='" & Request.Form(i) & "'" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("Session(""" & i & """)='" & Session(i) & "'" & vbCrLf)
	Next
	for each i in Request.Cookies
		if Request.Cookies(i).HasKeys then
			for each j in Request.Cookies(x)
				response.write("Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j))
			next
		else
			Response.Write("Cookies(""" & i & """)=" & Request.Cookies(i) & "<br>")
		end if
	next
	Response.Write("</pre>" & vbCrLF)
End If

If Len(Request.Form("LastName"))>0 Then
	LastName = Request.Form("LastName")
ElseIf Len(Request.QueryString("LastName"))>0 Then
	LastName = Request.QueryString("LastName")
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Add or Update A User</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" />
<script type="text/javascript">
	function validateForm()
	{
		if (document.UpdateUser.LastName.value.length>0)
		{
			return true;
		}
		else
		{
			alert("You must enter some text to use for searching existing users.");
			document.UpdateUser.LastName.focus();
			return false;
		}
	}
</script>
</head>
<body onload="document.UpdateUser.LastName.focus();">
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>
<div class="pagetag">Add or Update A User</div>
<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>
<div class="content">

<form name="UpdateUser" id="UpdateUser" method="post" action="UpdateUser2.asp" onsubmit="return validateForm();">
	<fieldset style="width: 640px;">
		<legend>Step 1: Search for User</legend>
		<label for="LastName">Last Name:</label>
		<input type="text" name="LastName" id="LastName" value="<%=LastName %>" size="24" maxlength="24" />
</fieldset>
<div style="text-align: center; width: 640px;">
<input type="submit" value="Submit" title="Submit last name for search" />
<input type="button" value="Cancel" title="Return to Home" onclick="location.href = '../Home/Default.asp';" />
</div>
</form>
</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepDB.asp"-->