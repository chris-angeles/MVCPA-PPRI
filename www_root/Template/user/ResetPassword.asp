<%@  language="VBScript" %>
<% Option Explicit%><% 
Dim Debug, i, UserID, Message, recaptcha_public_key
If Application("Instance") = "Test" Then
	recaptcha_public_key = "6Ld36okpAAAAAP35zWtqBsy0rx8d41pWK4utB6GG" ' test public key
Else
	recaptcha_public_key = "6LdCvh4UAAAAAEBmrEUz6t6yTdXN0ybyNaRQZH1i" ' your public key
End If

Debug = False

If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLF)
	Response.Write("Request.Form.Count=" & Request.Form.Count & vbCrLf)
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
	Response.Write("Time='" & Now() & "'" & vbCrLf)
	Response.Write("</pre>" & vbCrLF)
End If

If Request.QueryString.Count>0 Then
	' This is initial page load. Continue.
	UserID=Request.QueryString("UserID")
	Message = Request.QueryString("Message")
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Password Reset</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" />
<link rel="stylesheet" href="/styles/main.css" type="text/css" />
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" />
<script type="text/javascript">
	function validateForm()
	{
		if (document.PasswordReset.UserID.value.length <= 0) {
			alert("You must enter a User ID.");
			PasswordReset.UserID.focus();
			return false;
		}
		if (document.PasswordReset.UserID.value.length < 5) {
			alert("The User ID must be at least five characters.");
			PasswordReset.UserID.focus();
			return false;
		}
		if (grecaptcha.getResponse() == "") {
			alert("You must complete the CAPTCHA to submit this form!");
			return false;
		}
		return true;
	}

	function recaptchaCallback()
	{
		document.PasswordReset.Submit.removeAttribute("disabled");
	};
</script>
<script src="https://www.google.com/recaptcha/api.js" async defer></script>
</head>
<body>
	<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

	<div class="pagetag">
		The <strong>Motor Vehicle Crime Prevention Authority </strong>(MVCPA) 
	awards financial grants to agencies, organizations, and concerned parties in an effort to 
	raise public awareness of vehicle theft and burglary and implement education and prevention 
	initiatives.
	</div>

	<div class="content">

		<%	
If Len(message) > 0 Then
	Response.Write("<div class=""notice"">" & message & "</div>" & vbCrLf)
End If 
		%>
		<div class="section">
			Enter your email address. If a valid user, a new password will be emailed to that address. Use the new password to login. You may change your password to one that is easier to remember after you have logged in.<br />
			<form name="PasswordReset" id="PasswordReset" method="post" action="ResetPasswordSubmit.asp" onsubmit="return validateForm()">
				<fieldset style="width: 732px;">
					<legend>Password Reset</legend>
					<label for="UserID">User ID (email):</label>
					<input type="text" name="UserID" id="UserID" value="<%=UserID %>" size="40" maxlength="255" /><br />
				</fieldset>
				<div class="g-recaptcha" data-callback="recaptchaCallback" data-sitekey="<%=recaptcha_public_key %>"></div>
				<div style="text-align: center">
					<input type="submit" name="Submit" value="Reset Password" disabled /></div>
			</form>
		</div>
		<div class="section">
			<a href="/default.asp" class="plainlink" title="Return to login page.">Return</a> to login page.
		</div>
	</div>
	<div class="clearfix"></div>
	<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../includes/PrepDB.asp"-->
<!--#include file="../includes/Mail.asp"-->
<!--#include file="../includes/ResetPasswordInclude.asp"-->
