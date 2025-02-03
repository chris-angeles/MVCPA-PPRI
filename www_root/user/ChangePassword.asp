<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i
Dim OldPassword, NewPassword, Timestamp, message, strSecureURL
If Request.ServerVariables("SERVER_PORT")=80 Then
    strSecureURL = "https://"
    strSecureURL = strSecureURL & Request.ServerVariables("SERVER_NAME")
    strSecureURL = strSecureURL & Request.ServerVariables("URL")
	If Len(Request.ServerVariables("QUERY_STRING"))>0 Then
		strSecureURL = strSecureURL & "?" & Request.ServerVariables("QUERY_STRING")
	End If
    Response.Redirect strSecureURL
End If

debug = false
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
End If

If Request.Form.Count>0 Then

	Timestamp = Now()
	If Request.Form.Count>0 Then
		OldPassword = Request.Form("OldPassword")
		NewPassword = Request.Form("NewPassword")
		sql = "SELECT SystemID from System.Users WHERE SystemID=" & UserSystemID & " AND passwordhash=HASHBYTES('SHA2_256'," & prepUnicodeSQL(OldPassword) & ")"
		Set rs = con.Execute(sql)
		If rs.BOF = True Then
			Message = "What you typed for old password did not match password in system."
		Else
			sql = "UPDATE System.Users " & vbCrLF & _
				"SET PasswordHash=HASHBYTES('SHA2_256'," & prepUnicodeSQL(NewPassword) & "), LastPasswordChange=getdate() " & vbCrLf & _
				"WHERE SystemID=" & UserSystemID
			con.Execute(sql)
			Set rs = nothing
			Set Con = nothing
			Response.Redirect("../Home/Default.asp")
		End If
	End If
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Change Password</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" /> 
<SCRIPT TYPE="text/javascript">
	function validateForm()
	{
		if (document.ChangePassword.OldPassword.value == "") {
			alert("You must enter your existing password in order to change your password!");
			document.ChangePassword.OldPassword.focus();
			return false;
		}
		if (document.ChangePassword.NewPassword.value != document.ChangePassword.NewPassword2.value) {
			alert("The new password did not match the new password used for validation. Try again.");
			document.ChangePassword.NewPassword.value = "";
			document.ChangePassword.NewPassword2.value = "";
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		if (document.ChangePassword.OldPassword.value == document.ChangePassword.NewPassword.value) {
			alert("You cannot use the same password again. Try again.");
			document.ChangePassword.NewPassword.value = "";
			document.ChangePassword.NewPassword2.value = "";
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		return checkComplexity();
	}

	function checkComplexity()
	{
		if (document.ChangePassword.NewPassword.value.length < 8) {
			alert("The password must be at least eight characters in length.");
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		if (countUpper(document.ChangePassword.NewPassword.value) == 0) {
			alert("The password must be contain at least one uppercase letter.");
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		if (countLower(document.ChangePassword.NewPassword.value) == 0) {
			alert("The password must be contain at least one lowercase letter.");
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		if (countNum(document.ChangePassword.NewPassword.value) == 0) {
			alert("The password must be contain at least one number.");
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		if (countSpecial(document.ChangePassword.NewPassword.value) == 0) {
			alert("The password must be contain at least one special character from this list: (space) !#$%&'()*+,-./:;<=>?@[\]^_`{|}~.");
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		if (countInvalid(document.ChangePassword.NewPassword.value) > 0) {
			alert("The password must not contain an invalid character. Choose from uppercase letters, lowercase letters, numbers, and special characters in this list: (space) !#$%&'()*+,-./:;<=>?@[\]^_`{|}~.");
			document.ChangePassword.NewPassword.focus();
			return false;
		}
		return true;
	}

	function submitForm()
	{
		if (validateForm() == true) {
			document.ChangePassword.submit();
		}
	}
</SCRIPT>
<!--#include file="../includes/ComplexityValidation.asp"-->
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>
<div class="pagetag">Change Password for <%=UserName%></div>
<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">

<%
If Len(message)>0 Then
	Response.Write("<div class=""notice"">" & message & "</div>" & vbCrLf)
End If
%>

<form name="ChangePassword" id="ChangePassword" method="post" action="ChangePassword.asp" onSubmit="validateForm();">
	<fieldset style="width: 640px;">
		<legend>Change Password</legend>
		<label for="OldPassword">Old Password:</label>
		<input type="password" name="OldPassword" id="OldPassword" value="" size="36" maxlength="254" /><br />
		<label for="NewPassword">New Password:</label>
		<input type="password" name="NewPassword" id="NewPassword" value="" size="36" maxlength="254" onchange="return checkComplexity();" /><br />
		<label for="NewPassword2">Repeat New Password:</label>
		<input type="password" name="NewPassword2" id="NewPassword2" value="" size="36" maxlength="254" /><br />
	</fieldset>
<div style="text-align: center"><input type="button" value="Change Password" onclick="submitForm();" />
	<input type="button" value="Cancel" onclick="location.href = '../Default.asp';" />
</div>
</form>
</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/PrepDB.asp"-->