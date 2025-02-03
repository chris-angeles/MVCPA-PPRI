<%@ language=VBScript %><% Option Explicit

'Disable Page
'Response.End

Dim recaptcha_public_key
recaptcha_public_key = "6LdCvh4UAAAAAEBmrEUz6t6yTdXN0ybyNaRQZH1i" ' your public key
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Create a user account by entering your name and contact information</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" /> 
<script type="text/javascript">
	function validateForm()
	{
		if (document.SelfServeUser.FirstName.value.length == 0) {
			alert("You must enter a first name for the user.");
			document.SelfServeUser.FirstName.focus();
			return false;
		}
		if (document.SelfServeUser.LastName.value.length == 0) {
			alert("You must enter a last name for the user.");
			document.SelfServeUser.LastName.focus();
			return false;
		}
		if (document.SelfServeUser.email.value.length == 0) {
			alert("You must enter an email address that you have access to.");
			document.SelfServeUser.LastName.focus();
			return false;
		}
		if (document.SelfServeUser.email.value != document.SelfServeUser.email2.value) {
			alert("The email addresses must match.");
			document.SelfServeUser.LastName.focus();
			return false;
		}
		if (grecaptcha.getResponse() == "") {
			alert("You must complete the CAPTCHA to submit this form!");
			return false;
		}
		return true;
	}
</script>
<script src="https://www.google.com/recaptcha/api.js" async defer></script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Create a user account by entering your name and contact information.</div>


<div class="content">

<form name="SelfServeUser" id="SelfServeUser" method="post" action="SelfServeUserSubmit.asp" onSubmit="return validateForm();">
	<fieldset style="width: 100%;">
		<legend>Name and Title</legend>
		<label for="FirstName">First Name:</label>
		<input type="text" name="FirstName" id="FirstName" value="" size="24" maxlength="24" /><br />
		<label for="MiddleName">Middle Name:</label>
		<input type="text" name="MiddleName" id="MiddleName" value="" size="24" maxlength="24" /><br />
		<label for="LastName">LastName:</label>
		<input type="text" name="LastName" id="LastName" value="" size="24" maxlength="24" /><br />
		<label for="Suffix">Suffix:</label>
		<input type="text" name="Suffix" id="Suffix" value="" size="20" maxlength="20" /><br />
		<label for="Title">Position Title:</label>
		<input type="text" name="Title" id="Title" value="" size="50" maxlength="100" /><br />
	</fieldset>

	<fieldset style="width: 100%;">
		<legend>Contact Information</legend>
		<label for="email">E-Mail Address:</label>
		<input type="text" name="email" id="email" value="" size="50" maxlength="255" /><br />
		<label for="email2">Repeat E-Mail Address:</label>
		<input type="text" name="email2" id="email2" value="" size="50" maxlength="255" /><br />
		<label for="Address1">Address 1:</label>
		<input type="text" name="Address1" id="Address1" value="" size="50" maxlength="50" /><br />
		<label for="Address2">Address 2:</label>
		<input type="text" name="Address2" id="Address2" value="" size="50" maxlength="50" /><br />
		<label for="City">City:</label>
		<input type="text" name="City" id="City" value="" size="20" maxlength="20" /><br />
		<label for="State">State:</label>
		<input type="text" name="State" id="State" value="" size="2" maxlength="2" /><br />
		<label for="ZIP">ZIP:</label>
		<input type="text" name="ZIP" id="ZIP" value="" size="10" maxlength="10" /><br />
		<label for="Phone">Phone:</label>
		<input type="text" name="Phone" id="Phone" value="" size="20" maxlength="20" /><br />
		<label for="Fax">Fax:</label>
		<input type="text" name="Fax" id="Fax" value="" size="20" maxlength="20" /><br />
		<label for="Fax">Mobile:</label>
		<input type="text" name="Mobile" id="Mobile" value="" size="20" maxlength="20" /><br />
	</fieldset>
	<div class="g-recaptcha" data-callback="recaptchaCallback" data-sitekey="<%=recaptcha_public_key %>"></div>
	<div style="text-align: center;"><input type="submit" value="Submit" /></div>
</form>

</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../includes/prepDB.asp"-->