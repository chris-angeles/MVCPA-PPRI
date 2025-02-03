<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"--><% 
'Disable Page
'Response.End

Dim debug, i, SystemID, UserID, Name, FirstName, MiddleName, LastName, Suffix, Title, _
	email, email2, Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, _
	recaptcha_public_key, recaptcha_private_key, g_recaptcha_response, captcha_result
recaptcha_public_key = "6LdCvh4UAAAAAEBmrEUz6t6yTdXN0ybyNaRQZH1i" ' your public key
recaptcha_private_key = "6LdCvh4UAAAAAOsIcVBxRrKwLSN1EQW3kp3AgeDv" ' your private key
SystemID = 0
debug = False
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
End If

If Request.Form.Count>0 Then
	' validata reCAPTCHA
	g_recaptcha_response = Request.Form("g-recaptcha-response")
	captcha_result = recaptcha_confirm(g_recaptcha_response)
	If Debug = True Then
		Response.Write("<pre>captcha_result=" & captcha_result & "</pre>" & vbCrLf)
	End If
	If captcha_result = false Then
		Response.Write("Error: failed reCAPTCHA check.")
		Response.End
	End If
	FirstName = Request.Form("FirstName")
	MiddleName = Request.Form("MiddleName")
	LastName = Request.Form("LastName")
	Name = Request.Form("Name")
	Suffix = Request.Form("Suffix")
	Title = Request.Form("Title")
	email = Request.Form("email")
	email2 = Request.Form("email2")
	Address1 = Request.Form("Address1")
	Address2 = Request.Form("Address2")
	City = Request.Form("City")
	State = Request.Form("State")
	ZIP = Request.Form("ZIP")
	Phone = Request.Form("Phone")
	Fax = Request.Form("Fax")
	Mobile = Request.Form("Mobile")

	If email <> email2 Then
		Response.Write("The emails must match!")
		Response.End
	End If
	If Len(email)=0 Or Len(firstname)=0 Or Len(lastname)=0 Then
		Response.Write("An email address, first name, and last name must be provided.")
		Response.End
	End If
	If InStr(1, email, "@")=0 Or InStr(1, email, ".")=0 Or Len(email)<5 Then
		Response.Write("Invalid email address 1")
		Response.End
	End If
'	If checkEmail(email) Then
'		Response.Write("Invalid email address 2")
'		Response.End
'	End If
	If Len(UserID)>254 Then
		Response.Write("Invalid email address 3")
		Response.End
	End If
	If Len(FirstName)>24 Then
		FirstName = Left(FirstName,24)
	End If
	If Len(LastName)>24 Then
		LastName = Left(LastName,24)
	End If
	If Len(MiddleName)>24 Then
		MiddleName = Left(MiddleName,24)
	End If
	If Len(Suffix)>20 Then
		Suffix = Left(Suffix,20)
	End If
	If Len(Title)>100 Then
		Title = Left(Title,100)
	End If
	If Len(Address1)>100 Then
		Address1 = Left(Address1,100)
	End If
	If Len(Address2)>100 Then
		Address2 = Left(Address2,100)
	End If
	If Len(City)>100 Then
		City = Left(City,100)
	End If
	If Len(State)>2 Then
		State = Left(State,2)
	End If
	If Len(ZIP)>10 Then
		ZIP = Left(Zip,100)
	End If
	If Len(Phone)>25 Then
		Phone = Left(Phone,25)
	End If
	If Len(FAX)>25 Then
		FAX = Left(FAX,25)
	End If
	If Len(Mobile)>25 Then
		Mobile = Left(Mobile,25)
	End If
	sql = "SELECT UserID, Email, Name FROM System.Users WHERE UserID=" & prepStringSQL(email) & _
		" OR email=" & prepStringSQL(email)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then 
		If Debug = True Then
			Response.Write("<a href=""../Default.asp?Message=A user with this email address already exists in system. Please login to create a grant application. You may do a pasword reset if you cannot remember your password."">Redirect to login page</a>")
		Else
			Response.Redirect("../Default.asp?Message=A user with this email address already exists in system. Please login to create a grant application. You may do a pasword reset if you cannot remember your password.")
		End If
	End If

	sql = "SELECT UserID, Email, Name FROM System.Users WHERE FirstName=" & _
		prepStringSQL(firstname)
	If Len(middlename)=0 Then
		sql = sql & " AND MiddleName IS NULL "
	Else
		sql = sql & " AND MiddleName=" & prepStringSQL(middlename)
	End If
	sql = sql & " AND lastname=" & prepStringSQL(lastname)
	If Len(suffix) = 0 Then
		sql = sql & " AND Suffix IS NULL "
	Else
		sql = sql & " AND suffix=" & prepStringSQL(suffix)
	End If
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then 
		Response.Write("A user with this name already exists in system. Please login to create application if you are that person. Otherwise, use Back and add a middle initial or other change to make the name unique within the system.")
		Response.End
	End If

	sql = "INSERT INTO System.Users (UserID, FirstName, MiddleName, LastName, Suffix, Title, email, Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, UpdateID, UpdateTimestamp) " & vbCrLF & _
		"VALUES (" & prepStringSQL(Email) & _
		", " & prepStringSQL(FirstName) & _
		", " & prepStringSQL(MiddleName) & _
		", " & prepStringSQL(LastName) & _
		", " & prepStringSQL(Suffix) & _
		", " & prepStringSQL(Title) & _
		", " & prepStringSQL(email) & _
		", " & prepStringSQL(Address1) & _
		", " & prepStringSQL(Address2) & _
		", " & prepStringSQL(City) & _
		", " & prepStringSQL(State) & _
		", " & prepStringSQL(ZIP) & _
		", " & prepStringSQL(Phone) & _
		", " & prepStringSQL(Fax) & _
		", " & prepStringSQL(Mobile) & _
		", -1" & _
		", getdate()) "

	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
Else
	Response.Write("Error: Step 2 with no data submitted")
	Response.End

End If

sql = "SELECT SystemID, UserID, Name, Email " & vbCrLf & _
	"FROM System.Users " & vbCrLF & _
	"WHERE email=" & prepStringSQL(Email)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
If rs.EOF = False Then
	SystemID = rs.Fields("SystemID")
	UserID = rs.Fields("UserID")
	Name = rs.Fields("Name")
	If Debug = True Then
		Response.Write("<pre>SystemID=" & SystemID & ", UserID=" & UserID & ", Name=" & Name & "</pre>" & vbCrLF)
		Response.Flush
	End If
Else
	Response.Write("Error: Unable to retrive new user.")
	Response.End
End If

resetPassword(UserID)

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Step 2: Select or create your organizaton within the system</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function setFocus()
	{
		if (document.login.UserID.value.length == 0) {
			document.login.UserID.focus();
		}
		else {
			document.login.password.focus();
		}
	}

	function validate()
	{
		if (document.login.UserID.value.length <= 0) {
			alert("You must enter a user ID.");
			login.UserID.focus();
			return false;
		}
		if (document.login.password.value.length <= 0) {
			alert("You must enter a password");
			login.password.focus();
			return false;
		}
		if (document.login.UserID.value.length < 3) {
			alert("The user ID must be at least three characters.");
			login.UserID.focus();
			return false;
		}
		if (document.login.password.value.length < 3) {
			alert("The password must be at least three characters.");
			login.password.focus();
			return false;
		}
		return true;
	}

	function setClientInfo()
	{
		document.login.ScreenSize.value = window.screen.width + "x" + window.screen.height + "x" + window.screen.colorDepth;
		document.login.WindowSize.value = window.screen.availWidth + "x" + window.screen.availHeight;
		document.login.AvailableSize.value = window.innerWidth + "x" + window.innerHeight;
		document.login.appName.value = navigator.appName + "; " + navigator.userAgent + ";" + navigator.platform
		if (navigator.appName.indexOf('Netscape') != -1 || navigator.userAgent.indexOf('Opera') != -1) {
			document.login.AvailableSize.value = window.innerWidth + "x" + window.innerHeight;
		}
		else {
			document.login.AvailableSize.value = top.document.body.clientWidth + "x" + top.document.body.clientHeight;
		}
	}
</script>
</head>
<body onload="setClientInfo();setFocus()">
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Step 2: Select or create your organizaton within the system after you are able to login.</div>


<div class="content">
<p>A new user in the system has been created for you. Your username is your email address, 
<%=email %>.</p>

<p>A password reset has been initiated and that password has been emailed to you.
You should recieve the password within a few minutes. 
You may use that password to login and continue the application process. You may
change your password once you login to something easier for you to remember.
</p>

<div class="section">
	<form name="login" id="login" method="post" action="../ValidateLogin.asp">
	<fieldset style="width: 580px;">
		<legend>Login</legend>
		<label for="UserID">User ID (email):</label>
		<input type="text" name="UserID" id="UserID" value="<%=UserID%>" size="50" maxlength="255" /><br />
		<label for="password">Password:</label>
		<input type="password" name="password" id="password" value="" size="36" maxlength="255" /><br />
		<input type="submit" value="Logon" />
	</fieldset>
	<input type="hidden" name="ScreenSize" value="">
	<input type="hidden" name="AvailableSize" value="">
	<input type="hidden" name="WindowSize" value="">
	<input type="hidden" name="appName" value="">
	</form>
</div>
</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<%
function recaptcha_confirm(reresponse)
	Dim VarString, ResponseString
	VarString = _
			"secret=" & recaptcha_private_key & _
			"&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _
			"&response=" & reresponse

	Dim objXmlHttp
	Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objXmlHttp.open "POST", "https://www.google.com/recaptcha/api/siteverify", False
	objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXmlHttp.send VarString
	ResponseString = objXmlHttp.responseText

	If Debug = True Then
		Response.Write("<pre>" & objXmlHttp.responseText & "</pre>" & vbCrLf)
	End If
	Set objXmlHttp = Nothing

	if InStr(1, ResponseString, """success"": true") then
		'They answered correctly
		recaptcha_confirm = true
	else
		'They answered incorrectly
		recaptcha_confirm = false
	end if
end function
%>
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/ResetPasswordInclude.asp"-->
<!--#include file="../includes/Mail.asp"-->