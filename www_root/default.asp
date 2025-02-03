<%@ language=VBScript %>
<% Option Explicit
Dim Debug, i, UserID, message, strSecureURL
Debug = False

If Debug = True Then
	Response.Write("<pre>")
	for each i in Request.ServerVariables
	  Response.Write(i & " = " & Request.ServerVariables(i) & vbCrLf)
	next
	Response.Write("</pre>")
	Response.Flush
End If

If InStr(Request.ServerVariables("HTTP_URL"),"txwatchyourcar.com") Then
	Response.Redirect "https://www.txdmv.gov/motorists/consumer-protection/auto-theft-prevention/"
ElseIf Request.ServerVariables("SERVER_NAME") = "txwatchyourcar.com" Then
	Response.Redirect "https://www.txdmv.gov/motorists/consumer-protection/auto-theft-prevention/"
ElseIf Request.ServerVariables("SERVER_NAME") = "abtpa.tamu.edu" Then
	Response.Redirect "https://mvcpa.tamu.edu/" 
ElseIf Request.ServerVariables("SERVER_PORT")=80 Then
    strSecureURL = "https://"
    strSecureURL = strSecureURL & Request.ServerVariables("SERVER_NAME")
    strSecureURL = strSecureURL & Request.ServerVariables("URL")
	If Len(Request.ServerVariables("QUERY_STRING"))>0 Then
		strSecureURL = strSecureURL & "?" & Request.ServerVariables("QUERY_STRING")
	End If
    Response.Redirect strSecureURL
End If

Session.Abandon()

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

If Len(Request.QueryString("UserID"))>0 Then
	UserID = Request.QueryString("UserID")
ElseIf Len(Request.Cookies("UserID"))>0 Then
	UserID=Request.Cookies("UserID")
End If
If Len(Request.QueryString("message"))>0 Then
	message = Server.URLEncode(Request.QueryString("message"))
End IF
	
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Texas Motor Vehicle Crime Prevention Authority</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" /> 
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
<div class="header" title="Motor Vehicle Crime Prevention Authority logo banner. Outline of a car with eyes below and text Watch Your Car" onclick="window.open('https://www.txdmv.gov/about-us/MVCPA', '_blank');"></div>
<div class="pagetag">The <strong><a href="https://www.txdmv.gov/about-us/MVCPA" target="_blank" style="text-decoration: none; color: white; ">Motor Vehicle Crime Prevention Authority</a></strong> (MVCPA) 
	awards financial grants to agencies, organizations, and concerned parties in an effort to 
	raise public awareness of vehicle theft and burglary and implement education and prevention 
	initiatives.</div>

<div class="content">

<!--<div style="color: red; font-weight: bold; font-size: x-large">The Website is down temporarily. We are working to resolve the problem.</div>-->
<%
If Len(message)>0 Then
	Response.Write("<div class=""notice"">" & message & "</div>" & vbCrLf)
End If
%>
<div class="section">
	<form name="login" id="login" method="post" action="ValidateLogin.asp">
	<fieldset style="width: 100%;">
		<legend>Login</legend>
		<label for="UserID">User ID (email):</label>
		<input type="text" name="UserID" id="UserID" value="<%=UserID%>" size="45" maxlength="255" spellcheck="false" /><br />
		<label for="password">Password:</label>
		<input type="password" name="password" id="password" value="" size="36" maxlength="255" spellcheck="false" />
	</fieldset>
	<div style="text-align: center; width: 732px"><input type="submit" value="Logon" /></div>
	<input type="hidden" name="ScreenSize" value="">
	<input type="hidden" name="AvailableSize" value="">
	<input type="hidden" name="WindowSize" value="">
	<input type="hidden" name="appName" value="">
	</form>
</div>

<div class="section">
	<fieldset style="width: 100%">
		<legend>Password Reset</legend>
		<p>If you have forgotten your password, your password can be reset
<a href="/user/ResetPassword.asp" class="plainlink" title="Link to reset password page.">here</a>.</p>
</fieldset>
</div>
<%	If True = False Then %>
<div class="section">
	<fieldset style="width: 100%">
		<legend>New Applications</legend>
<% If Date() < "4/19/2019"  and Application("Instance") = "Production" Then %>
	
		<p>The system will be available on April 19th for the fiscal year 2020 grant cycle.</p>

		<p>A grant application workshop is being held April 18th from 8:30am - 4:30 pm at the 
		Norris Conference Center,  2525 W Anderson Lane, Austin, 78757.  To register please 
		call MVCPA at 1-800-CAR-WATCH (1-800-227-9282).</p>
		
		<p>The Request for Application can be viewed 
		<a href="/RFA/RFA2020-21.pdf" class="plainlink" title="Link to Request for Application">here</a>.</p>
</div>
<%	Else %>

		<p>Law enforcement agencies, local prosecutors, judicial agencies, and neighborhood, 
		community, business, and nonprofit organizations for programs designed to reduce the 
		incidence of economic automobile theft are eligible to apply for grants for automobile 
		burglary and theft prevention assistance projects.</p>

		<p>If you are an existing user of this system, please login above. If you represent an
		agency or organization that has not applied for one of these grants before, then you 
		must first create a user account for yourself. Then you create an agency/organization 
		within the system and then start to apply for a grant. Begin by creating a user
		<a href="/User/SelfServeUser.asp" class="plainlink" title="Link to begin application process">here</a>.</p>

		<p>The Request for Application can be viewed 
		<a href="/RFA/RFA2020-21.pdf" class="plainlink" target="_blank" title="Link to Request for Application">here</a>.</p>
<%	End If %>
	</fieldset>
</div>
<%	End If %>
</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>