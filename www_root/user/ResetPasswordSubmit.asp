<%@  language="VBScript" %>
<% Option Explicit%><!--#include file="../includes/adovbs.asp"--><!--#include file="../includes/OpenConnection.asp"--><% 
Dim Debug, i, SystemID, Name, email, UserID, Message, Source, _
	recaptcha_public_key, recaptcha_private_key, g_recaptcha_response, captcha_result
If Application("Instance") = "Test" Then
	recaptcha_public_key = "6Ld36okpAAAAAP35zWtqBsy0rx8d41pWK4utB6GG" ' your public key
	recaptcha_private_key = "6Ld36okpAAAAAO2RPpMbOiIxZBsNG_to-eh0DC5b" ' your private key
Else
	recaptcha_public_key = "6LdCvh4UAAAAAEBmrEUz6t6yTdXN0ybyNaRQZH1i" ' your public key
	recaptcha_private_key = "6LdCvh4UAAAAAOsIcVBxRrKwLSN1EQW3kp3AgeDv" ' your private key
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

If Request.Form.Count>0 Then
	UserID = Request.Form("UserID")
	Source = Request.Form("Source")
	g_recaptcha_response = Request.Form("g-recaptcha-response")
	If Len(UserID)=0 Then
		Response.Write("No UserID was provided. Operation cancelled.")
		Response.End
	End If
	If InStr(1,UserID,"<")>0 Then
		Response.Write("Invalid User ID. Processing stopped.")
		Response.End
	End If
	If InStr(1,UserID,">")>0 Then
		Response.Write("Invalid User ID. Processing stopped.")
		Response.End
	End If
	If InStr(1,UserID,"&")>0 Then
		Response.Write("Invalid User ID. Processing stopped.")
		Response.End
	End If
	If InStr(1,UserID,";")>0 Then
		Response.Write("Invalid User ID. Processing stopped.")
		Response.End
	End If
	If InStr(1, UserID, "@")=0 Or InStr(1, UserID, ".")=0 Or Len(UserID)<5 Then
		If Debug = True Then
			Response.Write("<a href=""ResetPassword.asp?UserID=" & UserID & "&Message=You must provide a valid User ID to do a password reset. It is generally your email address."">Continue</a")
		Else
			Response.Redirect("ResetPassword.asp?UserID=" & Server.URLEncode(UserID) & "&Message=You must provide a valid User ID to do a password reset. It is generally your email address.")
		End If
		Response.End
	End If
	' validata reCAPTCHA
	captcha_result = recaptcha_confirm(g_recaptcha_response)
	If Debug = True Then
		Response.Write("<pre>captcha_result=" & captcha_result & "</pre>" & vbCrLf)
	End If
	If captcha_result = false Then
		Response.Write("Error: failed reCAPTCHA check.")
		Response.End
	End If
	If checkEmail(UserID) Then
		Message = resetPassword(UserID)
	Else
		If Debug = True Then
			Response.Write("<a href=""ResetPassword.asp?UserID=" & UserID & "&Message=You must provide a valid User ID to do a password reset. It is generally your email address."">Continue</a")
		Else
			Response.Redirect("ResetPassword.asp?UserID=" & Server.URLEncode(UserID) & "&Message=You must provide a valid User ID to do a password reset. It is generally your email address.")
		End If
		Response.End
	End If
	If Source = "UpdateUser3.asp" Then
%><html lang="en-us">
<head>
	<title>Password Reset</title>
</head>
<body>
	The password has been reset and email sent to <%=UserID %>.
	<input type="button" value="Close Window" onclick="window.close();" />
	<script type="text/javascript">window.close();</script>
</body>
</html>
<%
		Response.End
	End If
	If Len(Message)>0 Then
		If Debug = True Then
			Response.Write("<a href=""../default.asp?UserID=" & UserID & "&Message=" & Message & """>../default.asp?UserID=" & UserID & "</a>" & vbCrLf)
			Response.End
		Else
			Response.Clear
			Response.Redirect("../default.asp?UserID=" & UserID & "&Message=" & Message)
		End If
	ElseIf Len(Message)=0 Then
		If Debug = True Then
			Response.Write("<a href=""../default.asp?UserID=" & UserID & "&Message=Your password has been sent to your email address" & email & """>../default.asp?UserID=" & UserID & "</a>" & vbCrLf)
			Response.End
		Else
			Response.Clear
			Response.Redirect("../default.asp?UserID=" & UserID & "&Message=Your password has been sent to your email address" & email & ". Although sent immediately, it can sometimes take several minutes or longer to arrive.")
		End If
	End If
End If

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

%><!--#include file="../includes/PrepDB.asp"-->
<!--#include file="../includes/Mail.asp"-->
<!--#include file="../includes/ResetPasswordInclude.asp"-->
