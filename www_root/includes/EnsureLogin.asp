<%
' Check to be sure that user is logged in. If they are logged in, the System_ID Session variable will be set.
' If they are not logged in, redirect to the login page.
' Check for Administrative and Member Status and set value for page.


Dim Developer, MVCPAAdministrator, MVCPAAuditor, MVCPAGrantCoordinator, _
	MVCPAAdministrativeAssistant, MVCPAScorer, MVCPAViewer, MVCPARights, _
	UserSystemID, UserName, UserEmail, UserFiscalYear, UserGranteeID, SessionID
UserSystemID = Session("SystemID")
if UserSystemID = 0 then
	If Request.Cookies("SystemID") = "" Then
		Response.Redirect("../default.asp?message=Please+login+to+site")
	End If
	UserSystemID = ReLogIn(Request.Cookies("SystemID"))
End If
UserName = Session("Name")
UserEmail = Session("Email")
Developer = Session("Developer")
MVCPAAdministrator = Session("MVCPAAdministrator")
MVCPAAuditor = Session("MVCPAAuditor")
MVCPAGrantCoordinator = Session("MVCPAGrantCoordinator")
MVCPAAdministrativeAssistant = Session("MVCPAAdministrativeAssistant")
MVCPARights = Session("MVCPARights")
MVCPAScorer = Session("MVCPAScorer")
MVCPAViewer = Session("MVCPAViewer")
If Len(Session("FiscalYear"))>0 Then
	UserFiscalYear = Session("FiscalYear")
End If
UserGranteeID = Session("GranteeID")

if UserSystemID = 0 then
	SendMessage

	Response.Redirect("../default.asp?message=Please+login+to+site")
End If

Function ReLogIn(vSystemID)
%><!--#include file="OpenConnection.asp"--><%
	Dim SystemID, UserID, Name, Title, EMail, AccountDisabled, _
	Developer, MVCPAAdministrator, MVCPAGrantCoordinator, MVCPAAdministrativeAssistant, Scorer, MVCPARights, LastLogin
	
	sql = "SELECT SystemID, UserID, Name, Title, EMail, AccountDisabled, " & vbCrLf & _
		"	Developer, MVCPAAdministrator, MVCPAAuditor, MVCPAGrantCoordinator, " & vbCrLF & _
		"	MVCPAAdministrativeAssistant, MVCPAScorer, MVCPAViewer, " & vbCrLF & _
		"	CAST(CASE WHEN Developer=1 OR MVCPAAdministrator=1 OR MVCPAAuditor=1 OR MVCPAGrantCoordinator=1 OR MVCPAAdministrativeAssistant=1 THEN 1 ELSE 0 END AS BIT) AS MVCPARights, " & vbCrLF & _
		"	(Select Max(LoginTime) FROM System.LoginLog WHERE System.LoginLog.SystemID=System.Users.SystemID) AS LastLogin " & _
		"FROM System.Users " & vbCrLF & _
		"WHERE SystemID=" & vSystemID
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		SystemID = rs.Fields("SystemID")
		UserID = rs.Fields("UserID")
		Name = rs.Fields("Name")
		Title = rs.Fields("Title")
		EMail = rs.Fields("EMail")
		AccountDisabled = rs.fields("AccountDisabled")
		Developer = rs.Fields("Developer")
		MVCPAAdministrator = rs.Fields("MVCPAAdministrator")
		MVCPAAuditor = rs.Fields("MVCPAAuditor")
		MVCPAGrantCoordinator = rs.Fields("MVCPAGrantCoordinator")
		MVCPAAdministrativeAssistant = rs.Fields("MVCPAAdministrativeAssistant")
		MVCPAScorer = rs.Fields("MVCPAScorer")
		MVCPAViewer = rs.Fields("MVCPAViewer")
		MVCPARights = rs.Fields("MVCPARights")
		LastLogin = rs.Fields("LastLogin")

		If LastLogin < Date() Then
			Response.Redirect("../default.asp?message=Your session has been inactive too long and you have been logged out.")
		End If
			
		If AccountDisabled = True Then
			Response.Redirect("../default.asp?message=The account has been disabled.")
		End If

		Session("SystemID") = SystemID
		Session("UserID") = UserID
		Session("Name") = Name
		Session("Title") = Title
		Session("EMail") = EMail
		Session("Developer") = Developer
		Session("MVCPAAdministrator") = MVCPAAdministrator
		Session("MVCPAAuditor") = MVCPAAuditor
		Session("MVCPAGrantCoordinator") = MVCPAGrantCoordinator
		Session("MVCPAAdministrativeAssistant") = MVCPAAdministrativeAssistant
		Session("MVCPAScorer") = MVCPAScorer
		Session("MVCPAViewer") = MVCPAViewer
		Session("MVCPARights") = MVCPARights
		Session("FiscalYear") = Request.Cookies("FiscalYear")
		Session("GranteeID") = Request.Cookies("GranteeID")

		rs.close
		sql = "INSERT INTO System.LoginLog (SessionID, SystemID, LoginTime, ipaddress, SessionRecovery) VALUES (" & _
			Session.SessionID & ", " & SystemID & ", getdate(), " & _
			prepStringSQL(Request.ServerVariables("REMOTE_ADDR")) & ", 1)"
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
	con.execute(sql)


	Else
		Response.Redirect("../default.asp?message=Please+login+to+site")
	End If
	ReLogIn = SystemID
	Session("GranteeID") = Request.Cookies("GranteeID")
End Function

Sub SendMessage(vMessage)
	'on error resume next
	'********************************
	'Send error message to webmaster
	'********************************
		dim ObjMail, Sender, Recipient, Recipient2, Subject, Body, strItem, strItemKey
		Body = "<table border=0>" & vbCrLf
		Body = Body & "<tr><td>Date/Time: </td><td>" & Now() & "</td></tr>" & vbCrLf
		Body = Body & "<tr><td>Site: </td><td>http://" & Request.ServerVariables("SERVER_NAME") &"</td></tr>" & vbCrLf
		Body = Body & "<tr><td>Script: </td><td>" & Request.ServerVariables("SCRIPT_NAME") &"</td></tr>" & vbCrLf
		Body = Body & "<tr><td>Message</td><td>" & vMessage & "</td></tr>" & vbCrLf
		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			Body = Body & "<tr><td>QueryString: </td><td>" & Request.ServerVariables("QUERY_STRING") & "</td></tr>" & vbCrLf
		End If
		If Request.QueryString.Count > 0 then
			Body = Body & vbCrLf & "<tr><td><b>QueryString:</b></td><td></td></tr>" & vbCrLF
		For Each strItem in Request.QueryString
			Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Request.QueryString(strItem) & "</td></tr>" & vbCrLf
		Next
			Body = Body & vbCrLf
		End If

		on error resume next		  	
		If Request.Form.Count > 0 then
			Body = Body & vbCrLf & "<tr><td><b>Form Variables:</b></td><td>" & "</td></tr>" & vbCrLF
			For Each strItem in Request.Form
				Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Request.Form(strItem) & "</td></tr>" & vbCrLf
			Next
			Body = Body & vbCrLf
		End If
		on error goto 0

		if Session.Contents.Count > 0 Then
			Body = Body & vbCrLf & "<tr><td><b>Session Variables:</b></td><td>" & vbCrLf
			For each strItem in Session.Contents
				Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Session.Contents(strItem) & "</td></tr>" & vbCrLf
			Next
		End If
		
		If Request.Cookies.Count > 0 Then
			Body = Body & vbCrLf & "<tr><td><b>Cookies:</b> (" & Request.Cookies.Count & ")</td><td></td></tr>" & vbCrLF
			For Each strItem in Request.Cookies
				If Request.Cookies(strItem).HasKeys Then
					For Each strItemKey in Request.Cookies(strItem)
						Body = Body & "<tr><td>    " & strItem & "(" & strItemKey & "): </td><td>" & Request.Cookies(strItem)(strItemKey) &"</td></tr>" &  vbCrLf
					Next
				Else
					Body = Body & "<tr><td>    " & strItem & ": </td><td>" & Request.Cookies(strItem) & "</td></tr>" & vbCrLf
				End If
			Next
			Body = Body & vbCrLf
		ELse
			Body = Body & vbCrLf & "<tr><td><b>Cookies:</b></td><td>None</td></tr>" & vbCrLF & vbCrLf
		End If
		  
		If Len(Trim(Request.ServerVariables("HTTP_REFERER"))) > 0 Then
			Body = Body & vbCrLf & "<tr><td>Referer: </td><td>" & Request.ServerVariables("HTTP_REFERER") & "</td></tr>" & vbCrLF
		End If
		  
		Body = Body & "<tr><td>Remote IP: </td><td>" & Request.ServerVariables("REMOTE_ADDR") & "</td></tr>" & vbCrLf
		Body = Body & "<tr><td>Browser: </td><td>" & Request.ServerVariables("HTTP_USER_AGENT") & "</td></tr>" & vbCrLf
		Body = Body & "</table>" & vbCrLF
		Sender = "NoReply@" & Request.ServerVariables("SERVER_NAME")
		Recipient = "mvcpa@tamu.edu"
		Recipient2 = ""
		If Len(Request.Cookies("System_ID"))>0 Then
			Subject = Request.ServerVariables("SERVER_NAME") & " Warning Message from " & Request.ServerVariables("REMOTE_ADDR")
		Else
			Subject = Request.ServerVariables("SERVER_NAME") & " Warning Message from " & Request.ServerVariables("REMOTE_ADDR")
		End If
		SendMail Sender, Recipient, Recipient2, Subject, Body	  
	'********************************
	on error goto 0 
End Sub

Function SendMail(vSender, vRecipient, vRecipient2, vSubject, vBody)
	'Messaging - build transport configuration
	Dim iMsg
	Dim iConf
	Dim Flds
	Dim strHTML

	Const cdoSendUsingPickup = 1	'Use local SMTP service using pickup directory
	Const cdoSendUsingPort = 2		'Use network SMTP service

	set iMsg = CreateObject("CDO.Message")
	set iConf = CreateObject("CDO.Configuration")

	Set Flds = iConf.Fields
	With Flds
		'Local SMTP service using pickup directory
		'.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
		'.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup"
		
		'Network SMTP service
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.tamu.edu"
	
		.Update
	End With

	'Messaging - build HTML
	strHTML = "<html lang=""en-us"">"
	strHTML = strHTML & "<head></head>"
	strHTML = strHTML & "<body>"
	strHTML = strHTML & vBody
	strHTML = strHTML & "</body>"
	strHTML = strHTML & "</html>"

	If debug = True then
		vRecipient = "mvcpa@tamu.edu"
	End if
		
	'Messaging - apply seetings to message
	With iMsg
		Set .Configuration = iConf
		.To = vRecipient
		'Messaging - determine/assign carbon copy
		If vRecipient <> "mvcpa@tamu.edu" and vRecipient2 <> "mvcpa@tamu.edu" then
			.BCC = "mvcpa@tamu.edu"		'assign to person monitoring system
		End If
		If Len(vRecipient2)>0 Then
			.CC = vRecipient2
		End if
		.From = vSender
		.Subject = vSubject
		.HTMLBody = strHTML
		If debug=False then 
			.Send
		End if
	End With

	'Cleanup variables
	Set iMsg = Nothing
	Set iConf = Nothing
	Set Flds = Nothing

End Function

Function sendWarning(vMessage)
	On Error Resume Next
	sendWarning = False
	Dim message
	message = "Page Message: " & vMessage & "<br /><br />" & vbCrLf
	message = message & "<pre>Dubugging Information: " & vbCrLf
	message = message & "Time: " & Now() & vbCrLf
	message = message & "Server: " & Request.ServerVariables("SERVER_NAME") & vbCrLf
	message = message & "Request Method: " & Request.ServerVariables("REQUEST_METHOD") & vbCrLf
	If Request.ServerVariables("SERVER_PORT_SECURE")="1" Then
		message = message & "Raw URL: : https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("HTTP_URL") & vbCrLf
	Else
		message = message & "Raw URL: : http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("HTTP_URL") & vbCrLf
	End If
	message = message & "Path: " & Request.ServerVariables("PATH_INFO") & vbCrLf
	message = message & "Path translated: " & Request.ServerVariables("PATH_TRANSLATED") & vbCrLf
	message = message & "Referrer: " & Request.ServerVariables("HTTP_REFERER") & vbCrLf
	message = message & "ConnectionString=" & Application("ConnectionString")
	For each i in Request.Form
		message = message & "Request.Form(""" & i & """)='" & Request.Form(i) & "'" & vbCrLf
	Next
	For each i in Request.QueryString
		message = message & "Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'" & vbCrLf
	Next
	For each i in Session.Contents
		message = message & "Session(""" & i & """)='" & Session(i) & "'" & vbCrLf
	Next
	for each i in Request.Cookies
		if Request.Cookies(i).HasKeys then
			for each j in Request.Cookies(x)
				message = message & "Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j)
			next
		else
			message = message & "Cookies(""" & i & """)=" & Request.Cookies(i) & "<br>"
		end if
	next
	If Len(sql)>0 Then
		message = message & "last sql:" & vbCrLf & sql & vbCrLf
	End If
	message = message & "</pre>" & vbCrLF
	SendMail "TxMVCPAWebsite <mvcpa@tamu.edu>", "mvcpa@tamu.edu", "", "Warning Message on " & LCase(Request.ServerVariables("SERVER_NAME")), message
	On Error Goto 0
	sendWarning = True
End Function
%>