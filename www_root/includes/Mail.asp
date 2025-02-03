<%
Sub SendMail(vSender, vRecipient, vRecipient2, vSubject, vBody)

	Dim Debug
	
	Debug = False
	
	'Messaging - build transport configuration
	Dim iMsg
	Dim iConf
	Dim Flds
	Dim strBody

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

	strBody = vBody

	If debug = True then
		vRecipient = "mvcpa@tamu.edu"
		vRecipient2 = ""
	End if
		
	'Messaging - apply seetings to message
	With iMsg
		Set .Configuration = iConf
		.To = vRecipient
		'Messaging - determine/assign carbon copy
		If vRecipient2 = "" then
			'
		Else
			.CC = vRecipient2
		End if
		.BCC = "mvcpa@tamu.edu"
		.From = vSender
		.Subject = vSubject
		.TextBody = strBody
		.Send
	End With

	If debug = True then
		Response.Write("<br><br>" & vbCrLf)
		Response.write("From: " & iMsg.From & "<br>")
		Response.write("To: " & iMsg.To & "<br>")
		Response.write("CC: " & iMsg.CC & "<br>")
		Response.write("Subject: " & iMsg.Subject & "<br>")
		Response.write("Body:" & vbCrLf & iMsg.HTMLBody & "<br>")
		Response.flush
	End if

	'Cleanup variables
	Set iMsg = Nothing
	Set iConf = Nothing
	Set Flds = Nothing

End Sub

Sub SendHTMLMail(vSender, vRecipient, vRecipient2, vSubject, vBody)

	Dim Debug
	
	Debug = False
	
	'Messaging - build transport configuration
	Dim iMsg
	Dim iConf
	Dim Flds
	Dim strBody

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

	strBody = vBody

	If debug = True then
		vRecipient = "mvcpa@tamu.edu"
		vRecipient2 = ""
	End if
		
	'Messaging - apply seetings to message
	With iMsg
		Set .Configuration = iConf
		.To = vRecipient
		'Messaging - determine/assign carbon copy
		If vRecipient2 = "" then
			'
		Else
			.CC = vRecipient2
		End if
		.BCC = "mvcpa@tamu.edu"
		.From = vSender
		.Subject = vSubject
		.HTMLBody = strBody
		.Send
	End With

	If debug = True then
		Response.Write("<br><br>" & vbCrLf)
		Response.write("From: " & iMsg.From & "<br>")
		Response.write("To: " & iMsg.To & "<br>")
		Response.write("CC: " & iMsg.CC & "<br>")
		Response.write("Subject: " & iMsg.Subject & "<br>")
		Response.write("Body:" & vbCrLf & iMsg.HTMLBody & "<br>")
		Response.flush
	End if

	'Cleanup variables
	Set iMsg = Nothing
	Set iConf = Nothing
	Set Flds = Nothing

End Sub
%>
