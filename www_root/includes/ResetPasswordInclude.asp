<%
function resetPassword(vUserID)
	Dim vsql, vrs, vpassword, vLastPasswordChange, vTimestamp, vMessage, vEmail, vName, vSystemID
	vTimeStamp = Now()

	If Debug = True Then
		Response.Write("<pre>UserID=" & vUserID & "</pre>" & vbCrLf)
		Response.Flush
	End If

	sql = "SELECT SystemID, UserID, Name, Email, AccountDisabled, LastPasswordChange, timestamp=getdate() " & vbCrLf & _
		"FROM System.Users " & vbCrLf & _
		"WHERE UserID=" & prepStringSQL(vUserID) 
	If Debug = True Then
		Response.Write("<pre>sql=" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs= Con.Execute(sql)

	If rs.EOF = False Then
		vSystemID = rs.Fields("SystemID")
		vLastPasswordChange = rs.Fields("LastPasswordChange")
		vtimestamp = rs.Fields("timestamp")

		If Debug = True Then
			Response.Write("<pre>Success. UserID found, SystemID=" & vSystemID & _
				", LastChange='" & vLastPasswordChange & "', timestamp='" & vtimestamp & "'.</pre>" & vbCrLf)
			Response.Flush
		End If

		' Don't allow update within next ten minutes.
		If DateAdd("n", 10, vLastPasswordChange) > vtimestamp Then
			If Debug = True Then
				Response.Write("<pre>LastPasswordChange=" & vLastPasswordChange & _
					", timestamp=" & vtimestamp & ". Wait 10 minutes</pre>" & vbCrLf)
				Response.Flush
			End If
			vMessage = "The password was recently reset. Check your email for new password. " & _
				"You must wait at least 10 minutes before changing password again. " & _
				"It could take several minutes or more for you to recieve the email although it was sent immediately. "

		Else

			' Reset password.
			vsql = "EXEC System.spResetPassword @SystemID=" & vSystemID
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLF)
				Response.Flush
			End If
			Set rs = con.Execute(vsql)
			If rs.EOF = True Then
				Response.WRite("Error: Password Reset Failed!")
				Response.End
			Else
				vpassword = rs.Fields("Password")
				vemail = rs.Fields("email")
				'vemail="mvcpa@tamu.edu" ' <---- This is temporary to stop password resests from going out.
				vName = rs.Fields("Name")
			End If
			If ISNull(email) = False Then
				SendHTMLMail "TxMVCPA website<mvcpa@tamu.edu>", vEmail, "", _
					"Password Reset Notification for mvcpa.tamu.edu (Texas Motor Vehicle Crime Prevention Authority)", _
					vName & ", <br>" & _
					"A password reset has been requested for your account. " & _
					"Your password has been reset to """ & vpassword & _
					""". You may change your password the next time that you login. " & _
					"If you did not request a password reset (and this is not a new account), " & _
					"please reply to this message.<br /><br />" & _
					"<a href=""" & Request.ServerVariables("SERVER_NAME")& "/default.asp?UserID=" & UserID & """>" & Request.ServerVariables("SERVER_NAME") & "</a>" 
			End If

			vmessage = ""
		End If
		Set rs = nothing
		'Set Con = nothing

	Else
		If Debug = True Then
			Response.Write("<pre>Failure. UserID was not found.</pre>" & vbCrLf)
			Response.Flush
		End If
		vmessage = "The User ID " & vUserID & " was not found in the system. " & _
			"If this was the incorrect User ID, then correct it and try again. " & _
			"Otherwise you will need to create a new account."
	End If
	resetPassword = vmessage
End Function
%>