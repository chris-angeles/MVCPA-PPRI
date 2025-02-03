<%@ language=VBScript %>
<% Option Explicit%>
<!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><%
Dim debug, i, SystemID, Email, UserID, Name, UserAccountDisabled, LastLogin, UserDefaultGrantee, _
	UserDeveloper, UserMVCPAAdministrator, UserMVCPAGrantCoordinator, UserMVCPAAdministrativeAssistant, _
	UserMVCPAScorer,  UserMVCPAAuditor, UserMVCPAViewer, UserMVCPARights, _
	ScreenSize, AvailableSize, WindowSize, AppName, _
	InitiatorID, ipaddress, Instance
Debug = False

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
	For each i in Application.Contents
		Response.Write("Application(""" & i & """)='" & Application(i) & "'" & vbCrLf)
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

SystemID = Request.QueryString("SystemID")
ipaddress = Request.ServerVariables("REMOTE_ADDR")
InitiatorID = UserSystemID
Instance = Application("Instance")

If Instance = "Production" Then
	If Developer = True Then
		' Proceed
	ElseIf MVCPAAdministrator = True Then
		' Proceed
	Else
		Response.Write("Error: Failed Impersonation Attempt for " & UserID & " by " & UserSystemID & " on " & Instance)
		SendMessage "Error: Failed Impersonation Attempt for " & UserID & " by " & Session("UserID") & " on " & Instance
		Response.End
	End If
ElseIf Instance = "Test" Then
	If Developer = True Then
		'Proceed
	ElseIf MVCPARights = True Then
		' Proceed
	ElseIf Left(ipaddress,10)="204.64.196" Then
		' Proceed
	ElseIf Left(ipaddress,10)="204.64.198" Then
		' Proceed
	ElseIf ipaddress="10.125.134.19" Then
		' Proceed
	Else
		Response.Write("Error: Failed Impersonation Attempt for " & UserID & " by " & UserSystemID & " on " & Instance)
		SendMessage "Error: Failed Impersonation Attempt for " & UserID & " by " & Session("UserID") & " on " & Instance
		Response.End
	End If
ElseIf Left(Instance,11) = "Development" Then
	' Proceed
Else
	Response.Write("Error: Failed Impersonation Attempt for " & UserID & " by " & UserSystemID & " on " & Instance)
	SendMessage "Error: Failed Impersonation Attempt for " & UserID & " by " & Session("UserID") & " on " & Instance
	Response.End
End If

If Len(SystemID)>0 Then
	SystemID = CInt(SystemID)
End If
If SystemID = 0 Then
	Response.Write("<html>")
	Response.Write("<head>" & vbCrLf)
	Response.Write("<title>Impersonate a user</title>" & vbCrLf)
	Response.Write("<link rel=""stylesheet"" href=""/styles/main.css"" type=""text/css"" />" & vbCrLf) 
	Response.Write("</head>" & vbCrLf)
	Response.Write("<body  style=""width: 100%; "">")
	Response.Write("<form name=""Impersonate"" method=""get"">" & vbCrLf)
	Response.Write("User to impersonate: <select name=""SystemID"" onchange=""document.Impersonate.submit();"">" & vbCrLf)
	Response.WRite("<option value=""0"">Select User</option>" & vbCrLf)
	If Developer = True Or Application("Instance") = "Test" Then
		sql = "SELECT SystemID, UserID, Name FROM [System].Users WHERE ISNULL(AccountDisabled,0)=0 And ISNULL(Developer,0)=0 ORDER BY LastName, FirstName"
	Else
		sql = "SELECT SystemID, UserID, Name FROM [System].Users WHERE ISNULL(AccountDisabled,0)=0 And ISNULL(Developer,0)=0 AND ISNULL(MVCPAAdministrator,0)=0 ORDER BY LastName, FirstName"
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("SystemID") & """>" & rs.Fields("Name") & ", " & rs.Fields("UserID") & "</option>" & vbCrLf)
		rs.MoveNext()
	Wend
	Response.Write("</select>")
	Response.Write("</form>")
	Response.Write("</body>")
	Response.Write("<html>")
	Response.End
End If

sql = "SELECT U.SystemID, U.UserID, U.EMail, U.Name, U.AccountDisabled, " & vbCrLf & _
	"	ISNULL(U.DefaultGrantee,ISNULL(GP.GranteeID,0)) AS DefaultGrantee, " & vbCRLF & _
	"	Developer, MVCPAAdministrator, MVCPAGrantCoordinator, MVCPAAdministrativeAssistant, " & vbCrLf & _
	"	MVCPAAuditor, MVCPAScorer, MVCPAViewer, " & vbCrLf & _
	"	CAST(CASE WHEN Developer=1 THEN 1 WHEN MVCPAAdministrator=1 THEN 1 WHEN MVCPAGrantCoordinator=1 THEN 1 WHEN MVCPAAdministrativeAssistant=1 THEN 1 ELSE 0 END AS BIT) AS MVCPARights, " & vbCrLF & _
	"	LastLogin=(SELECT MAX(LoginTime) FROM System.LoginLog AS L WHERE L.SystemID=U.SystemID) " & vbCrLF & _
	"FROM System.Users AS U" & vbCrLf & _
	"LEFT JOIN (SELECT SystemID, MIN(GranteeID) AS GranteeID FROM System.GranteePermissions GROUP BY SystemID) AS GP ON GP.SystemID=U.SystemID " & vbCrLF & _
	"WHERE U.SystemID=" & prepIntegerSQL(SystemID) & " AND ISNULL(AccountDisabled,0)=0 "
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs= Con.Execute(sql)

If rs.EOF = False Then
	If Debug = True Then
		Response.Write("Success")
	End If

	SystemID = rs.Fields("SystemID")
	UserID = rs.Fields("UserID")
	EMail = rs.Fields("EMail")
	Name = rs.Fields("Name")
	UserDefaultGrantee = rs.Fields("DefaultGrantee")
	UserDeveloper = rs.Fields("Developer")
	UserMVCPAAdministrator = rs.Fields("MVCPAAdministrator")
	UserMVCPAGrantCoordinator = rs.Fields("MVCPAGrantCoordinator")
	UserMVCPAAdministrativeAssistant = rs.Fields("MVCPAAdministrativeAssistant")
	UserMVCPAAuditor = rs.Fields("MVCPAAuditor")
	UserMVCPAScorer = rs.Fields("MVCPAScorer")
	UserMVCPAViewer = rs.Fields("MVCPAViewer")
	UserMVCPARights = rs.Fields("MVCPARights")
	UserAccountDisabled = rs.Fields("AccountDisabled")
	LastLogin = rs.Fields("LastLogin")

	' Remember system_id in case session times out.
	If Debug = False Then
		Response.Cookies("SystemID") = UserSystemID
		Response.Cookies("FiscalYear") = Application("DefaultFiscalYear")
	End If

	If UserAccountDisabled = True Then
		If Debug = True Then
			Response.Write("Account Disabled. <a href=""default.asp?userid=" & UserID & "&message=The username that you provided has been disabled. Contact TxMVCPA with questions."">Return to Login Page</a>") & vbCrLF
		Else
			Response.Redirect("default.asp?userid=" & UserID & "&message=The username that you provided has been disabled. Contact TxMVCPA with questions.")
		End If
	End If

	' Set Session Variables
	Session("SystemID") = SystemID
	Session("Name") = Name
	Session("email") = email
	Session("Developer") = UserDeveloper
	Session("MVCPAAdministrator") = UserMVCPAAdministrator
	Session("MVCPAGrantCoordinator") = UserMVCPAGrantCoordinator
	Session("MVCPAAdministrativeAssistant") = UserMVCPAAdministrativeAssistant
	Session("MVCPAAuditor") = UserMVCPAAuditor
	Session("MVCPAScorer") = UserMVCPAScorer
	Session("MVCPAViewer") = UserMVCPAViewer
	Session("MVCPARights") = UserMVCPARights
	Session("FiscalYear") = Application("DefaultFiscalYear")
	Session("GranteeID") = UserDefaultGrantee

	sql = "INSERT INTO [System].ImpersonationLog (SystemID, ImpersonizationID, LoginTime, ipaddress) " & vbCrLF & _
		"VALUES(" & prepIntegerSQL(InitiatorID) & ", " & prepIntegerSQL(SystemID) & ", " & prepStringSQL(Now()) & ", " & prepStringSQL(ipaddress) & ")"
	Con.Execute(sql)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	If Debug = True Then
		Response.Write("<a href=""../Home/default.asp?GranteeID=" & UserDefaultGrantee & """>Continue to Home Page</a>") & vbCrLF
	Else
		Response.Redirect("../Home/default.asp?GranteeID=" & UserDefaultGrantee)
	End If
End If


 %><!--#include file="../includes/PrepDB.asp"-->