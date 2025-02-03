<%@ language=VBScript %>
<% Option Explicit%>
<!--#include file="./includes/adovbs.asp"-->
<!--#include file="./includes/OpenConnection.asp"--><%
Dim debug, i, SystemID, Password, Name, email, UserID, AccountDisabled, LastLogin, _
	Applications, DefaultGrantee, _
	Developer, MVCPAAdministrator, MVCPAAuditor, MVCPAGrantCoordinator, _
	MVCPAAdministrativeAssistant, MVCPAScorer, MVCPAViewer, MVCPARights, _
	ScreenSize, AvailableSize, WindowSize, AppName, SessionID
Debug = False

If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
End If

UserID=Request.Form("UserID")
Password=Request.Form("Password")

If Len(UserID) = 0 Then
	Response.Redirect("default.asp?userid=" & UserID & "&message=You must provide a username to login. Contact TxMVCPA with questions.")	
End If
If Len(Password) < 8 Then
	Response.Redirect("default.asp?userid=" & UserID & "&message=You must provide password to login. Do a password reset if you have forgotten your password.")	
End If
If InStr(1, UserID, "@")=0 Or InStr(1, UserID, ".")=0 Or Len(UserID)<5 Then
	Response.Redirect("default.asp?userid=" & UserID & "&message=You must provide a valid username to login. Contact TxMVCPA with questions.")	
End If

ScreenSize = Request.Form("ScreenSize")
AvailableSize = Request.Form("AvailableSize")
WindowSize = Request.Form("WindowSize")
AppName = Request.Form("AppName")
SessionID = Session.SessionID

sql = "SELECT U.SystemID, U.UserID, U.EMail, U.Name, U.AccountDisabled, " & vbCrLf & _
	"	ISNULL(U.DefaultGrantee,ISNULL(GP.GranteeID,0)) AS DefaultGrantee, " & vbCRLF & _
	"	Developer, MVCPAAdministrator, MVCPAAuditor, MVCPAGrantCoordinator, MVCPAAdministrativeAssistant, MVCPAScorer, MVCPAViewer, " & vbCrLf & _
	"	CAST(CASE WHEN Developer=1 THEN 1 WHEN MVCPAAdministrator=1 THEN 1 WHEN MVCPAAuditor=1 THEN 1 WHEN MVCPAGrantCoordinator=1 THEN 1 WHEN MVCPAAdministrativeAssistant=1 THEN 1 ELSE 0 END AS BIT) AS MVCPARights, " & vbCrLF & _
	"	LastLogin=(SELECT MAX(LoginTime) FROM System.LoginLog AS L WHERE L.SystemID=U.SystemID), " & vbCrLf & _
	"	Applications=(SELECT ISNULL(COUNT(*),0) FROM Application.IDs WHERE GranteeID=U.DefaultGrantee) " & vbCrLf & _
	"FROM System.Users AS U" & vbCrLf & _
	"LEFT JOIN (SELECT SystemID, MIN(GranteeID) AS GranteeID FROM System.GranteePermissions GROUP BY SystemID) AS GP ON GP.SystemID=U.SystemID " & vbCrLF & _
	"WHERE UserID=" & prepStringSQL(UserID) & " AND passwordhash=HASHBYTES('SHA2_256'," & prepUnicodeSQL(Password) & ")"
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
	DefaultGrantee = rs.Fields("DefaultGrantee")
	Developer = rs.Fields("Developer")
	MVCPAAdministrator = rs.Fields("MVCPAAdministrator")
	MVCPAAuditor = rs.Fields("MVCPAAuditor")
	MVCPAGrantCoordinator = rs.Fields("MVCPAGrantCoordinator")
	MVCPAAdministrativeAssistant = rs.Fields("MVCPAAdministrativeAssistant")
	MVCPAScorer = rs.Fields("MVCPAScorer")
	MVCPAViewer = rs.Fields("MVCPAViewer")
	MVCPARights = rs.Fields("MVCPARights")
	AccountDisabled = rs.Fields("AccountDisabled")
	LastLogin = rs.Fields("LastLogin")
	Applications = rs.Fields("Applications")
	
	' Set cookie in users browsers for next login
	Response.Cookies("UserID")=UserID
	Response.Cookies("UserID").Expires=DateAdd("d", 95, Date())
	' Remember system_id in case session times out.
	Response.Cookies("SystemID") = SystemID
	Response.Cookies("FiscalYear") = Application("DefaultFiscalYear")
	Response.Cookies("GranteeID") = DefaultGrantee
	Response.Cookies("SessionID") = SessionID

	If AccountDisabled = True Then
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
	Session("Developer") = Developer
	Session("MVCPAAdministrator") = MVCPAAdministrator
	Session("MVCPAAuditor") = MVCPAAuditor
	Session("MVCPAGrantCoordinator") = MVCPAGrantCoordinator
	Session("MVCPAAdministrativeAssistant") = MVCPAAdministrativeAssistant
	Session("MVCPAScorer") = MVCPAScorer
	Session("MVCPAViewer") = MVCPAViewer
	Session("MVCPARights") = MVCPARights
	If MVCPARights = False And Applications=0 Then
		Session("FiscalYear") = Application("CurrentFiscalYear") + 1
	Else
		Session("FiscalYear") = Application("DefaultFiscalYear")
	End If
	Session("GranteeID") = DefaultGrantee
	Session("LastLogin") = LastLogin

	' Create Login Record
	rs.close
	sql = "INSERT INTO System.LoginLog (SessionID, SystemID, LoginTime, ScreenSize, AvailableSize, WindowSize, AppName, ipaddress, SessionRecovery) VALUES (" & _
		SessionID & ", " & SystemID & ", getdate(), " & prepStringSql(ScreenSize) & ", " & _
		prepStringSQL(AvailableSize) & ", " & prepStringSQL(WindowSize) & ", " & prepStringSQL(AppName) _
		& ", " & prepStringSQL(Request.ServerVariables("REMOTE_ADDR")) & ", 0)"
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)

	If Debug = True Then
		If IsNull(LastLogin) Then
			Response.Write("<a href=""/User/UpdateProfile.asp"">Go To Home Page</a>") & vbCrLF
		ElseIf DefaultGrantee>0 Then
			Response.Write("<a href=""/Home/Default.asp?GranteeID=" & GranteeID & """>Go To Home Page</a>") & vbCrLF
		Else
			Response.Write("<a href=""/Home/Default.asp"">Go To Home Page</a>") & vbCrLF
		End If
	Else
		If IsNull(LastLogin) Then
			Response.Redirect("/User/UpdateProfile.asp")
		Else
			Response.Redirect("/Home/Default.asp")
		End If
	End If

Else
	Response.Cookies("SessionID") = 0
	If Debug = True Then
		Response.Write("<pre>Failure</pre>" & vbCrLf)
	End If

	' Create Login Record
	rs.close
	sql = "INSERT INTO System.LoginFailureLog (UserID, Password, LoginTime, ScreenSize, AvailableSize, WindowSize, AppName, ipaddress) VALUES (" & _
		prepStringSQL(UserID) & ", " & prepStringSQL(password) & ", getdate(), " & prepStringSql(ScreenSize) & ", " & _
		prepStringSQL(AvailableSize) & ", " & prepStringSQL(WindowSize) & ", " & prepStringSQL(AppName) _
		& ", " & prepStringSQL(Request.ServerVariables("REMOTE_ADDR")) & ")"
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)

	If Debug = True Then
		Response.Write("<a href=""default.asp?userid=" & UserID & "&message=Username and password do not match an existing account."">Return to Login Page</a>") & vbCrLF
	Else
		Response.Redirect("default.asp?userid=" & UserID & "&message=Username and password do not match an existing account.")
	End If
End If


 %><!--#include file="includes/PrepDB.asp"-->