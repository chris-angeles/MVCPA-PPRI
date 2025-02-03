<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, UpdateSystemID, UserID, FirstName, MiddleName, LastName, Suffix, Title, email, _
	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, LicensedPeaceOfficer, TCOLEPID, _
	DeveloperRole, MVCPAAdministratorRole, MVCPAAuditorRole, MVCPAGrantCoordinatorRole, MVCPAAdministrativeAssistantRole, _
	MVCPAScorerRole,MVCPAViewerRole, MVCPAStaffRole, AccountDisabled, Comments, GranteePermissions, DefaultGrantee
debug = False
Timestamp = Now()
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

If Request.Form.Count>0 Then
	UpdateSystemID = Request.Form("UpdateSystemID")
	UserID = Request.Form("email")
	FirstName = Request.Form("FirstName")
	MiddleName = Request.Form("MiddleName")
	LastName = Request.Form("LastName")
	Suffix = Request.Form("Suffix")
	Title = Request.Form("Title")
	email = Request.Form("email")
	Address1 = Request.Form("Address1")
	Address2 = Request.Form("Address2")
	City = Request.Form("City")
	State = Request.Form("State")
	ZIP = Request.Form("ZIP")
	Phone = Request.Form("Phone")
	Fax = Request.Form("Fax")
	Mobile = Request.Form("Mobile")
	LicensedPeaceOfficer = Request.Form("LicensedPeaceOfficer")
	TCOLEPID = Request.Form("TCOLEPID")
	DeveloperRole = Request.Form("Developer")
	MVCPAAdministratorRole = Request.Form("MVCPAAdministrator")
	MVCPAAuditorRole = Request.Form("MVCPAAuditor")
	MVCPAGrantCoordinator = Request.Form("MVCPAGrantCoordinator")
	MVCPAAdministrativeAssistantRole = Request.Form("MVCPAAdministrativeAssistant")
	MVCPAScorerRole = Request.Form("MVCPAScorer")
	MVCPAViewerRole = Request.Form("MVCPAViewer")
	MVCPAStaffRole = Request.Form("MVCPAStaff")
	AccountDisabled = Request.Form("AccountDisabled")
	Comments = Request.Form("Comments")
	DefaultGrantee = Request.Form("DefaultGrantee")

	If UpdateSystemID="0" Then
		sql = "SELECT Name FROM System.Users WHERE UserID=" & prepStringSQL(UserID)
		Set rs=Con.Execute(sql)
		If rs.EOF = False Then
			Response.Write("Error: This UserID/Email is already used by " & rs.Fields("Name") & ". A username may not be duplicated. Use 'Back' to edit email address.")
			Response.End
		End If
		sql = "INSERT INTO System.Users (UserID, FirstName, MiddleName, LastName, Suffix, Title, email, " & vbCrLf & _
			"	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, LicensedPeaceOfficer, TCOLEPID, " & vbCrLf & _
			"	DefaultGrantee, Developer, MVCPAAdministrator, MVCPAAuditor, MVCPAGrantCoordinator, " & vbCrLf & _
			"	MVCPAAdministrativeAssistant, MVCPAScorer, MVCPAViewer, MVCPAStaff, AccountDisabled, " & vbCrLf & _
			"	Comments, LastPasswordChange, UpdateID, UpdateTimestamp) " & vbCrLf & _
			"OUTPUT Inserted.SystemID " & vbCrLf & _
			"VALUES (" & prepStringSQL(UserID) & ", " & _
			prepStringSQL(FirstName) & ", " & _
			prepStringSQL(MiddleName) & ", " & _
			prepStringSQL(LastName) & ", " & _
			prepStringSQL(Suffix) & ", " & _
			prepStringSQL(Title) & ", " & _
			prepStringSQL(email) & ", " & _
			prepStringSQL(Address1) & ", " & _
			prepStringSQL(Address2) & ", " & _
			prepStringSQL(City) & ", " & _
			prepStringSQL(State) & ", " & _
			prepStringSQL(ZIP) & ", " & _
			prepStringSQL(Phone) & ", " & _
			prepStringSQL(Fax) & ", " & _
			prepStringSQL(Mobile) & ", " & _
			prepBitRequiredSQL(LicensedPeaceOfficer) & ", " & _
			prepIntegerSQL(TCOLEPID) & ", " & _
			prepIntegerSQL(DefaultGrantee) & ", " & _
			prepBitRequiredSQL(DeveloperRole) & ", " & _
			prepBitRequiredSQL(MVCPAAdministratorRole) & ", " & _
			prepBitRequiredSQL(MVCPAAuditorRole) & ", " & _
			prepBitRequiredSQL(MVCPAGrantCoordinatorRole) & ", " & _
			prepBitRequiredSQL(MVCPAAdministrativeAssistantRole) & ", " & _
			prepBitRequiredSQL(MVCPAScorerRole) & ", " & _
			prepBitRequiredSQL(MVCPAViewerRole) & ", " & _
			prepBitRequiredSQL(MVCPAStaffRole) & ", " & _
			prepBitRequiredSQL(AccountDisabled) & ", " & _
			prepStringSQL(Comments) & ", null, " & vbCrLf & _
			UserSystemID & ", " & _
			prepStringSQL(Timestamp) & ") "
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		IF rs.EOF = False Then
			UpdateSystemID = rs.Fields("SystemID")
		Else
			Response.Write("Error: No SystemID returned form insert.")
			Response.End
		End If
	Else
		sql = "SELECT Name FROM System.Users WHERE UserID=" & prepStringSQL(UserID) & " AND SystemID<>" & prepIntegerSQL(UpdateSystemID)
		Set rs=Con.Execute(sql)
		If rs.EOF = False Then
			Response.Write("Error: This UserID/Email is already used by " & rs.Fields("Name") & ". A username may not be duplicated. Use 'Back' to edit email address.")
			Response.End
		End If

		sql = "UPDATE System.Users SET UserID=" & prepStringSQL(UserID) & _
			", FirstName=" & prepStringSQL(FirstName) & _
			", MiddleName=" & prepStringSQL(MiddleName) & _
			", LastName=" & prepStringSQL(LastName) & _
			", Suffix=" & prepStringSQL(Suffix) & _
			", Title=" & prepStringSQL(Title) & _
			", email=" & prepStringSQL(email) & _
			", Address1=" & prepStringSQL(Address1) & _
			", Address2=" & prepStringSQL(Address2) & _
			", City=" & prepStringSQL(City) & _
			", State=" & prepStringSQL(State) & _
			", ZIP=" & prepStringSQL(ZIP) & _
			", Phone=" & prepStringSQL(Phone) & _
			", Mobile=" & prepStringSQL(Mobile) & _
			", LicensedPeaceOfficer=" & prepBitRequiredSQL(LicensedPeaceOfficer) & _
			", TCOLEPID=" & prepIntegerSQL(TCOLEPID) & _
			", DefaultGrantee=" & prepIntegerSQL(DefaultGrantee) & _
			", Fax=" & prepStringSQL(Fax) & _
			", Developer=" & prepBitRequiredSQL(DeveloperRole) & _
			", MVCPAAdministrator=" & prepBitRequiredSQL(MVCPAAdministratorRole) & _
			", MVCPAAuditor=" & prepBitRequiredSQL(MVCPAAuditorRole) & _
			", MVCPAGrantCoordinator=" & prepBitRequiredSQL(MVCPAGrantCoordinator) & _
			", MVCPAAdministrativeAssistant=" & prepBitRequiredSQL(MVCPAAdministrativeAssistantRole) & _
			", MVCPAScorer=" & prepBitRequiredSQL(MVCPAScorerRole) & _
			", MVCPAViewer=" & prepBitRequiredSQL(MVCPAViewerRole) & _
			", MVCPAStaff=" & prepBitRequiredSQL(MVCPAStaffRole) & _
			", AccountDisabled=" & prepBitRequiredSQL(AccountDisabled) & _
			", Comments=" & prepStringSQL(Comments) & _
			", UpdateID=" & UserSystemID & _
			", UpdateTimestamp=" & prepStringSQL(Timestamp)& vbCrLF & _
			"WHERE SystemID=" & prepIntegerSQL(UpdateSystemID)
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		rs = Con.Execute(sql)
	End If
End If

' Update Permissions. Delete Existing and then add.
If CInt(UpdateSystemID)>0 Then
	GranteePermissions = Request.Form("GranteePermissions")
	sql = "DELETE FROM System.GranteePermissions WHERE SystemID=" & prepIntegerSQL(UpdateSystemID) 
	If Len(GranteePermissions)>0 Then
		sql = sql & " AND GranteeID NOT IN (" & GranteePermissions & ")"
	End If
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)

	For i = 1 To Request.Form("GranteePermissions").Count
		sql = "SELECT SystemID, GranteeID FROM System.GranteePermissions WHERE SystemID=" & _
			prepIntegerSQL(UpdateSystemID) & " AND GranteeID=" & prepIntegerSQL(Request.Form("GranteePermissions")(i))
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = con.execute(sql)
		IF rs.EOF = True Then
			sql = "INSERT INTO System.GranteePermissions(SystemID, GranteeID, UpdateID, UpdateTimeStamp) VALUES (" & _
				UpdateSystemID & ", " & prepIntegerSQL(Request.Form("GranteePermissions")(i)) & ", " & prepIntegerSQL(UserSystemID) & _
				", " & prepStringSQL(TimeStamp) & ")"
			If Debug = True Then
				Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			con.execute(sql)
		End If
	Next
End If
If Debug = True Then
	Response.Write("<a href=""../Home/default.asp"">../Home/default.asp</a>")
Else
	Response.Redirect("../Home/default.asp")
End If
%><!--#include file="../includes/prepDB.asp"-->