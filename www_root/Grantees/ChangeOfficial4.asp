<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, UpdateSystemID, Position, GranteeID, ReturnPage, SearchLastName, fieldname, _
	UserID, FirstName, MiddleName, LastName, Suffix, Name, Title, email, _
	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile
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

If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else 
	GranteeID = 0
End If
If Request.Form.Count>0 Then
	Position = Request.Form("Position")
	ReturnPage = Request.Form("ReturnPage")
	GranteeID = Request.Form("GranteeID")
	SearchLastName = Request.Form("SearchLastName")
	If Position = "Authorized Official" Then
		fieldname = "AuthorizedOfficialID"
	ElseIf Position = "Program Director" Then
		fieldname = "ProgramDirectorID" 
	ElseIf Position = "Program Manager" Then
		fieldname = "ProgramManagerID" 
	ElseIf Position = "Financial Officer" Then
		fieldname = "FinancialOfficerID"
	ElseIf Position = "Program Administrative Contact" Then
		fieldname = "ProgramAdministrativeContactID"
	ElseIf Position = "Financial Administrative Contact" Then
		fieldname = "FinancialAdministrativeContactID"
	ElseIf Position = "Taskforce Commander" Then
		fieldname = "TaskForceCommanderID"
	ElseIf Position = "PIO / Media Contact" Then
		fieldname = "PIOID"
	Else
		Response.Write("Error: Invalid position title")
		Response.End
	End If

	UpdateSystemID = Request.Form("UpdateSystemID")
	If UpdateSystemID = "" Then
		UpdateSystemID = 0
	End If
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

	sql = "SELECT Name FROM System.Users WHERE UserID=" & prepStringSQL(email) & " AND SystemID<>" & prepIntegerSQL(UpdateSystemID)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs=Con.Execute(sql)
	If rs.EOF = False Then
		Response.Write("Error: This UserID/Email is already used by " & rs.Fields("Name") & ". A username may not be duplicated. Use 'Back' to edit email address.")
		Response.End
	End If

	If UpdateSystemID=0 Then
		sql = "INSERT INTO System.Users (UserID, FirstName, MiddleName, LastName, Suffix, " & vbCrLF & _
			"	Title, email, Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, AccountDisabled, UpdateID, UpdateTimestamp) " & vbCrLF & _
			"VALUES (" & prepStringSQL(EMail) & ", " & _
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
			"0, " & vbCrLf & _
			UserSystemID & ", " & _
			prepStringSQL(Timestamp) & ")"

		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)

		'sql = "UPDATE Grantees SET " & fieldname & "=" & prepIntegerSQL(UpdateSystemID) & ", " & vbCrLf & _
		'	"UpdateID=" & UserSystemID & ", UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
		'	"WHERE GranteeID=" & prepIntegerSQL(GranteeID) 
		'If Debug = True Then
		'	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		'	Response.Flush
		'End If
		'Con.Execute(sql)

		' Retrieve UserID and generate a password.
		sql = "SELECT SystemID, UserID, Name, Email " & vbCrLf & _
			"FROM System.Users " & vbCrLF & _
			"WHERE email=" & prepStringSQL(Email)
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		If rs.EOF = False Then
			UpdateSystemID = rs.Fields("SystemID")
			UserID = rs.Fields("UserID")
			Name = rs.Fields("Name")
			If Debug = True Then
				Response.Write("<pre>SystemID=" & UpdateSystemID & ", UserID=" & UserID & ", Name=" & Name & "</pre>" & vbCrLF)
				Response.Flush
			End If
		Else
			Response.Write("Error: Unable to retrive new user.")
			Response.End
		End If

		'resetPassword(UserID)
	Else
		sql = "UPDATE System.Users SET FirstName=" & prepStringSQL(FirstName) & _
			", MiddleName=" & prepStringSQL(MiddleName) & _
			", LastName=" & prepStringSQL(LastName) & _
			", Suffix=" & prepStringSQL(Suffix) & _
			", Title=" & prepStringSQL(Title)
		If MVCPARights = True Then
			sql = sql & ", email=" & prepStringSQL(email)
		End If
		sql = sql & ", Address1=" & prepStringSQL(Address1) & _
			", Address2=" & prepStringSQL(Address2) & _
			", City=" & prepStringSQL(City) & _
			", State=" & prepStringSQL(State) & _
			", ZIP=" & prepStringSQL(ZIP) & _
			", Phone=" & prepStringSQL(Phone) & _
			", Fax=" & prepStringSQL(Fax) & _
			", Mobile=" & prepStringSQL(Mobile) & _
			", UpdateID=" & UserSystemID & _
			", UpdateTimestamp=" & prepStringSQL(Timestamp)& vbCrLF & _
			"WHERE SystemID=" & prepIntegerSQL(UpdateSystemID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
	' Whether an update or insert, update the position in grantees table.
	sql = "UPDATE Grantees SET " & fieldname & "=" & prepIntegerSQL(UpdateSystemID) & ", " & vbCrLf & _
		"UpdateID=" & UserSystemID & ", UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
		"WHERE GranteeID=" & prepIntegerSQL(GranteeID) 
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If


sql = "SELECT SystemID, GranteeID FROM System.GranteePermissions WHERE SystemID=" & _
	prepIntegerSQL(UpdateSystemID) & " AND GranteeID=" & prepIntegerSQL(GranteeID)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = con.execute(sql)
IF rs.EOF = True Then
	sql = "INSERT INTO System.GranteePermissions(SystemID, GranteeID, UpdateID, UpdateTimeStamp) VALUES (" & _
		UpdateSystemID & ", " & GranteeID & ", " & prepIntegerSQL(UserSystemID) & _
		", " & prepStringSQL(TimeStamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	con.execute(sql)
End If


If Debug = True Then
	Response.Write("<a href=""" & ReturnPage & "?GranteeID=" & GranteeID & """>" & ReturnPage & "</a>")
Else
	Response.Redirect(ReturnPage & "?GranteeID=" & GranteeID)
End If
%><!--#include file="../includes/prepDB.asp"-->