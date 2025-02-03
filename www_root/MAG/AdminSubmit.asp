<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 

Dim debug, i, j, LastCategory, Timestamp, PermitEdit, Button, MAGID, GrantResultID, _
	ResolutionConfirmedDate, ApplicationCertifiedCompleteDate, ApplicationConsideredDate, _
	GrantAwardAmount, CashMatch, GrantNumber, OfficialGrantAwardLetterDate, _
	GrantAwardCertifiedComplete, POIssueDate, GrantClosedDate, Notes
Debug = False

If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLf)
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
	Response.Write("Now=" & Now() & vbCrLf)
	Response.Write("</pre>" & vbCrLf)
End If

Timestamp = Now()
MAGID = Request.Form("MAGID")
GrantResultID = Request.Form("GrantResultID")
ResolutionConfirmedDate = Request.Form("ResolutionConfirmedDate")
ApplicationCertifiedCompleteDate = Request.Form("ApplicationCertifiedCompleteDate")
ApplicationConsideredDate = Request.Form("ApplicationConsideredDate")
GrantNumber = Request.Form("GrantNumber")
GrantAwardAmount = Request.Form("GrantAwardAmount")
CashMatch = Request.Form("CashMatch")
OfficialGrantAwardLetterDate = Request.Form("OfficialGrantAwardLetterDate")
GrantAwardCertifiedComplete = Request.Form("GrantAwardCertifiedComplete")
POIssueDate = Request.Form("POIssueDate")
Button = Request.Form("Button")
GrantClosedDate = Request.Form("GrantClosedDate")
Notes = Request.Form("Notes")

sql = "SELECT MAGID, GrantResultID, ResolutionConfirmedDate, ApplicationCertifiedCompleteDate, " & vbCrLf & _
	"	ApplicationConsideredDate, GrantNumber, GrantAwardAmount, CashMatch, OfficialGrantAwardLetterDate, " & vbCrLf & _
	"	POIssueDate, UpdateID, UpdateTimestamp " & vbCrLf & _
	"FROM MAG.Admin " & vbCrLF & _
	"WHERE MAGID=" & prepIntegerSQL(MagID) & " "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(SQL)
If rs.EOF = True Then
	' Insert
	sql = "INSERT INTO MAG.Admin(MAGID, GrantResultID, ResolutionConfirmedDate, ApplicationCertifiedCompleteDate, " & vbCrLf & _
	"	ApplicationConsideredDate, GrantNumber, GrantAwardAmount, CashMatch, OfficialGrantAwardLetterDate, " & vbCrLf & _
	"	POIssueDate, GrantAwardCertifiedComplete, GrantClosedDate, Notes, UpdateID, UpdateTimestamp) VALUES " & vbCrLf & _
		"(" & prepIntegerSQL(MAGID) & ", " & _
		prepIntegerSQL(GrantResultID) & ", " & _
		prepDateSQL(ResolutionConfirmedDate) & ", " & _
		prepDateSQL(ApplicationCertifiedCompleteDate) & ", " & _
		prepDateSQL(ApplicationConsideredDate) & ", " & _
		prepStringSQL(GrantNumber) & ", " & _
		prepNumberSQL(GrantAwardAmount) & ", " & _
		prepNumberSQL(CashMatch) & ", " & vbCrLf & _
		prepDateSQL(OfficialGrantAwardLetterDate) & ", " & _
		prepDateSQL(POIssueDate) & ", " & _
		prepDateSQL(GrantAwardCertifiedComplete) & ", " & _
		prepDateSQL(GrantClosedDate) & ", " & _
		prepStringSQL(Notes) & ", " & _
		prepIntegerSQL(UserSystemID) & ", " & _
		prepDateSQL(Timestamp) & ")"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
Else
	' Update

	sql = "UPDATE MAG.Admin SET " & _
		"GrantResultID=" & prepIntegerSQL(GrantResultID) & ", " & _
		"ResolutionConfirmedDate=" & prepDateSQL(ResolutionConfirmedDate) & ", " & _
		"ApplicationCertifiedCompleteDate=" & prepDateSQL(ApplicationCertifiedCompleteDate) & ", " & _
		"ApplicationConsideredDate=" & prepDateSQL(ApplicationConsideredDate) & ", " & _
		"GrantNumber=" & prepStringSQL(GrantNumber) & ", " & _
		"CashMatch=" & prepNumberSQL(CashMatch) & ", " & _
		"GrantAwardAmount=" & prepNumberSQL(GrantAwardAmount) & ", " & _
		"OfficialGrantAwardLetterDate=" & prepDateSQL(OfficialGrantAwardLetterDate) & ", " & _
		"GrantAwardCertifiedComplete=" & prepDateSQL(GrantAwardCertifiedComplete) & ", " & _
		"POIssueDate=" & prepDateSQL(POIssueDate) & ", " & _
		"GrantClosedDate=" & prepDateSQL(GrantClosedDate) & ", " & _
		"Notes=" & prepStringSQL(Notes) & ", " & _
		"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
		"UpdateTimestamp=" & prepDateSQL(Timestamp) & " " & vbCrLf & _
		"WHERE MagID=" & prepIntegerSQL(MagID) & " "
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
End If

If Debug = True Then
	Response.Write("<a href=""Admin.asp?MagID=" & MagID & """>Go To Page</a>")
Else
	Response.Redirect("Admin.asp?MagID=" & MagID)
End if
%><!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/prepDB.asp"-->