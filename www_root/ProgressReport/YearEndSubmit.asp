<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, GrantID, QuestionID, Version, update, months, submit, approval, _
	AdministrativeComments, AdministrativeCommentsDB, Text, dbtext, RecordPresent, _
	SubmitID, SubmitTimestamp, ApprovalTimestamp, ApprovalID, Unsubmit

TimeStamp = Now()

debug = False
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

GrantID = Request.Form("GrantID")
Version = Request.Form("Version")
If Request.Form("Submit") = "Submit" Then
	submit = True
	If Debug = True Then
		Response.Write("<pre>Submit=" & Submit & "</pre>")
	End If
Else
	submit = False
	If Debug = True Then
		Response.Write("<pre>Submit=" & Submit & "</pre>")
	End If
End If
If Request.Form("Approval") = "1" Then
	Approval = True
Else
	Approval = False
End If
If Request.Form("Unsubmit") = "1" Then
	Unsubmit = True
Else
	Unsubmit = False
End If
If Len(GrantID) = 0 Then
	Response.Write("Error: No GrantID provided.")
	sendWarning("Error: No GrantID provided.")
	Response.End
Else
	GrantID = CInt(GrantID)
End If
If Len(Version) = 0 Then
	Response.Write("Error: No Version provided.")
	sendWarning("Error: No Version provided.")
	Response.End
Else
	Version = CInt(Version)
End If
AdministrativeComments = Request.Form("AdministrativeComments")

sql = "SELECT GrantID, ISNULL(SubmitID,0) AS SubmitID, SubmitTimestamp, " & vbCrLf & _
	"	ISNULL(ApprovalID,0) AS ApprovalID, ApprovalTimestamp, AdministrativeComments  " & vbCrLF & _
	"FROM YE.Main " & vbCrLF & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	RecordPresent = False
	SubmitID = null
	SubmitTimestamp = null
	ApprovalID = null
	ApprovalTimestamp = null
	AdministrativeCommentsDB = null
Else
	RecordPresent = True
	SubmitID = rs.Fields("SubmitID")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	ApprovalID = rs.Fields("ApprovalID")
	ApprovalTimestamp = rs.Fields("ApprovalTimestamp")
	AdministrativeCommentsDB = rs.Fields("AdministrativeComments")
End If

sql = "SELECT A.QuestionID, A.Version, A.Identifier, REPLACE(A.Question, '{FiscalYear}','FY' + CAST((C.FiscalYear % 100) AS VARCHAR)) AS Question, " & vbCrLF & _
	"	A.QuestionSort, B.Response, CAST(CASE WHEN B.QuestionID IS NULL THEN 0 ELSE 1 END AS BIT) AS RecordPresent " & vbCrLf & _
	"FROM YE.Questions AS A " & vbCrLf & _
	"LEFT JOIN YE.Responses AS B ON B.QuestionID=A.QuestionID AND B.Version=A.Version AND B.GrantID=" & prepIntegerSQL(GrantID) & " " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS C ON C.GrantID=" & prepIntegerSQL(GrantID) & " " & vbCrLf & _
	"WHERE A.Version=" & Version
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
While rs.EOF = False
	' Check each question and record response if any. Delete response if necessary.
	sql = ""
	QuestionID = rs.Fields("QuestionID")
	Update=False

	Text = Request.Form("Response_" & rs.Fields("QuestionID"))
	dbtext = rs.Fields("Response")
	If Debug = True Then
		Response.Write("<pre>QuestionID=" & QuestionID & "; Response_" & QuestionID & "=" & _
			Text & "; dbvalue=" & rs.Fields("Response") & "</pre>")
	End If
	If Len(Text)=0 And IsNull(dbtext)=False Then
		update = True
	ElseIf Len(Text)>0 And IsNull(dbtext)=True Then
		update = True
	ElseIf Text <> dbtext Then
		update = True
	End If
	If update=True Then
		If rs.Fields("RecordPresent") = False Then
			' Do an insert.
			sql = "INSERT INTO YE.Responses (GrantID, QuestionID, Version, Response, UpdateID, UpdateTimestamp) " & vbCrLf & _
				"VALUES (" & prepStringSQL(GrantID) & ", " & _
				prepIntegerSQL(rs.Fields("QuestionID")) & ", " & _
				prepIntegerSQL(Version) & ", " & _
				prepStringSQL(Text) & ", " & _
				prepIntegerSQL(UserSystemID) & ", " & _
				prepStringSQL(Timestamp) & ")"
		Else
			' Do an update.
			sql = "UPDATE YE.Responses SET " & vbCRLF & _
				"Response=" & prepStringSQL(Text) & ", " & _
				"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
				"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
				"WHERE GrantID=" & prepIntegerSQL(GrantID) & _
				" AND QuestionID=" & prepIntegerSQL(rs.Fields("QuestionID")) 
		End If
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
	rs.MoveNext
Wend

If Debug = True Then
	Response.Write("<pre>Submit=" & Submit & "</pre>")
End If
If Submit = True  Then
	If RecordPresent = False Then ' Do an insert
		sql = "INSERT INTO YE.Main (GrantID, SubmitID, SubmitTimestamp) " & vbCrLf & _
			"VALUES (" & prepIntegerSQL(GrantID) & ", " & prepIntegerSQL(UserSystemID)& ", " & prepStringSQL(Timestamp) &  ")" & vbCrLf
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	Else
		sql = "UPDATE YE.Main " & vbCrLf & _
			"SET SubmitID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
			"	SubmitTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
			"WHERE GrantID=" & prepIntegerSQL(GrantID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
ElseIf MVCPARights = True Then
	sql = "UPDATE YE.Main SET " & vbCrLf
	If Unsubmit = True Then
		sql = sql & "SubmitTimestamp=null, " & vbCrLf & _
			"SubmitID=null, " & vbCrLf & _ 
			"ApprovalTimestamp=null, " & vbCrLf & _
			"ApprovalID=null, " & vbCrLf
	ElseIf (Approval = False And ApprovalID>0) Then
		sql = sql & "ApprovalTimestamp=null, " & vbCrLf & _
			"ApprovalID=null, " & vbCrLf
	ElseIf ApprovalID=0 And Approval = True Then 
		sql = sql & "ApprovalTimestamp=" & PrepStringSQL(Timestamp) & ", " & vbCrLf & _
			"ApprovalID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf
	End If
	If (Len(AdministrativeComments)=0 And IsNull(AdministrativeCommentsDB) = False) Or _
		(Len(AdministrativeComments)>0 And IsNull(AdministrativeCommentsDB) = True) Or _
		AdministrativeComments<>AdministrativeCommentsDB Then
		sql = sql & "AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ", " & vbCrLf
	End If
	sql = sql & "UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
		"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

If Debug = True Then
	Response.Write("<a href=""YearEnd.asp?GrantID=" & GrantID & """>return to year-end progress report</a><br />" & vbCrLf)
	Response.Write("<a href=""../Home/Default.asp?GrantID=" & GrantID & """>return to Home</a><br />" & vbCrLf)
Else
	Response.Redirect("YearEnd.asp?GrantID=" & GrantID)
End If




%><!--#include file="../includes/prepDB.asp"-->