<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, changes, changedfields, parts, Column, AppID, QuestionID, TextSectionVersion, ScoringVersion, Score, Comments, Timestamp, URL
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
	Response.Write("</pre>" & vbCrLf)
	Response.Flush
End If
Changes = Request.Form("Changes")
changedfields = split(Changes,vbCrLf)
TextSectionVersion = Request.Form("TextSectionVersion")
ScoringVersion = Request.Form("ScoringVersion")

url = "Scoring.asp?FiscalYear=" & Request.Form("FiscalYear") & "&GrantClassID=" & Request.Form("GrantClassID") & "&GrantTypeID=" & Request.Form("GrantTypeID") & "&AppID=" & Request.Form("AppID") & "&QuestionID=" & Request.Form("QuestionID")
Response.Write("<pre>URL=""" & URL & """</pre>" & vbCrLf)

For each i in changedfields
	If Debug = True Then
		Response.Write("<pre>" & i & " changed.</pre>")
	End If
	parts = split(i,"_")
	'Response.Write("UBound=" & UBound(parts))
	If UBound(parts)=2 Then
		Column = parts(0)
		AppID = parts(1)
		QuestionID = parts(2)
		If Debug = True Then
			Response.Write("<pre>AppID="  & AppID & "; QuestionID=" & QuestionID & "</pre>")
		End If
		Score = Request.Form("score_" & AppID & "_" & QuestionID)
		Comments = Request.Form("Comments_" & AppID & "_" & QuestionID)
		If Debug = True Then
			Response.Write("<pre>Score="  & Score & "; Comments=" & Comments & "</pre>")
		End If
		sql = "SELECT AppID, QuestionID, Version, SystemID, Score, Comments, UpdateTimestamp " & vbCrLF & _
			"FROM Scoring.Scores " & vbCrLf & _
			"WHERE AppID=" & prepIntegerSQL(AppID) & " AND QuestionID=" & prepIntegerSQL(QuestionID) & " AND SystemID=" & prepIntegerSQL(UserSystemID) & " AND Version=" & prepIntegerSQL(ScoringVersion)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		If rs.EOF = True Then
			If Len(Score)>0 OR Len(Comments)>0 Then
				' Do an insert
				sql = "INSERT INTO Scoring.Scores (AppID, QuestionID, Version, SystemID, Score, Comments, UpdateTimestamp) VALUES " & vbCrLF & _
					"(" & prepIntegerSQL(AppID) & ", " & _
					prepIntegerSQL(QuestionID) & ", " & _
					prepIntegerSQL(ScoringVersion) & ", " & _
					prepIntegerSQL(UserSystemID) & ", " & _
					prepIntegerSQL(Score) & ", " & _
					prepStringSQL(Comments) & ", " & _
					prepStringSQL(Timestamp) & ")"
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			End If
		ElseIf Len(Score)=0 And Len(Comments)=0 Then
			' Do a delete
			sql = "DELETE FROM Scoring.Scores " & vbCrLf & _
			"WHERE AppID=" & prepIntegerSQL(AppID) & " AND QuestionID=" & prepIntegerSQL(QuestionID) & " AND Version=" & prepIntegerSQL(ScoringVersion) & " AND SystemID=" & prepIntegerSQL(UserSystemID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		ElseIf rs.Fields("Score")=Score And rs.Fields("Comments")=Comments Then
			' Do nothing.
			If Debug = True Then
				Response.Write("<pre>There has been no change.</pre>" & vbCrLf)
				Response.Flush
			End If
		Else
			' Do an update
			sql = "UPDATE Scoring.Scores " & vbCrLf & _
			"SET Score=" & prepIntegerSQL(Score) & ", Comments=" & prepStringSQL(Comments) & ", UpdateTimestamp=" & prepStringSQL(Timestamp) & vbCrLF & _
			"WHERE AppID=" & prepIntegerSQL(AppID) & " AND QuestionID=" & prepIntegerSQL(QuestionID) & " AND Version=" & prepIntegerSQL(ScoringVersion) & " AND SystemID=" & prepIntegerSQL(UserSystemID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		End If
	End If
Next

If Debug = False Then
	Response.Redirect(url)
Else
	Response.Write("<a href=""" & URL & """>" & URL & "</a>")
End If
%><!--#include file="../includes/prepDB.asp"-->