<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, AppURL, AppID, GrantClassID, MatchTypeID, FiscalYear, RowCount, ButtonChoice,  _
	MatchID, Source, MatchSourceID, Amount, UpdateID, UpdateTimestamp, Narrative, NextCategory
ReDim ProgramCategory(5)
TimeStamp = Now()
GrantClassID = 4

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
				response.write("Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j) & vbCrLf)
			next
		else
			Response.Write("Cookies(""" & i & """)=" & Request.Cookies(i) & vbCrLf)
		end if
	next
	Response.Write("</pre>" & vbCrLf)
	Response.Flush
End If

AppID = Request.Form("AppID")
MatchTypeID = Request.Form("MatchTypeID")
FiscalYear = CInt(Request.Form("FiscalYear"))
RowCount = Request.Form("RowCount")
ButtonChoice = Request.Form("ButtonChoice")
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

If Debug = True Then
	Response.Write("<pre>Made it to Line 43</pre>" & vbCrLf)
	Response.Flush
End If

If GrantClassID = 1 And FiscalYear<2022 Then
	AppURL = "Negotiation.asp"
Else
	AppURL = getHomeNegotiationReferenceByAppID(AppID)
End If

If Debug = True Then
	Response.Write("<pre>Made it to Line 49</pre>" & vbCrLf)
	Response.Flush
End If

For i = 1 to RowCount
	MatchID = Request.Form("MatchID_" & i)
	Source = Request.Form("Source_" & i)
	MatchSourceID = Request.Form("MatchSourceID_" & i)
	Amount = Request.Form("Amount_" & i)
	If Len(Source)>0 Or Len(Amount)>0 Then
		If CInt(MatchID) = 0 Then ' Insert
			sql  = "INSERT INTO Negotiation.Matches (AppID, MatchTypeID, Source, MatchSourceID, Amount, " & vbCrLf & _
				"	UpdateID, UpdateTimestamp)" & vbCrLf & _
				"VALUES (" & _
				prepIntegerSQL(AppID) & ", " & _
				prepIntegerSQL(MatchTypeID) & ", " & _
				prepStringSQL(Source) & ", " & _
				prepIntegerSQL(MatchSourceID) & ", " & _
				prepNumberSQL(Amount) & ", " & _
				prepIntegerSQL(UpdateID) & ", " & _
				prepStringSQL(TimeStamp) & ")"
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		Else ' Update
			sql = "UPDATE Negotiation.Matches SET " & vbCrLF & _
				"AppID=" & prepIntegerSQL(AppID) & ", " & _
				"MatchTypeID=" & prepIntegerSQL(MatchTypeID) & ", " & _
				"Source=" & prepStringSQL(Source) & ", " & _
				"MatchSourceID=" & prepIntegerSQL(MatchSourceID) & ", " & _
				"Amount=" & prepNumberSQL(Amount) & ", " & _
				"UpdateID=" & prepIntegerSQL(UpdateID) & ", " & _
				"UpdateTimeStamp=" & prepStringSQL(TimeStamp) & " " & _
				"WHERE MatchID=" & prepIntegerSQL(MatchID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		End If
	ElseIf Cint(MatchID)>0 Then ' Delete
		sql = "DELETE FROM Negotiation.Matches WHERE MatchID=" & prepIntegerSQL(MatchID)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs=Con.Execute(sql)
	End If
Next

If Debug = True Then
Response.Write("<pre>ButtonChoice=" & ButtonChoice & "</pre>" & vbCrLf)
End If

If ButtonChoice = "save" Then
	If Debug = True Then
		Response.Write("<a href=""Matches.asp?AppID=" & AppID & "&MatchTypeID=" & MatchTypeID & _
		""">return</a>")
	Else
		Response.Redirect("Matches.asp?AppID=" & AppID & "&MatchTypeID=" & MatchTypeID)
	End If
Else
	If Debug = True Then
		Response.Write("<a href=""" & AppURL & """>return</a>")
	Else
		Response.Redirect(AppURL)
	End If
End If
%><!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/HomeRef.asp"-->
