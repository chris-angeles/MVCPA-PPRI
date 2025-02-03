<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, GrantClassID, AppURL, ApplicationSchema, AppID, FiscalYear, RowCount, ButtonChoice,  _
	StatisticsID, Jurisdiction, CCTheft1, CCTheft2, UpdateID, UpdateTimestamp, Narrative, NextCategory
ReDim ProgramCategory(5)
TimeStamp = Now()

debug = False
GrantClassID=4

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
	Response.Write("</pre>" & vbCrLf)
	Response.Flush
End If


AppID = Request.Form("AppID")
FiscalYear = CInt(Request.Form("FiscalYear"))
RowCount = CInt(Request.Form("RowCount"))
ButtonChoice = Request.Form("ButtonChoice")
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

If Len(Request.Form("ApplicationSchema")) > 0 Then
	ApplicationSchema = Request.Form("ApplicationSchema")
ElseIf Len(Request.QueryString("ApplicationSchema")) > 0 Then
	ApplicationSchema = Request.QueryString("ApplicationSchema")
End If

If ApplicationSchema = "Application" Or ApplicationSchema = "Negotiation" Then
	If Debug = True Then
		Response.Write("<pre>ApplicationSchema'" & ApplicationSchema & "</pre>" & vbCrLf)
	End If
Else
	Response.Write("Error in Application Schema: " & ApplicationSchema)
	Response.End
End If

IF GrantClassID = 1 And FiscalYear<2022 Then
	AppURL = "Application.asp"
Else
	AppURL = getHomeNegotiationReferenceByGrantClass(GrantClassID, AppID)
End If

For i = 1 to RowCount
	StatisticsID = CInt(Request.Form("StatisticsID_" & i))
	Jurisdiction = Request.Form("Jurisdiction_" & i)
	CCTheft1 = Request.Form("CCTheft1_" & i)
	CCTheft2 = Request.Form("CCTheft2_" & i)

	If Len(Jurisdiction)>0 Or Len(CCTheft1)>0 Or Len(CCTheft2)>0 Then
		If CInt(StatisticsID) = 0 Then ' Insert
			sql  = "INSERT INTO CC." & ApplicationSchema & "Statistics (AppID, Jurisdiction, CCTheft1, CCTheft2, " & vbCrLf & _
				"	UpdateID, UpdateTimestamp)" & vbCrLf & _
				"VALUES (" & _
				prepIntegerSQL(AppID) & ", " & _
				prepStringSQL(Jurisdiction) & ", " & _
				prepNumberSQL(CCTheft1) & ", " & _
				prepNumberSQL(CCTheft2) & ", " & _
				prepIntegerSQL(UpdateID) & ", " & _
				prepStringSQL(TimeStamp) & ")"
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		Else ' Update
			sql = "UPDATE CC." & ApplicationSchema & "Statistics SET " & vbCrLf & _
				"AppID=" & prepIntegerSQL(AppID) & ", " & _
				"Jurisdiction=" & prepStringSQL(Jurisdiction) & ", " & _
				"CCTheft1=" & prepNumberSQL(CCTheft1) & ", " & _
				"CCTheft2=" & prepNumberSQL(CCTheft2) & ", " & _
				"UpdateID=" & prepIntegerSQL(UpdateID) & ", " & _
				"UpdateTimeStamp=" & prepStringSQL(TimeStamp) & " " & _
				"WHERE StatisticsID=" & prepIntegerSQL(StatisticsID)
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		End If
	ElseIf Cint(StatisticsID)>0 Then ' Delete
		sql = "DELETE FROM CC." & ApplicationSchema & "Statistics WHERE StatisticsID=" & prepIntegerSQL(StatisticsID)
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
		Response.Write("<a href=""Statistics.asp?AppID=" & AppID & "&ApplicationSchema=" & ApplicationSchema & """>return</a>")
	Else
		Response.Redirect("Statistics.asp?AppID=" & AppID & "&ApplicationSchema=" & ApplicationSchema)
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

