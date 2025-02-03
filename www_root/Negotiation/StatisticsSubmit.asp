<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, AppURL, ApplicationSchema, AppID, FiscalYear, RowCount, ButtonChoice,  _
	StatisticsID, Jurisdiction, MVT1, BMV1, FRMVC1, MVT2, BMV2, FRMVC2, UpdateID, UpdateTimestamp, Narrative, NextCategory
ReDim ProgramCategory(5)
TimeStamp = Now()

debug = False
ApplicationSchema = "Negotiation"

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
RowCount = Request.Form("RowCount")
ButtonChoice = Request.Form("ButtonChoice")
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

AppURL = "TFGApplication.asp"


For i = 1 to RowCount
	StatisticsID = Request.Form("StatisticsID_" & i)
	Jurisdiction = Request.Form("Jurisdiction_" & i)
	MVT1 = Request.Form("MVT1_" & i)
	BMV1 = Request.Form("BMV1_" & i)
	FRMVC1 = Request.Form("FRMVC1_" & i)
	MVT2 = Request.Form("MVT2_" & i)
	BMV2 = Request.Form("BMV2_" & i)
	FRMVC2 = Request.Form("FRMVC2_" & i)
	If Len(Jurisdiction)>0 Or Len(MVT1)>0 Or Len(BMV1)>0 Or Len(FRMVC1)>0 Or Len(MVT2)>0 Or Len(BMV2)>0 Or Len(FRMVC2)>0 Then
		If CInt(StatisticsID) = 0 Then ' Insert
			sql  = "INSERT INTO " & ApplicationSchema & ".[Statistics] (AppID, Jurisdiction, MVT1, BMV1, FRMVC1, MVT2, BMV2, FRMVC2, " & vbCrLf & _
				"	UpdateID, UpdateTimestamp)" & vbCrLf & _
				"VALUES (" & _
				prepIntegerSQL(AppID) & ", " & _
				prepStringSQL(Jurisdiction) & ", " & _
				prepNumberSQL(MVT1) & ", " & _
				prepNumberSQL(BMV1) & ", " & _
				prepNumberSQL(FRMVC1) & ", " & _
				prepNumberSQL(MVT2) & ", " & _
				prepNumberSQL(BMV2) & ", " & _
				prepNumberSQL(FRMVC2) & ", " & _
				prepIntegerSQL(UpdateID) & ", " & _
				prepStringSQL(TimeStamp) & ")"
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		Else ' Update
			sql = "UPDATE " & ApplicationSchema & ".[Statistics] SET " & vbCrLf & _
				"AppID=" & prepIntegerSQL(AppID) & ", " & _
				"Jurisdiction=" & prepStringSQL(Jurisdiction) & ", " & _
				"MVT1=" & prepNumberSQL(MVT1) & ", " & _
				"BMV1=" & prepNumberSQL(BMV1) & ", " & _
				"FRMVC1=" & prepNumberSQL(FRMVC1) & ", " & _
				"MVT2=" & prepNumberSQL(MVT2) & ", " & _
				"BMV2=" & prepNumberSQL(BMV2) & ", " & _
				"FRMVC2=" & prepNumberSQL(FRMVC2) & ", " & _
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
		sql = "DELETE FROM " & ApplicationSchema & ".[Statistics] WHERE StatisticsID=" & prepIntegerSQL(StatisticsID)
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
		Response.Write("<a href=""Statistics.asp?AppID=" & AppID & """>return</a>")
	Else
		Response.Redirect("Statistics.asp?AppID=" & AppID)
	End If
Else
	If Debug = True Then
		Response.Write("<a href=""" & AppURL & "?AppID=" & AppID & """>return</a>")
	Else
		Response.Redirect(AppURL & "?AppID=" & AppID)
	End If
End If
%><!--#include file="../includes/prepDB.asp"-->

