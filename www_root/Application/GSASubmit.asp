<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, AppID, GrantClassID, FiscalYear, MeasureNumber, Target, Version, AppURL, ApplicationSchema
ReDim ProgramCategory(5)
TimeStamp = Now()

debug = False
ApplicationSchema = "Application"

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
	Response.Flush
End If

AppID = Request.Form("AppID")
GrantClassID = CInt(Request.Form("GrantClassID"))
Version = CInt(Request.Form("Version"))
FiscalYear = CInt(Request.Form("FiscalYear"))

If FiscalYear<2022 Then
	AppURL = "Application.asp"
Else
	AppURL = getHomeApplicationReferenceByGrantClass(GrantClassID, AppID)
End If

sql = "SELECT G.Version, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
	"	T.IntegerResponse, T.DecimalResponse " & vbCrLf & _
	"FROM Lookup.Goals AS G " & vbCrLf & _
	"LEFT JOIN Lookup.Strategies AS S ON S.Version=G.Version AND S.GoalID=G.GoalID " & vbCrLf & _
	"LEFT JOIN Lookup.Activities AS A ON A.Version=S.Version AND A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".GSATargets AS T ON T.AppID=" & prepIntegerSQL(AppID) & " AND T.Version=A.Version " & vbCrLf & _
	" AND T.GoalID=G.GoalID AND T.StrategyID=S.StrategyID AND T.ActivityID=A.ActivityID AND T.MeasureID=A.MeasureID " & vbCrLF & _
	"WHERE G.Version=" & prepIntegerSQL(Version) & "  AND A.MeasureID IS NOT NULL " & vbCrLf & _
	"ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

While rs.EOF = False
	' Check each question and record response if any. Delete response if necessary.
	If IsNull(rs.Fields("MeasureNumber")) = False Then
		MeasureNumber = Replace(rs.Fields("MeasureNumber"),".","_")
	End If
	Target = Request.Form("Target_" & MeasureNumber)
	If Debug = True Then
		Response.Write("<pre>Target_" & MeasureNumber & "=""" & Target & """; Length(Target)=" & Len(Target) & "; ResponseTypeID=" & rs.Fields("ResponseTypeID") & "; DB Value=" & rs.Fields("IntegerResponse") & "/" & rs.Fields("DecimalResponse") & " </pre>")
	End If
	If rs.Fields("ResponseTypeID")=1 Then
		If Len(Target)=0 And IsNull(rs.Fields("IntegerResponse"))=True Then
			' Do nothing.
		ElseIf Len(Target)>0 And IsNull(rs.Fields("IntegerResponse"))=True Then
			' Do an insert.
			If rs.Fields("ResponseTypeID") = 1 Then
				sql = "INSERT INTO " & ApplicationSchema & ".GSATargets (AppID, Version, GoalID, StrategyID, ActivityID, MeasureID, IntegerResponse, UpdateID, UpdateTimestamp) " & vbCrLf & _
					"VALUES (" & prepIntegerSQL(AppID) & ", " & _
					prepIntegerSQL(rs.Fields("Version")) & ", " & _
					prepIntegerSQL(rs.Fields("GoalID")) & ", " & _
					prepIntegerSQL(rs.Fields("StrategyID")) & ", " & _
					prepIntegerSQL(rs.Fields("ActivityID")) & ", " & _
					prepIntegerSQL(rs.Fields("MeasureID")) & ", " & _
					prepIntegerSQL(Target) & ", " & _
					prepIntegerSQL(UserSystemID) & ", " & _
					prepStringSQL(Timestamp) & ")"
			Else
			End If
		ElseIf Len(Target)=0 And IsNull(rs.Fields("IntegerResponse")) = False Then
			' Do a delete.
			sql = "DELETE FROM " & ApplicationSchema & ".GSATargets " & vbCrLF & _
				"WHERE AppID=" & prepIntegerSQL(AppID) & _
				" AND Version=" & prepIntegerSQL(rs.Fields("Version")) & _
				" AND GoalID=" & prepIntegerSQL(rs.Fields("GoalID")) & _
				" AND StrategyID=" & prepIntegerSQL(rs.Fields("StrategyID")) & _
				" AND ActivityID=" & prepIntegerSQL(rs.Fields("ActivityID")) & _
				" AND MeasureID=" & prepIntegerSQL(rs.Fields("MeasureID"))
		ElseIf CLng(Target) = rs.Fields("IntegerResponse") Then
			' Do nothing. Response are the same.
		Else
			' Do an update.
			sql = "UPDATE " & ApplicationSchema & ".GSATargets SET " & vbCRLF & _
				"IntegerResponse=" & prepIntegerSQL(Target) & ", " & _
				"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
				"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
				"WHERE AppID=" & prepIntegerSQL(AppID) & _
				" AND Version=" & prepIntegerSQL(rs.Fields("Version")) & _
				" AND GoalID=" & prepIntegerSQL(rs.Fields("GoalID")) & _
				" AND StrategyID=" & prepIntegerSQL(rs.Fields("StrategyID")) & _
				" AND ActivityID=" & prepIntegerSQL(rs.Fields("ActivityID")) & _
				" AND MeasureID=" & prepIntegerSQL(rs.Fields("MeasureID"))
		End If
	ElseIf rs.Fields("ResponseTypeID")=2 Or rs.Fields("ResponseTypeID")=3 Then
		If Len(Target)>0 Then
			Target = CDbl(Replace(Target,"$",""))
		End If
		If Len(Target)=0 And IsNull(rs.Fields("DecimalResponse"))=True Then
			' Do nothing.
		ElseIf Len(Target)>0 And IsNull(rs.Fields("DecimalResponse"))=True Then
			' Do an insert.
			sql = "INSERT INTO " & ApplicationSchema & ".GSATargets (AppID, Version, GoalID, StrategyID, ActivityID, MeasureID, DecimalResponse, UpdateID, UpdateTimestamp) " & vbCrLf & _
				"VALUES (" & prepIntegerSQL(AppID) & ", " & _
				prepIntegerSQL(rs.Fields("Version")) & ", " & _
				prepIntegerSQL(rs.Fields("GoalID")) & ", " & _
				prepIntegerSQL(rs.Fields("StrategyID")) & ", " & _
				prepIntegerSQL(rs.Fields("ActivityID")) & ", " & _
				prepIntegerSQL(rs.Fields("MeasureID")) & ", " & _
				prepNumberSQL(Target) & ", " & _
				prepIntegerSQL(UserSystemID) & ", " & _
				prepStringSQL(Timestamp) & ")"
		ElseIf Len(Target)=0 And IsNull(rs.Fields("DecimalResponse")) = False Then
			' Do a delete.
			sql = "DELETE FROM " & ApplicationSchema & ".GSATargets " & vbCrLF & _
				"WHERE AppID=" & prepIntegerSQL(AppID) & _
				" AND Version=" & prepIntegerSQL(rs.Fields("Version")) & _
				" AND GoalID=" & prepIntegerSQL(rs.Fields("GoalID")) & _
				" AND StrategyID=" & prepIntegerSQL(rs.Fields("StrategyID")) & _
				" AND ActivityID=" & prepIntegerSQL(rs.Fields("ActivityID")) & _
				" AND MeasureID=" & prepIntegerSQL(rs.Fields("MeasureID"))
		ElseIf Target = CDbl(rs.Fields("DecimalResponse")) Then
			' Do nothing. Response are the same.
		Else
			' Do an update.
			sql = "UPDATE " & ApplicationSchema & ".GSATargets SET " & vbCRLF & _
				"DecimalResponse=" & prepNumberSQL(Target) & ", " & _
				"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
				"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
				"WHERE AppID=" & prepIntegerSQL(AppID) & _
				" AND Version=" & prepIntegerSQL(rs.Fields("Version")) & _
				" AND GoalID=" & prepIntegerSQL(rs.Fields("GoalID")) & _
				" AND StrategyID=" & prepIntegerSQL(rs.Fields("StrategyID")) & _
				" AND ActivityID=" & prepIntegerSQL(rs.Fields("ActivityID")) & _
				" AND MeasureID=" & prepIntegerSQL(rs.Fields("MeasureID"))
		End If
	End If
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	If Len(sql)>0 Then
		Con.Execute(sql)
	End If
	sql = ""
	rs.MoveNext
Wend

If Debug = True Then
	If GrantClassID = 4 Then
		Response.Write("<a href=""/CatalyticConverter/Application.asp?AppID=" & AppID & """>return to App</a><br />" & vbCrLf)
	ElseIf FiscalYear >=2022 Then
		Response.Write("<a href=""TFGApplication.asp?AppID=" & AppID & """>return to App</a><br />" & vbCrLf)
	Else
		Response.Write("<a href=""Application.asp?AppID=" & AppID & """>return to App</a><br />" & vbCrLf)
	End If
	Response.Write("<a href=""GSA.asp?AppID=" & AppID & """>return to GSA</a><br />" & vbCrLf)
Else
	If GrantClassID = 4 Then
		Response.Redirect("/CatalyticConverter/Application.asp?AppID=" & AppID)
	ElseIf FiscalYear >= 2022 Then
		Response.Redirect("TFGApplication.asp?AppID=" & AppID)
	Else
		Response.Redirect("Application.asp?AppID=" & AppID)
	End If
End If

%><!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/HomeRef.asp"-->