<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, MAGID, Quarter, Version, QuestionID, update, months, submit, confirmed, _
	GranteeID, FiscalYear, BorderCounty, PortCounty, _
	value1, value2, value3, text, dbvalue1, dbvalue2, dbvalue3, dbtext, _
	AdministrativeComments, ApprovalDate, ApprovalID, AdministrativeUpdate, Unsubmit

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

MAGID = Request.Form("MAGID")
Quarter = Request.Form("Quarter")
Version = Request.Form("Version")
If Request.Form("Confirmed") = "1" Then
	confirmed = true
Else
	confirmed = false
End If
If Request.Form("action") = "submit" Then
	submit = True
Else
	submit = False
End If
If Request.Form("Unsubmit") = "1" Then
	Unsubmit = True
Else
	Unsubmit = False
End If
If Len(MAGID) = 0 Then
	Response.Write("Error: No MAGID provided.")
	sendWarning("Error: No MAGID provided.")
	Response.End
Else
	MAGID = CInt(MAGID)
End If

If Len(Quarter) = 0 Then
	Response.Write("Error: No Quarter provided.")
	sendWarning("Error: No Quarter provided.")
	Response.End
Else
	Quarter = CInt(Quarter)
End If

sql = "SELECT H.MAGID, H.FiscalYear, G.GranteeID, G.GranteeName, 'MVCPA Auxiliary Grant' AS ProgramName, " & vbCrLf & _
	"	ISNULL(G.BorderCounty,0) AS BorderCounty, ISNULL(G.PortCounty,0) AS PortCounty " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN [MAG].Main AS H ON G.GranteeID=H.GranteeID " & vbCrLf & _
	"WHERE H.MAGID=" & prepIntegerSQL(MAGID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF Then
	Response.Write("Error: Grant not found.")
	Response.End
Else
	MAGID = rs.Fields("MAGID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	BorderCounty = rs.Fields("BorderCounty")
	PortCounty = rs.Fields("PortCounty")
End If

If Debug = True Then
	Response.Write("<pre>Quarter=" & Quarter & vbCrLf & "Submit=" & Submit & "</pre>")
End If

sql = "SELECT A.QuestionID, ISNULL(R.Quarter," & prepIntegerSQL(Quarter) & ") AS Quarter, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.MAGSpecial, A.ResponseTypeID, " & vbCrLf & _
	"	IntegerResponse_M1, IntegerResponse_M2, IntegerResponse_M3, " & vbCrLf & _
	"	DecimalResponse_M1, DecimalResponse_M2, DecimalResponse_M3, " & vbCrLf & _
	"	TextResponse, " & vbCrLf & _
	"	CAST(CASE WHEN R.MAGID IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS RecordPresent " & vbCrLf & _
	"FROM PR.Goals AS G " & vbCrLf & _
	"LEFT JOIN PR.Strategies AS S ON S.GoalID=G.GoalID AND S.Version=G.Version " & vbCrLf & _
	"LEFT JOIN PR.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID AND A.Version=G.Version " & vbCrLf & _
	"LEFT JOIN MAG.ProgressReportResponses AS R ON R.MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND R.QuestionID=A.QuestionID " & vbCrLf
	If BorderCounty = True Or PortCounty = True Then
		sql = sql & "WHERE G.Version=" & prepIntegerSQL(Version) & " AND (ISNULL(A.Mandatory,0)=1 OR G.GoalID IN (4,7) OR ISNULL(MAGSpecial,0)=1) AND A.QuestionID NOT IN (552, 557) " & vbCrLf
	Else
		sql = sql & "WHERE G.Version=" & prepIntegerSQL(Version) & " AND (ISNULL(A.Mandatory,0)=1 OR G.GoalID IN (7) OR ISNULL(MAGSpecial,0)=1) AND A.QuestionID NOT IN (552, 557) " & vbCrLf
	End If
	sql = sql & "ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
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
	'If Quarter = 1 Then
		If rs.Fields("ResponseTypeID")=1 or rs.Fields("ResponseTypeID")=6 or rs.Fields("ResponseTypeID")=7 Then
			value1 = Request.Form("Response_M1_" & QuestionID)
			value2 = Request.Form("Response_M2_" & QuestionID)
			value3 = Request.Form("Response_M3_" & QuestionID)
			dbvalue1 = rs.Fields("IntegerResponse_M1")
			dbvalue2 = rs.Fields("IntegerResponse_M2")
			dbvalue3 = rs.Fields("IntegerResponse_M3")
			If Debug = True Then
				Response.Write("<pre>QuestionID=" & QuestionID & "; ResponseTypeID=" & rs.Fields("ResponseTypeID") & "; Response_M1_" & QuestionID & "=" & _
					Request.Form("Response_M1_" & QuestionID) & "; dbvalue=" & rs.Fields("IntegerResponse_M1") & _
					"; value1='" & value1 & "'</pre>")
			End If
			If Len(value1)=0 And IsNull(dbvalue1)=False Then
				update = True
			ElseIf Len(value2)=0 And IsNull(dbvalue2)=False Then
				update = True
			ElseIf Len(value3)=0 And IsNull(dbvalue3)=False Then
				update = True
			ElseIf Len(value1)>0 And IsNull(dbvalue1)=True Then
				update = True
			ElseIf Len(value2)>0 And IsNull(dbvalue2)=True Then
				update = True
			ElseIf Len(value3)>0 And IsNull(dbvalue3)=True Then
				update = True
			ElseIf value1 <> dbvalue1 Then
				update = True
			ElseIf value2 <> dbvalue2 Then
				update = True
			ElseIf value3 <> dbvalue3 Then
				update = True
			End If
			If update=True Then
				If rs.Fields("RecordPresent") = False Then
					' Do an insert.
					sql = "INSERT INTO MAG.ProgressReportResponses (MAGID, Quarter, QuestionID, " & vbCrLf & _
						"	IntegerResponse_M1, IntegerResponse_M2, IntegerResponse_M3, UpdateID, UpdateTimestamp) " & vbCrLf & _
						"VALUES (" & prepIntegerSQL(MAGID) & ", " & _
						prepIntegerSQL(rs.Fields("Quarter")) & ", " & _
						prepIntegerSQL(rs.Fields("QuestionID")) & ", " & _
						prepIntegerSQL(value1) & ", " & _
						prepIntegerSQL(value2) & ", " & _
						prepIntegerSQL(value3) & ", " & _
						prepIntegerSQL(UserSystemID) & ", " & _
						prepStringSQL(Timestamp) & ")"
				Else
					' Do an update.
					sql = "UPDATE MAG.ProgressReportResponses SET " & vbCRLF & _
						"IntegerResponse_M1=" & prepIntegerSQL(value1) & ", " & _
						"IntegerResponse_M2=" & prepIntegerSQL(value2) & ", " & _
						"IntegerResponse_M3=" & prepIntegerSQL(value3) & ", " & _
						"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
						"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
						"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter) & _
						" AND QuestionID=" & prepIntegerSQL(rs.Fields("QuestionID")) 
				End If
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			End If
		ElseIf rs.Fields("ResponseTypeID")=2 Or rs.Fields("ResponseTypeID")=3 Then
			value1 = Trim(Request.Form("Response_M1_" & QuestionID))
			value2 = Trim(Request.Form("Response_M2_" & QuestionID))
			value3 = Trim(Request.Form("Response_M3_" & QuestionID))
			dbvalue1 = rs.Fields("DecimalResponse_M1")
			dbvalue2 = rs.Fields("DecimalResponse_M2")
			dbvalue3 = rs.Fields("DecimalResponse_M3")
			If Debug = True Then
				Response.Write("<pre>QuestionID=" & QuestionID & "; Response_M1_" & QuestionID & "=" & _
					Request.Form("Response_M1_" & QuestionID) & "; dbvalue=" & rs.Fields("DecimalResponse_" & Months(0)) & _
					"; value1='" & value1 & "'; Len(value1)=" & Len(value1) & "; ISNULL(dbvalue)=" & IsNull(rs.Fields("DecimalResponse_M1")) & vbCrLf)
				Response.Write("QuestionID=" & QuestionID & "; Response_M2_" & QuestionID & "=" & _
					Request.Form("Response_M2_" & QuestionID) & "; dbvalue=" & rs.Fields("DecimalResponse_M2") & _
					"; value2='" & value2 & "'; Len(value2)=" & Len(value2) & "; ISNULL(dbvalue)=" & IsNull(rs.Fields("DecimalResponse_M2")) & vbCrLf)
				Response.Write("QuestionID=" & QuestionID & "; Response_M3_" & QuestionID & "=" & _
					Request.Form("Response_M3_" & QuestionID) & "; dbvalue=" & rs.Fields("DecimalResponse_M3") & _
					"; value2='" & value3 & "'; Len(value3)=" & Len(value3) & "; ISNULL(dbvalue)=" & IsNull(rs.Fields("DecimalResponse_M3")) & "</pre>")
			End If

			If CompareNumbers(value1, dbvalue1) = False Then
				update = True
			ElseIf CompareNumbers(value2, dbvalue2) = False Then
				update = True
			ElseIf CompareNumbers(value3, dbvalue3) = False Then
				update = True
			Else
				update = False
			End If 

			If update=True Then
				If rs.Fields("RecordPresent") = False Then
					' Do an insert.
					sql = "INSERT INTO MAG.ProgressReportResponses (MAGID, Quarter, QuestionID, DecimalResponse_M1, DecimalResponse_M2, DecimalResponse_M3, UpdateID, UpdateTimestamp) " & vbCrLf & _
						"VALUES (" & prepIntegerSQL(MAGID) & ", " & _
						prepIntegerSQL(rs.Fields("Quarter")) & ", " & _
						prepIntegerSQL(rs.Fields("QuestionID")) & ", " & _
						prepNumberSQL(value1) & ", " & _
						prepNumberSQL(value2) & ", " & _
						prepNumberSQL(value3) & ", " & _
						prepIntegerSQL(UserSystemID) & ", " & _
						prepStringSQL(Timestamp) & ")"
				Else
					' Do an update.
					sql = "UPDATE MAG.ProgressReportResponses SET " & vbCRLF & _
						"DecimalResponse_M1=" & prepNumberSQL(value1) & ", " & _
						"DecimalResponse_M2=" & prepNumberSQL(value2) & ", " & _
						"DecimalResponse_M3=" & prepNumberSQL(value3) & ", " & _
						"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
						"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
						"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter) & _
						" AND QuestionID=" & prepIntegerSQL(rs.Fields("QuestionID")) 
				End If
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			End If
		ElseIf rs.Fields("ResponseTypeID")=5 Then
			text = Request.Form("Response_" & QuestionID)
			dbtext = rs.Fields("TextResponse")
			If Debug = True Then
				Response.Write("<pre>QuestionID=" & QuestionID & "; Response_" & QuestionID & "=" & _
					Request.Form("Response_" & QuestionID) & "; dbvalue=" & rs.Fields("TextResponse") & _
					"; value1='" & text & "'</pre>")
			End If
			If Len(text)=0 And IsNull(dbtext)=False Then
				update = True
			ElseIf Len(text)>0 And IsNull(dbtext)=True Then
				update = True
			ElseIf text <> dbtext Then
				update = True
			End If
			If update=True Then
				If rs.Fields("RecordPresent") = False Then
					' Do an insert.
					sql = "INSERT INTO MAG.ProgressReportResponses (MAGID, Quarter, QuestionID, TextResponse, UpdateID, UpdateTimestamp) " & vbCrLf & _
						"VALUES (" & prepStringSQL(MAGID) & ", " & _
						prepIntegerSQL(Quarter) & ", " & _
						prepIntegerSQL(rs.Fields("QuestionID")) & ", " & _
						prepStringSQL(text) & ", " & _
						prepIntegerSQL(UserSystemID) & ", " & _
						prepStringSQL(Timestamp) & ")"
				Else
					' Do an update.
					sql = "UPDATE MAG.ProgressReportResponses SET " & vbCRLF & _
						"TextResponse=" & prepStringSQL(text) & ", " & _
						"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
						"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
						"WHERE MAGID=" & prepIntegerSQL(MAGID) & _
						" AND QuestionID=" & prepIntegerSQL(rs.Fields("QuestionID")) 
				End If
				If Debug = True Then
					Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
					Response.Flush
				End If
				Con.Execute(sql)
			End If
		End If
	'End If
	rs.MoveNext
Wend

If True = False Then
	sql = "SELECT MAGID, Quarter, ISNULL(Confirmed,0) AS Confirmed " & vbCrLF & _
		"FROM MAG.ProgressReportResponses " & vbCrLF & _
		"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then ' Do an insert
		sql = "INSERT INTO MAG.ProgressReportResponses (MAGID, Quarter, Confirmed) " & vbCrLf & _
			"VALUES (" & prepIntegerSQL(MAGID) & ", " & prepIntegerSQL(Quarter) & ", " & prepBitSQL(confirmed)& ")" & vbCrLf
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	ElseIf rs.Fields("Confirmed") <> Confirmed Then ' Do an update
		sql = "UPDATE MAG.ProgressReportResponses " & vbCrLF & _
			"SET Confirmed=" & prepBitSQL(Confirmed) & " " & vbCrLF & _
			"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
End If


If Submit = True Then
	sql = "SELECT MAGID, Quarter, Confirmed, SubmitID, SubmitTimestamp " & vbCrLF & _
		"FROM MAG.ProgressReportSubmissions " & vbCrLF & _
		"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then ' Do an insert
		sql = "INSERT INTO MAG.ProgressReportSubmissions (MAGID, Quarter, Confirmed, SubmitID, SubmitTimestamp) " & vbCrLf & _
			"VALUES (" & prepIntegerSQL(MAGID) & ", " & prepIntegerSQL(Quarter) & ", " & prepBitSQL(Confirmed) & ", " & prepIntegerSQL(UserSystemID) & ", " & prepStringSQL(Timestamp) & ")" & vbCrLf
	Else ' Do an update
		sql = "UPDATE MAG.ProgressReportSubmissions " & vbCrLF & _
			"SET Confirmed=" & prepBitSQL(Confirmed) & ", SubmitID=" & prepIntegerSQL(UserSystemID) & ", SubmitTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLF & _
			"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	End If
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

If MVCPARights = True Then
	AdministrativeUpdate = False
	ApprovalDate = Request.Form("ApprovalDate")
	AdministrativeComments = Request.Form("AdministrativeComments")

	sql = "SELECT MAGID, Quarter, SubmitID, AdministrativeComments, ApprovalID, ApprovalDate " & vbCrLF & _
		"FROM MAG.ProgressReportSubmissions " & vbCrLF & _
		"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If Len(ApprovalDate)>0 And IsNull(rs.Fields("ApprovalDate")) = True Then
		ApprovalID = UserSystemID
	ElseIf Len(ApprovalDate)=0 Then
		ApprovalID = null
	Else
		ApprovalID = rs.Fields("ApprovalID")
	End If
	If Len(ApprovalDate)>0 And IsNull(rs.Fields("ApprovalDate")) = True Then
		AdministrativeUpdate = True
	ElseIf Len(ApprovalDate) = 0 And IsNull(rs.Fields("ApprovalDate")) = False Then
		AdministrativeUpdate = True
	ElseIf Len(ApprovalDate)>0 And IsNull(rs.Fields("ApprovalDate")) = False Then
		If CDate(ApprovalDate) <> CDate(rs.Fields("ApprovalDate")) Then
			AdministrativeUpdate = True
		End If
	End If

	If rs.EOF = False Then
		If Len(AdministrativeComments)>0 And IsNull(rs.Fields("AdministrativeComments")) = True Then
			AdministrativeUpdate = True
		ElseIf Len(AdministrativeComments)=0 And IsNull(rs.Fields("AdministrativeComments")) = False Then
			AdministrativeUpdate = True
		ElseIf Len(AdministrativeComments)>0 And IsNull(rs.Fields("AdministrativeComments")) = False Then
			If AdministrativeComments <> rs.Fields("AdministrativeComments") Then
				AdministrativeUpdate = True
			End If
		End If
	ElseIf Len(AdministrativeComments)>0 Then
		AdministrativeUpdate = True
	End If

	If Unsubmit = True Then
		sql = "UPDATE MAG.ProgressReportSubmissions " & vbCrLF & _
			"SET Confirmed=null, SubmitID=null, SubmitTimestamp=null, AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ",ApprovalID=null, ApprovalDate= null " & vbCrLF & _
			"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	ElseIf AdministrativeUpdate = True Then
		If rs.EOF = True Then ' Do an insert
			sql = "INSERT INTO MAG.ProgressReportSubmissions (MAGID, Quarter, AdministrativeComments, ApprovalID, ApprovalDate) " & vbCrLf & _
				"VALUES (" & prepIntegerSQL(MAGID) & ", " & prepIntegerSQL(Quarter) & ", " & prepStringSQL(AdministrativeComments) & ", " & prepIntegerSQL(ApprovalID) & ", " & prepStringSQL(ApprovalDate) & ")"
		Else ' Do an update
			sql = "UPDATE MAG.ProgressReportSubmissions " & vbCrLF & _
				"SET AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ", ApprovalID=" & prepIntegerSQL(ApprovalID) & ", ApprovalDate=" & prepStringSQL(ApprovalDate) & " " & vbCrLF & _
				"WHERE MAGID=" & prepIntegerSQL(MAGID) & " AND Quarter=" & prepIntegerSQL(Quarter)
		End If
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
End If

If Debug = True Then
	Response.Write("<a href=""ProgressReport.asp?MAGID=" & MAGID & "&Quarter=" & Quarter & """>return to progress report</a><br />" & vbCrLf)
	Response.Write("<a href=""../Home/Default.asp?GranteeID=" & GranteeID & """>return to Home</a><br />" & vbCrLf)
	Response.End
Else
	Response.Redirect("ProgressReport.asp?MAGID=" & MAGID & "&Quarter=" & Quarter)
End If

Function CompareNumbers(a, b)
	' compare two values for equality. Treat null and empty string as equivalent for comparison.
	If Debug = True Then
		Response.Write("<pre>a=" & a & "; b=" & b & "; VarType(a)=" & VarType(a) & "; VarType(b)=" & VarType(b) & "</pre>")
		Response.Flush
	End If
	If IsNull(a) = True and IsNull(b) = True Then
		CompareNumbers = True
	ElseIf IsNull(a) = True and IsNull(b) = False Then
		If Len(b)>0 Then
			CompareNumbers = False
		Else
			CompareNumbers = True
		End If
	ElseIf IsNull(a) = False and IsNull(b) = True Then
		If Len(a)>0 Then
			CompareNumbers = False
		Else
			CompareNumbers = True
		End If
	ElseIf Len(a)=0 and Len(b)=0 Then
		CompareNumbers = True
	ElseIf Len(a)=0 and Len(b)>0 Then
		CompareNumbers = False
	ElseIf Len(a)>0 and Len(b)=0 Then
		CompareNumbers = False
	ElseIf CDbl(a)=CDbl(b) Then
		CompareNumbers = True
	Else
		CompareNumbers = False
	End If
End Function

Function CompareString(a, b)
	' compare two values for equality. Treat null and empty string as equivalent for comparison.
	If Debug = True Then
		Response.Write("<pre>a=" & a & "; b=" & b & "; VarType(a)=" & VarType(a) & "; VarType(b)=" & VarType(b) & "</pre>")
		Response.Flush
	End If
	If IsNull(a) = True and IsNull(b) = True Then
		CompareString = True
	ElseIf IsNull(a) = True and IsNull(b) = False Then
		If Len(b)>0 Then
			CompareString = False
		Else
			CompareString = True
		End If
	ElseIf IsNull(a) = False and IsNull(b) = True Then
		If Len(a)>0 Then
			CompareString = False
		Else
			CompareString = True
		End If
	ElseIf Len(a)=0 and Len(b)=0 Then
		CompareString = True
	ElseIf CStr(a)=CStr(b) Then
		CompareString = True
	Else
		CompareString = False
	End If
End Function

%><!--#include file="../includes/prepDB.asp"-->