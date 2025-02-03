<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, GrantID, Quarter, Version, ShowOneQuarter, QuestionID, update, months, submit, confirmed, _
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

GrantID = Request.Form("GrantID")
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
If Len(GrantID) = 0 Then
	Response.Write("Error: No GrantID provided.")
	sendWarning("Error: No GrantID provided.")
	Response.End
Else
	GrantID = CInt(GrantID)
End If
If Request.Form("ShowOneQuarter") = "True" Then
	ShowOneQuarter = True
Else
	ShowOneQuarter = False
End If

If Len(Quarter) = 0 Then
	Response.Write("Error: No Quarter provided.")
	sendWarning("Error: No Quarter provided.")
	Response.End
Else
	Quarter = CInt(Quarter)
End If

If Quarter = 1 Then
	Months = Array("Sep","Oct","Nov")
ElseIf Quarter = 2 Then
	Months = Array("Dec", "Jan", "Feb")
ElseIf Quarter = 3 Then
	Months = Array("Mar", "Apr", "May")
ElseIf Quarter = 4 Then
	Months = Array("Jun", "Jul", "Aug")
Else
	Response.Write("Error: Invalid quarter provided.")
	sendWarning("Error: Invalid quarter provided.")
	Response.End
End If

If Debug = True Then
	Response.Write("<pre>Quarter=" & Quarter & vbCrLf & "Submit=" & Submit & "</pre>")
End If

sql = "SELECT A.QuestionID, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
	"	Q.IntegerTarget, Q.DecimalTarget, " & vbCrLF & _
	"	IntegerResponse_Sep, IntegerResponse_Oct, IntegerResponse_Nov, " & vbCrLf & _
	"	IntegerResponse_Dec, IntegerResponse_Jan, IntegerResponse_Feb, " & vbCrLf & _
	"	IntegerResponse_Apr, IntegerResponse_May, IntegerResponse_Jun, " & vbCrLf & _
	"	IntegerResponse_Mar, IntegerResponse_Jul, IntegerResponse_Aug, " & vbCrLF & _
	"	DecimalResponse_Sep, DecimalResponse_Oct, DecimalResponse_Nov, " & vbCrLf & _
	"	DecimalResponse_Dec, DecimalResponse_Jan, DecimalResponse_Feb, " & vbCrLf & _
	"	DecimalResponse_Mar, DecimalResponse_Apr, DecimalResponse_May, " & vbCrLf & _
	"	DecimalResponse_Jun, DecimalResponse_Jul, DecimalResponse_Aug, " & vbCrLf & _
	"	TextResponse_Q1, TextResponse_Q2, TextResponse_Q3, TextResponse_Q4, " & vbCrLf & _
	"	CAST(CASE WHEN R.GrantID IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS RecordPresent " & vbCrLf & _
	"FROM PR.Goals AS G " & vbCrLf & _
	"LEFT JOIN PR.Strategies AS S ON S.GoalID=G.GoalID AND S.Version=G.Version " & vbCrLf & _
	"LEFT JOIN PR.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID AND A.Version=G.Version " & vbCrLf & _
	"LEFT JOIN PR.GrantQuestions AS Q ON Q.GrantID=" & prepIntegerSQL(GrantID) & " AND Q.QuestionID=A.QuestionID " & vbCrLF & _
	"LEFT JOIN PR.Responses AS R WITH (NOLOCK) ON R.GrantID=" & prepIntegerSQL(GrantID) & " AND R.QuestionID=A.QuestionID " & vbCrLf & _
	"WHERE G.Version=" & prepIntegerSQL(Version) & " " & vbCrLf & _
	"ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
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
			value1 = Request.Form("Response_" & Months(0) & "_" & QuestionID)
			value2 = Request.Form("Response_" & Months(1) & "_" & QuestionID)
			value3 = Request.Form("Response_" & Months(2) & "_" & QuestionID)
			dbvalue1 = rs.Fields("IntegerResponse_" & Months(0))
			dbvalue2 = rs.Fields("IntegerResponse_" & Months(1))
			dbvalue3 = rs.Fields("IntegerResponse_" & Months(2))
			If Debug = True Then
				Response.Write("<pre>QuestionID=" & QuestionID & "; ResponseTypeID=" & rs.Fields("ResponseTypeID") & "; Response_" & Months(0) & "_" & QuestionID & "=" & _
					Request.Form("Response_" & Months(0) & "_" & QuestionID) & "; dbvalue=" & rs.Fields("IntegerResponse_" & Months(0)) & _
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
					sql = "INSERT INTO PR.Responses (GrantID, QuestionID, IntegerResponse_" & Months(0) & ", IntegerResponse_" & Months(1) & ", IntegerResponse_" & Months(2) & ", UpdateID, UpdateTimestamp) " & vbCrLf & _
						"VALUES (" & prepIntegerSQL(GrantID) & ", " & _
						prepIntegerSQL(rs.Fields("QuestionID")) & ", " & _
						prepIntegerSQL(value1) & ", " & _
						prepIntegerSQL(value2) & ", " & _
						prepIntegerSQL(value3) & ", " & _
						prepIntegerSQL(UserSystemID) & ", " & _
						prepStringSQL(Timestamp) & ")"
				Else
					' Do an update.
					sql = "UPDATE PR.Responses SET " & vbCRLF & _
						"IntegerResponse_" & Months(0) & "=" & prepIntegerSQL(value1) & ", " & _
						"IntegerResponse_" & Months(1) & "=" & prepIntegerSQL(value2) & ", " & _
						"IntegerResponse_" & Months(2) & "=" & prepIntegerSQL(value3) & ", " & _
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
		ElseIf rs.Fields("ResponseTypeID")=2 Or rs.Fields("ResponseTypeID")=3 Then
			value1 = Trim(Request.Form("Response_" & Months(0) & "_" & QuestionID))
			value2 = Trim(Request.Form("Response_" & Months(1) & "_" & QuestionID))
			value3 = Trim(Request.Form("Response_" & Months(2) & "_" & QuestionID))
			dbvalue1 = rs.Fields("DecimalResponse_" & Months(0))
			dbvalue2 = rs.Fields("DecimalResponse_" & Months(1))
			dbvalue3 = rs.Fields("DecimalResponse_" & Months(2))
			If Debug = True Then
				Response.Write("<pre>QuestionID=" & QuestionID & "; Response_" & Months(0) & "_" & QuestionID & "=" & _
					Request.Form("Response_" & Months(0) & "_" & QuestionID) & "; dbvalue=" & rs.Fields("DecimalResponse_" & Months(0)) & _
					"; value1='" & value1 & "'; Len(value1)=" & Len(value1) & "; ISNULL(dbvalue)=" & ISNull(rs.Fields("DecimalResponse_" & Months(0))) & vbCrLf)
				Response.Write("QuestionID=" & QuestionID & "; Response_" & Months(1) & "_" & QuestionID & "=" & _
					Request.Form("Response_" & Months(1) & "_" & QuestionID) & "; dbvalue=" & rs.Fields("DecimalResponse_" & Months(1)) & _
					"; value2='" & value2 & "'; Len(value2)=" & Len(value2) & "; ISNULL(dbvalue)=" & ISNull(rs.Fields("DecimalResponse_" & Months(1))) & vbCrLf)
				Response.Write("QuestionID=" & QuestionID & "; Response_" & Months(2) & "_" & QuestionID & "=" & _
					Request.Form("Response_" & Months(2) & "_" & QuestionID) & "; dbvalue=" & rs.Fields("DecimalResponse_" & Months(2)) & _
					"; value2='" & value3 & "'; Len(value3)=" & Len(value3) & "; ISNULL(dbvalue)=" & ISNull(rs.Fields("DecimalResponse_" & Months(2))) & "</pre>")
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
					sql = "INSERT INTO PR.Responses (GrantID, QuestionID, DecimalResponse_" & Months(0) & ", DecimalResponse_" & Months(1) & ", DecimalResponse_" & Months(2) & ", UpdateID, UpdateTimestamp) " & vbCrLf & _
						"VALUES (" & prepIntegerSQL(GrantID) & ", " & _
						prepIntegerSQL(rs.Fields("QuestionID")) & ", " & _
						prepNumberSQL(value1) & ", " & _
						prepNumberSQL(value2) & ", " & _
						prepNumberSQL(value3) & ", " & _
						prepIntegerSQL(UserSystemID) & ", " & _
						prepStringSQL(Timestamp) & ")"
				Else
					' Do an update.
					sql = "UPDATE PR.Responses SET " & vbCRLF & _
						"DecimalResponse_" & Months(0) & "=" & prepNumberSQL(value1) & ", " & _
						"DecimalResponse_" & Months(1) & "=" & prepNumberSQL(value2) & ", " & _
						"DecimalResponse_" & Months(2) & "=" & prepNumberSQL(value3) & ", " & _
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
		ElseIf rs.Fields("ResponseTypeID")=5 Then
			text = Request.Form("Response_Q" & Quarter & "_" & QuestionID)
			dbtext = rs.Fields("TextResponse_Q" & Quarter)
			If Debug = True Then
				Response.Write("<pre>QuestionID=" & QuestionID & "; Response_Q" & Quarter & "_" & QuestionID & "=" & _
					Request.Form("Response_Q" & Quarter & "_" & QuestionID) & "; dbvalue=" & rs.Fields("TextResponse_Q" & Quarter) & _
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
					sql = "INSERT INTO PR.Responses (GrantID, QuestionID, TextResponse_Q" & Quarter & ", UpdateID, UpdateTimestamp) " & vbCrLf & _
						"VALUES (" & prepStringSQL(GrantID) & ", " & _
						prepIntegerSQL(rs.Fields("QuestionID")) & ", " & _
						prepStringSQL(text) & ", " & _
						prepIntegerSQL(UserSystemID) & ", " & _
						prepStringSQL(Timestamp) & ")"
				Else
					' Do an update.
					sql = "UPDATE PR.Responses SET " & vbCRLF & _
						"TextResponse_Q" & Quarter & "=" & prepStringSQL(text) & ", " & _
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
		End If
	'End If
	rs.MoveNext
Wend


sql = "SELECT GrantID, Quarter, ISNULL(Confirmed,0) AS Confirmed " & vbCrLF & _
	"FROM PR.Main " & vbCrLF & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then ' Do an insert
	sql = "INSERT INTO PR.Main (GrantID, Quarter, Confirmed) " & vbCrLf & _
		"VALUES (" & prepIntegerSQL(GrantID) & ", " & prepIntegerSQL(Quarter) & ", " & prepBitSQL(confirmed)& ")" & vbCrLf
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
ElseIf rs.Fields("Confirmed") <> Confirmed Then ' Do an update
	sql = "UPDATE PR.Main " & vbCrLF & _
		"SET Confirmed=" & prepBitSQL(Confirmed) & " " & vbCrLF & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If



If Submit = True Then
	sql = "SELECT GrantID, Quarter, SubmitID, SubmitTimestamp " & vbCrLF & _
		"FROM PR.Main " & vbCrLF & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then ' Do an insert
		sql = "INSERT INTO PR.Main (GrantID, Quarter, SubmitID, SubmitTimestamp) " & vbCrLf & _
			"VALUES (" & prepIntegerSQL(GrantID) & ", " & prepIntegerSQL(Quarter) & ", " & prepIntegerSQL(UserSystemID) & ", " & prepStringSQL(Timestamp) & ")" & vbCrLf
	Else ' Do an update
		sql = "UPDATE PR.Main " & vbCrLF & _
			"SET SubmitID=" & prepIntegerSQL(UserSystemID) & ", SubmitTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLF & _
			"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
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

	sql = "SELECT GrantID, Quarter, SubmitID, AdministrativeComments, ApprovalID, ApprovalDate " & vbCrLF & _
		"FROM PR.Main " & vbCrLF & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
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
		sql = "UPDATE PR.Main " & vbCrLF & _
			"SET Confirmed=null, SubmitID=null, SubmitTimestamp=null, AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ",ApprovalID=null, ApprovalDate= null " & vbCrLF & _
			"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	ElseIf AdministrativeUpdate = True Then
		If rs.EOF = True Then ' Do an insert
			sql = "INSERT INTO PR.Main (GrantID, Quarter, AdministrativeComments, ApprovalID, ApprovalDate) " & vbCrLf & _
				"VALUES (" & prepIntegerSQL(GrantID) & ", " & prepIntegerSQL(Quarter) & ", " & prepIntegerSQL(ApprovalID) & ", " & prepStringSQL(AdministrativeComments) & ", " & prepStringSQL(ApprovalDate) & ")"
		Else ' Do an update
			sql = "UPDATE PR.Main " & vbCrLF & _
				"SET AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ", ApprovalID=" & prepIntegerSQL(ApprovalID) & ", ApprovalDate=" & prepStringSQL(ApprovalDate) & " " & vbCrLF & _
				"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
		End If
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
End If

If Debug = True Then
	Response.Write("<a href=""Report.asp?GrantID=" & GrantID & "&Quarter=" & Quarter & "&ShowOneQuarter=" & prepBitRequiredSQL(ShowOneQuarter) & """>return to progress report</a><br />" & vbCrLf)
	Response.Write("<a href=""../Home/Default.asp?GrantID=" & GrantID & """>return to Home</a><br />" & vbCrLf)
	Response.End
Else
	Response.Redirect("Report.asp?GrantID=" & GrantID & "&Quarter=" & Quarter & "&ShowOneQuarter=" & prepBitRequiredSQL(ShowOneQuarter))
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