<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, TimeStamp, Action, _
	GrantID, AdjustmentID, FiscalYear, ProgramChange, BudgetChange, _
	ProgramModificationExplanation, BudgetModificationExplanation, _
	TotalExpenditures, MVCPAExpenditures, MatchExpenditures, _
	TotalExpendituresChange, MatchExpendituresChange, MVCPAExpendituresChange, _
	NewProgramIncome, ProgramIncomeToBeAddedToBudget, Confirmed, SubmitID, SubmitTimestamp, _
	FirstApprovalID, FirstApprovalDate, SecondApprovalID, SecondApprovalDate, _
	DenialDate, AdministrativelyClosedDate, ExternalComments, InternalComments, _
	InLieuOfNICBBudget, ReimbursementRate, CashMatchRate, OvertimeRate, _
	ClearSubmit, Submit, ChangesApplied, ApplyChanges
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

Action = Request.Form("Action")
GrantID = Request.Form("GrantID")
AdjustmentID = Request.Form("AdjustmentID")
FiscalYear = Request.Form("FiscalYear")
ProgramChange = Request.Form("ProgramChange")
BudgetChange = Request.Form("BudgetChange")
ProgramModificationExplanation = Request.Form("ProgramModificationExplanation")
BudgetModificationExplanation = Request.Form("BudgetModificationExplanation")
TotalExpenditures = Request.Form("TotalExpenditures")
MatchExpenditures = Request.Form("MatchExpenditures")
MVCPAExpenditures = Request.Form("MVCPAExpenditures")
TotalExpendituresChange = Request.Form("TotalExpendituresChange")
MatchExpendituresChange = Request.Form("MatchExpendituresChange")
MVCPAExpendituresChange = Request.Form("MVCPAExpendituresChange")
NewProgramIncome = Request.Form("NewProgramIncome")
If Request.Form("Confirmed") = "1" Then
	Confirmed = True 
Else
	Confirmed =  False
End If
ProgramIncomeToBeAddedToBudget = Request.Form("ProgramIncomeToBeAddedToBudget")
If CInt(FiscalYear)>2022 Then
	InLieuOfNICBBudget = Request.Form("InLieuOfNICBBudget")
	ReimbursementRate = Request.Form("ReimbursementRate")
	CashMatchRate = Request.Form("CashMatchRate")
	OvertimeRate = Request.Form("OvertimeRate")
Else
	InLieuOfNICBBudget = NULL
	ReimbursementRate = NULL
	CashMatchRate = NULL
	OvertimeRate = NULL
End If
ExternalComments = Request.Form("ExternalComments")
InternalComments = Request.Form("InternalComments")
FirstApprovalID = Request.Form("FirstApprovalID")
FirstApprovalDate = Request.Form("FirstApprovalDate")
If FirstApprovalDate = "" Then
	FirstApprovalID = ""
End If
SecondApprovalID = Request.Form("SecondApprovalID")
SecondApprovalDate = Request.Form("SecondApprovalDate")
DenialDate = Request.Form("DenialDate")
AdministrativelyClosedDate = Request.Form("AdministrativelyClosedDate")
If Request.Form("ApplyChanges") = "1" Then
	ApplyChanges = True 
Else
	ApplyChanges =  False
End If

If SecondApprovalDate = "" Then
	SecondApprovalID = ""
End If
If Request.Form("ClearSubmit") = "1" Then
	ClearSubmit = True
Else
	ClearSubmit = False
End If
If Action = "Submit" Then
	SubmitID = UserSystemID
	SubmitTimeStamp = Timestamp
	Submit = True
Else
	SubmitID = Request.Form("SubmitID")
	SubmitTimestamp = Request.Form("SubmitTimestamp")
	Submit = False
End If
If ClearSubmit = True Then
	SubmitId = null
	SubmitTimestamp = null
	FirstApprovalID = null
	FirstApprovalDate = null
	SecondApprovalID = null
	SecondApprovalDate = null
	DenialDate = null
	Confirmed = null
End If
If Len(GrantID) > 0 Then
	If IsNumeric(GrantID) Then
		GrantID = CInt(GrantID)
	Else
		Response.Write("Error: Non-numeric GrantID Specified")
		SendMessage "Error: Non-numeric GrantID Specified"
		Response.End
	End If
Else
	Response.Write("Error: No GrantID Specified")
	SendMessage "Error: No GrantID Specified"
	Response.End
End If

If Len(AdjustmentID) > 0 Then
	If IsNumeric(AdjustmentID) Then
		AdjustmentID = CInt(AdjustmentID)
	Else
		Response.Write("Error: Non-numeric AdjustmentID Specified")
		SendMessage "Error: Non-numeric AdjustmentID Specified"
		Response.End
	End If
Else
	Response.Write("Error: No AdjustmentID Specified")
	SendMessage "Error: No AdjustmentID Specified"
	Response.End
End If

If AdjustmentID = 0 Then
	sql = "INSERT INTO [Grants].Adjustments (GrantID, ProgramChange, BudgetChange, BudgetModificationExplanation, ProgramModificationExplanation, NewProgramIncome,  ProgramIncomeToBeAddedToBudget, Confirmed, SubmitID, SubmitTimestamp, InLieuOfNICBBudget, ReimbursementRate, CashMatchRate, OvertimeRate, UpdateID, UpdateTimestamp) " & vbCrLF & _
		"VALUES (" & prepIntegerSQL(GrantID) & ", " & _
		prepBitRequiredSQL(ProgramChange) & ", " & _
		prepBitRequiredSQL(BudgetChange) & ", " & _
		prepStringSQL(cleanMS(BudgetModificationExplanation)) & ", " & _
		prepStringSQL(cleanMS(ProgramModificationExplanation)) & ", " & _
		prepNumberSQL(NewProgramIncome) & ", " & vbCrLf & _
		prepNumberSQL(ProgramIncomeToBeAddedToBudget) & ", " & _
		prepBitSQL(Confirmed) & ", " & _
		prepIntegerSQL(SubmitID) & ", " & _
		prepStringSQL(SubmitTimestamp) & ", " & _
		prepNumberSQL(InLieuOfNICBBudget) & ", " & _ 
		prepNumberSQL(ReimbursementRate) & ", " & _ 
		prepNumberSQL(CashMatchRate) & ", " & _
		prepNumberSQL(OvertimeRate) & ", " & _
		prepIntegerSQL(UserSystemID) & ", " & _
		prepStringSQL(TimeStamp) & ") "

	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)

	If AdjustmentID = 0 Then
		sql = "SELECT SCOPE_IDENTITY() AS AdjustmentID"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set rs = Con.Execute(sql)
		If rs.EOf = False Then
			AdjustmentID = rs.Fields("AdjustmentID")
			If Len(AdjustmentID)>0 Then
				AdjustmentID = CInt(AdjustmentID)
			End If
			If Debug = True Then
				Response.Write("<pre>AdjustmentID=" & AdjustmentID & "</pre>" & vbCrLf)
				Response.Flush
			End If
		Else
			Response.Write("Error: Identity value for AdjustmentID not recovered")
			SendMessage "Error: Identity value for AdjustmentID not recovered"
			Response.End
		End If
	End If
Else
	sql = "UPDATE [Grants].Adjustments SET ProgramChange=" & prepBitRequiredSQL(ProgramChange) & ", " & _
		"BudgetChange=" & prepBitRequiredSQL(BudgetChange) & ", " & _
		"BudgetModificationExplanation=" & prepStringSQL(cleanMS(BudgetModificationExplanation)) & ", " & _
		"ProgramModificationExplanation=" & prepStringSQL(cleanMS(ProgramModificationExplanation)) & ", " & _
		"NewProgramIncome=" & prepNumberSQL(NewProgramIncome) & ", " & _
		"ProgramIncomeToBeAddedToBudget=" & prepNumberSQL(ProgramIncomeToBeAddedToBudget) & ", " & _
		"Confirmed=" & prepBitSQL(Confirmed) & ", " 
	If Submit = True Or ClearSubmit = True Then
		sql = sql & "SubmitID=" & prepIntegerSQL(SubmitID) & ", " & _
		"SubmitTimestamp=" & prepStringSQL(SubmitTimestamp) & ", "
	End If
	sql = sql & "InLieuOfNICBBudget=" & prepNumberSQL(InLieuOfNICBBudget) & ", " & _
		"ReimbursementRate=" & prepNumberSQL(ReimbursementRate) & ", " & _
		"CashMatchRate=" & prepNumberSQL(CashMatchRate) & ", " & _
		"OvertimeRate=" & prepNumberSQL(OvertimeRate) & ", " & _
		"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _		
		"UpdateTimeStamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
		"WHERE AdjustmentID=" & prepIntegerSQL(AdjustmentID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)

End If

If MVCPARights = True Then
	sql = "UPDATE [Grants].Adjustments SET  " & _
		"FirstApprovalID=" & prepIntegerSQL(FirstApprovalID) & ", " & _
		"FirstApprovalDate=" & prepStringSQL(FirstApprovalDate) & ", " & _
		"SecondApprovalID=" & prepIntegerSQL(SecondApprovalID) & ", " & _
		"SecondApprovalDate=" & prepStringSQL(SecondApprovalDate) & ", " & _
		"DenialDate=" & prepStringSQL(DenialDate) & ", " & _
		"AdministrativelyClosedDate=" & prepStringSQL(AdministrativelyClosedDate) & ", " & _
		"ExternalComments=" & prepStringSQL(ExternalComments) & ", " & _
		"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _		
		"UpdateTimeStamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
		"WHERE AdjustmentID=" & prepIntegerSQL(AdjustmentID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)

	' Add Internal Comments if any
	If Debug = True Then
		Response.Write("<pre>Len(InternalComments)=" & Len(InternalComments) & ": " & InternalComments & "</pre>" & vbCrLF) 
	End If
	If Len(InternalComments)>0 Then
		sql = "INSERT INTO [Grants].AdjustmentComments (AdjustmentID, InternalComments, UpdateID, UpdateTimestamp) VALUES (" & _
			prepIntegerSQL(AdjustmentID) & ", " & prepStringSQL(InternalComments) & ", " & prepIntegerSQL(UserSystemID) & ", " & prepStringSQL(TimeStamp) & ")"
		Set rs = Con.Execute(sql)
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
	End If
End If

For i = 1 to 7
	TotalExpendituresChange = Request.Form("TotalExpendituresChange_" & i)
	If Len(TotalExpendituresChange) = 0 Then 
		TotalExpendituresChange = 0
	Else
		TotalExpendituresChange = CDbl(TotalExpendituresChange)
	End If
	MatchExpendituresChange = Request.Form("MatchExpendituresChange_" & i)
	If Len(MatchExpendituresChange) = 0 Then 
		MatchExpendituresChange = 0
	Else
		MatchExpendituresChange = CDbl(MatchExpendituresChange)
	End If
	MVCPAExpendituresChange = Request.Form("MVCPAExpendituresChange_" & i)
	If Len(MVCPAExpendituresChange) = 0 Then 
		MVCPAExpendituresChange = 0
	Else
		MVCPAExpendituresChange = CDbl(MVCPAExpendituresChange)
	End If

	TotalExpenditures = Request.Form("TotalExpenditures_" & i)
	If Len(TotalExpenditures) = 0 Then 
		TotalExpenditures = 0
	Else
		TotalExpenditures = CDbl(TotalExpenditures)
	End If
	MatchExpenditures = Request.Form("MatchExpenditures_" & i)
	If Len(MatchExpenditures) = 0 Then 
		MatchExpenditures = 0
	Else
		MatchExpenditures = CDbl(MatchExpenditures)
	End If
	MVCPAExpenditures = Request.Form("MVCPAExpenditures_" & i)
	If Len(MVCPAExpenditures) = 0 Then 
		MVCPAExpenditures = 0
	Else
		MVCPAExpenditures = CDbl(MVCPAExpenditures)
	End If

	sql = "SELECT AdjustmentID, BudgetCategoryID, " & vbCrLf & _
		"	ISNULL(TotalExpenditures,0) AS TotalExpenditures, " & vbCrLf & _
		"	ISNULL(MVCPAExpenditures,0) AS MVCPAExpenditures, " & vbCrLf & _
		"	ISNULL(MatchExpenditures,0) AS MatchExpenditures, " & vbCrLf & _
		"	ISNULL(TotalExpendituresChange,0) AS TotalExpendituresChange, " & vbCrLf & _
		"	ISNULL(MVCPAExpendituresChange,0) AS MVCPAExpendituresChange, " & vbCrLf & _
		"	ISNULL(MatchExpendituresChange,0) AS MatchExpendituresChange " & vbCrLf & _
		"FROM [Grants].AdjustmentDetails " & vbCrLf & _
		"WHERE AdjustmentID=" & AdjustmentID & " AND BudgetCategoryID=" & i
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)

	If rs.EOF = True Then ' Do an Insert if there are any values
		If TotalExpendituresChange<>0 Or MVCPAExpendituresChange<>0 Or MatchExpendituresChange<>0 Or Submit = True Then
			sql = "INSERT INTO [Grants].AdjustmentDetails (AdjustmentID, BudgetCategoryID, " & vbCrLf & _
				"	TotalExpenditures, MVCPAExpenditures, MatchExpenditures, " & vbCrLf & _
				"	TotalExpendituresChange, MVCPAExpendituresChange, MatchExpendituresChange, UpdateID, UpdateTimestamp)" & vbCrLF & _
				"VALUES (" & AdjustmentID & ", " & i & _
				", " & prepNumberSQL(TotalExpenditures) & _
				", " & prepNumberSQL(MVCPAExpenditures) & _
				", " & prepNumberSQL(MatchExpenditures) & _
				", " & prepNumberSQL(TotalExpendituresChange) & _
				", " & prepNumberSQL(MVCPAExpendituresChange) & _
				", " & prepNumberSQL(MatchExpendituresChange) & _
				", " & UserSystemID & ", " & prepStringSQL(Timestamp) & ")"
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		End If
	Else ' Do an update if necessary
		If Submit = True Or _
			TotalExpenditures<>rs.Fields("TotalExpenditures") Or _
			MVCPAExpenditures<>rs.Fields("MVCPAExpenditures") Or _
			MatchExpenditures<>rs.Fields("MatchExpenditures") Or _
			TotalExpendituresChange<>rs.Fields("TotalExpendituresChange") Or _
			MVCPAExpendituresChange<>rs.Fields("MVCPAExpendituresChange") Or _
			MatchExpendituresChange<>rs.Fields("MatchExpendituresChange") Then
			sql = "UPDATE [Grants].AdjustmentDetails SET " & vbCrLf & _
				"TotalExpenditures=" & prepNumberSQL(TotalExpenditures) & ", " & _
				"MVCPAExpenditures= " & prepNumberSQL(MVCPAExpenditures) & ", " & _
				"MatchExpenditures=" & prepNumberSQL(MatchExpenditures) & ", " & _
				"TotalExpendituresChange=" & prepNumberSQL(TotalExpendituresChange) & ", " & _
				"MVCPAExpendituresChange= " & prepNumberSQL(MVCPAExpendituresChange) & ", " & _
				"MatchExpendituresChange=" & prepNumberSQL(MatchExpendituresChange) & ", " & _
				"UpdateID= " & UserSystemID & ", " & _
				"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
				"WHERE AdjustmentID=" & prepIntegerSQL(AdjustmentID) & _
				" AND BudgetCategoryID=" & i
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		ElseIf Debug = True Then
			Response.Write("<pre>Nothing to update</pre>")
		End If
			
	End If
Next

If Submit = True Then
	Dim From, Recipient, CC, BCC, Subject, Body
	From = "TXMVCPA Website Automated email-Do Not Reply<mvcpa@tamu.edu>"
	CC = "grantsMVCPA@txdmv.gov"
	If Debug = True Then
		Response.Write("<pre>Instance: " & Application("Instance") & vbCrLf)
		Response.Write("AdjustmentID='" & AdjustmentID & "'" & vbCrLf)
		Response.Write("IsNumeric(AdjustmentID)='" & IsNumeric(AdjustmentID) & "'" & vbCrLf)
		Response.Write("VarType(AdjustmentID)='" & VarType(AdjustmentID) & "'</pre>" & vbCrLf)
		Response.Flush
	End If
	sql = "SELECT A.GranteeID, B.GrantID, A.GranteeName, B.ProgramName, C.AdjustmentID, CAST(C.SubmitTimestamp AS DATE) AS SubmitDate, H.Name AS SubmitBy, " & vbCrLf & _
	"	D.email AS AO, " & vbCrLf & _
	"	CASE WHEN E.email NOT IN (D.email) THEN E.Email ELSE NULL END AS PD, " & vbCrLf & _
	"	CASE WHEN F.email NOT IN (D.email, E.email) THEN F.Email ELSE NULL END AS PM, " & vbCrLf & _
	"	CASE WHEN G.email NOT IN (D.email, E.email, F.email) THEN G.email ELSE NULL END AS FO, " & vbCrLf & _
	"	CASE WHEN H.email NOT IN (D.email, E.email, F.email, G.email) THEN H.email ELSE NULL end AS SB " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"JOIN [Grants].Main AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
	"LEFT JOIN [Grants].Adjustments AS C ON C.GrantID=B.GrantID AND C.AdjustmentID=" & prepIntegerSQL(AdjustmentID) & " " & vbCrLf & _
	"LEFT JOIN [System].Users AS D ON D.SystemID=A.AuthorizedOfficialID " & vbCrLf & _
	"LEFT JOIN [System].Users AS E ON E.SystemID=A.ProgramDirectorID " & vbCrLf & _
	"LEFT JOIN [System].Users AS F ON F.SystemID=A.ProgramManagerID " & vbCrLf & _
	"LEFT JOIN [System].Users AS G ON G.SystemID=A.FinancialOfficerID " & vbCrLf & _
	"LEFT JOIN [System].USers AS H ON H.SystemID=C.SubmitID " & vbCrLf & _
	"WHERE B.GrantID=" & prepIntegerSQL(GrantID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		Recipient = rs.Fields("AO")
		If IsNull(rs.Fields("PD")) = False Then
			Recipient = Recipient & "; " & rs.Fields("PD")
		End If
		If IsNull(rs.Fields("PM")) = False Then
			Recipient = Recipient & "; " & rs.Fields("PM")
		End If
		If IsNull(rs.Fields("FO")) = False Then
			Recipient = Recipient & "; " & rs.Fields("FO")
		End If
		If IsNull(rs.Fields("SB")) = False Then
			Recipient = Recipient & "; " & rs.Fields("SB")
		End If
		Subject = "Grant Adjustment Request Submitted"
		Body = "<p>A grant adjustment request for " & rs.Fields("ProgramName") & _
			" was submitted by " & rs.Fields("SubmitBy") & " on " & FormatDateTime(rs.Fields("SubmitDate"), vbShortDate) & ".</p>" & vbCrLf & _
			"<p>You may review the request by logging into https://MVCPA.tamu.edu. " & _
			"If the request is not an authorized grant adjustment request please respond " & _
			"immediately by sending an e-mail to grantsMVCPA@txdmv.gov.</p>" & vbCrLF & _
			"<p>The request will be reviewed and considered by MVCPA staff generally within " & _
			"five (5) business days.</p>" & vbCrLf & _
			"<p>Thank you.</p>" & vbCrLf & _
			"<p>Motor Vehicle Crime Prevention Authority<br />" & vbCrLf & _
			"Txwatchyourcar.com<br />" & vbCrLf & _
			"800-Car Watch</p>"
		If Debug = True Then
		End If
		If Debug = True Then
			Response.Write("<pre>Recipient: " & Recipient & vbCrLf)
			Response.Write("CC: " & CC & vbCrLF)
			Response.Write("From: " & From & vbCrLf)
			Response.Write("Subject: " & subject & vbCrLf)
			Response.Write("body: " & body & "</pre>" & vbCrLf)
		End If
	End If
	If Application("Instance")="Production" Then
		SendHTMLMail From, Recipient, CC, Subject, Body
	End If
End If

If ApplyChanges = True Then
	' Ensure that changes have not been applied!
	sql = "SELECT AdjustmentID, ChangesApplied FROM [Grants].Adjustments WHERE AdjustmentID=" & AdjustmentID
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		ChangesApplied = rs.Fields("ChangesApplied")
	Else
		Response.Write("Error retrieving Changes Applied Status.")
		SendMessage "Error: No GrantID Specified"
		Response.End
	End If
	If ChangesApplied = True Then
		Response.Write("The budget changes have already been applied. Operation cancelled.")
		SendMessage "The budget changes have already been applied. Operation cancelled."
		Response.End
	End If
	sql = "EXEC spApplyBudgetAdjustment @ADjustmentID=" & prepIntegerSQL(AdjustmentID) 
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
End If

If Debug = True Then
	Response.Write("<a href=""Adjustment.asp?GrantID=" & GrantID & "&AdjustmentID=" & AdjustmentID & """>return</a>")
Else
	Response.Redirect("Adjustment.asp?GrantID=" & GrantID & "&AdjustmentID=" & AdjustmentID)
End If
%><!--#include file="../includes/PrepDB.asp"-->
<!--#include file="../includes/Mail.asp"-->
