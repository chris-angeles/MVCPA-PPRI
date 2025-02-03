<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, AppID, NewAppID, FiscalYear, Changes, PRVersion, _
	StaffShortProgramDescription, ResolutionConfirmedDate, _
	ResolutionFundsProvided, ResolutionReturnFunds, ResolutionDesignateOfficals, _
	ResolutionGoverningBody, ResolutionDelegationSupported, ApplicationCertifiedCompleteDate, _
	ApplicationConsideredDate, GrantResultID, GrantAwardAmount, GrantNumber, POIssueDate, _
	InterlocalAgreementConfirmedDate, InterlocalAgreementConfirmedBy, _
	ProsecutorAgreementConfirmedDate, ProsecutorAgreementConfirmedID, _
	OperationalPlanApprovalDate, OperationalPlanApprovalID, _
	InitialAwardTransmissionDate, CreateNegotiationRecords, AwardAcceptanceDate, RevisionsAcceptedDate, _
	OfficialGrantAwardLetterDate, AwardLetterTransmissionMethodID, _
	NegotiationLocked, SignedGrantAwardLetterDate, _
	AwardAcceptanceSignatureConfirmedDate, GrantAwardCertifiedComplete, _
	GrantAwardDeclineLetterReceived, Notes, ExcludeFromConsideration, ConsiderationNotes, Timestamp, _
	ClearSubmit, RevisedClearSubmit, CreateGrantRecord, CreateStubRecord, records, ApplicationSchema
debug = False
Timestamp = now()

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
End If

AppID = Request.Form("AppID")
NewAppID = Request.Form("NewAppID")
If Len(NewAppID)=0 Then
	NewAppID = AppID
End If
FiscalYear = CInt(Request.Form("FiscalYear"))
Changes = Request.Form("Changes")

ApplicationSchema = getCCApplicationSchema(FiscalYear)

' Update annualy if new progress report/target versions.
If FiscalYear=2024 Then
	PRVersion = 1001
ElseIf FiscalYear = 2025 Then
	PRVersion = 1001
Else
	PRVersion = 1001
End If

If Request.Form("ClearSubmit") = "1" Then
	ClearSubmit = True
	ApplicationCertifiedCompleteDate = null
Else
	ClearSubmit = False 
	ApplicationCertifiedCompleteDate = Request.Form("ApplicationCertifiedCompleteDate")
End If

If Request.Form("RevisedClearSubmit") = "1" Then
	RevisedClearSubmit = True
	RevisionsAcceptedDate = null
Else
	RevisedClearSubmit = False 
	RevisionsAcceptedDate = Request.Form("RevisionsAcceptedDate")
End If

If Len(Changes) > 2 Or ClearSubmit = True Or RevisedClearSubmit = True Then
	StaffShortProgramDescription= Request.Form("StaffShortProgramDescription")
	ResolutionConfirmedDate = Request.Form("ResolutionConfirmedDate")
	ResolutionFundsProvided = Request.Form("ResolutionFundsProvided")
	ResolutionReturnFunds = Request.Form("ResolutionReturnFunds")
	ResolutionDesignateOfficals = Request.Form("ResolutionDesignateOfficals")
	ResolutionGoverningBody = Request.Form("ResolutionGoverningBody")
	ResolutionDelegationSupported = Request.Form("ResolutionDelegationSupported")
	ApplicationConsideredDate = Request.Form("ApplicationConsideredDate")
	GrantResultID = Request.Form("GrantResultID")
	GrantAwardAmount = Request.Form("GrantAwardAmount")
	GrantNumber = Request.Form("GrantNumber")
	POIssueDate = Request.Form("POIssueDate")
	InterlocalAgreementConfirmedDate = Request.Form("InterlocalAgreementConfirmedDate")
	InterlocalAgreementConfirmedBy = Request.Form("InterlocalAgreementConfirmedBy")
	If InterlocalAgreementConfirmedDate = "" Then InterlocalAgreementConfirmedBy = null
	ProsecutorAgreementConfirmedDate = Request.Form("ProsecutorAgreementConfirmedDate")
	ProsecutorAgreementConfirmedID = Request.Form("ProsecutorAgreementConfirmedID")
	If ProsecutorAgreementConfirmedDate = "" Then ProsecutorAgreementConfirmedID = null
	OperationalPlanApprovalDate = Request.Form("OperationalPlanApprovalDate")
	OperationalPlanApprovalID = Request.Form("OperationalPlanApprovalID")
	If OperationalPlanApprovalDate = "" Then OperationalPlanApprovalID = null
	InitialAwardTransmissionDate = Request.Form("InitialAwardTransmissionDate")
	If Request.Form("CreateNegotiationRecords") = "1" Then
		CreateNegotiationRecords = True
	Else
		CreateNegotiationRecords = False
	End If
	AwardAcceptanceDate = Request.Form("AwardAcceptanceDate")
	OfficialGrantAwardLetterDate = Request.Form("OfficialGrantAwardLetterDate")
	AwardLetterTransmissionMethodID = Request.Form("AwardLetterTransmissionMethodID")
	If Request.Form("NegotiationLocked") = "1" Then
		NegotiationLocked = True
	Else
		NegotiationLocked = False
	End If
	SignedGrantAwardLetterDate = Request.Form("SignedGrantAwardLetterDate")
	AwardAcceptanceSignatureConfirmedDate = Request.Form("AwardAcceptanceSignatureConfirmedDate")
	GrantAwardCertifiedComplete = Request.Form("GrantAwardCertifiedComplete")
	Notes = Request.Form("Notes")
	ExcludeFromConsideration = Request.Form("ExcludeFromConsideration")
	ConsiderationNotes = Request.Form("ConsiderationNotes")
	If Request.Form("CreateGrantRecord") = "1" Then
		CreateGrantRecord = True
	Else
		CreateGrantRecord = False
	End If
	If Request.Form("CreateStubRecord") = "1" Then
		CreateStubRecord = True
	Else
		CreateStubRecord = False
	End If
	
	sql = "SELECT * FROM CC.Admin WHERE AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		If IsNull(rs.Fields("InterlocalAgreementConfirmedDate")) = True And Len(InterlocalAgreementConfirmedDate)>0 Then
			InterlocalAgreementConfirmedBy = UserSystemID
		End If
		If IsNull(rs.Fields("ProsecutorAgreementConfirmedDate")) = True And Len(ProsecutorAgreementConfirmedDate)>0 Then
			ProsecutorAgreementConfirmedID = UserSystemID
		End If
		If IsNull(rs.Fields("OperationalPlanApprovalDate")) = True And Len(OperationalPlanApprovalDate)>0 Then
			OperationalPlanApprovalID = UserSystemID
		End If
	End If
	If rs.EOF = True Then 
		'Insert
		sql = "INSERT INTO CC.Admin (AppID, StaffShortProgramDescription, ResolutionConfirmedDate, ResolutionFundsProvided, ResolutionReturnFunds, ResolutionDesignateOfficals, ResolutionGoverningBody, ResolutionDelegationSupported, ApplicationCertifiedCompleteDate, ApplicationConsideredDate, GrantResultID, GrantAwardAmount, GrantNumber, POIssueDate, InterlocalAgreementConfirmedDate, InterlocalAgreementConfirmedBy, ProsecutorAgreementConfirmedDate, ProsecutorAgreementConfirmedID, OperationalPlanApprovalDate, OperationalPlanApprovalID, InitialAwardTransmissionDate, AwardAcceptanceDate, RevisionsAcceptedDate, OfficialGrantAwardLetterDate, AwardLetterTransmissionMethodID, NegotiationLocked, SignedGrantAwardLetterDate, AwardAcceptanceSignatureConfirmedDate, GrantAwardCertifiedComplete, GrantAwardDeclineLetterReceived, Notes, ExcludeFromConsideration, ConsiderationNotes, UpdateID, UpdateTimestamp) " & vbCrLf & _
			"VALUES (" & prepIntegerSQL(AppID) & ", " & _
					prepStringSQL(StaffShortProgramDescription) & ", " & vbCrLf & _
			prepDateSQL(ResolutionConfirmedDate) & ", " & vbCrLf & _
			prepBitSQL(ResolutionFundsProvided) & ", " & vbCrLf & _
			prepBitSQL(ResolutionReturnFunds) & ", " & vbCrLf & _
			prepBitSQL(ResolutionDesignateOfficals) & ", " & vbCrLf & _
			prepBitSQL(ResolutionGoverningBody) & ", " & vbCrLf & _
			prepBitSQL(ResolutionDelegationSupported) & ", " & vbCrLf & _
			prepDateSQL(ApplicationCertifiedCompleteDate) & ", " & vbCrLf & _
			prepDateSQL(ApplicationConsideredDate) & ", " & vbCrLf & _
			prepIntegerSQL(GrantResultID) & ", " & vbCrLf & _
			prepNumberSQL(GrantAwardAmount) & ", " & vbCrLf & _
			prepStringSQL(GrantNumber) & ", " & vbCrLf & _
			prepDateSQL(POIssueDate) & ", " & vbCrLf & _
			prepDateSQL(InterlocalAgreementConfirmedDate) & ", " & vbCrLf & _
			prepIntegerSQL(InterlocalAgreementConfirmedBy) & ", " & vbCrLf & _
			prepDateSQL(ProsecutorAgreementConfirmedDate) & ", " & vbCrLf & _
			prepIntegerSQL(ProsecutorAgreementConfirmedID) & ", " & vbCrLf & _
			prepDateSQL(OperationalPlanApprovalDate) & ", " & vbCrLf & _
			prepIntegerSQL(OperationalPlanApprovalID) & ", " & vbCrLf & _
			prepDateSQL(InitialAwardTransmissionDate) & ", " & vbCrLf & _
			prepDateSQL(AwardAcceptanceDate) & ", " & vbCrLf & _
			prepDateSQL(RevisionsAcceptedDate) & ", " & vbCrLf & _
			prepDateSQL(OfficialGrantAwardLetterDate) & ", " & vbCrLf & _
			prepIntegerSQL(AwardLetterTransmissionMethodID) & ", " & vbCrLf & _
			prepBitSQL(NegotiationLocked) & ", " & vbCrLf & _
			prepDateSQL(SignedGrantAwardLetterDate) & ", " & vbCrLf & _
			prepDateSQL(AwardAcceptanceSignatureConfirmedDate) & ", " & vbCrLf & _
			prepDateSQL(GrantAwardCertifiedComplete) & ", " & vbCrLf & _
			prepDateSQL(GrantAwardDeclineLetterReceived) & ", " & vbCrLf & _
			prepStringSQL(Notes) & ", " & vbCrLf & _
			prepBitSQL(ExcludeFromConsideration) & ", " & vbCrLf & _
			prepStringSQL(ConsiderationNotes) & ", " & vbCrLf & _
			UserSystemID & ", " & vbCrLf & _
			prepDateSQL(Timestamp) & ") "

		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	Else
		' Update
		sql = "UPDATE [CC].[Admin] " & vbCrLf & _
			"SET [StaffShortProgramDescription] = " & prepStringSQL(StaffShortProgramDescription) & ", " & vbCrLf & _
			"	[ResolutionConfirmedDate] = " & prepDateSQL(ResolutionConfirmedDate) & ", " & vbCrLf & _
			"	[ResolutionFundsProvided] = " & prepBitSQL(ResolutionFundsProvided) & ", " & vbCrLf & _
			"	[ResolutionReturnFunds] = " & prepBitSQL(ResolutionReturnFunds) & ", " & vbCrLf & _
			"	[ResolutionDesignateOfficals] = " & prepBitSQL(ResolutionDesignateOfficals) & ", " & vbCrLf & _
			"	[ResolutionGoverningBody] = " & prepBitSQL(ResolutionGoverningBody) & ", " & vbCrLf & _
			"	[ResolutionDelegationSupported] = " & prepBitSQL(ResolutionDelegationSupported) & ", " & vbCrLf & _
			"	[ApplicationCertifiedCompleteDate] = " & prepDateSQL(ApplicationCertifiedCompleteDate) & ", " & vbCrLf & _
			"	[ApplicationConsideredDate] = " & prepDateSQL(ApplicationConsideredDate) & ", " & vbCrLf & _
			"	[GrantResultID] = " & prepIntegerSQL(GrantResultID) & ", " & vbCrLf & _
			"	[GrantAwardAmount] = " & prepNumberSQL(GrantAwardAmount) & ", " & vbCrLf & _
			"	[GrantNumber] = " & prepStringSQL(GrantNumber) & ", " & vbCrLf & _
			"	[POIssueDate] = " & prepDateSQL(POIssueDate) & ", " & vbCrLf & _
			"	[InterlocalAgreementConfirmedDate] = " & prepDateSQL(InterlocalAgreementConfirmedDate) & ", " & vbCrLf & _
			"	[InterlocalAgreementConfirmedBy] = " & prepIntegerSQL(InterlocalAgreementConfirmedBy) & ", " & vbCrLf & _
			"	[ProsecutorAgreementConfirmedDate] = " & prepDateSQL(ProsecutorAgreementConfirmedDate) & ", " & vbCrLf & _
			"	[ProsecutorAgreementConfirmedID] = " & prepIntegerSQL(ProsecutorAgreementConfirmedID) & ", " & vbCrLf & _
			"	[OperationalPlanApprovalDate] = " & prepDateSQL(OperationalPlanApprovalDate) & ", " & vbCrLf & _
			"	[OperationalPlanApprovalID] =  " & prepIntegerSQL(OperationalPlanApprovalID) & ", " & vbCrLf & _
			"	[InitialAwardTransmissionDate] = " & prepDateSQL(InitialAwardTransmissionDate) & ", " & vbCrLf & _
			"	[AwardAcceptanceDate] = " & prepDateSQL(AwardAcceptanceDate) & ", " & vbCrLf & _
			"	[RevisionsAcceptedDate] = " & prepDateSQL(RevisionsAcceptedDate) & ", " & vbCrLf & _
			"	[OfficialGrantAwardLetterDate] = " & prepDateSQL(OfficialGrantAwardLetterDate) & ", " & vbCrLf & _
			"	[AwardLetterTransmissionMethodID] = " & prepIntegerSQL(AwardLetterTransmissionMethodID) & ", " & vbCrLf & _
			"	[NegotiationLocked] = " & prepBitSQL(NegotiationLocked) & ", " & vbCrLf & _
			"	[SignedGrantAwardLetterDate] = " & prepDateSQL(SignedGrantAwardLetterDate) & ", " & vbCrLf & _
			"	[AwardAcceptanceSignatureConfirmedDate] = " & prepDateSQL(AwardAcceptanceSignatureConfirmedDate) & ", " & vbCrLf & _
			"	[GrantAwardCertifiedComplete] = " & prepDateSQL(GrantAwardCertifiedComplete) & ", " & vbCrLf & _
			"	[GrantAwardDeclineLetterReceived] = " & prepDateSQL(GrantAwardDeclineLetterReceived) & ", " & vbCrLf & _
			"	[Notes] = " & prepStringSQL(Notes) & ", " & vbCrLf & _
			"	[ExcludeFromConsideration] = " & prepBitSQL(ExcludeFromConsideration) & ", " & vbCrLf & _
			"	[ConsiderationNotes] = " & prepStringSQL(ConsiderationNotes) & ", " & vbCrLf & _
			"	[UpdateID] = " & UserSystemID & ", " & vbCrLf & _
			"	[UpdateTimestamp] = " & prepDateSQL(Timestamp) & " " & vbCrLf & _
			"WHERE AppID=" & prepIntegerSQL(AppID)

		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
	End If
Else
	If Debug = True Then
		Response.Write("<pre>There were no changes to make.</pre>")
	End If
End If

If CreateNegotiationRecords = True Then
	If Debug = True Then
		Response.Write("<pre>Create Negotiation Records.</pre>")
	End If
	sql = "EXECUTE dbo.spCopyApplicationToNegotiation @AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
End If

If CreateGrantRecord = True Or CreateStubRecord = True Then
	sql = "INSERT INTO [Grants].Main (GrantClassID, FiscalYear, GranteeID, GrantNumber, POIssueDate, ProgramName, " & vbCrLf & _
		"	AwardAmount, MatchAmount, InLieuOfDPSBudget, InLieuOfNICBBudget, ProgramIncomeBudget, " & vbCrLf & _
		"	ReimbursementRate, AppID, UpdateID, UpdateTimestamp) " & vbCrLf & _
		"SELECT 4 AS GrantClassID, A.FiscalYear, A.GranteeID, B.GrantNumber, A.POIssueDate, A.ProgramName, " & vbCrLf & _
		"	B.GrantAwardAmount AS AwardAmount, " & vbCrLf & _
		"	MatchAmount = M.CashMatch, " & vbCrLf & _
		"	D.InLieuOfDPSBudget, D.InLieuOfNICBBudget, D.ProgramIncomeBudget, " & vbCrLf & _
		"	ReimbursementRate = 100.0 * B.GrantAwardAmount / (ISNULL(B.GrantAwardAmount,0.0) + ISNULL(M.CashMatch,0.0) - ISNULL(D.InLieuOfDPSBudget,0.0) - ISNULL(D.InLieuOfNICBBudget,0.0)), " & vbCrLf & _
		"	A.AppID, " & UserSystemID & " AS UpdateID, '" & timestamp & "' AS UpdateTimestamp " & vbCrLf & _
		"FROM Application.IDs AS I " & vbCrLf & _
		"LEFT JOIN [CC].[" & ApplicationSchema & "] AS A ON A.AppID=I.AppID " & vbCrLf & _
		"LEFT JOIN Application.Admin AS B ON B.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN [Grants].Main AS C ON C.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN (SELECT AppID, " & vbCrLf & _
		"	SUM(CASE WHEN MatchSourceID=3 AND MatchTypeID=1 THEN Amount ELSE NULL END) AS InLieuOfDPSBudget, " & vbCrLf & _
		"	SUM(CASE WHEN MatchSourceID=4 AND MatchTypeID=1 THEN Amount ELSE NULL END) AS InLieuOfNICBBudget, " & vbCrLf & _
		"	SUM(CASE WHEN MatchSourceID=5 AND MAtchTypeID=1 THEN Amount ELSE NULL END) AS ProgramIncomeBudget " & vbCrLf & _
		"	FROM " & ApplicationSchema & ".Matches  " & vbCrLf & _
		"	WHERE MatchTypeID=1 " & vbCrLf & _
		"	GROUP BY AppID) AS D ON D.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN (SELECT AppID, SUM(CashMatch) AS CashMatch FROM " & ApplicationSchema & ".BudgetDetails GROUP BY AppID) AS M ON M.AppID=A.AppID " & vbCrLf & _
		"WHERE A.AppID=" & prepIntegerSQL(AppID) & " AND C.AppID IS NULL "


	sql = "MERGE [Grants].Main AS [Target] " & vbCrLf & _
		"USING ( " & vbCrLf & _
		"	SELECT I.GrantClassID, I.FiscalYear, I.GranteeID, B.GrantNumber, B.POIssueDate, A.ProgramName, B.GrantAwardAmount AS AwardAmount, " & vbCrLf & _
		"		MatchAmount = M.CashMatch, " & vbCrLf & _
		"		D.InLieuOfDPSBudget, D.InLieuOfNICBBudget, D.ProgramIncomeBudget, " & vbCrLf & _
		"		ReimbursementRate = 100.0 * B.GrantAwardAmount / (ISNULL(B.GrantAwardAmount,0.0) + ISNULL(M.CashMatch,0.0) - ISNULL(D.InLieuOfDPSBudget,0.0) - ISNULL(D.InLieuOfNICBBudget,0.0)), " & vbCrLf & _
		"		I.AppID, PY.GrantID AS PreviousYEarGrantID, " & vbCrLf & _
		"		" & UserSystemID & " AS UpdateID, '" & timestamp & "' AS UpdateTimestamp " & vbCrLf & _
		"	FROM [Application].[IDs] AS I " & vbCrLf & _
		"	LEFT JOIN [CC].[" & ApplicationSchema & "] AS A ON A.AppID=I.AppID " & vbCrLf & _
		"	LEFT JOIN CC.Admin AS B ON B.AppID=I.AppID " & vbCrLf & _
		"	LEFT JOIN [Grants].Main AS C ON C.AppID=I.AppID " & vbCrLf & _
		"	LEFT JOIN (SELECT AppID, " & vbCrLf & _
		"		SUM(CASE WHEN MatchSourceID=3 AND MatchTypeID=1 THEN Amount ELSE NULL END) AS InLieuOfDPSBudget, " & vbCrLf & _
		"		SUM(CASE WHEN MatchSourceID=4 AND MatchTypeID=1 THEN Amount ELSE NULL END) AS InLieuOfNICBBudget, " & vbCrLf & _
		"		SUM(CASE WHEN MatchSourceID=5 AND MatchTypeID=1 THEN Amount ELSE NULL END) AS ProgramIncomeBudget " & vbCrLf & _
		"		FROM " & ApplicationSchema & ".Matches  " & vbCrLf & _
		"		GROUP BY AppID) AS D ON D.AppID=I.AppID " & vbCrLf & _
		"	LEFT JOIN (SELECT AppID, SUM(CashMatch) AS CashMatch FROM " & ApplicationSchema & ".BudgetDetails GROUP BY AppID) AS M ON M.AppID=I.AppID " & vbCrLf & _
		"	LEFT JOIN [Grants].Main AS PY ON PY.GranteeID=I.GranteeID AND PY.FiscalYear=I.FiscalYear-1 AND PY.GrantClassID=I.GrantClassID " & vbCrLf & _
		"	WHERE I.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
		") AS [Source] " & vbCrLf & _
		"ON [Target].AppID=[Source].AppID " & vbCrLf & _
		"WHEN MATCHED THEN UPDATE SET GrantNumber=[Source].GrantNumber, " & vbCrLf & _
		"	POIssueDate = [Source].POIssueDate, " & vbCrLf & _
		"	ProgramName = [Source].ProgramName, " & vbCrLf & _
		"	AwardAmount = [Source].AwardAmount, " & vbCrLf & _
		"	MatchAmount = [Source].MatchAmount, " & vbCrLf & _
		"	InLieuOfDPSBudget = [Source].InLieuOfDPSBudget, " & vbCrLf & _
		"	InLieuOfNICBBudget = [Source].InLieuOfNICBBudget, " & vbCrLf & _
		"	ProgramIncomeBudget = [Source].ProgramIncomeBudget, " & vbCrLf & _
		"	ReimbursementRate = [Source].ReimbursementRate, " & vbCrLf & _
		"	PreviousYearGrantID = [Source].PreviousYearGrantID, " & vbCrLf & _
		"	UpdateID = [Source].UpdateID, " & vbCrLf & _
		"	UpdateTimestamp = [Source].UpdateTimestamp " & vbCrLf & _
		"WHEN NOT MATCHED THEN INSERT (GrantClassID, FiscalYear, GranteeID, GrantNumber, POIssueDate, ProgramName, " & vbCrLf & _
		"	AwardAmount, MatchAmount, InLieuOfDPSBudget, InLieuOfNICBBudget, ProgramIncomeBudget, " & vbCrLf & _
		"	ReimbursementRate, AppID, PreviousYearGrantID, UpdateID, UpdateTimestamp) " & vbCrLf & _
		"	VALUES ([Source].GrantClassID, [Source].FiscalYear, [Source].GranteeID, [Source].GrantNumber, [Source].POIssueDate, " & vbCrLF & _
		"	[Source].ProgramName, [Source].AwardAmount, [Source].MatchAmount, " & vbCrLf & _
		"	[Source].InLieuOfDPSBudget, [Source].InLieuOfNICBBudget, [Source].ProgramIncomeBudget, " & vbCrLf & _
		"	[Source].ReimbursementRate, [Source].AppID, [Source].PreviousYearGrantID, " & vbCrLf & _
		"	[Source].UpdateID, [Source].UpdateTimestamp);"
	If Debug = True Then
		Response.Write("<pre>Create Grant Record." & vbCrLf & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql, records)
	If Debug = True Then
		Response.Write("<pre>" & records & " records affected in grant creation." & "</pre>")
	End If

End If
If CreateGrantRecord = True Or CreateStubRecord = True Then

	' Copy Budget to Grant
	sql = "INSERT INTO [Grants].Budget (GrantID, BudgetCategoryID, TotalExpenditures, MVCPAExpenditures, MatchExpenditures, InKindExpenditures, UpdateID, UpdateTimestamp) " & vbCrLf & _
	"SELECT C.GrantID, A.BudgetCategoryID, " & vbCrLf & _
	"	SUM(LineTotal) AS TotalExpenditures, " & vbCrLf & _
	"	SUM(A.MVCPAFunds) AS MVCPAExpenditures, " & vbCrLf & _
	"	SUM(CashMatch) AS MatchExpenditures, " & vbCrLf & _
	"	SUM(InKindMatch) AS InKindExpenditures, " & vbCrLf & _
	"	2147483647 AS UpdateID, " & vbCrLf & _
	"	getdate() AS UpdateTimestamp " & vbCrLf & _
	"FROM " & ApplicationSchema & ".BudgetDetails AS A " & vbCrLf & _
	"JOIN [CC].Admin AS B ON B.AppID=A.AppID " & vbCrLf & _
	"JOIN [Grants].Main AS C ON C.AppID=B.AppID " & vbCrLf & _
	"WHERE B.AppID=" & prepIntegerSQL(AppID) & " " & vbCrLf & _
	"GROUP BY C.GrantID, A.BudgetCategoryID " & vbCrLf & _
	"ORDER BY C.GrantID, A.BudgetCategoryID"
	If Debug = True Then
		Response.Write("<pre>Create Grant Record." & vbCrLf & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql, records)
	If Debug = True Then
		Response.Write("<pre>" & records & " records affected in grant creation." & "</pre>")
	End If

	' Copy Particpating Agencies
	sql = "INSERT INTO [Grants].ParticipatingAgencies " & vbCrLf & _
		"SELECT A.GrantID, B.ORI, B.UpdateID, B.UpdateTimestamp " & vbCrLf & _
		"FROM [Grants].Main AS A " & vbCrLf & _
		"JOIN " & ApplicationSchema & ".ParticipatingAgencies AS B ON B.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN [Grants].ParticipatingAgencies AS C ON C.GrantID=A.GrantID AND C.ORI=B.ORI " & vbCrLf & _
		"WHERE A.AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>Create Grant Record." & vbCrLf & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql, records)
	If Debug = True Then
		Response.Write("<pre>" & records & " records affected in grant creation." & "</pre>")
	End If
	
	' Copy Coverage Agencies
	sql = "INSERT INTO [Grants].CoverageAgencies " & vbCrLf & _
		"SELECT A.GrantID, B.ORI, B.UpdateID, B.UpdateTimestamp " & vbCrLf & _
		"FROM [Grants].Main AS A " & vbCrLf & _
		"JOIN " & ApplicationSchema & ".CoverageAgencies AS B ON B.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN [Grants].ParticipatingAgencies AS C ON C.GrantID=A.GrantID AND C.ORI=B.ORI " & vbCrLf & _
		"WHERE A.AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>Create Grant Record." & vbCrLf & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql, records)
	If Debug = True Then
		Response.Write("<pre>" & records & " records affected in grant creation." & "</pre>")
	End If

	' Copy GSATargets for progress report questions.
	' *******************************************************
	' This needs to be updated to include verions and re-enabled. Currently grant questions done with manual query--Add Progress Report Questions.sql
	' moved to last so, even if enabled, it doesn't stop other parts.
	' *******************************************************

	sql = "INSERT INTO PR.GrantQuestions (GrantID, QuestionID, Version, IntegerTarget, DecimalTarget) " & vbCrLf & _
		"SELECT A.GrantID, B.QuestionID, B.Version, E.IntegerResponse AS IntegerTarget, E.DecimalResponse AS DecimalTarget " & vbCrLf & _
		"FROM [Grants].Main AS A " & vbCrLf & _
		"CROSS JOIN (SELECT * FROM PR.Activities WHERE Version=" & PRVersion & ") AS B " & vbCrLf & _
		"LEFT JOIN PR.GrantQuestions AS C ON C.QuestionID=B.QuestionID AND C.GrantID = A.GrantID " & vbCrLf & _
		"LEFT JOIN Application.Admin AS D ON D.AppID=A.AppID " & vbCrLf & _
		"LEFT JOIN Negotiation.GSATargets AS E ON E.AppID=D.AppID AND E.GoalID=B.GoalID AND E.StrategyID=B.StrategyID AND E.ActivityID=B.ActivityID AND E.MeasureID=B.MeasureID AND E.Version=B.Version " & vbCrLf & _
		"WHERE A.AppID=" & prepIntegerSQL(AppID) & " AND C.QuestionID IS NULL " & vbCrLf & _
		"ORDER BY GrantID, QuestionID "

	If Debug = True Then
		Response.Write("<pre>Add Progress Report Questions." & vbCrLf & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql, records)
	If Debug = True Then
		Response.Write("<pre>" & records & " records affected in grant creation." & "</pre>")
	End If

	sql = "UPDATE A " & vbCrLf & _
	"	SET  " & vbCrLf & _
	"	IntegerTarget = CASE " & vbCrLf & _
	"		WHEN EN.IntegerResponse IS NOT NULL THEN EN.IntegerResponse " & vbCrLf & _
	"		WHEN EA.IntegerResponse IS NOT NULL THEN EA.IntegerResponse " & vbCrLf & _
	"		ELSE NULL END, " & vbCrLf & _
	"	DecimalTarget = CASE  " & vbCrLf & _
	"		WHEN EN.DecimalResponse IS NOT NULL THEN EN.DecimalResponse " & vbCrLf & _
	"		WHEN EA.DecimalResponse IS NOT NULL THEN EA.DecimalResponse " & vbCrLf & _
	"		ELSE NULL END " & vbCrLf & _
	"FROM PR.GrantQuestions AS A " & vbCrLf & _
	"LEFT JOIN Grants.Main AS B ON B.GrantID=A.GrantID " & vbCrLf & _
	"LEFT JOIN PR.Activities AS C ON C.QuestionID=A.QuestionID " & vbCrLf & _
	"LEFT JOIN Lookup.Activities AS D ON D.QuestionID=C.QuestionID " & vbCrLf & _
	"LEFT JOIN Application.GSATargets AS EA ON EA.AppID=B.AppID AND EA.Version=D.Version AND EA.GoalID=D.GoalID AND EA.StrategyID=D.StrategyID AND EA.ActivityID=D.ActivityID AND EA.MeasureID=D.MeasureID " & vbCrLf & _
	"LEFT JOIN Negotiation.GSATargets AS EN ON EN.AppID=B.AppID AND EN.Version=D.Version AND EN.GoalID=D.GoalID AND EN.StrategyID=D.StrategyID AND EN.ActivityID=D.ActivityID AND EN.MeasureID=D.MeasureID " & vbCrLf & _
	"WHERE B.AppID=" & prepIntegerSQL(AppID) & " AND B.GrantClassID=4 "

	If Debug = True Then
		Response.Write("<pre>Add targets to progress report questions." & vbCrLf & sql & "</pre>")
		Response.Flush
	End If
	Set rs = Con.Execute(sql, records)
	If Debug = True Then
		Response.Write("<pre>" & records & " records affected in grant creation." & "</pre>")
	End If


End If

If ClearSubmit = True Then
	sql = "UPDATE CC.Application " & vbCrLf & _
		"SET ConfirmationNumber=null, SubmitID=null, SubmitTimeStamp=null, UpdateID=" & prepIntegerSQL(UserSystemID) & ", UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

If RevisedClearSubmit = True Then
	sql = "UPDATE CC." & ApplicationSchema & " " & vbCrLf & _
		"SET ConfirmationNumber=null, SubmitID=null, SubmitTimeStamp=null, UpdateID=" & prepIntegerSQL(UserSystemID) & ", UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE AppID=" & prepIntegerSQL(AppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

If Debug = True Then
	Response.Write("<a href=""AppAdmin.asp?AppID=" & NewAppID & "&FiscalYear=" & FiscalYear & """>AppAdmin.asp?AppID=" & AppID & "&FiscalYear=" & FiscalYear & "</a>")
Else
	Response.Redirect("AppAdmin.asp?AppID=" & NewAppID & "&FiscalYear=" & FiscalYear)
End If
%>

<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/getApplicationSchema.asp"-->