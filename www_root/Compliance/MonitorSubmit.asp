<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, Timestamp, Changes, ParticipantsChanged, Participants
debug = False
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

Dim MonitorID, GranteeID, FiscalYear, YearsReviewedStart, YearsReviewedEnd, DateOfNotice, _
	InformationOrFilesRequested, RequestedInformationReceivedDate, _
	StartDate, EndDate, ExitInterview, DataCollectionCompleteDate, DraftReportToGranteeDate, _
	GranteeResponseToDraftDueDate, GranteeResponseToDraftReceivedDate, FinalReportCompleteDate, ReportReceivedDate, ManagementLetterReceivedDate, _
	MVCPAFundsTested, MVCPAFundsTestedFinding, MVCPAStaffReviewDate, DeskReview, SiteVisit, _
	MonitoringVisit, CAFR, ExternalAudit, OtherStateAgencyAudit, OtherAudit, OtherAuditDescription, _
	SubgranteeReview, ProgramReview, FiscalReview, SpecialOrTargetReview, _
	SpecialOrTargetReviewText, OtherAgenciesOnVisit, _
	ActionPlanRequired, ActionPlanDueDate, ActionPlanFollowupDate, ActionPlanCompleteDate, _
	RiskLevelAssigned, CompletionClosedDate, UpdateID, UpdateTimestamp, Note

Timestamp = now()

If Len(Request.Form("MonitorID"))>0 Then
	MonitorID = Request.Form("MonitorID")
	If IsNumeric(MonitorID) Then
		MonitorID = CInt(MonitorID)
	Else
		Response.Write("Error: Invalid MonitorID.")
		SendMessage "Error: Invalid MonitorID."
		Response.End
	End If
Else
	Response.Write("Error: No MonitorID provided.")
	SendMessage "Error: No MonitorID provided."
	Response.End
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = Request.Form("FiscalYear")
	If IsNumeric(FiscalYear) Then
		FiscalYear = CInt(FiscalYear)
	Else
		Response.Write("Error: Invalid FiscalYear.")
		SendMessage "Error: Invalid FiscalYear."
		Response.End
	End If
Else
	Response.Write("Error: No FiscalYear provided.")
	SendMessage "Error: No FiscalYear provided."
	Response.End
End If

If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = Request.Form("GranteeID")
	If IsNumeric(GranteeID) Then
		GranteeID = CInt(GranteeID)
	Else
		Response.Write("Error: Invalid GranteeID.")
		SendMessage "Error: Invalid GranteeID."
		Response.End
	End If
Else
	Response.Write("Error: No GranteeID provided.")
	SendMessage "Error: No GranteeID provided."
	Response.End
End If

Changes = Request.Form("Changes")
ParticipantsChanged = Request.Form("ParticipantsChanged")
Participants = Request.Form("Participants")
YearsReviewedStart = Request.Form("YearsReviewedStart")
YearsReviewedEnd = Request.Form("YearsReviewedEnd")
DateOfNotice = Request.Form("DateOfNotice")
InformationOrFilesRequested = Request.Form("InformationOrFilesRequested")
RequestedInformationReceivedDate = Request.Form("RequestedInformationReceivedDate")
GranteeResponseToDraftReceivedDate = Request.Form("GranteeResponseToDraftReceivedDate")
StartDate = Request.Form("StartDate")
ExitInterview = Request.Form("ExitInterview")
EndDate = Request.Form("EndDate")
DataCollectionCompleteDate = Request.Form("DataCollectionCompleteDate")
DraftReportToGranteeDate = Request.Form("DraftReportToGranteeDate")
GranteeResponseToDraftDueDate = Request.Form("GranteeResponseToDraftDueDate")
FinalReportCompleteDate = Request.Form("FinalReportCompleteDate")
ReportReceivedDate = Request.Form("ReportReceivedDate")
ManagementLetterReceivedDate = Request.Form("ManagementLetterReceivedDate")
MVCPAFundsTested = Request.Form("MVCPAFundsTested")
MVCPAFundsTestedFinding = Request.Form("MVCPAFundsTestedFinding")
MVCPAStaffReviewDate = Request.Form("MVCPAStaffReviewDate")
If Request.Form("DeskReview") = "1" Then
	DeskReview = True
Else
	DeskReview = False
End If
If Request.Form("DeskReview") = "1" Then
	DeskReview = True
Else
	DeskReview = False
End If
If Request.Form("SiteVisit") = "1" Then
	SiteVisit = True
Else
	SiteVisit = False
End If
If Request.Form("MonitoringVisit") = "1" Then
	MonitoringVisit = True
Else
	MonitoringVisit = False
End If
If Request.Form("CAFR") = "1" Then
	CAFR = True
Else
	CAFR = False
End If
If Request.Form("ExternalAudit") = "1" Then
	ExternalAudit = True
Else
	ExternalAudit = False
End If
If Request.Form("OtherStateAgencyAudit") = "1" Then
	OtherStateAgencyAudit = True
Else
	OtherStateAgencyAudit = False
End If
If Request.Form("OtherAudit") = "1" Then
	OtherAudit = True
Else
	OtherAudit = False
End If
OtherAuditDescription = Request.Form("OtherAuditDescription")
If Request.Form("SubgranteeReview") = "1" Then
	SubgranteeReview = True
Else
	SubgranteeReview = False
End If
If Request.Form("ProgramReview") = "1" Then
	ProgramReview = True
Else
	ProgramReview = False
End If
If Request.Form("FiscalReview") = "1" Then
	FiscalReview = True
Else
	FiscalReview = False
End If
If Request.Form("SpecialOrTargetReview") = "1" Then
	SpecialOrTargetReview = True
Else
	SpecialOrTargetReview = False
End If
SpecialOrTargetReviewText = Request.Form("SpecialOrTargetReviewText")
OtherAgenciesOnVisit = Request.Form("OtherAgenciesOnVisit")
ActionPlanRequired = Request.Form("ActionPlanRequired")
ActionPlanDueDate = Request.Form("ActionPlanDueDate")
ActionPlanFollowupDate = Request.Form("ActionPlanFollowupDate")
ActionPlanCompleteDate = Request.Form("ActionPlanCompleteDate")
RiskLevelAssigned = Request.Form("RiskLevelAssigned")
Note = Request.Form("Note")
CompletionClosedDate = Request.Form("CompletionClosedDate")

If MonitorID=0 Then
	sql = "INSERT INTO Monitor.Main(GranteeID, FiscalYear, YearsReviewedStart, YearsReviewedEnd, " & vbCrLf & _
		"DateOfNotice, InformationOrFilesRequested, RequestedInformationReceivedDate, " &  vbCrLf & _
		"StartDate, EndDate, ExitInterview, DataCollectionCompleteDate, DraftReportToGranteeDate, " & vbCrLf & _
		"GranteeResponseToDraftDueDate, GranteeResponseToDraftReceivedDate, FinalReportCompleteDate, ReportReceivedDate, ManagementLetterReceivedDate, MVCPAFundsTested, " & vbCrLf & _
		"MVCPAFundsTestedFinding, MVCPAStaffReviewDate, DeskReview, " & vbCrLf & _
		"SiteVisit, MonitoringVisit, CAFR, ExternalAudit, OtherStateAgencyAudit, OtherAudit, OtherAuditDescription, SubgranteeReview, " & vbCrLF & _
		"ProgramReview, FiscalReview, SpecialOrTargetReview, SpecialOrTargetReviewText, OtherAgenciesOnVisit, ActionPlanRequired, ActionPlanDueDate, ActionPlanFollowupDate, ActionPlanCompleteDate, RiskLevelAssigned, CompletionClosedDate, UpdateID, UpdateTimestamp) " & vbCrLf & _
		"VALUES (" & vbCrLf & _
		prepIntegerSQL(GranteeID) & ", " & vbCrLf & _
		prepIntegerSQL(FiscalYear) & ", " & vbCrLf & _
		prepIntegerSQL(YearsReviewedStart) & ", " & vbCrLf & _
		prepIntegerSQL(YearsReviewedEnd) & ", " & vbCrLf & _
		prepDateSQL(DateOfNotice) & ", " & vbCrLf & _
		prepIntegerSQL(InformationOrFilesRequested) & ", " & vbCrLf & _
		prepDateSQL(RequestedInformationReceivedDate) & ", " & vbCrLf & _
		prepDateSQL(StartDate) & ", " & vbCrLf & _
		prepDateSQL(EndDate) & ", " & vbCrLf & _
		prepIntegerSQL(ExitInterview) & ", " & vbCrLf & _
		prepDateSQL(DataCollectionCompleteDate) & ", " & vbCrLf & _
		prepDateSQL(DraftReportToGranteeDate) & ", " & vbCrLf & _
		prepDateSQL(GranteeResponseToDraftDueDate) & ", " & vbCrLf & _
		prepDateSQL(GranteeResponseToDraftReceivedDate) & ", " & vbCrLf & _
		prepDateSQL(FinalReportCompleteDate) & ", " & vbCrLf & _
		prepDateSQL(ReportReceivedDate) & ", " & vbCrLf & _
		prepDateSQL(ManagementLetterReceivedDate) & ", " & vbCrLf & _
		prepIntegerSQL(MVCPAFundsTested) & ", " & vbCrLf & _
		prepStringSQL(MVCPAFundsTestedFinding) & ", " & vbCrLf & _
		prepDateSQL(MVCPAStaffReviewDate) & ", " & vbCrLf & _
		prepBitRequiredSQL(DeskReview) & ", " & vbCrLf & _
		prepBitRequiredSQL(SiteVisit) & ", " & vbCrLf & _
		prepBitRequiredSQL(MonitoringVisit) & ", " & vbCrLf & _
		prepBitRequiredSQL(CAFR) & ", " & vbCrLf & _
		prepBitRequiredSQL(ExternalAudit) & ", " & vbCrLf & _
		prepBitRequiredSQL(OtherStateAgencyAudit) & ", " & vbCrLf & _
		prepBitRequiredSQL(OtherAudit) & ", " & vbCrLf & _
		prepStringSQL(OtherAuditDescription) & ", " & vbCrLf & _
		prepBitRequiredSQL(SubgranteeReview) & ", " & vbCrLf & _
		prepBitRequiredSQL(ProgramReview) & ", " & vbCrLf & _
		prepBitRequiredSQL(FiscalReview) & ", " & vbCrLf & _
		prepBitRequiredSQL(SpecialOrTargetReview) & ", " & vbCrLf & _
		prepStringSQL(SpecialOrTargetReviewText) & ", " & vbCrLf & _
		prepStringSQL(OtherAgenciesOnVisit) & ", " & vbCrLf & _
		prepIntegerSQL(ActionPlanRequired) & ", " & vbCrLf & _
		prepDateSQL(ActionPlanDueDate) & ", " & vbCrLf & _
		prepDateSQL(ActionPlanFollowupDate) & ", " & vbCrLf & _
		prepDateSQL(ActionPlanCompleteDate) & ", " & vbCrLf & _
		prepIntegerSQL(RiskLevelAssigned) & ", " & vbCrLf & _
		prepDateSQL(CompletionClosedDate) & ", " & vbCrLf & _
		prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _
		prepStringSQL(Timestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Con.Execute(sql)
	sql = "SELECT IDENT_CURRENT('Monitor.Main') AS MonitorID"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: No Monitor ID retrieved after insert.")
		sendWarning("Error: No Monitor ID retrieved after insert.")
		Response.End
	Else
		MonitorID = CLng(rs.Fields("MonitorID"))
		If Debug = True Then
			Response.Write("<pre>MoitorID='" & MonitorID & "'</pre>" & vbCrLf)
			Response.Write("<pre>IsNumeric(MonitorID)=" & IsNumeric(MonitorID) & "</pre>" & vbCrLf)
			Response.Flush
		End If
	End If
ElseIf Len(Changes)>0 Then
	sql ="UPDATE Monitor.Main SET " & vbCrLf & _
		"GranteeID=" & prepIntegerSQL(GranteeID) & ", " & vbCrLF & _
		"FiscalYear=" & prepIntegerSQL(FiscalYear) & ", " & vbCrLF & _
		"YearsReviewedStart=" & prepIntegerSQL(YearsReviewedStart) & ", " & vbCrLf & _
		"YearsReviewedEnd=" & prepIntegerSQL(YearsReviewedEnd) & ", " & vbCrLf & _
		"DateOfNotice=" & prepDateSQL(DateOfNotice) & ", " & vbCrLf & _
		"InformationOrFilesRequested=" & prepIntegerSQL(InformationOrFilesRequested) & ", " & vbCrLf & _
		"RequestedInformationReceivedDate=" & prepDateSQL(RequestedInformationReceivedDate) & ", " & vbCrLf & _
		"StartDate=" & prepDateSQL(StartDate) & ", " & vbCrLF & _
		"EndDate=" & prepDateSQL(EndDate) & ", " & vbCrLf & _
		"ExitInterview=" & prepIntegerSQL(ExitInterview) & ", " & vbCrLf & _
		"DataCollectionCompleteDate=" & prepDateSQL(DataCollectionCompleteDate) & ", " & vbCrLF & _
		"DraftReportToGranteeDate=" & prepDateSQL(DraftReportToGranteeDate) & ", " & vbCrLF & _
		"GranteeResponseToDraftDueDate=" & prepDateSQL(GranteeResponseToDraftDueDate) & ", " & vbCrLf & _
		"GranteeResponseToDraftReceivedDate=" & prepDateSQL(GranteeResponseToDraftReceivedDate) & ", " & vbCrLf & _
		"FinalReportCompleteDate=" & prepDateSQL(FinalReportCompleteDate) & ", " & vbCrLF & _
		"ReportReceivedDate=" & prepDateSQL(ReportReceivedDate) & ", " & vbCrLF & _
		"ManagementLetterReceivedDate=" & prepDateSQL(ManagementLetterReceivedDate) & ", " & vbCrLF & _
		"MVCPAFundsTested=" & prepIntegerSQL(MVCPAFundsTested) & ", " & vbCrLF & _
		"MVCPAFundsTestedFinding=" & prepStringSQL(MVCPAFundsTestedFinding) & ", " & vbCrLF & _
		"MVCPAStaffReviewDate=" & prepDateSQL(MVCPAStaffReviewDate) & ", " & vbCrLF & _
		"DeskReview=" & prepBitRequiredSQL(DeskReview) & ", " & vbCrLF & _
		"SiteVisit=" & prepBitRequiredSQL(SiteVisit) & ", " & vbCrLF & _
		"MonitoringVisit=" & prepBitSQL(MonitoringVisit) & ", " & vbCrLf & _
		"CAFR=" & prepBitRequiredSQL(CAFR) & ", " & vbCrLF & _
		"ExternalAudit=" & prepBitRequiredSQL(ExternalAudit) & ", " & vbCrLF & _
		"OtherStateAgencyAudit=" & prepBitRequiredSQL(OtherStateAgencyAudit) & ", " & vbCrLF & _
		"OtherAudit=" & prepBitRequiredSQL(OtherAudit) & ", " & vbCrLF & _
		"OtherAuditDescription=" & prepStringSQL(OtherAuditDescription) & ", " & vbCrLF & _
		"SubgranteeReview=" & prepIntegerSQL(SubgranteeReview) & ", " & vbCrLF & _
		"ProgramReview=" & prepBitRequiredSQL(ProgramReview) & ", " & vbCrLF & _
		"FiscalReview=" & prepBitRequiredSQL(FiscalReview) & ", " & vbCrLF & _
		"SpecialOrTargetReview=" & prepBitRequiredSQL(SpecialOrTargetReview) & ", " & vbCrLF & _
		"SpecialOrTargetReviewText=" & prepStringSQL(SpecialOrTargetReviewText) & ", " & vbCrLF & _
		"OtherAgenciesOnVisit=" & prepStringSQL(OtherAgenciesOnVisit) & ", " & vbCrLF & _
		"ActionPlanRequired=" & prepIntegerSQL(ActionPlanRequired) & ", " & vbCrLF & _
		"ActionPlanDueDate=" & prepDateSQL(ActionPlanDueDate) & ", " & vbCrLf & _
		"ActionPlanFollowupDate=" & prepDateSQL(ActionPlanFollowupDate) & ", " & vbCrLF & _
		"ActionPlanCompleteDate=" & prepDateSQL(ActionPlanCompleteDate) & ", " & vbCrLF & _
		"RiskLevelAssigned=" & prepIntegerSQL(RiskLevelAssigned) & ", " & vbCrLF & _
		"CompletionClosedDate=" & prepDateSQL(CompletionClosedDate) & ", " & vbCrLF & _
		"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLF & _
		"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE MonitorID=" & prepIntegerSQL(MonitorID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Con.Execute(sql)
End If

If Len(Note)>0 Then
	sql = "INSERT INTO Monitor.Notes (MonitorID, Note, UpdateID, UpdateTimestamp) VALUES " & vbCrLF & _
		"(" & prepIntegerSQL(MonitorID) & ", " & prepStringSQL(Note) & ", " & prepIntegerSQL(UserSystemID) & ", " & prepStringSQL(Timestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Con.Execute(sql)
End If
If ParticipantsChanged="1" Then
	sql = "DELETE FROM Monitor.Participants WHERE MonitorID=" & prepIntegerSQL(MonitorID) 
	If Len(Participants) > 0 Then
		sql = sql & " AND SystemID NOT IN (" & Participants & ")"
	End If
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>")
		Response.Flush
	End If
	Con.Execute(sql)

	If Len(Participants)>0 Then
		sql = "INSERT INTO Monitor.Participants (MonitorID, SystemID) " & vbCrLf & _
			"SELECT " & prepIntegerSQL(MonitorID) & " AS MonitorID, A.SystemID " & vbCrLf & _
			"FROM [System].Users AS A " & vbCrLf & _
			"LEFT JOIN Monitor.Participants AS B ON A.SystemID=B.SystemID AND B.MonitorID=" & prepIntegerSQL(MonitorID) & " " & vbCrLf & _
			"WHERE MVCPAStaff=1 AND A.SystemID IN (" & Participants & ") AND B.SystemID IS NULL"
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>")
			Response.Flush
		End If
		Con.Execute(sql)
	End If
End If

If Debug = True Then
	Response.Write("<a href=""Monitor.asp?MonitorID=" & MonitorID & """>Return</a>")
Else
	Response.Redirect("Monitor.asp?MonitorID=" & MonitorID)
End If
%><!--#include file="../includes/prepDB.asp"-->