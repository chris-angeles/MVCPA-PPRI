<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, TimeStamp, GrantID, Quarter, Action, Submit, Unsubmit, _
CashExpenditureTotal, InKindExpenditureTotal, InLieuOfDPS, InLieuOfNICB, UnbudgetedPI, _
ReimbursableExpenditures, ReimbursementRate, ReimbursementYTD, PriorAmountPaid, Reimbursement, _
PriorYearAllocation, AwardAmount, _
InLieuOfDPSBudget, InLieuOfNICBBudget, ProgramIncomeBudget, PriorInLieuOfDPS, PriorInLieuOfNICB, PriorProgramIncome, _
BPOMVCPA, BPOLocal, BPOPI, COVIDMVCPA, COVIDLocal, COVIDPI, COVIDNote, SupplementaryComments, _
BeginningBalance, EarnedThisQuarter, ExpendedThisQuarter, EndingBalance, Confirmed, _
SubmitID, SubmitTimestamp, AmountPaid, PriorYearFunds, CurrentYearFunds, DatePaid, _
AdministrativeComments, ReviewID, ReviewDate, AuditApprovalID, AuditApprovalDate, _
DirectorApprovalID, DirectorApprovalDate, NewDirectorApproval, UpdateID, UpdateTimestamp

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
If Len(Quarter)=0 Then
	Quarter = Request.Form("LoadedQuarter")
End If
If Len(GrantID) = 0 Then
	Response.Write("Error: No GrantID provided.")
	sendWarning("Error: No GrantID provided.")
	Response.End
Else
	GrantID = CInt(GrantID)
End If

If Len(Quarter) = 0 Then
	Response.Write("Error: No Quarter provided.")
	sendWarning("Error: No Quarter provided.")
	Response.End
Else
	Quarter = CInt(Quarter)
End If

Action = Request.Form("Action")
If Action = "submit" Then
	Submit = True
	SubmitID = UserSystemID
	SubmitTimestamp = Timestamp
Else
	Submit = False
End If
CashExpenditureTotal = Request.Form("CashExpenditure_Total")
InKindExpenditureTotal = Request.Form("InKindExpenditure_Total")
InLieuOfDPS = Request.Form("InLieuOfDPS")
InLieuOfNICB = Request.Form("InLieuOfNICB")
UnbudgetedPI =  Request.Form("UnbudgetedPI")
ReimbursableExpenditures = Request.Form("ReimbursableExpenditures")
ReimbursementRate = Request.Form("ReimbursementRate")
ReimbursementYTD = Request.Form("ReimbursementYTD")
PriorAmountPaid = Request.Form("PriorAmountPaid")
Reimbursement = Request.Form("Reimbursement")
PriorYearAllocation = Request.Form("PriorYearAllocation")
AwardAmount = Request.Form("AwardAmount")
BeginningBalance = Request.Form("BeginningBalance")
EarnedThisQuarter = Request.Form("EarnedThisQuarter")
ExpendedThisQuarter = Request.Form("ExpendedThisQuarter")
EndingBalance = Request.Form("EndingBalance")
InLieuOfDPSBudget = Request.Form("InLieuOfDPSBudget")
InLieuOfNICBBudget = Request.Form("InLieuOfNICBBudget")
ProgramIncomeBudget = Request.Form("ProgramIncomeBudget")
PriorInLieuOfDPS = Request.Form("PriorInLieuOfDPS")
PriorInLieuOfNICB = Request.Form("PriorInLieuOfNICB")
PriorProgramIncome = Request.Form("PriorProgramIncome")
BPOMVCPA = Request.Form("BPOMVCPA")
BPOLocal = Request.Form("BPOLocal")
BPOPI = Request.Form("BPOPI")
COVIDMVCPA = Request.Form("COVIDMVCPA")
COVIDLocal = Request.Form("COVIDLocal")
COVIDPI = Request.Form("COVIDPI")
COVIDNote = Request.Form("COVIDNote")
SupplementaryComments = Request.Form("SupplementaryComments")
If Request.Form("Confirmed") = "1" Then
	Confirmed = True
Else
	Confirmed = False
End If
AmountPaid = Request.Form("AmountPaid")
PriorYearFunds = Request.Form("PriorYearFunds")
CurrentYearFunds = Request.Form("CurrentYearFunds")
DatePaid = Request.Form("DatePaid") 
AdministrativeComments = Request.Form("AdministrativeComments")

ReviewID = Request.Form("ReviewID")
ReviewDate = Request.Form("ReviewDate")
If ReviewDate = "" Then
	ReviewID = null
ElseIf (ReviewID = "" OR ReviewID="0") And Len(ReviewDate)>0 Then
	ReviewID = UserSystemID
End If

AuditApprovalID = Request.Form("AuditApprovalID")
AuditApprovalDate = Request.Form("AuditApprovalDate")
If AuditApprovalDate = "" Then
	AuditApprovalID = null
ElseIf (AuditApprovalID="" Or AuditApprovalID="0") And Len(AuditApprovalDate)>0 Then
	AuditApprovalID = UserSystemID
End If

DirectorApprovalID = Request.Form("DirectorApprovalID")
DirectorApprovalDate = Request.Form("DirectorApprovalDate")
NewDirectorApproval = False
If DirectorApprovalDate = "" Then
	DirectorApprovalID = null
ElseIf (DirectorApprovalID="" Or DirectorApprovalID="0") And Len(DirectorApprovalDate)>0 Then
	DirectorApprovalID = UserSystemID
	NewDirectorApproval = True
End If

If Request.Form("Unsubmit") = "1" Then
	Unsubmit = True
	SubmitID = null
	SubmitTimestamp = null
	Confirmed = null
	ReviewID = null
	ReviewDate = null
	AuditApprovalID = null
	AuditApprovalDate = null
	DirectorApprovalID = null
	DirectorApprovalDate = null
Else
	Unsubmit = False
End If

If Debug = True Then
	Response.Write("<pre>SubmitID=" & SubmitID & "</pre>")
	Response.Write("<pre>SubmitTimestamp=" & SubmitTimestamp & "</pre>")
End If
UpdateID = UserSystemID
UpdateTimestamp = Timestamp

sql = "SELECT GrantID, Quarter, CashExpenditureTotal, InKindExpenditureTotal, " & vbCrLf & _
	"	InLieuOfDPS, InLieuOfNICB, UnbudgetedPI, " & vbCrLf & _
	"	ReimbursableExpenditures, ReimbursementRate, ReimbursementYTD, PriorAmountPaid, Reimbursement, " & vbCrLf & _
	"	PriorYearAllocation, AwardAmount, " & vbCrLf & _
	"	InLieuOfDPSBudget, InLieuOfNICBBudget, ProgramIncomeBudget, " & vbCrLf & _
	"	PriorInLieuOfDPS, PriorInLieuOfNICB, PriorProgramIncome, " & vbCrLf & _
	"	BPOMVCPA, BPOLocal, BPOPI, COVIDMVCPA, COVIDLocal, COVIDPI, COVIDNote, " & vbCrLF & _
	"	SupplementaryComments, Confirmed, " & vbCrLF & _
	"	SubmitID, SubmitTimestamp, AmountPaid, DatePaid, AdministrativeComments, " & vbCrLf & _
	"	ReviewID, ReviewDate, AuditApprovalID, AuditApprovalDate, DirectorApprovalID, " & vbCrLF & _
	"	DirectorApprovalDate, UpdateID, UpdateTimestamp " & vbCrLF & _
	"FROM ER.Main " & vbCrLf & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	' Initial submit. Insert record
	sql = "INSERT INTO ER.Main (GrantID, Quarter, CashExpenditureTotal, InKindExpenditureTotal, " & vbCrLf & _
		"	InLieuOfDPS, InLieuOfNICB, UnbudgetedPI, ReimbursableExpenditures, " & vbCrLf & _
		"	ReimbursementRate, ReimbursementYTD, PriorAmountPaid, Reimbursement, " & vbCrLf & _
		"	PriorYearAllocation, AwardAmount, " & vbCrLf & _
		"	BeginningBalance, EarnedThisQuarter, ExpendedThisQuarter, EndingBalance, " & vbCrLf & _
		"	InLieuOfDPSBudget, InLieuOfNICBBudget, ProgramIncomeBudget, " & vbCrLf & _
		"	PriorInLieuOfDPS, PriorInLieuOfNICB, PriorProgramIncome, " & vbCrLf & _
		"	BPOMVCPA, BPOLocal, BPOPI, COVIDMVCPA, COVIDLocal, COVIDPI, COVIDNote, SupplementaryComments, " & vbCrLf & _
		"	Confirmed, SubmitID, SubmitTimestamp, UpdateID, UpdateTimestamp) VALUES (" & vbCrLf & _
		prepIntegerSQL(GrantID) & ", " & _
		prepIntegerSQL(Quarter) & ", " & _
		prepNumberSQL(CashExpenditureTotal) & ", " & _
		prepNumberSQL(InKindExpenditureTotal) & ", " & _
		prepNumberSQL(InLieuOfDPS) & ", " & _
		prepNumberSQL(InLieuOfNICB) & ", " & _
		prepNumberSQL(UnbudgetedPI) & ", " & _
		prepNumberSQL(ReimbursableExpenditures) & ", " & _
		prepNumberSQL(ReimbursementRate) & ", " & _
		prepNumberSQL(ReimbursementYTD) & ", " & _
		prepNumberSQL(PriorAmountPaid) & ", " & _
		prepNumberSQL(Reimbursement) & ", " & _
		prepNumberSQL(PriorYearAllocation) & ", " & _
		prepNumberSQL(AwardAmount) & ", " & _
		prepNumberSQL(BeginningBalance) & ", " & _
		prepNumberSQL(EarnedThisQuarter) & ", " & _
		prepNumberSQL(ExpendedThisQuarter) & ", " & _
		prepNumberSQL(EndingBalance) & ", " & _
		prepNumberSQL(InLieuOfDPSBudget) & ", " & _
		prepNumberSQL(InLieuOfNICBBudget) & ", " & _
		prepNumberSQL(ProgramIncomeBudget) & ", " & _
		prepNumberSQL(PriorInLieuOfDPS) & ", " & _
		prepNumberSQL(PriorInLieuOfNICB) & ", " & _
		prepNumberSQL(PriorProgramIncome) & ", " & _
		prepNumberSQL(BPOMVCPA) & ", " & _
		prepNumberSQL(BPOLocal) & ", " & _
		prepNumberSQL(BPOPI) & ", " & _
		prepNumberSQL(COVIDMVCPA) & ", " & _
		prepNumberSQL(COVIDLocal) & ", " & _
		prepNumberSQL(COVIDPI) & ", " & _
		prepStringSQL(COVIDNote) & ", " & _
		prepStringSQL(SupplementaryComments) & ", " & _
		prepBitSQL(Confirmed) & ", " & _
		prepIntegerSQL(SubmitID) & ", " & _
		prepStringSQL(SubmitTimestamp) & ", " & _
		prepIntegerSQL(UpdateID) & ", " & _
		prepStringSQL(Timestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
Else
	sql = "UPDATE ER.Main SET " & vbCrLf & _
		"CashExpenditureTotal=" & prepNumberSQL(CashExpenditureTotal) & ", " & vbCrLf & _
		"InKindExpenditureTotal=" & prepNumberSQL(InKindExpenditureTotal) & ", " & vbCrLf & _
		"InLieuOfDPS=" & prepNumberSQL(InLieuOfDPS) & ", " & vbCrLf & _
		"InLieuOfNICB=" & prepNumberSQL(InLieuOfNICB) & ", " & vbCrLf & _
		"UnbudgetedPI=" & prepNumberSQL(UnbudgetedPI) & ", " & vbCrLf & _
		"ReimbursableExpenditures=" & prepNumberSQL(ReimbursableExpenditures) & ", " & vbCrLf & _
		"ReimbursementRate=" & prepNumberSQL(ReimbursementRate) & ", " & vbCrLf & _
		"ReimbursementYTD=" & prepNumberSQL(ReimbursementYTD) & ", " & vbCrLf & _
		"PriorAmountPaid=" & prepNumberSQL(PriorAmountPaid) & ", " & vbCrLf & _
		"Reimbursement=" & prepNumberSQL(Reimbursement) & ", " & vbCrLf & _
		"PriorYearAllocation=" & prepNumberSQL(PriorYearAllocation) & ", " & vbCrLf & _
		"AwardAmount=" & prepNumberSQL(AwardAmount) & ", " & vbCrLf & _
		"BeginningBalance=" & prepNumberSQL(BeginningBalance) & ", " & vbCrLf & _
		"EarnedThisQuarter=" & prepNumberSQL(EarnedThisQuarter) & ", " & vbCrLf & _
		"ExpendedThisQuarter=" & prepNumberSQL(ExpendedThisQuarter) & ", " & vbCrLf & _
		"EndingBalance=" & prepNumberSQL(EndingBalance) & ", " & vbCrLf & _
		"InLieuOfDPSBudget=" & prepNumberSQL(InLieuOfDPSBudget) & ", " & vbCrLf & _
		"InLieuOfNICBBudget=" & prepNumberSQL(InLieuOfNICBBudget) & ", " & vbCrLf & _
		"ProgramIncomeBudget=" & prepNumberSQL(ProgramIncomeBudget) & ", " & vbCrLf & _
		"PriorInLieuOfDPS=" & prepNumberSQL(PriorInLieuOfDPS) & ", " & vbCrLf & _
		"PriorInLieuOfNICB=" & prepNumberSQL(PriorInLieuOfNICB) & ", " & vbCrLf & _
		"PriorProgramIncome=" & prepNumberSQL(PriorProgramIncome) & ", " & vbCrLf & _
		"BPOMVCPA=" & prepNumberSQL(BPOMVCPA) & ", " & vbCrLf & _
		"BPOLocal=" & prepNumberSQL(BPOLocal) & ", " & vbCrLf & _
		"BPOPI=" & prepNumberSQL(BPOPI) & ", " & vbCrLf & _
		"COVIDMVCPA=" & prepNumberSQL(COVIDMVCPA) & ", " & vbCrLf & _
		"COVIDLocal=" & prepNumberSQL(COVIDLocal) & ", " & vbCrLf & _
		"COVIDPI=" & prepNumberSQL(COVIDPI) & ", " & vbCrLf & _
		"COVIDNote=" & prepStringSQL(COVIDNote) & ", " & vbCrLf & _
		"SupplementaryComments=" & prepStringSQL(SupplementaryComments) & ", " & vbCrLf & _
		"Confirmed=" & prepBitSQL(Confirmed) & ", " & vbCrLf
	If Submit = True Or Unsubmit = True Then
		sql = sql & "SubmitID=" & prepIntegerSQL(SubmitID) & ", " & vbCrLf & _
		"SubmitTimestamp=" & prepStringSQL(SubmitTimestamp) & ", " & vbCrLf
	End If
		sql = sql & "UpdateID=" & prepIntegerSQL(UpdateID) & ", " & vbCrLf & _
		"UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

IF MVCPARights = True Then
	sql = "UPDATE ER.Main SET " & vbCrLf & _
		"	AdministrativeComments=" & prepStringSQL(AdministrativeComments) & ", " & vbCrLf & _
		"	ReviewID=" & prepIntegerSQL(ReviewID) & ", " & vbCrLF & _
		"	ReviewDate=" & prepDateSQL(ReviewDate) & ", " & vbCrLF & _
		"	AuditApprovalID=" & prepIntegerSQL(AuditApprovalID) & ", " & vbCrLF & _
		"	AuditApprovalDate=" & prepStringSQL(AuditApprovalDate) & ", " & vbCrLF & _
		"	DirectorApprovalID=" & prepIntegerSQL(DirectorApprovalID) & ", " & vbCrLf & _
		"	DirectorApprovalDate=" & prepDateSQL(DirectorApprovalDate) & ", " & vbCrLf & _
		"	AmountPaid=" & prepNumberSQL(AmountPaid) & ", " & vbCrLf & _
		"	PriorYearFunds=" & prepNumberSQL(PriorYearFunds) & ", " & vbCrLf & _
		"	CurrentYearFunds=" & prepNumberSQL(CurrentYearFunds) & ", " & vbCrLf & _
		"	DatePaid=" & prepDateSQL(DatePaid) & ", " & vbCrLf & _
		"	UpdateID=" & prepIntegerSQL(UpdateID) & ", " & vbCrLf & _
		"	UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

Dim CashExpenditure, InKindExpenditure, YTDExpenditure, BudgetExpenditure, RemainingBudget, ExcludedAmount
For i = 1 to 7
	CashExpenditure = Request.Form("CashExpenditure_" & i)
	If Len(CashExpenditure) = 0 Then 
		CashExpenditure = 0
	Else
		CashExpenditure = CDbl(CashExpenditure)
	End If
	InKindExpenditure = Request.Form("InKindExpenditure_" & i)
	If Len(InKindExpenditure) = 0 Then 
		InKindExpenditure = 0
	Else
		InKindExpenditure = CDbl(InKindExpenditure)
	End If
	YTDExpenditure = Request.Form("YTDExpenditure_" & i)
	If Len(YTDExpenditure) = 0 Then 
		YTDExpenditure = 0
	Else
		YTDExpenditure = CDbl(YTDExpenditure)
	End If
	BudgetExpenditure = Request.Form("BudgetExpenditure_" & i)
	If Len(BudgetExpenditure) = 0 Then 
		BudgetExpenditure = 0
	Else
		BudgetExpenditure = CDbl(BudgetExpenditure)
	End If
	RemainingBudget = Request.Form("RemainingBudget_" & i)
	If Len(RemainingBudget) = 0 Then 
		RemainingBudget = 0
	Else
		RemainingBudget = CDbl(RemainingBudget)
	End If
	ExcludedAmount = Request.Form("ExcludedAmount_" & i)
	If Len(ExcludedAmount) = 0 Then 
		ExcludedAmount = 0
	Else
		ExcludedAmount = CDbl(ExcludedAmount)
	End If

	sql = "SELECT GrantID, Quarter, BudgetCategoryID, " & vbCrLf & _
		"	ISNULL(CashExpenditure,0) AS CashExpenditure, " & vbCrLf & _
		"	ISNULL(InKindExpenditure,0) AS InKindExpenditure, " & vbCrLf & _
		"	ISNULL(YTDExpenditure,0) AS YTDExpenditure, " & vbCrLf & _
		"	ISNULL(BudgetExpenditure,0) AS BudgetExpenditure, " & vbCrLf & _
		"	ISNULL(RemainingBudget,0) AS RemainingBudget, " & vbCrLf & _
		"	ISNULL(ExcludedAmount,0) AS ExcludedAmount " & vbCrLf & _
		"FROM [ER].Detail " & vbCrLf & _
		"WHERE GrantID=" & GrantID & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND BudgetCategoryID=" & i
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)

	If rs.EOF = True Then ' Do an Insert if there are any values
		If CashExpenditure<>0 Or InKindExpenditure<>0 Or YTDExpenditure<>0 Or BudgetExpenditure<>0 Then
			sql = "INSERT INTO [ER].Detail (GrantID, Quarter, BudgetCategoryID, " & vbCrLf & _
				"	CashExpenditure, InKindExpenditure, YTDExpenditure, BudgetExpenditure, " & vbCrLf & _
				"	RemainingBudget, ExcludedAmount, UpdateID, UpdateTimestamp)" & vbCrLF & _
				"VALUES (" & GrantID & ", " & prepIntegerSQL(quarter) & ", " & i & _
				", " & prepNumberSQL(CashExpenditure) & _
				", " & prepNumberSQL(InKindExpenditure) & _
				", " & prepNumberSQL(YTDExpenditure) & _
				", " & prepNumberSQL(BudgetExpenditure) & _
				", " & prepNumberSQL(RemainingBudget) & _
				", " & prepNumberSQL(ExcludedAmount) & _
				", " & UserSystemID & ", " & prepStringSQL(Timestamp) & ")"
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Con.Execute(sql)
		End If
	Else ' Do an update if necessary
		If Submit = True Or _
			CashExpenditure<>rs.Fields("CashExpenditure") Or _
			InKindExpenditure<>rs.Fields("InKindExpenditure") Or _
			YTDExpenditure<>rs.Fields("YTDExpenditure") Or _
			BudgetExpenditure<>rs.Fields("BudgetExpenditure") Or _
			RemainingBudget<>rs.Fields("RemainingBudget") Or _
			ExcludedAmount<>rs.Fields("ExcludedAmount") Then
			sql = "UPDATE [ER].Detail SET " & vbCrLf & _
				"CashExpenditure=" & prepNumberSQL(CashExpenditure) & ", " & _
				"InKindExpenditure= " & prepNumberSQL(InKindExpenditure) & ", " & _
				"YTDExpenditure=" & prepNumberSQL(YTDExpenditure) & ", " & _
				"BudgetExpenditure= " & prepNumberSQL(BudgetExpenditure) & ", " & _
				"RemainingBudget=" & prepNumberSQL(RemainingBudget) & ", " & _
				"ExcludedAmount=" & prepNumberSQL(ExcludedAmount) & ", " & _
				"UpdateID= " & UserSystemID & ", " & _
				"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLf & _
				"WHERE GrantID=" & prepIntegerSQL(GrantID) & _
				" 	AND Quarter=" & prepIntegerSQL(Quarter) & _
				" 	AND BudgetCategoryID=" & i
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

sql = "SELECT GrantID, Quarter, BeginningBalance, EarnedThisQuarter, ExpendedThisQuarter, EndingBalance " & vbCrLF & _
	"FROM [Grants].ProgramIncome " & vbCrLF & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs=Con.Execute(sql)
If rs.EOF = True Then ' Do an insert
	sql = "INSERT INTO [Grants].ProgramIncome (GrantID, Quarter, BeginningBalance, EarnedThisQuarter, ExpendedThisQuarter, EndingBalance, UpdateID, UpdateTimestamp) " & VBcRlF & _
		"VALUES (" & prepIntegerSQL(GrantID) & ", " & prepIntegerSQL(Quarter) & ", " & _
		prepNumberSQL(BeginningBalance) & ", " & _
		prepNumberSQL(EarnedThisQuarter) & ", " & _
		prepNumberSQL(ExpendedThisQuarter) & ", " & _
		prepNumberSQL(EndingBalance) & ", " & _
		prepIntegerSQL(UserSystemID) & ", " & _
		prepStringSQL(Timestamp) & ")"
Else ' Do an update
	sql = "UPDATE [Grants].ProgramIncome SET " & vbCrLf & _
		"	BeginningBalance=" & prepNumberSQL(BeginningBalance) & ", " & _
		"	EarnedThisQuarter=" & prepNumberSQL(EarnedThisQuarter) & ", " & _
		"	ExpendedThisQuarter=" & prepNumberSQL(ExpendedThisQuarter) & ", " & _
		"	EndingBalance=" & prepNumberSQL(EndingBalance) & ", " & _
		"	UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
		"	UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter)
End If
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Con.Execute(sql)

If NewDirectorApproval = True Then
	sql = "SELECT EquipmentID FROM ER.EquipmentDetail READONLY WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND ISNULL(Cost,0.0)>0.0 "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		sql = "EXEC ER.CreateInventoryFromEquipment @EquipmentID=" & rs.Fields("EquipmentID")
		If Debug = True Then
			Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Con.Execute(sql)
		rs.MoveNext
	Wend
End If
If Action = "Equipment" Then
	If Debug = True Then
		Response.Write("<a href=""Equipment.asp?GrantID=" & GrantID & "&Quarter=" & Quarter & """>return to expenditure report</a><br />" & vbCrLf)
	Else
		Response.Redirect("Equipment.asp?GrantID=" & GrantID & "&Quarter=" & Quarter)
	End If
Else
	If Debug = True Then
		Response.Write("<a href=""Report.asp?GrantID=" & GrantID & "&Quarter=" & Quarter & """>return to expenditure report</a><br />" & vbCrLf)
		Response.Write("<a href=""../Home/Default.asp?GrantID=" & GrantID & """>return to Home</a><br />" & vbCrLf)
	Else
		Response.Redirect("Report.asp?GrantID=" & GrantID & "&Quarter=" & Quarter)
	End If
End If

%><!--#include file="../includes/prepDB.asp"-->