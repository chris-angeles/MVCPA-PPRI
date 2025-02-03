<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim MAXQUARTER
Dim debug, i, j, PermitEdit, ViewDocuments, UploadDirectory, _
	GrantID, FiscalYear, Quarter, GranteeID, GranteeName, ProgramName, GrantNumber, _
	Confirmed, SubmitID, SubmitTimeStamp, SubmitName, SubmitterEMail, _
	ReviewID, ReviewName, ReviewDate, AuditApprovalID, AuditApprovalName, AuditApprovalDate, _
	DirectorApprovalID, DirectorApprovalName, DirectorApprovalDate, _
	AmountPaid, PriorYearFunds, CurrentYearFunds, DatePaid, AdministrativeComments, _
	PriorYearAllocation, AwardAmount, ReimbursementRate, Reimbursement, _
	InLieuOfDPSBudget, InLieuOfNICBBudget, ProgramIncomeBudget, _
	InLieuOfDPS, InLieuOfNICB, UnbudgetedPI, _
	BeginningBalance, EarnedThisQuarter, ExpendedThisQuarter, EndingBalance, _
	BPOMVCPA, BPOLocal, BPOPI, BPOTotal, COVIDMVCPA, COVIDLocal, COVIDPI, COVIDTotal, COVIDNote, _
	SupplementaryComments, PriorInLieuOfDPS, PriorInLieuOfNICB, PriorAmountPaid, PriorProgramIncome, _
	UpdateID, UpdateName, UpdateTimestamp, PriorNonApproval, LaterSubmission, _
	CanSubmit, CanApprove, CanInvoice, StartDate, PRSubmitted, PRApproved

debug = False

'MVCPAGrantCoordinator=True
'MVCPAAdministrator = True
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

If Len(Request.Form("GrantID"))>0 Then
	GrantID = Request.Form("GrantID")
ElseIf Len(Request.QueryString("GrantID"))>0 Then
	GrantID = Request.QueryString("GrantID")
Else
	GrantID = Session("GrantID")
End If
If Len(Request.Form("Quarter"))>0 Then
	Quarter = Request.Form("Quarter")
ElseIf Len(Request.Querystring("Quarter"))>0 Then
	Quarter = Request.Querystring("Quarter")
Else
	Quarter = 1
End If

sql = "SELECT A.GranteeID, B.GrantID, B.FiscalYear, ISNULL(C.Quarter, " & prepIntegerSQL(Quarter) & ") AS Quarter , " & vbCrLf & _
	"	A.GranteeName, B.ProgramName, B.GrantNumber, " & vbCrLf & _
	"	CAST(CASE WHEN C.SubmitID IS NULL THEN 0 ELSE ISNULL(C.Confirmed,0) END AS BIT) AS Confirmed,  " & vbCrLf & _
	"	ISNULL(C.SubmitID,0) AS SubmitID, D.Name AS SubmitName, D.Email AS SubmitterEMail, C.SubmitTimestamp, " & vbCrLf & _
	"	ISNULL(C.ReviewID,0) AS ReviewID, L.Name As ReviewName, C.ReviewDate, " & vbCrLf & _
	"	ISNULL(C.AuditApprovalID,0) AS AuditApprovalID, E.Name AS AuditApprovalName, C.AuditApprovalDate, " & vbCrLf & _
	"	ISNULL(DirectorApprovalID,0) AS DirectorApprovalID, F.Name AS DirectorApprovalName, DirectorApprovalDate, " & vbCrLf & _
	"	CASE WHEN C.SubmitID IS NOT NULL AND C.AmountPaid IS NULL THEN C.Reimbursement ELSE C.AmountPaid END AS AmountPaid, " & vbCrLf & _
	"	ISNULL(C.PriorYearFunds, 0.0) AS PriorYearFunds, ISNULL(C.CurrentYearFunds, 0.0) AS CurrentYearFunds, " & vbCrLf & _
	"	C.DatePaid, C.AdministrativeComments, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.PriorYearAllocation, 0.0) ELSE ISNULL(B.PriorYearAllocation,0.0) END AS PriorYearAllocation, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.AwardAmount, 0.0) ELSE B.AwardAmount END AS AwardAmount, " & vbCrLf & _
	"	CASE WHEN C.SubmitID IS NULL THEN B.ReimbursementRate ELSE C.ReimbursementRate END AS ReimbursementRate, " & vbCrLf & _
	"	C.Reimbursement, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.InLieuOfDPSBudget,0.0) ELSE ISNULL(B.InLieuOfDPSBudget,0.0) END AS InLieuOfDPSBudget, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.InLieuOfNICBBudget,0.0) ELSE ISNULL(B.InLieuOfNICBBudget,0.0) END AS InLieuOfNICBBudget, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.ProgramIncomeBudget,0.0) ELSE ISNULL(B.ProgramIncomeBudget, 0.0) END AS ProgramIncomeBudget, " & vbCrLf & _
	"	ISNULL(C.InLieuOfDPS,0.0) AS InLieuOfDPS, " & vbCrLf & _
	"	ISNULL(C.InLieuOfNICB, 0.0) AS InLieuOfNICB, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.BeginningBalance,0.0) WHEN H.EndingBalance IS NOT NULL THEN H.EndingBalance ELSE ISNULL(I.BeginningBalance,0.0) END AS BeginningBalance, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.EarnedThisQuarter,0.0) ELSE ISNULL(I.EarnedThisQuarter,0.0) END AS EarnedThisQuarter, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.ExpendedThisQuarter,0.0) ELSE ISNULL(I.ExpendedThisQuarter,0.0) END AS ExpendedThisQuarter, " & vbCrLf & _
	"	ISNULL(I.EndingBalance,0.0) AS EndingBalance, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN C.UnbudgetedPI " & vbCrLf & _
	"		WHEN ISNULL(K.PriorProgramIncome, 0.0)+ISNULL(I.EarnedThisQuarter,0.0)-ISNULL(B.ProgramIncomeBudget,0.0) < 0.0 THEN 0.0 " & vbCrLf & _
	"		WHEN ISNULL(K.PriorProgramIncome, 0.0)+ISNULL(I.EarnedThisQuarter,0.0)-ISNULL(B.ProgramIncomeBudget,0.0) > 1000 THEN 1000.0 " & vbCrLf & _
	"		ELSE ISNULL(K.PriorProgramIncome, 0.0)+ISNULL(I.EarnedThisQuarter,0.0)-ISNULL(B.ProgramIncomeBudget,0.0) END AS UnbudgetedPI, " & vbCrLf & _
	"	BPOMVCPA, " & vbCrLf & _
	"	BPOLocal, " & vbCrLf & _
	"	BPOPI, " & vbCrLf & _
	"	ISNULL(BPOMVCPA,0.0)+ISNULL(BPOLocal,0.0)+ISNULL(BPOPI,0.0) AS BPOTotal, " & vbCrLf & _
	"	COVIDMVCPA, " & vbCrLf & _
	"	COVIDLocal, " & vbCrLf & _
	"	COVIDPI, " & vbCrLf & _
	"	ISNULL(COVIDMVCPA,0.0)+ISNULL(COVIDLocal,0.0)+ISNULL(COVIDPI,0.0) AS COVIDTotal, " & vbCrLf & _
	"	COVIDNote, SupplementaryComments, " & vbCrLf & _
	"	C.UpdateID, G.Name AS UpdateName, C.UpdateTimestamp, " & vbCrLf & _
	"	ISNULL(J.PriorNonApproval,0) AS PriorNonApproval, ISNULL(N.LaterSubmission,0) AS LaterSubmission, " & vbCrLf & _
	"	CAST(CASE WHEN " & UserSystemID & " IN (A.FinancialOfficerId, A.FinancialAdministrativeContactID) THEN 1 ELSE 0 END AS BIT) AS CanSubmit, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.PriorInLieuOfDPS,0.0) ELSE ISNULL(J.PriorInLieuOfDPS,0.0) END AS PriorInLieuOfDPS, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.PriorInLieuOfNICB,0.0) ELSE ISNULL(J.PriorInLieuOfNICB,0.0) END AS PriorInLieuOfNICB, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.PriorProgramIncome,0.0) ELSE ISNULL(K.PriorProgramIncome, 0.0) END AS PriorProgramIncome, " & vbCrLf & _
	"	CASE WHEN C.SubmitID>0 THEN ISNULL(C.PriorAmountPaid,0.0) ELSE ISNULL(J.PriorAmountPaid,0.0) END AS PriorAmountPaid, " & vbCrLf & _
	"	CAST(CASE WHEN M.SubmitID>0 THEN 1 ELSE 0 END AS BIT) AS PRSubmitted, " & vbCrLf & _
	"	CAST(CASE WHEN M.ApprovalID>0 THEN 1 ELSE 0 END AS BIT) AS PRApproved " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"LEFT JOIN [Grants].Main AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
	"LEFT JOIN ER.Main AS C ON C.GrantID=B.GrantID AND C.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"LEFT JOIN [System].Users AS D ON D.SystemID=C.SubmitID " & vbCrLf & _
	"LEFT JOIN [System].Users AS E ON E.SystemID=C.AuditApprovalID " & vbCrLf & _
	"LEFT JOIN [System].Users AS G ON G.SystemID=C.UpdateID " & vbCrLf & _
	"LEFT JOIN [System].Users AS F ON F.SystemID=C.DirectorApprovalID " & vbCrLf & _
	"LEFT JOIN [Grants].ProgramIncome AS H ON " & vbCrLf & _
	"	H.GrantID=CASE WHEN ISNULL(C.Quarter, " & prepIntegerSQL(Quarter) & ")=1 THEN B.PreviousYearGrantID ELSE B.GrantID END AND " & vbCrLf & _
	"	H.Quarter=CASE WHEN ISNULL(C.Quarter, " & prepIntegerSQL(Quarter) & ")=1 THEN 4 ELSE ISNULL(C.Quarter, " & prepIntegerSQL(Quarter) & ")-1 END " & vbCrLf & _
	"LEFT JOIN [Grants].ProgramIncome AS I ON I.GrantID=B.GrantID AND I.Quarter=ISNULL(C.Quarter, " & prepIntegerSQL(Quarter) & ") " & vbCrLf & _
	"LEFT JOIN (SELECT GrantID, SUM(InLieuOfDPS) AS PriorInLieuOfDPS, " & vbCrLf & _
	"	SUM(InLieuOfNICB) AS PriorInLieuOfNICB, SUM(AmountPaid) AS PriorAmountPaid, " & vbCrLf & _
	"	CAST(MAX(CASE WHEN DirectorApprovalID IS NULL THEN 1 ELSE 0 END) AS BIT) AS PriorNonApproval " & vbCrLf & _
	"	FROM ER.Main WHERE Quarter<" & prepIntegerSQL(Quarter) & " GROUP BY GrantID) AS J ON J.GrantID=B.GrantID " & vbCrLf & _
	"LEFT JOIN (SELECT GrantID, SUM(ExpendedThisQuarter) AS PriorProgramIncome " & vbCrLf & _
	"	FROM [Grants].ProgramIncome WHERE Quarter<" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"	GROUP BY GrantID) AS K ON K.GrantID=B.GrantID " & vbCrLf & _
	"LEFT JOIN [System].Users AS L ON L.SystemID=C.ReviewID " & vbCrLf & _
	"LEFT JOIN [PR].Main AS M ON M.GrantID=B.GrantID AND M.Quarter=C.Quarter " & vbCrLf & _
	"LEFT JOIN (SELECT GrantID, CAST(MAX(CASE WHEN SubmitID>0 THEN 1 ELSE 0 END) AS BIT) AS LaterSubmission " & vbCrLf & _
	"	FROM [ER].Main WHERE Quarter>" & prepIntegerSQL(Quarter) & " GROUP BY GrantID) AS N ON N.GrantID=B.GrantID " & vbCrLf & _
	"WHERE B.GrantID=" & prepIntegerSQL(GrantID) 
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Expenditure Report record retrieved")
	SendMessage "Error: No Expenditure Report record retrieved"
	Response.End
Else
	FiscalYear = rs.Fields("FiscalYear")
	Quarter = rs.Fields("Quarter")
	GranteeID = rs.Fields("GranteeID")
	PriorYearAllocation = rs.Fields("PriorYearAllocation")
	AwardAmount = rs.Fields("AwardAmount")
	InLieuOfDPSBudget = rs.Fields("InLieuOfDPSBudget")
	InLieuOfNICBBudget = rs.Fields("InLieuOfNICBBudget")
	ProgramIncomeBudget = rs.Fields("ProgramIncomeBudget")
	ReimbursementRate = rs.Fields("ReimbursementRate")
	Reimbursement = rs.Fields("Reimbursement")
	InLieuOfDPS = rs.Fields("InLieuOfDPS")
	InLieuOfNICB = rs.Fields("InLieuOfNICB")
	BeginningBalance = rs.Fields("BeginningBalance")
	EarnedThisQuarter = rs.Fields("EarnedThisQuarter")
	ExpendedThisQuarter = rs.Fields("ExpendedThisQuarter")
	EndingBalance = rs.Fields("EndingBalance")
	UnbudgetedPI = rs.Fields("UnbudgetedPI")
	BPOMVCPA = rs.Fields("BPOMVCPA")
	BPOLocal = rs.Fields("BPOLocal")
	BPOPI = rs.Fields("BPOPI")
	BPOTotal = rs.Fields("BPOTotal")
	COVIDMVCPA = rs.Fields("COVIDMVCPA")
	COVIDLocal = rs.Fields("COVIDLocal")
	COVIDPI = rs.Fields("COVIDPI")
	COVIDTotal = rs.Fields("COVIDTotal")
	COVIDNote = rs.Fields("COVIDNote")
	SupplementaryComments = rs.Fields("SupplementaryComments")
	GranteeName = rs.Fields("GranteeName")
	ProgramName = rs.Fields("ProgramName")
	GrantNumber = rs.Fields("GrantNumber")
	Confirmed = rs.Fields("Confirmed")
	SubmitID = rs.Fields("SubmitID")
	SubmitTimeStamp = rs.Fields("SubmitTimeStamp")
	SubmitName = rs.Fields("SubmitName")
	SubmitterEMail = rs.Fields("SubmitterEMail")
	ReviewID = rs.Fields("ReviewID")
	ReviewName = rs.Fields("ReviewName")
	ReviewDate = rs.Fields("ReviewDate")
	AuditApprovalID = rs.Fields("AuditApprovalID")
	AuditApprovalName = rs.Fields("AuditApprovalName")
	AuditApprovalDate = rs.Fields("AuditApprovalDate")
	DirectorApprovalID = rs.Fields("DirectorApprovalID")
	DirectorApprovalName = rs.Fields("DirectorApprovalName")
	DirectorApprovalDate = rs.Fields("DirectorApprovalDate")
	AmountPaid = rs.Fields("AmountPaid")
	PriorYearFunds = rs.Fields("PriorYearFunds")
	CurrentYearFunds = rs.Fields("CurrentYearFunds")
	DatePaid = rs.Fields("DatePaid")
	AdministrativeComments = rs.Fields("AdministrativeComments")
	UpdateID = rs.Fields("UpdateID")
	UpdateName = rs.Fields("UpdateName")
	UpdateTimestamp = rs.Fields("UpdateTimestamp")
	PriorNonApproval = rs.Fields("PriorNonApproval")
	LaterSubmission = rs.Fields("LaterSubmission")
	CanSubmit = rs.Fields("CanSubmit")
	PriorInLieuOfDPS = rs.Fields("PriorInLieuOfDPS")
	PriorInLieuOfNICB = rs.Fields("PriorInLieuOfNICB")
	PriorAmountPaid = rs.Fields("PriorAmountPaid")
	PriorProgramIncome = rs.Fields("PriorProgramIncome")
	PRSubmitted = rs.Fields("PRSubmitted")
	PRApproved = rs.Fields("PRApproved")
End If

If FiscalYear > 2021 Then
	Response.Redirect("Report2.asp?GrantID=" & GrantID & "&Quarter=" & Quarter)
End If
'If Developer Then 
'	MAXQUARTER = 4
'Else
	'MAXQUARTER = 4
	If Date() > CDate("8/1/" & (FiscalYear)) Then
		MAXQUARTER = 4
	ElseIf Date() > CDate("5/1/" & (FiscalYear)) Then
		MAXQUARTER = 3
	ElseIf Date() > CDate("2/1/" & (FiscalYear)) Then
		MAXQUARTER = 2
	Else
		MAXQUARTER = 1
	End If
'End If

If Quarter = 1 Then
	StartDate = CDate("12/1/" & (FiscalYear-1))
ElseIf Quarter = 2 Then
	StartDate = CDate("3/1/" & FiscalYear)
ElseIf Quarter = 3 Then
	StartDate = CDate("6/1/" & FiscalYear)
ElseIf Quarter = 4 Then
	StartDate = CDate("9/1/" & FiscalYear)
End If

If SubmitID = 0 Then
	PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
ElseIf SubmitID > 0 Then
	PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
Else
	PermitEdit = False
End If
ViewDocuments = CheckPermissions(UserSystemID, GranteeID, True)
' Changed to give state auditor access.
If MVCPAViewer = True Then
	ViewDocuments = True
End If

' If equipment, use sum of equipment from ER.EquipmentDetail rather than cashexpenditures until submitted.
sql = "SELECT C.BudgetCategoryID, C.BudgetCategory, " & vbCrLf & _
	"	CASE WHEN B.SubmitID>0 THEN D.CashExpenditure " & vbCrLf & _
	"		WHEN C.BudgetCategoryID=6 THEN G.Equipment " & vbCrLf & _
	"		WHEN D.CashExpenditure IS NOT NULL THEN D.CashExpenditure " & vbCrLf & _
	"		WHEN E.TotalExpenditures>0.0 THEN 0.00 " & vbCrLf & _
	"		ELSE NULL END AS CashExpenditure, " & vbCrLf & _
	"	CASE WHEN B.SubmitID>0 THEN D.ExcludedAmount " & vbCrLf & _
	"		WHEN D.ExcludedAmount IS NOT NULL THEN  D.ExcludedAmount " & vbCrLf & _
	"		ELSE NULL END AS ExcludedAmount, " & vbCrLf & _
	"	CASE WHEN B.SubmitID>0 THEN D.InKIndExpenditure " & vbCrLf & _
	"		WHEN D.InKindExpenditure IS NOT NULL THEN D.InKindExpenditure " & vbCrLf & _
	"		WHEN E.InKindExpenditures>0.0 THEN 0.00 " & vbCrLf & _
	"		ELSE NULL END AS InKindExpenditure, " & vbCrLf & _
	"	CASE WHEN B.SubmitID>0 THEN ISNULL(D.YTDExpenditure,0.0)-ISNULL(D.CashExpenditure, 0.0)+ISNULL(ExcludedAmount,0.0) " & vbCrLf & _
	"		ELSE ISNULL(F.PriorExpenditure,0.0) END AS PriorExpenditure, " & vbCrLf & _
	"	CASE WHEN B.SubmitID>0 THEN ISNULL(D.YTDExpenditure,0.0) " & vbCrLf & _
	"		WHEN C.BudgetCategoryID=6 THEN ISNULL(F.PriorExpenditure,0.0) + ISNULL(G.Equipment,0) " & vbCrLf & _
	"		ELSE ISNULL(F.PriorExpenditure,0.0)+ ISNULL(D.CashExpenditure,0) END AS YTDExpenditure, " & vbCrLf & _
	"	CASE WHEN B.SubmitID>0 THEN ISNULL(D.BudgetExpenditure,0.0) " & vbCrLf & _
	"		ELSE ISNULL(E.TotalExpenditures,0.0) END AS BudgetExpenditure, " & vbCrLf & _
	"	CASE WHEN B.SubmitID>0 THEN ISNULL(D.RemainingBudget,0.0) " & vbCrLf & _
	"	WHEN C.BudgetCategoryID=6 THEN ISNULL(E.TotalExpenditures,0.0)-ISNULL(G.Equipment,0.0)-ISNULL(F.PriorExpenditure,0.0)" & vbCrLf & _
	"		ELSE ISNULL(E.TotalExpenditures,0.0)-ISNULL(D.CashExpenditure,0.0)-ISNULL(F.PriorExpenditure,0.0) END " & vbCrLf & _
	"		AS RemainingBudget, " & vbCrLf & _
	"	CAST(CASE WHEN ISNULL(E.TotalExpenditures,0.0)=0.0 THEN 0.0 " & vbCrLf & _
	"		WHEN C.BudgetCategoryID IN (3) THEN ISNULL(E.TotalExpenditures, 0.0) + 1000.0 " & vbCrLf & _
	"		WHEN C.BudgetCategoryID IN (6) THEN ISNULL(E.TotalExpenditures, 0.0) " & vbCrLf & _
	"		ELSE ISNULL(E.TotalExpenditures, 0.0) + Round(0.05*(ISNULL(A.MatchAmount,0.0)+ISNULL(A.AwardAmount,0.0)),2) END AS MONEY) + " & vbCrLf & _
	"		ISNULL(F.AllowedOverage,0.0) AS MaxExpenditure, " & vbCrLf & _
	"	ISNULL(F.AllowedOverage,0.0) AS AllowedOverage " & vbCrLf & _
	"FROM [Grants].Main AS A " & vbCrLf & _
	"LEFT JOIN ER.Main AS B ON B.GrantID=A.GrantID AND B.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"CROSS JOIN lookup.BudgetCategories AS C " & vbCrLf & _
	"LEFT JOIN ER.Detail AS D ON D.GrantID=A.GrantID AND D.BudgetCategoryID=C.BudgetCategoryID AND D.Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"LEFT JOIN [Grants].Budget AS E ON E.GrantID=A.GrantID AND E.BudgetCategoryID=C.BudgetCategoryID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT GrantID, BudgetCategoryID, " & vbCrLf & _
	"		ISNULL(SUM(CashExpenditure),0.0)-ISNULL(SUM(ExcludedAmount),0.0) AS PriorExpenditure, " & vbCrLf & _
	"		SUM(AllowedOverage) AS AllowedOverage " & vbCrLf & _
	"	FROM ER.Detail " & vbCrLf & _
	"	WHERE Quarter<" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"	GROUP BY GrantID, BudgetCategoryID " & vbCrLf & _
	") AS F ON F.GrantID=A.GrantID AND F.BudgetCategoryID=C.BudgetCategoryID " & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT GrantID, Quarter, SUM(Cost) AS Equipment " & vbCrLf & _
	"	FROM ER.EquipmentDetail " & vbCrLf & _
	"	GROUP BY GrantID, Quarter " & vbCrLf & _
	") AS G ON G.GrantID=A.GrantID AND G.Quarter=" & prepIntegerSQL(Quarter) & " AND C.BudgetCategoryID=6 " & vbCrLf & _
	"WHERE A.GrantID=" & prepIntegerSQL(GrantID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Budget Item records retrieved")
	SendMessage "Error: No Budget Item records retrieved"
	Response.End
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Quarterly Expenditure Report for <%=GranteeName %></title>
<meta http-equiv="cache-control" content="no-cache" />
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	var MaxExpenditure = [0<% 
	while rs.EOF = false
		Response.Write(", " & rs.Fields("MaxExpenditure"))
		rs.MoveNext
	Wend
	rs.MoveFirst
%>];
	var YTDExpenditure = [0<%
	while rs.EOF = false
		Response.Write(", " & rs.Fields("PriorExpenditure"))
		rs.MoveNext
	Wend
	rs.MoveFirst
%>];
	var AllowedOverage = [0<%
	while rs.EOF = false
		Response.Write(", " & rs.Fields("AllowedOverage"))
	rs.MoveNext
	Wend
	rs.MoveFirst
%>];
	var reimbursementrate=<%=ReimbursementRate%>, awardamount=<%=awardamount%>, prioramountpaid=<%=PriorAmountPaid%>, prioryearallocation=<%=PriorYearAllocation%>;
	function updateTotals()
	{
		var reimbursementytd, reimbursement;
		var cashexpendituretotal = 0.0, excludedamounttotal = 0.0; inkindexpendituretotal = 0.0, ytdexpendituretotal = 0.0, budgetexpendituretotal = 0.0, remainingbudgettotal = 0.0, reimbursabletotal = 0.0;
		for (var i = 1; i < 8; i++)
		{
<%	If SubmitID = 0 Then %>
			if (YTDExpenditure[i]+getNumericValue(document.ER["CashExpenditure_" + i].value) >= MaxExpenditure[i]) {
				document.ER["ExcludedAmount_" + i].value = currency(YTDExpenditure[i]+getNumericValue(document.ER["CashExpenditure_" + i].value) - MaxExpenditure[i]);}
			else
		{document.ER["ExcludedAmount_" + i].value = "$0.00";}
<%	End If %>
			document.ER["YTDExpenditure_" + i].value = currency(YTDExpenditure[i] + getNumericValue(document.ER["CashExpenditure_" + i].value) - getNumericValue(document.ER["ExcludedAmount_" + i].value));
			document.ER["RemainingBudget_" + i].value = currency(getNumericValue(document.ER["BudgetExpenditure_" + i].value) - getNumericValue(document.ER["YTDExpenditure_" + i].value));
			cashexpendituretotal = cashexpendituretotal + getNumericValue(document.ER["CashExpenditure_" + i].value);
			excludedamounttotal = excludedamounttotal + getNumericValue(document.ER["ExcludedAmount_" + i].value);
			inkindexpendituretotal = inkindexpendituretotal + getNumericValue(document.ER["InKindExpenditure_" + i].value);
			ytdexpendituretotal = ytdexpendituretotal + getNumericValue(document.ER["YTDExpenditure_" + i].value);
			budgetexpendituretotal = budgetexpendituretotal + getNumericValue(document.ER["BudgetExpenditure_" + i].value);
			remainingbudgettotal = remainingbudgettotal + getNumericValue(document.ER["RemainingBudget_" + i].value);
			reimbursabletotal = reimbursabletotal + getNumericValue(document.ER["YTDExpenditure_" + i].value);
		}
		document.ER.CurrentInLieuOfDPS.value = document.ER.InLieuOfDPS.value;
		document.ER.CurrentInLieuOfNICB.value = document.ER.InLieuOfNICB.value;
		document.ER.CurrentProgramIncome.value =  document.ER.ExpendedThisQuarter.value;
		document.ER.BalanceDPS.value = currency(<%=InLieuOfDPSBudget%> - <%=PriorInLieuOfDPS%> - getNumericValue(document.ER.CurrentInLieuOfDPS.value));
		document.ER.BalanceNICB.value = currency(<%=InLieuOfNICBBudget%> - <%=PriorInLieuOfNICB%> - getNumericValue(document.ER.CurrentInLieuOfNICB.value));
		document.ER.BalanceProgramIncome.value = currency(<%=ProgramIncomeBudget%> - <%=PriorProgramIncome%> - getNumericValue(document.ER.CurrentProgramIncome.value));
		if (getNumericValue(document.ER.ExpendedThisQuarter.value) >= <%=(ProgramIncomeBudget-PriorProgramIncome)%>) {
		document.ER.UnbudgetedPI.value = currency((getNumericValue(document.ER.ExpendedThisQuarter.value) - <%=(ProgramIncomeBudget-PriorProgramIncome)%>));
		} else {
			document.ER.UnbudgetedPI.value =  currency(0.0);
		}
		reimbursabletotal = reimbursabletotal - getNumericValue(document.ER["PriorInLieuOfDPS"].value) - getNumericValue(document.ER["PriorInLieuOfNICB"].value) -
			getNumericValue(document.ER["InLieuOfDPS"].value) - getNumericValue(document.ER["InLieuOfNICB"].value) - getNumericValue(document.ER.UnbudgetedPI.value);
		document.ER.CashExpenditure_Total.value = currency(cashexpendituretotal);
		document.ER.ExcludedAmount_Total.value = currency(excludedamounttotal);
		document.ER.InKindExpenditure_Total.value = currency(inkindexpendituretotal);
		document.ER.YTDExpenditure_Total.value = currency(ytdexpendituretotal);
		document.ER.ExpendituresYTD.value = currency(ytdexpendituretotal);
		document.ER.BudgetExpenditure_Total.value = currency(budgetexpendituretotal);
		document.ER.RemainingBudget_Total.value = currency(remainingbudgettotal);
		document.ER.ReimbursableExpenditures.value = currency(reimbursabletotal);
		reimbursementytd = 0.01* Math.round(reimbursabletotal * reimbursementrate);
		if (reimbursementytd>awardamount) reimbursementytd = awardamount;
		document.ER.ReimbursementYTD.value = currency(reimbursementytd);
		reimbursement = reimbursementytd - prioramountpaid;
		document.ER.Reimbursement.value = currency(reimbursement);
		document.ER.EndingBalance.value = currency(getNumericValue(document.ER.BeginningBalance.value) +
			getNumericValue(document.ER.EarnedThisQuarter.value) - getNumericValue(document.ER.ExpendedThisQuarter.value));
		document.ER.BPOTotal.value = currency(getNumericValue(document.ER.BPOMVCPA.value) +
			getNumericValue(document.ER.BPOLocal.value) + getNumericValue(document.ER.BPOPI.value));
		<%	If FiscalYear=2020 Or FiscalYear=2021 Then %>
		document.ER.COVIDTotal.value = currency(getNumericValue(document.ER.COVIDMVCPA.value) +
			getNumericValue(document.ER.COVIDLocal.value) + getNumericValue(document.ER.COVIDPI.value));
		<%	End If %>
		checkLimits();
		<% If MVCPARights = True And SubmitID>0 Then %>
		document.ER.AmountPaid.value = document.ER.Reimbursement.value;
		if (getNumericValue(document.ER.AmountPaid.value) > 0.0)
		{
			if(prioramountpaid>=prioryearallocation) {
				document.ER.PriorYearFunds.value = "$0.00";
				document.ER.CurrentYearFunds.value = document.ER.AmountPaid.value;
			} else if (getNumericValue(document.ER.AmountPaid.value) + prioramountpaid <= prioryearallocation) {
				document.ER.PriorYearFunds.value = document.ER.AmountPaid.value;
				document.ER.CurrentYearFunds.value = "$0.00";
			} else {
				document.ER.PriorYearFunds.value = currency(prioryearallocation - prioramountpaid);
				document.ER.CurrentYearFunds.value = currency(getNumericValue(document.ER.AmountPaid.value) - (prioryearallocation - prioramountpaid));
			}
		}
		<%	End If %>
	}

	function checkLimits()
	{
		var passed = true;
		for (var i = 1; i < 8; i++) {
			if (getNumericValue(document.ER["ExcludedAmount_" + i].value) > 0.0) {
				document.ER["CashExpenditure_" + i].setAttribute("class", "warning");
				document.ER["ExcludedAmount_" + i].setAttribute("class", "warning");
				passed = false;
			}
			else {
				if (document.ER["CashExpenditure_" + i].getAttribute("class") == "warning")
					document.ER["CashExpenditure_" + i].setAttribute("class", "");
					document.ER["ExcludedAmount_" + i].setAttribute("class", "");
			}
		}
		if (getNumericValue(document.ER["InLieuOfDPS"].value) > <%=(Round(InLieuOfDPSBudget/4.0,2))%>) {
			document.ER["InLieuOfDPS"].setAttribute("class", "warning");
		}
		else {
			if (document.ER["InLieuOfDPS"].getAttribute("class") == "warning")
				document.ER["InLieuOfDPS"].setAttribute("class", "");
		}
		if (getNumericValue(document.ER["InLieuOfNICB"].value) > <%=(Round(InLieuOfNICBBudget/4.0,2))%>) {
			document.ER["InLieuOfNICB"].setAttribute("class", "warning");
		}
		else {
			if (document.ER["InLieuOfNICB"].getAttribute("class") == "warning")
				document.ER["InLieuOfNICB"].setAttribute("class", "");
		}
		if (getNumericValue(document.ER.ExpendedThisQuarter.value) > <%=(ProgramIncomeBudget-PriorProgramIncome+1000)%>) {
			document.ER["UnbudgetedPI"].setAttribute("class", "warning");
		}
		else {
			if (document.ER["UnbudgetedPI"].getAttribute("class") == "warning")
				document.ER["UnbudgetedPI"].setAttribute("class", "");
		}
		if (getNumericValue(document.ER.BalanceProgramIncome.value) < -1000.0) {
			document.ER["BalanceProgramIncome"].setAttribute("class", "warning");
			document.ER["ExpendedThisQuarter"].setAttribute("class", "warning");
		}
		else {
			if (document.ER["BalanceProgramIncome"].getAttribute("class") == "warning")
				document.ER["BalanceProgramIncome"].setAttribute("class", "");
			if (document.ER["ExpendedThisQuarter"].getAttribute("class") == "warning")
				document.ER["ExpendedThisQuarter"].setAttribute("class", "");
		}
		return passed;
	}

	function submitForm(action) {
		if (validateForm() == false) {
			return false;
		}
		if (action == "submit") {
			if (document.ER["Confirmed"].checked == false) {
				alert("You must read certification and check box to submit expenditure report.");
				return false;
			}
		}
		if (getNumericValue(document.ER.BalanceProgramIncome.value) < -1000.0) {
			alert("You have entered an amount in program income totaling greater than $1,000 over the MVCPA Grant Budget. Please contact the grant coordinator.");
			return false;
		}
		document.ER.Action.value = action;
		document.ER.submit();
	}

	function validateForm(action)
	{
		if (checkLimits() == false) {
			alert("Some expenditure categories exceeded the limits established by the grant award statement. (See figures in yellow)\n\nThe reimbursable expenditures used to calculate the reimbursement will be less than the actual total cash expenditures.");
		}
		for (var i = 1; i < 8; i++) {
			if (checkCurrency(document.ER["CashExpenditure_"+i]) == false)
				return false;
		}
		if (checkCurrency(document.ER["InLieuOfDPS"]) == false)
			return false;
		if (checkCurrency(document.ER["InLieuOfNICB"]) == false)
			return false;
		if (checkCurrency(document.ER["BeginningBalance"]) == false)
			return false;
		if (checkCurrency(document.ER["EarnedThisQuarter"]) == false)
			return false;
		if (checkCurrency(document.ER["ExpendedThisQuarter"]) == false)
			return false;
		if (checkCurrency(document.ER["BPOMVCPA"]) == false)
			return false;
		if (checkCurrency(document.ER["BPOLocal"]) == false)
			return false;
		if (checkCurrency(document.ER["BPOPI"]) == false)
			return false;
	<%	If FiscalYear=2020 Or FiscalYear=2021 Then %>
		if (checkCurrency(document.ER["COVIDMVCPA"]) == false)
			return false;
		if (checkCurrency(document.ER["COVIDLocal"]) == false)
			return false;
		if (checkCurrency(document.ER["COVIDPI"]) == false)
			return false;
	<%	End If %>
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body onload="updateTotals();">

<div class="sectiontitle">Motor Vehicle Crime Prevention Authority</div>
<div class="sectiontitle">FY <%=FiscalYear %> Quarterly Expenditure Report</div>
<%
If SubmitID > 0 Then
	Response.Write("<div style=""text-align: center;"">Submitted by " & SubmitName & ", " & SubmitTimeStamp & "</div>" & vbCrLf)
End If
If IsNull(DatePaid) = False Then ' Display paid date if any.
	Response.Write("<div style=""text-align: center;"">Date Paid: " & formatDateTime(DatePaid,vbShortDate) & "</div>" & vbCrLf)
End If
If PriorNonApproval = True Then
	Response.Write("<div style=""text-align: center; color: red;"">There is a prior Expenditure Report that has not been approved. This will prevent submission.</div>" & vbCrLf)
End If
%>
<br />
<form name="ER" id="ER" method="post" action="ReportSubmit.asp">
<%=HiddenField("GrantID",GrantID) %><%=HiddenField("LoadedQuarter",Quarter) %><%=HiddenField("Action","save") %><%=HiddenField("PriorYearAllocation", PriorYearAllocation) %><%=HiddenField("AwardAmount",AwardAmount) %>
<table>
	<tr><td>Grantee Name</td><td><%=GranteeName %></td></tr>
	<tr><td>Program Name</td><td><%=ProgramName %></td></tr>
	<tr><td>Grant Number</td><td><%=GrantNumber %></td></tr>
	<tr><td>Fiscal Year</td><td><%=FiscalYear %></td></tr>
	<tr><td>Grant Award Amount</td><td><%=prepCurrencyWeb(AwardAmount) %></td></tr>
	<tr><td></td><td></td></tr>
	<tr><td>Quarter:</td><td><select name="ReportingPeriod" onchange="location.href='Report.asp?GrantID=<%=GrantID%>&Quarter='+this.options[this.selectedIndex].value;">
<%
	'If Quarter > 0 Then
	'	Response.Write(vbTab & SelectOption(Quarter, ReportingPeriodDates(FiscalYear, Quarter), Quarter))
	'Else
		for i = 1 to MAXQUARTER
			Response.Write(vbTab & SelectOption(i, ReportingPeriodDates(FiscalYear, i), Quarter))
		next
	'End If
%></select></td></tr>

</table>
<br />
<div class="singleborder">
<table>
	<caption>Expenditures by Category</caption>
	<thead>
	<tr>
		<th rowspan="2">Budget Category</th>
		<th colspan="3">Quarterly Expenditures</th>
		<th colspan="3">Year to Date</th>
	</tr>
	<tr>
		<th>Total Cash Expenses:<br />MVCPA & Match</th>
		<th>Excluded Amount</th>
		<th>In-Kind<br />Expenditures</th>
		<th>YTD<br />Expenditures<br /><div class="detailnote">(less excluded)</div></th>
		<th>Total<br />Grant<br />Budget</th>
		<th>Grant<br />Budget<br />Remaining</th>
	</tr>
	</thead>
	<tbody>
<%
rs.MoveFirst
While rs.EOF = False
	Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
	If rs.Fields("BudgetCategoryID") = 6 Then
		Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("CashExpenditure_" & rs.Fields("BudgetCategoryID"), rs.Fields("CashExpenditure"), 12, 15, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
	Else
		Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("CashExpenditure_" & rs.Fields("BudgetCategoryID"), rs.Fields("CashExpenditure"), 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
	End If
	Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("ExcludedAmount_" & rs.Fields("BudgetCategoryID"), rs.Fields("ExcludedAmount"), 12, 15, False, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("InKindExpenditure_" & rs.Fields("BudgetCategoryID"), rs.Fields("InKindExpenditure"), 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("YTDExpenditure_" & rs.Fields("BudgetCategoryID"), rs.Fields("YTDExpenditure"), 12, 15, False, "") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("BudgetExpenditure_" & rs.Fields("BudgetCategoryID"), rs.Fields("BudgetExpenditure"), 12, 15, False, "") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("RemainingBudget_" & rs.Fields("BudgetCategoryID"), rs.Fields("RemainingBudget"), 12, 15, False, "") & "</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	rs.MoveNext
Wend
Response.Write("<tr>" & vbCrLf)
Response.Write(vbTab & "<td style=""font-weight: bold; "">Totals</td>" & vbCrLf)
Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("CashExpenditure_Total", 0, 12, 15, False, "") & "</td>" & vbCrLf)
Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("ExcludedAmount_Total", 0, 12, 15, False, "") & "</td>" & vbCrLf)
Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("InKindExpenditure_Total", 0, 12, 15, False, "") & "</td>" & vbCrLf)
Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("YTDExpenditure_Total", 0, 12, 15, False, "") & "</td>" & vbCrLf)
Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("BudgetExpenditure_Total", 0, 12, 15, False, "") & "</td>" & vbCrLf)
Response.Write(vbTab & "<td style=""text-align: center; "">" & CurrencyField("RemainingBudget_Total", 0, 12, 15, False, "") & "</td>" & vbCrLf)
Response.Write("</tr>" & vbCrLf)
'Response.Write("</table>" & vbCrLf)
%>	</tbody>
</table>
<br />
<table style="margin: auto; ">
	<thead>
		<tr><th colspan="7">Equipment Detail</th></tr>
		<tr>
			<th colspan="2">Asset Class</th>
			<th>Description</th>
			<th>Serial No. / VIN</th>
			<th>Date</th>
			<th>Cost</th>
			<th>Use</th>
		</tr>
	</thead>
	<tbody>
<%
sql = "SELECT A.EquipmentID, A.AssetClassID, B.AssetClassShort, " & vbCrLf & _
	"	A.ItemDescription + CASE WHEN ModelYear IS NOT NULL THEN ', ' + CAST(ModelYear AS VARCHAR) ELSE '' END + " & vbCrLf & _
	"	CASE WHEN MakeManufacturer IS NOT NULL THEN ', ' + MakeManufacturer ELSE '' END + " & vbCrLf & _
	"	CASE WHEN Model IS NOT NULL THEN ', ' + Model ELSE '' END AS Description, " & vbCrLf & _
	"	A.SerialNo, CONVERT(VARCHAR,AcquisitionDate,1) AS AcquisitionDate, A.Cost, C.[Use]  " & vbCrLf & _
	"FROM ER.EquipmentDetail AS A " & vbCrLf & _
	"LEFT JOIN Lookup.InventoryAssetClass AS B ON B.AssetClassID=A.AssetClassID " & vbCrLf & _
	"LEFT JOIN Lookup.InventoryUse AS C ON C.UseID=A.UseID " & vbCrLf & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"ORDER BY AssetClassID, EquipmentID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("<tr><td colspan=""7"" style=""text-align: center"">No Equipment Detail Records Found</td></tr>" & vbCrLf)
Else
	'If Quarter = 4 Then
	'	Response.Write("<tr><td colspan=""7"">The following items were added after certification because reimbursement was made in the 4th quarter expenditure report after certification</td></tr>" & vbCrLf)
	'End If

	While rs.EOF = False
		Response.Write("<tr>" & vbCrLf)
		Response.Write(vbTab & "<td style=""white-space: nowrap; "">" & rs.Fields("AssetClassID") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("AssetClassShort") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("Description") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("SerialNo") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center;"">" & rs.Fields("AcquisitionDate") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("Cost")) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center;"">" & rs.Fields("Use") & "</td>" & vbCrLf)
		Response.Write("<tr>" & vbCrLf)
		rs.MoveNext
	Wend
End If

%>
	</tbody>
<%	If SubmitID = 0 Then %>
	<tfoot>
		<tr><td colspan="6" style="text-align: center;"><a href="Javascript: submitForm('Equipment');" 
			class="plainlink">Add/Edit Equipment</a></td></tr>
	</tfoot>
<%	End If %>
</table>
<br />
<table style="margin: auto; ">
	<caption>Reimbursement Calculation</caption>
	<tr>
		<td>Year To Date Expenditures</td>
		<td style="text-align: center; "><%=CurrencyField("ExpendituresYTD", 0, 12, 15, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>less In lieu of DPS for prior quarters</td>
		<td style="text-align: center; "><%=CurrencyField("PriorInLieuOfDPS2", PriorInLieuOfDPS, 12, 15, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>less In lieu of NICB for prior quarters</td>
		<td style="text-align: center; "><%=CurrencyField("PriorInLieuOfNICB2", PriorInLieuOfNICB, 12, 15, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>less In lieu of DPS for quarter</td>
		<td style="text-align: center; "><%=CurrencyField("InLieuOfDPS", InLieuOfDPS, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>less In lieu of NICB for quarter</td>
		<td style="text-align: center; "><%=CurrencyField("InLieuOfNICB", InLieuOfNICB, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr title="Unbudgeted program income used is not reimbursed and limited to $1000 per fiscal year.">
		<td>less unbudgeted program income used</td>
		<td style="text-align: center; "><%=CurrencyField("UnbudgetedPI", UnbudgetedPI, 12, 15, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Reimbursable Expenditures</td>
		<td style="text-align: center; "><%=CurrencyField("ReimbursableExpenditures", 0, 12, 15, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Reimbursement Rate</td>
		<td style="text-align: center; white-space: nowrap; "><%=NumberField("ReimbursementRate", ReimbursementRate, 12, 15, False, "") %>%</td>
	</tr>
	<tr>
		<td style="white-space: nowrap; ">Reimbursement on YTD Expenditures</td>
		<td style="text-align: center; "><%=CurrencyField("ReimbursementYTD", 0, 12, 15, False, "") %></td>
	</tr>
	<tr>
		<td>less Prior Quarter Payments</td>
		<td style="text-align: center; "><%=CurrencyField("PriorAmountPaid", PriorAmountPaid, 12, 15, False, "") %></td>
	</tr>

	<tr>
		<td>Reimbursement for this quarter</td>
		<td style="text-align: center; "><%=CurrencyField("Reimbursement", Reimbursement, 12, 15, False, "") %></td>
	</tr>
	</tbody>
</table></div>
<br />
<table style="margin: auto; ">
	<caption>Program Income Account Reconciliation</caption>
	<tr>
		<td>Beginning Balance</td>
		<td><%=CurrencyField("BeginningBalance", BeginningBalance, 12, 15, False, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Earned This Quarter</td>
		<td><%=CurrencyField("EarnedThisQuarter", EarnedThisQuarter, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Expended This Quarter</td>
		<td><%=CurrencyField("ExpendedThisQuarter", ExpendedThisQuarter, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Ending Balance</td>
		<td><%=CurrencyField("EndingBalance", EndingBalance, 12, 15, False, "") %></td>
	</tr>
</table>
<br />
<table style="margin: auto; ">
	<caption>Sources of Cash Match Budget Summary</caption>
	<thead>
		<tr>
			<th></th>
			<th>In Lieu of DPS</th>
			<th>In Lieu of NICB</th>
			<th>Program Income</th>
		</tr>
	</thead>
	<tbody>
	<tr>
		<td>Budgeted Amount</td>
		<td><%=CurrencyField("InLieuOfDPSBudget", InLieuOfDPSBudget, 12, 15, False, "") %></td>
		<td><%=CurrencyField("InLieuOfNICBBudget", InLieuOfNICBBudget, 12, 15, False, "") %></td>
		<td><%=CurrencyField("ProgramIncomeBudget", ProgramIncomeBudget, 12, 15, False, "") %></td>
	</tr>
	<tr>
		<td>Prior Recognized Amount</td>
		<td><%=CurrencyField("PriorInLieuOfDPS", PriorInLieuOfDPS, 12, 15, False, "") %></td>
		<td><%=CurrencyField("PriorInLieuOfNICB", PriorInLieuOfNICB, 12, 15, False, "") %></td>
		<td><%=CurrencyField("PriorProgramIncome", PriorProgramIncome, 12, 15, False, "") %></td>
	</tr>
	<tr>
		<td>Current Recognized Amount</td>
		<td><%=CurrencyField("CurrentInLieuOfDPS", (InLieuOfDPS), 12, 15, False, "") %></td>
		<td><%=CurrencyField("CurrentInLieuOfNICB", (InLieuOfNICB), 12, 15, False, "") %></td>
		<td><%=CurrencyField("CurrentProgramIncome", (ExpendedThisQuarter), 12, 15, False, "") %></td>
	</tr>
	<tr>
		<td>Remaining Balance</td>
		<td><%=CurrencyField("BalanceDPS", (InLieuOfDPSBudget-PriorInLieuOfDPS-InLieuOfDPS), 12, 15, False, "") %></td>
		<td><%=CurrencyField("BalanceNICB", (InLieuOfNICBBudget-PriorInLieuOfNICB-InLieuOfNICB), 12, 15, False, "") %></td>
		<td><%=CurrencyField("BalanceProgramIncome", (ProgramIncomeBudget-PriorProgramIncome-ExpendedThisQuarter), 12, 15, False, "") %></td>
	</tr>
	</tbody>
</table>
<br />
<table style="margin: auto; ">
	<caption>Additional Direct Costs for Border / Port Operations</caption>
	<tr>
		<td>MVCPA</td>
		<td><%=CurrencyField("BPOMVCPA", BPOMVCPA, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Local</td>
		<td><%=CurrencyField("BPOLocal", BPOLocal, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Program Income</td>
		<td><%=CurrencyField("BPOPI", BPOPI, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Total</td>
		<td><%=CurrencyField("BPOTotal", BPOTotal, 12, 15, False, "") %></td>
	</tr>
</table>
<hr />
<%
If FiscalYear=2020 Or FiscalYear=2021 Then
%>
<table style="margin: auto; ">
	<caption>Additional Direct Costs for COVID</caption>
	<tr>
		<td>MVCPA</td>
		<td><%=CurrencyField("COVIDMVCPA", COVIDMVCPA, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Local</td>
		<td><%=CurrencyField("COVIDLocal", COVIDLocal, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Program Income</td>
		<td><%=CurrencyField("COVIDPI", COVIDPI, 12, 15, PermitEdit, "checkCurrency(this); updateTotals()") %></td>
	</tr>
	<tr>
		<td>Total</td>
		<td><%=CurrencyField("COVIDTotal", COVIDTotal, 12, 15, False, "") %></td>
	</tr>
	<tr>
		<td title="Please limit to 250 characters." colspan="2">Expense Description: <span style="font-size: smaller; ">(please limit to 250 characters.)</span><br />
		<%=TextArea("COVIDNote", COVIDNote, 2, 60, 250, PermitEdit, "if (this.value.length>250) alert('Please limit length to 250 characters.')") %></td>
	</tr>
</table>
<%
Else
	Response.Write(HiddenField("COVIDMVCPA",COVIDMVCPA))
	Response.Write(HiddenField("COVIDlocal",COVIDlocal))
	Response.Write(HiddenField("COVIDPI",COVIDPI))
	Response.Write(HiddenField("COVIDTotal",COVIDTotal))
	Response.Write(HiddenField("COVIDNote",COVIDNote))
End If
%>
<hr />
<%
If ViewDocuments = True Then
	Dim Folder, file, files, DocumentFolder, fso, counter
	counter=0
	DocumentFolder = Application("DocumentRoot") & "\Grant\" & GrantID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemOBject")
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		If PErmitEdit = True Then
			Response.Write("<tr style=""vertical-align: top; ""><td>Current Documents in folder: <a href=""../Upload/Upload.asp?fid=6&quarter=" & Quarter & "&GrantID=" & GrantID & """ class=""plainlink"" target=""_blank"">Upload</a></td>" & vbCrLf)
		End If
		If files.count>0 Then 
			Response.Write("<tr><td>")
			For Each file in files
				If Left(file.Name,3)="ER"&Quarter Then
					Response.Write("<a href=""../Documents/Grant/" & GrantID & "/" & file.Name & _
						""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
					counter = counter + 1
				End If
			Next
			Response.Write("</td></tr>" & vbCrLf)
		End If
	End If
	If counter = 0 Then
		Response.Write("<tr style=""vertical-align: top; ""><td style=""text-align: center; "">No Documents in folder</td></tr>" & vbCrLf)
	End If
	Response.Write("</table>" & vbCrLf)
	Response.Write("<hr />" & vbCrLf)
End If
%>
<div style="margin: auto; ">
Supplementary Comments. Provide additional informaton that would be helpful in evaluating this expenditure report.<br />
<%=TextArea2("SupplementaryComments", SupplementaryComments, 8, 960, 8000, PermitEdit, "") %>
<p><%
Response.Write(CheckBoxField2("Confirmed", Confirmed, PermitEdit) & vbCrLf)

If FiscalYear >= 2022 Then%>
By signing this report, I certify to the best of my knowledge and belief that the report is true, 
complete, and accurate, and the expenditures, disbursements and cash receipts are for the purposes 
and objectives set forth in the terms and conditions of the state award. I am aware that any false, 
fictitious, or fraudulent information, or the omission of any material fact, may subject me to 
criminal, civil or administrative penalties for fraud, false statements, false claims or otherwise.
<%
Else
%>
I acknowledge that I have reviewed and confirmed the accuracy of the information in this report, 
and I attest that this report is  correct and complete and that the costs incurred as stated 
herein are for allowable purposes as set forth in the Statement of Grant Award and, (1) pursuant 
to 43 TAC &sect; 57.9, I hereby further certify that MVCPA funds have not been used to replace 
state or local funds and (2) any false, fictitious, or fraudulent information provided herein 
may subject me to criminal, civil, or administrative penalties or sanctions. 
<%
End If

If SubmitID > 0 Then
	Response.Write(" <i>Submitted by " & SubmitName & ", " & SubmitTimeStamp & "</i>")
End If
%>
</p>
</div>
<br />
<%

If SubmitID = 0 Then
	CanApprove = False
	CanInvoice = False
Else
	CanApprove = True
	CanInvoice = True ' Keep true if all conditions are passed.
End If

sql = "SELECT * FROM [Grants].vwGrantStatus WHERE Grant_ID=" & prepIntegerSQL(GrantID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	If Not(rs.Fields("ER Q1")="P" Or rs.Fields("ER Q1")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">First Quarter Expenditure Report has not been approved.</div>" & vbCrLf)
		If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		End If
	End If
	If Quarter>=2 And Not(rs.Fields("ER Q2")="P" Or rs.Fields("ER Q2")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">Second Quarter Expenditure Report has not been approved.</div>" & vbCrLf)
		If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		End If
	End If
	If Quarter>=3 And Not(rs.Fields("ER Q3")="P" Or rs.Fields("ER Q3")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">Third Quarter Expenditure Report has not been approved.</div>" & vbCrLf)
		If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		End If
	End If
	If Quarter>=4 And rs.Fields("ER Q4")="" Then
		Response.Write("<div style=""text-align: center; color: red; "">Fourth Quarter Expenditure Report has not been submitted.</div>" & vbCrLf)
		'CanApprove = False
		CanInvoice = False
	End If
	If Quarter>=1 And Not(rs.Fields("PR Q1")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">First Quarter Progress Report has not been approved.</div>" & vbCrLf)
		If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		End If
	End If
	If Quarter>=2 And Not(rs.Fields("PR Q2")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">Second Quarter Progress Report has not been approved.</div>" & vbCrLf)
		If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		End If
	End If
	If Quarter>=3 And Not(rs.Fields("PR Q3")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">Third Quarter Progress Report has not been approved.</div>" & vbCrLf)
		If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		End If
	End If
	If Quarter=4 And Not(rs.Fields("PR Q4")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">Fourth Quarter Progress Report has not been approved.</div>" & vbCrLf)
		'CanApprove = False
		CanInvoice = False
	End If
	If Quarter=4 And Not(rs.Fields("PR YE")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">Fourth Quarter Progress Report has not been approved.</div>" & vbCrLf)
		'CanApprove = False
		CanInvoice = False
	End If
	If  Quarter=4 And Not(rs.Fields("IC_Status")="A") Then
		Response.Write("<div style=""text-align: center; color: red; "">The Inventory Certification has not been approved.</div>" & vbCrLf)
		'CanApprove = False
		CanInvoice = False
	End If
	If CanApprove = False Then
		Response.Write("<div style=""text-align: center; color: red; "">There are outstanding requirements and this Expenditure Report is not ready to be approved.</div>" & vbCrLf)
	End If
'	If ReviewID=0 Then
'		If CanApprove = True Then
'			Response.Write("<div style=""text-align: center; color: red; "">The expenditure report has not been approved and cannot be invoiced.</div>" & vbCrLf)
'		End If
'		CanInvoice = False
'	End If
	If DirectorApprovalID=0 Then
		If CanApprove = True Then
			Response.Write("<div style=""text-align: center; color: red; "">The expenditure report has not received director approval and cannot be invoiced.</div>" & vbCrLf)
		End If
		'If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		'End If
	End If
	If PRSubmitted = False Then
		'If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		'End If
	End If
	If PRApproved = False Then
		'If Quarter=4 Then
			'CanApprove = False
			CanInvoice = False
		'End If
	End If
	If PriorNonApproval = True Then
		Response.Write("<div style=""text-align: center; color: red;"">There is a prior Expenditure Report that has not been approved. This will prevent submission.</div>" & vbCrLf)
	End If
	If LaterSubmission = True And DirectorApprovalID=0 Then
		Response.Write("<div style=""text-align: center; color: red;"">Warning: A later Expenditure Report has been submitted.</div>" & vbCrLf)
	End If
Else
	Response.Write("<div style=""text-align: center; color: red; "">Grant Not found in Status Query.</div>" & vbCrLf)
	CanApprove = False
	CanInvoice = False
End If


If MVCPARights = True Or MVCPAViewer = True Then
	Response.Write(HiddenField("ReviewID", ReviewID))
	Response.Write(HiddenField("AuditApprovalID", AuditApprovalID))
	Response.Write(HiddenField("DirectorApprovalID", DirectorApprovalID))
	Response.Write("<table style=""margin: auto; "">" & vbCrLf)
	Response.Write("<tr><td colspan=""2"">Administrative Comments:<br />" & vbCrLf)
	Response.Write(TextArea2("AdministrativeComments", AdministrativeComments, 6, 920, 8000, MVCPARights, "") & "</td></tr>" & vbCrLf)
	If SubmitID > 0 Then
		Response.Write("<tr><td>MVCPA Grant Coordinator Review Date:</td>" & vbCrLf)
		If ReviewID > 0 And IsNull(DirectorApprovalDate) = False And CanApprove = True Then
			Response.Write("<td>" & DateField("ReviewDate", ReviewDate, False))
		ElseIf IsNull(DirectorApprovalDate) = False And CanApprove = True Then
			Response.Write("<td>" & DateField("ReviewDate", ReviewDate, True))
		Else
			Response.Write("<td>" & DateField("ReviewDate", ReviewDate, MVCPARights))
		End If
		If IsNull(ReviewName) = False Then
			Response.Write(" by " & ReviewName)
		End If
		Response.Write("<tr><td>MVCPA Audit Approval Date:</td>" & vbCrLf)
		If IsNull(DirectorApprovalDate) = False And CanApprove = True Then ' no edit once director approval
			Response.Write("<td>" & DateField("AuditApprovalDate", AuditApprovalDate, False))
		Else
			Response.Write("<td>" & DateField("AuditApprovalDate", AuditApprovalDate, MVCPAAuditor))
		End If
		If IsNull(AuditApprovalName) = False Then
			Response.Write(" by " & AuditApprovalName)
		End If
		Response.Write("<tr><td>MVCPA Director Approval Date:</td>" & vbCrLf)
		If IsNull(DirectorApprovalDate) = False And CanApprove = True Then ' no edit once director approval
			Response.Write("<td>" & DateField("DirectorApprovalDate", DirectorApprovalDate, False))
		Else
			Response.Write("<td>" & DateField("DirectorApprovalDate", DirectorApprovalDate, MVCPAAdministrator))
		End If
		If IsNull(DirectorApprovalName) = False Then
			Response.Write(" by " & DirectorApprovalName)
		End If
		Response.Write("</td>")
		Response.Write("<tr><td>Amount Paid</td><td>" & CurrencyField("AmountPaid", AmountPaid, 12, 15, False, "checkCurrency(this); updateTotals();") & "</td>" & vbCrLf)
		Response.Write("<tr><td>&nbsp;&nbsp;From Prior Year Appropriation (" & prepCurrencyWeb(PriorYearAllocation) & ")</td><td>" & CurrencyField("PriorYearFunds", PriorYearFunds, 12, 15, False, "checkCurrency(this);") & "</td>" & vbCrLf)
		Response.Write("<tr><td>&nbsp;&nbsp;From Current Year Appropriation</td><td>" & CurrencyField("CurrentYearFunds", CurrentYearFunds, 12, 15, False, "checkCurrency(this);") & "</td>" & vbCrLf)
		If IsNull(DirectorApprovalDate) = True and IsNull(DatePaid)=True Then
			' Do not show date paid.
			Response.Write(HiddenField("DatePaid", DatePaid))
		ElseIf IsNull(DatePaid) = False Then ' no edit once there is a value.
			Response.Write("<tr><td>Date Paid</td><td>" & DateField("DatePaid", DatePaid, False))
		Else
			Response.Write("<tr><td>Date Paid</td><td>" & DateField("DatePaid", DatePaid, MVCPARights))
		End If
		'If ReviewID>0 And AuditApprovalID>0 And DirectorApprovalID>0 And PRSubmitted=True AND PRApproved=True Then
		If CanInvoice = True Then
			Response.Write(" <a href=""Voucher.asp?GrantID=" & GrantID & "&Quarter=" & Quarter & """ target=""_blank"">invoice</a>" & vbCrLf)
		End If
		Response.Write("</td></tr>" & vbCrLf)

		If IsNull(DirectorApprovalDate) = True Then ' no edit once director approval
			Response.Write("<tr><td colspan=""2"">" &  CheckBoxField("Unsubmit", False) & " Unsubmit Expenditure Report (Clears submission and approval.)</td></tr>" & vbCrLf)
		Else
			Response.Write("<tr><td colspan=""2"">This Expenditure Report may not be unsubmitted because it has already received director approval.</td></tr>" & vbCrLf)
		End If
		Response.Write("</table>" & vbCrLf)
	Else
		Response.Write("</table>" & vbCrLf)
		Response.Write(HiddenField("ReviewDate", ReviewDate))
		Response.Write(HiddenField("AuditApprovalDate", AuditApprovalDate))
		Response.Write(HiddenField("DirectorApprovalDate", DirectorApprovalDate))
		Response.Write(HiddenField("AmountPaid", AmountPaid))
		Response.Write(HiddenField("DatePaid", DatePaid))
		Response.Write(HiddenField("PriorYearFunds", PriorYearFunds))
		Response.Write(HiddenField("CurrentYearFunds", CurrentYearFunds))
	End If
End If

If MVCPARights = True Or MVCPAViewer = True Then
	sql = "SELECT A.SubmitTimestamp, B.Name AS SubmitName " & vbCrLf & _
		"FROM ER.Main AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT A.SubmitTimestamp, B.Name AS SubmitName " & vbCrLf & _
		"FROM ER.Main_Log AS A " & vbCrLf & _
		"LEFT JOIN [System].Users AS B ON A.SubmitID=B.SystemID " & vbCrLf & _
		"WHERE GrantID=" & prepIntegerSQL(GrantID) & " AND Quarter=" & prepIntegerSQL(Quarter) & " AND SubmitTimestamp IS NOT NULL " & vbCrLf & _
		"ORDER BY 1 DESC "
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then
		Response.Write("<br><table style=""margin: auto; "">" & vbCrLf)
		Response.WRite("<thead><tr><th>Submission History</th></tr></thead>" & vbCrLf & "<tbody>" & vbCrLf)
		While rs.EOF = False
			Response.Write("<tr><td style=""text-align: center"">Submitted By " & rs.Fields("SubmitName") & ", " & rs.Fields("SubmitTimestamp") & "</td></tr>" & vbCrLf)
			rs.MoveNext()
		Wend
		Response.Write("</tbody>" & vbCrLf & "</table><br />" & vbCrLf)
	End If
End If
'CanSubmit = True
If Debug = True Then
	Response.Write("<pre>")
	Response.Write("PermitEdit=" & PermitEdit & vbCrLf)
	Response.Write("CanSubmit=" & CanSubmit & vbCrLf)
	Response.Write("PriorNonApproval=" & PriorNonApproval & vbCrLf)
	Response.Write("LaterSubmission=" & LaterSubmission & vbCrLf)
	Response.Write("</pre>")
End If
%>
<div style="text-align: center">
<%	If PermitEdit = True or SubmitID = 0 Or MVCPARights = True Then %>
<input type="button" name="Save" value="Save" onclick="return submitForm('save');" />
<%	
	End If	
	If SubmitID = 0 And CanSubmit=True And PriorNonApproval=False And Date()>=StartDate  Then 
%>
<input type="button" name="Submit" value="Submit" onclick="return submitForm('submit');" />
<%	 
	End If 
%>
<input type="reset" name="Reset" value="Reset" />
<input type="button" value="Close" onclick="window.close();" />
</div>

</form>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<%
Function ReportingPeriodDates(vFiscalYear, vReportingPeriod)
	If vReportingPeriod = 1 Then
		ReportingPeriodDates = "September 1, " & (vFiscalYear-1) & " - November 30, " & (vFiscalYear-1)
	ElseIf vReportingPeriod = 2 and (vFiscalYear Mod 4 = 0) Then
		ReportingPeriodDates = "December 1, " & (vFiscalYear-1) & " - February 29, " & vFiscalYear
	ElseIf vReportingPeriod = 2 Then
		ReportingPeriodDates = "December 1, " & (vFiscalYear-1) & " - February 28, " & vFiscalYear
	ElseIf vReportingPeriod = 3 Then
		ReportingPeriodDates = "March 1, " & vFiscalYear & " - May 31, " & vFiscalYear
	ElseIf vReportingPeriod = 4 Then
		ReportingPeriodDates = "June 1, " & vFiscalYear & " - August 31, " & vFiscalYear
	Else
		ReportingPeriodDates = "Error in Reporting Period"
	End If
End Function

Function ReportingPeriodForDate(vDate)
	Select Case Month(vDate)
		Case 1, 2 
			ReportingPeriodForDate = 2
		Case 3, 4, 5
			ReportingPeriodForDate = 3
		Case 6, 7, 8
			ReportingPeriodForDate = 4
		Case 9, 10, 11
			ReportingPeriodForDate = 1
		Case 12
			ReportingPeriodForDate = 2
	End Select
End Function
%>
