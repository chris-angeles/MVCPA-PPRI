<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, Quarter, FirstQuarter, OrderByDescription, QuarterDescription, OrderByField, _
	ShowExcel, ShowExcluded, ShowYTD, ShowInKind, ShowBPO, Show, ShowDescription, ShowClause, _
	ShowCOVID, ShowRemaining, ShowPercent, ShowCashMatch
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number")
QuarterDescription = Array("", "September 1 - November 30","December 1 - February 28", "March 1 - May 31", "June 1 - August 31")
OrderByField = Array("GrantID", "GranteeSort", "Grant_Number")
ShowDescription = Array ("All", "Border", "Port", "Port 2", "Border and Port", "Border, Port, and Port 2")
ShowClause = Array ("1=1", "A.BorderCounty=1", "A.PortCounty=1", "A.Port2County=1", "(A.BorderCounty=1 OR A.PortCounty=1)", "(A.BorderCounty=1 OR A.PortCounty=1 OR A.Port2County=1)")
debug = False

If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("<pre>Session(""" & i & """)='" & Session(i) & "'</pre>" & vbCrLf)
	Next
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	If Month(Date()) > 9 Then
		FiscalYear = Year(Date)+1
	Else
		FiscalYear = Year(Date)
	End If
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.Querystring("OrderBy"))>0 Then
	OrderBy = CInt(Request.Querystring("OrderBy"))
End If

If Len(Request.Form("Quarter"))>0 Then
	Quarter = CInt(Request.Form("Quarter"))
ElseIf Len(Request.QueryString("Quarter"))>0 Then
	Quarter = CInt(Request.QueryString("Quarter"))
Else
	Quarter = 1
End If

If Request.Form("ShowExcluded")="1" Then
	ShowExcluded = True
ElseIf Request.QueryString("ShowExcluded")="1" Then
	ShowExcluded = True
Else
	ShowExcluded = False
End If

If Request.Form("ShowYTD")="1" Then
	ShowYTD = True
	FirstQuarter = 1
ElseIf Request.QueryString("ShowYTD")="1" Then
	ShowYTD = True
	FirstQuarter = 1
Else
	ShowYTD = False
	FirstQuarter = Quarter
End If

If Request.Form("ShowInKind")="1" Then
	ShowInKind = True
ElseIf Request.QueryString("ShowInKind")="1" Then
	ShowInKind = True
Else
	ShowInKind = False
End If

If Len(Request.Form("ShowBPO"))="1" Then
	ShowBPO = True
ElseIf Request.QueryString("ShowBPO")="1" Then
	ShowBPO = True
Else
	ShowBPO = False
End If

If Len(Request.Form("Show")) > 0 Then
	Show = CInt(Request.Form("Show"))
ElseIf Len(Request.QueryString("Show")) > 0 Then
	Show = CInt(Request.QueryString("Show"))
Else
	Show = 0
End If

If Len(Request.Form("ShowCOVID"))="1" Then
	ShowCOVID = True
ElseIf Request.QueryString("ShowCOVID")="1" Then
	ShowCOVID = True
Else
	ShowCOVID = False
End If

If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

If Len(Request.Form("ShowRemaining"))="1" Then
	ShowRemaining = True
ElseIf Request.QueryString("ShowRemaining")="1" Then
	ShowRemaining = True
Else
	ShowRemaining = False
End If

If Len(Request.Form("ShowPercent"))="1" Then
	ShowPercent = True
ElseIf Request.QueryString("ShowPercent")="1" Then
	ShowPercent = True
Else
	ShowPercent = False
End If

If Len(Request.Form("ShowCashMatch"))="1" Then
	ShowCashMatch = True
ElseIf Request.QueryString("ShowCashMatch")="1" Then
	ShowCashMatch = True
Else
	ShowCashMatch = False
End If

If ShowRemaining = True Then
	ShowBPO = False
	ShowCOVID = False
	ShowYTD = True
End If

sql = "DECLARE @Quarter AS INT=" & Quarter & "; " & vbCrLf & _
	"DECLARE @StartQuarter AS INT=" & FirstQuarter & "; " & vbCrLf & _
	"DECLARE @FiscalYear AS INT=" & FiscalYear & "; " & vbCrLf & _
	"WITH CTE AS " & vbCrLf & _
	"( " & vbCrLf & _
	"	SELECT A.GranteeID, B.GrantID, @Quarter AS Quarter, " & vbCrLf & _
	"		B.FiscalYear AS Fiscal_Year, A.GranteeName AS Grantee_Name, B.ProgramName AS Program_Name,  " & vbCrLf & _
	"		B.GrantNumber AS Grant_Number, " & vbCrLf & _
	"		D.Personnel, D.Fringe, D.Overtime, D.Professional_And_Contract_Services,   " & vbCrLf & _
	"		D.Travel, D.Equipment, D.Supplies_And_DOE, D.Total_Expenditures,  Total_Excluded,  " & vbCrLf & _
	"		C.Reimbursement, D.In_Kind_Expenditure, " & vbCrLf & _
	"		E.Personnel_Budget, E.Fringe_Budget, E.Overtime_Budget, E.Professional_And_Contract_Services_Budget, " & vbCrLf & _
	"		E.Travel_Budget, E.Equipment_Budget, E.Supplies_And_DOE_Budget, E.Total_Budget, " & vbCrLf & _
	"		E.MVCPA_Budget, E.In_Kind_Expenditure_Budget, " & vbCrLf & _
	"		Personnel_Remaining = CASE WHEN E.Personnel_Budget IS NULL AND D.Personnel IS NULL THEN NULL ELSE " & vbCrLF & _
	"			ISNULL(E.Personnel_Budget,0.0) - ISNULL(D.Personnel,0.0) END, " & vbCrLf & _
	"		Fringe_Remaining = CASE WHEN E.Fringe_Budget IS NULL AND D.Fringe IS NULL THEN NULL ELSE " & vbCrLF & _
	"			ISNULL(E.Fringe_Budget,0.0) - ISNULL(D.Fringe,0.0) END, " & vbCrLf & _
	"		Overtime_Remaining = CASE WHEN E.Overtime_Budget IS NULL AND D.Overtime IS NULL THEN NULL ELSE " & vbCrLf & _
	"			ISNULL(E.Overtime_Budget,0.0) - ISNULL(D.Overtime,0.0) END, " & vbCrLf & _
	"		Professional_And_Contract_Services_Remaining = CASE WHEN E.Professional_And_Contract_Services_Budget IS NULL AND D.Professional_And_Contract_Services IS NULL THEN NULL ELSE " & vbCrLF & _
	"			ISNULL(E.Professional_And_Contract_Services_Budget,0.0) - ISNULL(D.Professional_And_Contract_Services,0.0) END, " & vbCrLf & _
	"		Travel_Remaining = CASE WHEN E.Travel_Budget IS NULL AND D.Travel IS NULL THEN NULL ELSE " & vbCrLf & _
	"			ISNULL(E.Travel_Budget,0.0) - ISNULL(D.Travel,0.0) END, " & vbCrLf & _
	"		Equipment_Remaining = CASE WHEN E.Equipment_Budget IS NULL AND D.Equipment IS NULL THEN NULL ELSE " & vbCrLf & _
	"			ISNULL(E.Equipment_Budget,0.0) - ISNULL(D.Equipment,0.0) END, " & vbCrLf & _
	"		Supplies_And_DOE_Remaining = CASE WHEN E.Supplies_And_DOE_Budget IS NULL AND D.Supplies_And_DOE IS NULL THEN NULL ELSE " & vbCrLf & _
	"			ISNULL(E.Supplies_And_DOE_Budget,0.0) - ISNULL(D.Supplies_And_DOE,0.0) END, " & vbCrLf & _
	"		Total_Remaining = ISNULL(E.Total_Budget,0.0) - ISNULL(D.Total_Expenditures,0.0), " & vbCrLf & _
	"		Reimbursement_Remaining = ISNULL(E.MVCPA_Budget,0.0) - ISNULL(C.Reimbursement,0.0), " & vbCrLf & _
	"		Cash_Match, Cash_Match_Percentage, " & vbCrLf & _
	"		In_Kind_Remaining = CASE WHEN E.In_Kind_Expenditure_Budget IS NULL AND D.In_Kind_Expenditure IS NULL THEN NULL ELSE " & vbCrLf & _
	"			ISNULL(E.In_Kind_Expenditure_Budget,0.0) - ISNULL(D.In_Kind_Expenditure,0.0) END, " & vbCrLf & _
	"		BPO_MVCPA, BPO_Local, BPO_PI, BPO_Total, " & vbCrLf & _
	"		COVID_MVCPA, COVID_Local, COVID_PI, COVID_Total " & vbCrLf & _
	"	FROM Grantees AS A   " & vbCrLf & _
	"	LEFT JOIN [Grants].Main AS B ON B.GranteeID=A.GranteeID  " & vbCrLf & _
	"	LEFT JOIN ( " & vbCrLf & _
	"		SELECT GrantID, SUM(Reimbursement) AS Reimbursement,  " & vbCrLf & _
	"			SUM(BPOMVCPA) AS BPO_MVCPA, SUM(BPOLocal) AS BPO_Local, SUM(BPOPI) AS BPO_PI,  " & vbCrLf & _
	"			SUM(CASE WHEN BPOMVCPA IS NULL AND BPOLocal IS NULL AND BPOPI IS NULL THEN NULL  " & vbCrLf & _
	"			ELSE ISNULL(BPOMVCPA,0.0)+ISNULL(BPOLocal,0.0)+ISNULL(BPOPI,0.0) END) AS BPO_Total,  " & vbCrLf & _
	"			SUM(COVIDMVCPA) AS COVID_MVCPA, SUM(COVIDLocal) AS COVID_Local, SUM(COVIDPI) AS COVID_PI,  " & vbCrLf & _
	"			SUM(CASE WHEN COVIDMVCPA IS NULL AND COVIDLocal IS NULL AND COVIDPI IS NULL THEN NULL  " & vbCrLf & _
	"			ELSE ISNULL(COVIDMVCPA,0.0)+ISNULL(COVIDLocal,0.0)+ISNULL(COVIDPI,0.0) END) AS COVID_Total  " & vbCrLf & _
	"		FROM ER.Main  " & vbCrLf & _
	"		WHERE Quarter>=@StartQuarter AND Quarter<=@Quarter " & vbCrLf & _
	"		GROUP BY GrantID  " & vbCrLf & _
	"	) AS C ON C.GrantID=B.GrantID " & vbCrLf & _
	"	LEFT JOIN ( " & vbCrLf & _
	"		SELECT A.GrantID, MAX(A.Quarter) AS Quarter,  " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=1 THEN CashExpenditure-ISNULL(ExcludedAmount,0) ELSE null END) AS Personnel,  " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=2 THEN CashExpenditure-ISNULL(ExcludedAmount,0) ELSE null END) AS Fringe,  " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=3 THEN CashExpenditure-ISNULL(ExcludedAmount,0) ELSE null END) AS Overtime,  " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=4 THEN CashExpenditure-ISNULL(ExcludedAmount,0) ELSE null END) AS Professional_And_Contract_Services,  " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=5 THEN CashExpenditure-ISNULL(ExcludedAmount,0) ELSE null END) AS Travel,  " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=6 THEN CashExpenditure-ISNULL(ExcludedAmount,0) ELSE null END) AS Equipment,  " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=7 THEN CashExpenditure-ISNULL(ExcludedAmount,0) ELSE null END) AS Supplies_And_DOE,  " & vbCrLf & _
	"			SUM(ExcludedAmount) AS Total_Excluded,  " & vbCrLf & _
	"			SUM(CashExpenditure-ISNULL(ExcludedAmount,0)) AS Total_Expenditures,  " & vbCrLf & _
	"			SUM(InKindExpenditure) AS In_Kind_Expenditure " & vbCrLf & _
	"		FROM ER.Detail AS A  " & vbCrLf & _
	"		JOIN ER.Main AS B ON A.GrantID=B.GrantID AND A.Quarter=B.Quarter " & vbCrLf & _
	"		WHERE A.Quarter>=@StartQuarter AND A.Quarter<=@Quarter AND B.SubmitID>0 " & vbCrLf & _
	"		GROUP BY A.GrantID " & vbCrLf & _
	"	) AS D ON D.GrantID = B.GrantID " & vbCrLf & _
	"	LEFT JOIN ( " & vbCrLf & _
	"		SELECT GrantID, " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=1 THEN TotalExpenditures ELSE null END) AS Personnel_Budget, " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=2 THEN TotalExpenditures ELSE null END) AS Fringe_Budget, " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=3 THEN TotalExpenditures ELSE null END) AS Overtime_Budget, " & vbCrLf & _ 
	"			SUM(CASE WHEN BudgetCategoryID=4 THEN TotalExpenditures ELSE null END) AS Professional_And_Contract_Services_Budget, " & vbCrLf & _
	"			SUM(CASE WHEN BudgetCategoryID=5 THEN TotalExpenditures ELSE null END) AS Travel_Budget, " & vbCrLf & _ 
	"			SUM(CASE WHEN BudgetCategoryID=6 THEN TotalExpenditures ELSE null END) AS Equipment_Budget, " & vbCrLf & _ 
	"			SUM(CASE WHEN BudgetCategoryID=7 THEN TotalExpenditures ELSE null END) AS Supplies_And_DOE_Budget, " & vbCrLf & _ 
	"			SUM(TotalExpenditures) AS Total_Budget, " & vbCrLf & _ 
	"			SUM(MVCPAExpenditures) AS MVCPA_Budget, " & vbCrLf & _ 
	"			SUM(InKindExpenditures) AS In_Kind_Expenditure_Budget " & vbCrLf & _ 
	"		FROM [Grants].Budget " & vbCrLf & _ 
	"		GROUP BY GrantID " & vbCrLf & _ 
	"	) AS E ON E.GrantID=B.GrantID " & vbCrLf & _
	"	LEFT JOIN (" & vbCrLf & _
	"		SELECT GrantID, Quarter, ReimbursableExpenditures AS Reimbursable_Expenditures, " & vbCrLF & _
	"			ReimbursementYTD AS Reimbursement, " & vbCrLf & _
	"			ISNULL(ReimbursableExpenditures,0.0)-ISNULL(ReimbursementYTD,0.0) AS Cash_Match, " & vbCrLf & _
	"			CASE WHEN ISNULL(ReimbursementYTD,0.0) > 0 THEN (ISNULL(ReimbursableExpenditures,0.0)-ISNULL(ReimbursementYTD,0.0)) / ISNULL(ReimbursementYTD,0.0)" & vbCrLf & _
	"			ELSE NULL END AS Cash_Match_Percentage " & vbCrLf & _
	"		FROM ER.Main) AS F ON F.GrantID=B.GrantID AND F.Quarter=@Quarter " & vbCrLf & _
	"	WHERE B.FiscalYear=@FiscalYear AND " & ShowClause(Show) & " " & vbCrLf & _
	") " & vbCrLf & _
	"SELECT GranteeID, GrantID, Quarter, Fiscal_Year, Grantee_Name, Program_Name, Grant_Number,  " & vbCrLf
	If ShowRemaining = False And ShowPercent = False Then
		sql = sql & "	Personnel, Fringe, Overtime, Professional_And_Contract_Services, Travel, Equipment, Supplies_And_DOE, Total_Expenditures,  " & vbCrLf
	ElseIf ShowRemaining = False And ShowPercent = True Then
		sql = sql & "	CASE WHEN Personnel_Budget<>0 THEN Personnel/Personnel_Budget ELSE NULL END AS Personnel_Percent, " & vbCrLf & _
			"	CASE WHEN Fringe_Budget<>0 THEN Fringe/Fringe_Budget ELSE NULL END AS Fringe_Percent, " & vbCrLf & _
			"	CASE WHEN Overtime_Budget<>0 THEN Overtime/Overtime_Budget ELSE NULL END AS Overtime_Percent,  " & vbCrLf & _
			"	CASE WHEN Professional_And_Contract_Services_Budget<>0 THEN Professional_And_Contract_Services/Professional_And_Contract_Services_Budget ELSE NULL END AS Professional_And_Contract_Services_Percent, " & vbCrLf & _
			"	CASE WHEN Travel_Budget<>0 THEN Travel/Travel_Budget ELSE NULL END AS Travel_Percent,  " & vbCrLf & _
			"	CASE WHEN Equipment_Budget<>0 THEN Equipment/Equipment_Budget ELSE NULL END  AS Equipment_Percent, " & vbCrLf & _
			"	CASE WHEN Supplies_And_DOE_Budget<>0 THEN Supplies_And_DOE/Supplies_And_DOE_Budget ELSE NULL END AS Supplies_And_DOE_Percent, " & vbCrLf & _
			"	CASE WHEN Total_Budget <>0 THEN Total_Expenditures/Total_Budget ELSE NULL END AS Total_Expenditures_Percent, " & vbCrLf
	ElseIf ShowRemaining = True And ShowPercent = False Then
		sql = sql & "	Personnel_Remaining, Fringe_Remaining, Overtime_Remaining, " & vbCrLf & _
			"	Professional_And_Contract_Services_Remaining, Travel_Remaining, Equipment_Remaining, " & vbCrLf & _
			"	Supplies_And_DOE_Remaining, Total_Remaining,  " & vbCrLf
	ElseIf ShowRemaining = True And ShowPercent = True Then
		sql = sql & "	CASE WHEN Personnel_Budget<>0 THEN Personnel_Remaining/Personnel_Budget ELSE NULL END AS Personnel_Remaining_Percent, " & vbCrLf & _
			"	CASE WHEN Fringe_Budget<>0 THEN Fringe_Remaining/Fringe_Budget ELSE NULL END AS Fringe_Remaining_Percent, " & vbCrLf & _
			"	CASE WHEN Overtime_Budget<>0 THEN Overtime_Remaining/Overtime_Budget ELSE NULL END AS Overtime_Remaining_Percent,  " & vbCrLf & _
			"	CASE WHEN Professional_And_Contract_Services_Budget<>0 THEN Professional_And_Contract_Services_Remaining/Professional_And_Contract_Services_Budget ELSE NULL END AS Professional_And_Contract_Services_Remaining_Percent, " & vbCrLf & _
			"	CASE WHEN Travel_Budget<>0 THEN Travel_Remaining/Travel_Budget ELSE NULL END AS Travel_Remaining_Percent,  " & vbCrLf & _
			"	CASE WHEN Equipment_Budget<>0 THEN Equipment_Remaining/Equipment_Budget ELSE NULL END  AS Equipment_Remaining_Percent, " & vbCrLf & _
			"	CASE WHEN Supplies_And_DOE_Budget<>0 THEN Supplies_And_DOE_Remaining/Supplies_And_DOE_Budget ELSE NULL END AS Supplies_And_DOE_Remaining_Percent, " & vbCrLf & _
			"	CASE WHEN Total_Budget <>0 THEN Total_Remaining/Total_Budget ELSE NULL END AS Total_Remaining_Percent, " & vbCrLf
	End If
	If ShowExcluded = True Then
		sql = sql & "	Total_Excluded, "
	Else
		sql = sql & "	"
	End If
	If ShowRemaining = False Then
		sql = sql & "Reimbursement, "
	Else
		sql = sql & "Reimbursement_Remaining, "
	End If
	If ShowCashMatch = True Then
		sql = sql & "Cash_Match, Cash_Match_Percentage, "
	End If
	If ShowInKind = True Then
		If ShowRemaining = False Then
			sql = sql & "In_Kind_Expenditure, "
		Else
			sql = sql & "In_Kind_Remaining, "
		End If
	End If
	If ShowBPO = True Then
		sql = sql & "BPO_MVCPA, BPO_Local, BPO_PI, BPO_Total, " & vbCrLf
	End If
	If ShowCOVID = True Then
		sql = sql & "	COVID_MVCPA, COVID_Local, COVID_PI, COVID_Total, " & vbCrLf
	End If
	sql = sql & "1 AS Sorting, REPLACE(Grantee_Name,'City of ','') As GranteeSort " & vbCrLf & _
	"FROM CTE  " & vbCrLf & _
	"UNION " & vbCrLf & _
	"SELECT NULL AS GranteeID, NULL AS GrantID, @Quarter AS Quarter, @FiscalYEar AS Fiscal_Year, 'Total' AS Grantee_Name,  " & vbCrLf & _
	"	NULL AS Program_Name, NULL AS Grant_Number, " & vbCrLf
	If ShowRemaining = False And ShowPercent = False Then
		sql = sql & "	SUM(Personnel) AS Personnel, SUM(Fringe) AS Fringe, SUM(Overtime) AS Overtime,  " & vbCrLf & _
			"	SUM(Professional_And_Contract_Services) AS Professional_And_Contract_Services, " & vbCrLf & _
			"	SUM(Travel) AS Travel,  " & vbCrLf & _
			"	SUM(Equipment) AS Equipment, " & vbCrLf & _
			"	SUM(Supplies_And_DOE) AS Supplies_And_DOE, " & vbCrLf & _
			"	SUM(Total_Expenditures) AS Total_Expenditures, " & vbCrLf
	ElseIf ShowRemaining = False And ShowPercent = True Then
		sql = sql & "	SUM(Personnel)/SUM(Personnel_Budget) AS Personnel_Remaining_Percent, " & vbCrLf & _
			"	SUM(Fringe)/SUM(Fringe_Budget) AS Fringe_Remaining_Percent, " & vbCrLf & _
			"	SUM(Overtime)/SUM(Overtime_Budget) AS Overtime_Remaining_Percent,  " & vbCrLf & _
			"	SUM(Professional_And_Contract_Services)/SUM(Professional_And_Contract_Services_Budget) AS Professional_And_Contract_Services_Remaining_Percent, " & vbCrLf & _
			"	SUM(Travel)/SUM(Travel_Budget) AS Travel_Remaining_Percent,  " & vbCrLf & _
			"	SUM(Equipment)/SUM(Equipment_Budget) AS Equipment_Remaining_Percent, " & vbCrLf & _
			"	SUM(Supplies_And_DOE)/SUM(Supplies_And_DOE_Budget) AS Supplies_And_DOE_Remaining_Percent, " & vbCrLf & _
			"	SUM(Total_Expenditures)/sum(Total_Budget) AS Total_Remaining_Percent, " & vbCrLf
	ElseIf ShowRemaining = True And ShowPercent = False Then
		sql = sql & "	SUM(Personnel_Remaining) AS Personnel_Remaining, " & vbCrLf & _
			"	SUM(Fringe_Remaining) AS Fringe_Remaining, " & vbCrLf & _
			"	SUM(Overtime_Remaining) AS Overtime_Remaining,  " & vbCrLf & _
			"	SUM(Professional_And_Contract_Services_Remaining) AS Professional_And_Contract_Services_Remaining, " & vbCrLf & _
			"	SUM(Travel_Remaining) AS Travel_Remaining,  " & vbCrLf & _
			"	SUM(Equipment_Remaining) AS Equipment_Remaining, " & vbCrLf & _
			"	SUM(Supplies_And_DOE_Remaining) AS Supplies_And_DOE_Remaining, " & vbCrLf & _
			"	SUM(Total_Remaining) AS Total_Remaining, " & vbCrLf
	ElseIf ShowRemaining = True And ShowPercent = True Then
		sql = sql & "	SUM(Personnel_Remaining)/SUM(Personnel_Budget) AS Personnel_Remaining_Percent, " & vbCrLf & _
			"	SUM(Fringe_Remaining)/SUM(Fringe_Budget) AS Fringe_Remaining_Percent, " & vbCrLf & _
			"	SUM(Overtime_Remaining)/SUM(Overtime_Budget) AS Overtime_Remaining_Percent,  " & vbCrLf & _
			"	SUM(Professional_And_Contract_Services_Remaining)/SUM(Professional_And_Contract_Services_Budget) AS Professional_And_Contract_Services_Remaining_Percent, " & vbCrLf & _
			"	SUM(Travel_Remaining)/SUM(Travel_Budget) AS Travel_Remaining_Percent,  " & vbCrLf & _
			"	SUM(Equipment_Remaining)/SUM(Equipment_Budget) AS Equipment_Remaining_Percent, " & vbCrLf & _
			"	SUM(Supplies_And_DOE_Remaining)/SUM(Supplies_And_DOE_Budget) AS Supplies_And_DOE_Remaining_Percent, " & vbCrLf & _
			"	SUM(Total_Remaining)/sum(Total_Budget) AS Total_Remaining_Percent, " & vbCrLf
	End If
	If ShowExcluded = True Then
		sql = sql & "	SUM(Total_Excluded) AS Total_Excluded, "  & vbCrLf
	End If
	If ShowRemaining = False Then
		sql = sql & "	SUM(Reimbursement) AS Reimbursement, " & vbCrLf
	Else
		sql = sql & "	SUM(Reimbursement_Remaining) AS Reimbursement_Remaining, " & vbCrLf
	End If
	If ShowCashMatch = True Then
		sql = sql & "	SUM(Cash_Match) AS Cash_Match, SUM(Cash_Match)/SUM(Reimbursement) AS Cash_Match_Percentage, " & vbCrLf
	End If
	If ShowInkind = True Then
		If ShowRemaining = False Then
			sql = sql & "SUM(In_Kind_Expenditure) AS In_Kind_Expenditure,  " & vbCrLf
		Else
			sql = sql & "SUM(In_Kind_Remaining) AS In_Kind_Remaining,  " & vbCrLf
		End If
	End If
	If ShowBPO = True Then
		sql = sql &	"	SUM(BPO_MVCPA) AS BPO_MVCPA, SUM(BPO_Local) AS BPO_Local, SUM(BPO_PI) AS BPO_PI, SUM(BPO_Total) AS BPO_Total, " & vbCrLf
	End If
	If ShowCOVID = True Then
		sql = sql & "	SUM(COVID_MVCPA) AS COVID_MVCPA, SUM(COVID_Local) AS COVID_Local, SUM(COVID_PI) AS COVID_PI, SUM(COVID_Total) AS COVID_Total, " & vbCrLf
	End If
	sql = sql & "2 AS Sorting, NULL AS GranteeSort " & vbCrLf & _
	"FROM CTE " & vbCrLf & _
	"ORDER BY Sorting, " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If rs.EOF = False Then
' passed here.
End If

If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=Summary" & FiscalYear & "_" & Quarter & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Expenditure Summary</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="Quarter">Quarter:</label> <select name="Quarter" id="Quarter" onchange="Selection.submit();">
<%
	For i = 1 to 4
		Response.Write("<option value=""" & i & """" & selected(Quarter, i) & ">" & QuarterDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="Show">Show:</label> <select name="Show" id="Show" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowDescription)
	Response.Write("<option value=""" & i & """" & Selected(Show, i) & ">" & ShowDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select><br />
<input type="checkbox" name="ShowExcluded" id="ShowExcluded" value="1" <%=Checked(ShowExcluded, true) %>  onchange="Selection.submit();" /> Show Excluded Total&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowYTD" id="ShowYTD" value="1" <%=Checked(ShowYTD, true) %>  onchange="Selection.submit();" /> Show Year-To-Date values&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowRemaining" id="ShowRemaining" value="1" <%=Checked(ShowRemaining, true) %>  onchange="Selection.submit();" /> Show Remaining&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowCashMatch" id="ShowCashMatch" value="1" <%=Checked(ShowCashMatch, true) %>  onchange="Selection.submit();" /> Show Cash Match&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowInKind" id="ShowInKind" value="1" <%=Checked(ShowInKind, true) %>  onchange="Selection.submit();" /> Show In-Kind Totals&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowBPO" id="ShowBPO" value="1" <%=Checked(ShowBPO, true) %>  onchange="Selection.submit();" /> Show BPO&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowCOVID" id="ShowCOVID" value="1" <%=Checked(ShowCOVID, true) %>  onchange="Selection.submit();" /> Show COVID&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="ShowPercent" id="ShowPercent" value="1" <%=Checked(ShowPercent, true) %>  onchange="Selection.submit();" /> Show Percent&nbsp;&nbsp;&nbsp;
<a href="Summary.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&Quarter=<%=Quarter %>&ShowExcluded=<%=prepBitSQL(ShowExcluded) %>&OrderBy=<%=OrderBy %>&ShowYTD=<%=prepBitSQL(ShowYTD) %>&ShowInKind=<%=prepBitSQL(ShowInKind) %>&Show=<%=prepIntegerSQL(Show) %>&ShowCOVID=<%=prepBitSQL(ShowCOVID) %>&ShowRemaining=<%=prepBitSQL(ShowRemaining) %>&ShowCashMatch=<%=prepBitSQL(ShowCashMatch) %>" target="_blank">Excel</a>
</form>

<br />
<%
End If
%>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	Response.Write("<th>Quarter</th>" & vbCrLf)
	Response.Write("<th>Grantee</th>" & vbCrLf)
	For i = 6 To (rs.Fields.Count-3)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		If ShowExcel = False Then
			Response.Write("<td style=""text-align: center; ""><a href=Report.asp?GrantID=" & rs.Fields("GrantID") & "&Quarter=" & rs.Fields("Quarter") & " target=""_blank"">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</a></td>" & vbCrLf)
		Else
			Response.Write("<td style=""text-align: center; "">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</td>" & vbCrLf)
		End If
		Response.Write("<td style=""text-align: left; white-space: nowrap; "" title=""" & rs.Fields("Program_Name") & ", " & rs.Fields("Grant_Number") & """>" & rs.Fields("Grantee_Name") & "</td>" & vbCrLf)
		Response.Write("<td style=""text-align: left; white-space: nowrap; "">" & rs.Fields("Grant_Number") & "</td>" & vbCrLf)
		For i = 7 To (rs.Fields.Count-3)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf InStr(rs.Fields(i).Name,"Percent")>0 Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(100.0*rs.Fields(i).value,2, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf InStr(1, rs.Fields(i).Name, "date", vbTextCompare) > 0 Then
				Response.Write("<td style=""text-align: right"">" & formatDate(rs.Fields(i).value) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'Response.Write("<td>" & rs.Fields("Approval_Date").Type & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%
If ShowExcel = False Then
%>
<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->