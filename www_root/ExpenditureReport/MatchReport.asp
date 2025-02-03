<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, Quarter, OrderByDescription, QuarterDescription, OrderByField, _
	ShowCategoryDetail, ShowYTD, Show, ShowDescription, ShowClause, StartQuarter, ShowExcel, CurrentDate
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number", "ORI", "Award Amount")
QuarterDescription = Array("", "September 1 - November 30","December 1 - February 28", "March 1 - May 31", "June 1 - August 31")
OrderByField = Array("B.GrantID", "REPLACE(A.GranteeName,'City of ','')", "B.GrantNumber", "A.ORI", "B.AwardAmount")
ShowDescription = Array ("All", "Border", "Port", "Port 2", "Border and Port", "Border, Port, and Port 2")
ShowClause = Array ("1=1", "A.BorderCounty=1", "A.PortCounty=1", "A.Port2County=1", "(A.BorderCounty=1 OR A.PortCounty=1)", "(A.BorderCounty=1 OR A.PortCounty=1 OR A.Port2County=1)")
debug = False
CurrentDate = Date()
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'</pre>" & vbCrLf)
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
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
End If
If Len(Request.Form("Quarter"))>0 Then
	Quarter = CInt(Request.Form("Quarter"))
ElseIf Len(Request.QueryString("Quarter"))>0 Then
	Quarter = CInt(Request.QueryString("Quarter"))
ElseIf CurrentDate < CDate("3/1/" & FiscalYear) Then
	Quarter = 1
ElseIf CurrentDate < CDate("5/1/" & FiscalYear) Then
	Quarter = 2
ElseIf CurrentDate < CDate("8/1/" & FiscalYear) Then
	Quarter = 3
Else
	Quarter = 4
End If
If Request.Form("ShowCategoryDetail") = "1" Then
	ShowCategoryDetail = True
ElseIf Request.QueryString("ShowCategoryDetail") = "1" Then
	ShowCategoryDetail = True
Else
	ShowCategoryDetail = False
End If
'If Request.Form("ShowYTD") = "1" Then
	ShowYTD = True
	StartQuarter=1
'ElseIf Request.QueryString("ShowYTD") = "1" Then
'	ShowYTD = True
'	StartQuarter=1
'Else
'	StartQuarter = Quarter
'	ShowYTD = False
'End If
If Len(Request.Form("Show")) > 0 Then
	Show = CInt(Request.Form("Show"))
ElseIf Len(Request.QueryString("Show")) > 0 Then
	Show = CInt(Request.QueryString("Show"))
Else
	Show = 0
End If
If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "DECLARE @FiscalYear INT = " & prepIntegerSQL(FiscalYear) & "  " & vbCrLf & _
	"DECLARE @Quarter INT = " & prepIntegerSQL(Quarter) & " " & vbCrLf & _
	"DECLARE @StartQuarter INT = " & prepIntegerSQL(StartQuarter) & " " & vbCrLf & _
	"SELECT B.FiscalYear AS Fiscal_Year, @Quarter AS Qtr, A.GranteeID AS Grantee_ID, B.GrantID AS Grant_ID, " & vbCrLf & _
	"	A.GranteeName AS Grantee_Name, " & vbCrLf
If OrderByDescription(OrderBY) = "ORI" Then
	sql = sql & "	A.ORI, " & vbCrLf
End If
If OrderByDescription(OrderBy) = "Grant Number" Then
	sql = sql & "	B.GrantNumber AS Grant_Number, " & vbCrLf
End If
If OrderByDescription(OrderBy) = "Award Amount" Then
	sql = sql & "	B.AwardAmount AS Award_Amount, " & vbCrLf
End If
If ShowCategoryDetail = True Then
	sql = sql & "	D.Personnel, D.Fringe, D.Overtime, D.Professional_And_Contract_Services, " & vbCrLf & _
	"	D.Travel, D.Equipment, D.Supplies_And_DOE, " & vbCrLf
End If
sql = sql & "	D.[Total_Expenditures_(less_excluded)], " & vbCrLF & _
	"	C.Reimbursement AS MVCPA_Reimbursements, " & vbCrLf & _
	"	E.ReimbursableExpenditures-E.ReimbursementYTD AS Reimbursement_Cash_Match, " & vbCrLf & _
	"	D.Excluded_Amount AS Over_Budget_Category, " & vbCrLf & _
	"	CASE WHEN ISNULL(D.YTD_Expenditure,0.0) > ISNULL(B.AwardAmount,0.0)+ISNULL(B.MatchAmount,0.0) " & vbCrLf & _
	"	THEN ISNULL(D.YTD_Expenditure,0.0) - ISNULL(B.AwardAmount,0.0) + ISNULL(B.MatchAmount,0.0) " & vbCrLf & _
	"		ELSE NULL END AS Over_Budget_Total, " & vbCrLf & _
	"	D.Allowed_Overage AS Manual_Adjustment, " & vbCrLf & _
	"	ISNULL(D.YTD_Expenditure,0.0) - ISNULL(E.ReimbursableExpenditures,0.0) - " & vbCrLF & _
	"		ISNULL(DPS_Reported,0.0) - ISNULL(NICB_Reported,0.0) AS Total_Unallowed, " & vbCrLf & _
	"	C.Program_Income AS Included_PI, " & vbCrLf & _
	"	C.PI_Excluded AS Over_Budget_PI, " & vbCrLf & _
	"	DPS_Reported, NICB_Reported, " & vbCrLf & _
	"	ISNULL(C.PI_Excluded,0.0) + ISNULL(DPS_Reported,0.0) + ISNULL(NICB_Reported,0.0) AS Total, " & vbCrLf & _
	"	ISNULL(E.ReimbursementYTD,0) + ISNULL(D.YTD_Expenditure,0.0) - ISNULL(C.PI_Excluded,0.0) + " & vbCrLf & _
	"		ISNULL(DPS_Reported,0.0) + ISNULL(NICB_Reported,0.0) AS Grand_Total_Match " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"JOIN [Grants].Main AS B ON A.GranteeID=B.GranteeID " & vbCrLf & _
	"JOIN (SELECT GrantID, @Quarter AS Quarter, " & vbCrLf & _
	"		SUM(Reimbursement) AS Reimbursement, " & vbCrLf & _
	"		SUM(UnbudgetedPI) AS PI_Excluded, " & vbCrLf & _
	"		SUM(ExpendedThisQuarter) AS Program_Income, " & vbCrLf & _
	"		SUM(InLieuOfDPS) AS DPS_Reported, " & vbCrLf & _
	"		SUM(InLieuOfNICB) AS NICB_Reported " & vbCrLf & _
	"	FROM ER.Main " & vbCrLf & _
	"	WHERE Quarter BETWEEN @StartQuarter AND @Quarter " & vbCrLf & _
	"	GROUP BY GrantID " & vbCrLf & _
	")AS C ON C.GrantID=B.GrantID " & vbCrLf & _
	"JOIN ( " & vbCrLf & _
	"	SELECT GrantID, @Quarter AS Quarter, " & vbCrLf & _
	"		SUM(CASE WHEN BudgetCategoryID=1 THEN ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0) ELSE NULL END) AS Personnel, " & vbCrLf & _
	"		SUM(CASE WHEN BudgetCategoryID=2 THEN ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0) ELSE NULL END) AS Fringe, " & vbCrLf & _
	"		SUM(CASE WHEN BudgetCategoryID=3 THEN ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0) ELSE NULL END) AS Overtime, " & vbCrLf & _
	"		SUM(CASE WHEN BudgetCategoryID=4 THEN ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0) ELSE NULL END) AS Professional_And_Contract_Services, " & vbCrLf & _
	"		SUM(CASE WHEN BudgetCategoryID=5 THEN ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0) ELSE NULL END) AS Travel, " & vbCrLf & _
	"		SUM(CASE WHEN BudgetCategoryID=6 THEN ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0) ELSE NULL END) AS Equipment, " & vbCrLf & _
	"		SUM(CASE WHEN BudgetCategoryID=7 THEN ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0) ELSE NULL END) AS Supplies_And_DOE, " & vbCrLf & _
	"		SUM(CashExpenditure) AS YTD_Expenditure, " & vbCrLf & _
	"		SUM(ExcludedAmount) AS Excluded_Amount, " & vbCrLf & _
	"		SUM(AllowedOverage) AS Allowed_Overage, " & vbCrLf & _
	"		SUM(ISNULL(CashExpenditure,0.0)) AS [Total_Expenditures], " & vbCrLf & _
	"		SUM(ISNULL(CashExpenditure,0.0)-ISNULL(ExcludedAmount,0.0)) AS [Total_Expenditures_(less_excluded)] " & vbCrLf & _
	"	FROM ER.Detail " & vbCrLf & _
	"	WHERE Quarter BETWEEN @StartQuarter AND @Quarter " & vbCrLf & _
	"	GROUP BY GrantID " & vbCrLf & _
	") AS D ON D.GrantID=C.GrantID " & vbCrLf & _
	"LEFT JOIN ER.Main AS E ON E.GrantID=C.GrantID AND E.Quarter=@Quarter " & vbCrLf & _
	"WHERE B.FiscalYear=@FiscalYear AND C.Quarter=@Quarter AND " & ShowClause(Show) & " " & vbCrLf & _
	"ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)


If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ERStatus" & FiscalYear & ".xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Match Report</title>
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
</select>&nbsp;&nbsp;
<label for="Quarter">Quarter:</label> <select name="Quarter" id="Quarter" onchange="Selection.submit();">
<%
	For i = 1 to 4
		Response.Write("<option value=""" & i & """" & selected(Quarter, i) & ">" & QuarterDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<input type="checkbox" name="ShowCategoryDetail" id="ShowCategoryDetail" value="1" <%
If ShowCategoryDetail=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowCategoryDetail">Show Category Detail</label>&nbsp;&nbsp;
<!--<input type="checkbox" name="ShowYTD" id="ShowYTD" value="1" <%
If ShowYTD=True Then 
	Response.Write(" Checked")
End If  
%> onchange="Selection.submit();" /><label for="ShowYTD">YTD</label>&nbsp;&nbsp;-->
<label for="Show">Show:</label> <select name="Show" id="Show" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowDescription)
	Response.Write("<option value=""" & i & """" & Selected(Show, i) & ">" & ShowDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<a href="MatchReport.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&Quarter=<%=Quarter %>&ShowCategoryDetail=<%=prepBitSQL(ShowCategoryDetail) %>&OrderBy=<%=OrderBy %>&ShowYTD=<%=prepBitSQL(ShowYTD) %>" target="_blank">Excel</a>
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
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("<thead>" & vbCrLf)

	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		For i = 0 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "Grant_ID" Then
				Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf (InStr(rs.Fields(i).Name,"Rate")>0 OR InStr(rs.Fields(i).Name,"Percent")>0) and rs.Fields(i).Name<>"Reimbursement_Rate_Cash_Match" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,4, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf InStr(1, rs.Fields(i).Name, "date", vbTextCompare) > 0 Then
				Response.Write("<td style=""text-align: right"">" & formatDate(rs.Fields(i).value) & "</td>")
			ElseIf rs.Fields(i).Name = "First_Submit" Then
				Response.Write("<td style=""text-align: right"">" & formatDate(rs.Fields(i).value) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		'Response.Write("<td>" & rs.Fields(7).Type & "</td>")
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