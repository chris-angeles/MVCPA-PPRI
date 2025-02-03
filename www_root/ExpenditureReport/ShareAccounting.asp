<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, Quarter, OrderByDescription, QuarterDescription, OrderByField, _
	ShowExcel, CurrentDate, NonDisplayColumns
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number", "ORI")
QuarterDescription = Array("", "September 1 - November 30","December 1 - February 28", "March 1 - May 31", "June 1 - August 31")
OrderByField = Array("B.GrantID", "REPLACE(GranteeName,'City of ','')", "GrantNumber", "ORI")
debug = False
NonDisplayColumns = 1
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
If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "DECLARE @Quarter INT=" & Quarter & "; " & vbCrLf & _
	"DECLARE @FiscalYear INT=" & prepIntegerSQL(FiscalYear) & "; " & vbCrLf & _
	"SELECT B.GrantID AS Grant_ID,  " & vbCrLf & _
	"	REPLACE(GranteeName, 'City of ', '') AS Grantee,  " & vbCrLf & _
	"	ORI, " & vbCrLf & _
	"	B.FiscalYear AS Fiscal_Year,  " & vbCrLf & _
	"	C.Quarter,  " & vbCrLf & _
	"	D.Cash_Expenditure,  " & vbCrLf & _
	"	D.Excluded_Amount,  " & vbCrLf & _
	"	D.YTD_Expenditure_Less_Excluded,  " & vbCrLf & _
	"	ISNULL(C.PriorInLieuOfDPS, 0.0) + ISNULL(C.InLieuOfDPS, 0.0) AS DPS,  " & vbCrLf & _
	"	ISNULL(C.PriorInLieuOfNICB, 0.0) + ISNULL(C.InLieuOfNICB, 0.0) AS NICB,  " & vbCrLf & _
	"	C.UnbudgetedPI AS Unbudgeted_PI,  " & vbCrLf & _
	"	C.ReimbursementYTD AS MVCPA,  " & vbCrLf & _
	"	C.ReimbursableExpenditures AS Reimbursable_Expenditures,  " & vbCrLf & _
	"	D.YTD_Expenditure_Less_Excluded - (ISNULL(C.PriorInLieuOfDPS, 0.0) + ISNULL(C.InLieuOfDPS, 0.0)) - (ISNULL(C.PriorInLieuOfNICB, 0.0) + ISNULL(C.InLieuOfNICB, 0.0)) - C.UnbudgetedPI - C.ReimbursementYTD AS Reimbursable_County_Match,  " & vbCrLf & _
	"	D.YTD_Expenditure_Less_Excluded - (ISNULL(C.PriorInLieuOfDPS, 0.0) + ISNULL(C.InLieuOfDPS, 0.0)) - (ISNULL(C.PriorInLieuOfNICB, 0.0) + ISNULL(C.InLieuOfNICB, 0.0)) - C.UnbudgetedPI - C.ReimbursementYTD + D.Excluded_Amount + UnbudgetedPI AS County_Funds,  " & vbCrLf & _
	"	(D.YTD_Expenditure_Less_Excluded - (ISNULL(C.PriorInLieuOfDPS, 0.0) + ISNULL(C.InLieuOfDPS, 0.0)) - (ISNULL(C.PriorInLieuOfNICB, 0.0) + ISNULL(C.InLieuOfNICB, 0.0)) - C.UnbudgetedPI - C.ReimbursementYTD + D.Excluded_Amount + UnbudgetedPI) + -- County Funds " & vbCrLf & _
	"	+(ISNULL(C.PriorInLieuOfDPS, 0.0) + ISNULL(C.InLieuOfDPS, 0.0)) -- DPS " & vbCrLf & _
	"	+ (ISNULL(C.PriorInLieuOfNICB, 0.0) + ISNULL(C.InLieuOfNICB, 0.0))  -- NICB " & vbCrLf & _
	"	+ C.ReimbursementYTD AS Total_Of_Fund_Sources, " & vbCrLf & _
	"	CASE WHEN C.GrantID IS NOT NULL THEN 1 ELSE 0 END AS ERPresent " & vbCrLf & _
	"FROM Grantees AS A " & vbCrLf & _
	"     JOIN [Grants].Main AS B ON B.GranteeID = A.GranteeID " & vbCrLf & _
	"     JOIN ER.Main AS C ON C.GrantID = B.GrantID " & vbCrLf & _
	"     LEFT JOIN " & vbCrLf & _
	"( " & vbCrLf & _
	"    SELECT GrantID,  " & vbCrLf & _
	"           SUM(CashExpenditure) AS Cash_Expenditure,  " & vbCrLf & _
	"           SUM(ExcludedAmount) AS Excluded_Amount,  " & vbCrLf & _
	"           SUM(CASE " & vbCrLf & _
	"                   WHEN Quarter = @Quarter " & vbCrLf & _
	"                   THEN YTDExpenditure " & vbCrLf & _
	"                   ELSE NULL " & vbCrLf & _
	"               END) AS YTD_Expenditure_Less_Excluded,  " & vbCrLf & _
	"           SUM(AllowedOverage) AS Allowed_Overage " & vbCrLf & _
	"    FROM ER.Detail " & vbCrLf & _
	"    WHERE Quarter <= @Quarter " & vbCrLf & _
	"    GROUP BY GrantID " & vbCrLf & _
	") AS D ON D.GrantID = C.GrantID " & vbCrLf & _
	"WHERE C.Quarter = @Quarter AND FiscalYear=@FiscalYear " & vbCrLf & _
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
<title>Grant Report</title>
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
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<a href="ShareAccounting.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&Quarter=<%=Quarter %>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
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
	For i = 0 To (rs.Fields.Count-1-NonDisplayColumns)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("<thead>" & vbCrLf)

	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		If rs.Fields("ERPresent") Then
			Response.Write("<td style=""text-align: center; ""><a href=Report.asp?GrantID=" & rs.Fields("Grant_ID") & "&Quarter=" & rs.Fields("Quarter") & " target=""_blank"">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</a></td>" & vbCrLf)
		Else
			Response.Write("<td style=""text-align: center; "">" & rs.Fields("Fiscal_Year") & "/" & rs.Fields("Quarter") & "</td>" & vbCrLf)
		End If
		For i = 0 To (rs.Fields.Count-1-NonDisplayColumns)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "Grant_ID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grants\Grant.asp?GrantID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Reimbursement_Rate" Then
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