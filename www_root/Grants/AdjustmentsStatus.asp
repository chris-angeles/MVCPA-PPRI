<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, OrderByDescription, OrderByField
OrderByDescription = Array("GrantID", "Grantee Name", "Grant Number", "ORI", "Award Amount", "Submission Date (Ascending)", "Submission Date (Descending)", "Approval Date")
OrderByField = Array("H.GrantID", "REPLACE(G.GranteeName,'City of ','')", "H.GrantNumber", "G.ORI", "H.AwardAmount", "A.SubmitTimeStamp ASC, AdjustmentID ASC", "A.SubmitTimeStamp DESC, AdjustmentID DESC", "SecondApprovalDate, AdjustmentID")
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
ElseIf Len(Session("FiscalYear"))>0 Then
	FiscalYear = CInt(Session("FiscalYear"))
Else
	If Month(Date()) > 8 Then
		FiscalYear = Year(Date)+1
	Else
		FiscalYear = Year(Date)
	End If
End If
If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
Else
	OrderBy = 6
End If

sql = "SELECT H.FiscalYear AS Fiscal_Year, A.AdjustmentID, H.GrantID, " & vbCrLf & _
	"	G.GranteeName AS Grantee_Name, H.ProgramName AS Program_Name, H.GrantNumber AS Grant_Number, " & vbCrLf & _
	"	A.ProgramChange AS Program_Change, A.BudgetChange AS Budget_Change, " & vbCrLf & _
	"	CAST(ISNULL(T.SumAbsChanges,0.0)/2.0 + ABS(ISNULL(A.ProgramIncomeToBeAddedToBudget,0.0))/2 AS MONEY) AS Grant_Adjustment_Total, " & vbCrLf & _
	"	CASE WHEN S.FirstSubmit<A.SubmitTimestamp THEN CONVERT(VARCHAR,S.FirstSubmit, 101) ELSE NULL END AS First_Submit, " & vbCrLf & _
	"	CONVERT(VARCHAR,A.SubmitTimestamp, 101) AS Submit_Date, " & vbCrLf & _
	"	CONVERT(VARCHAR,FirstApprovalDate, 101) AS First_Approval_Date, " & vbCrLf & _
	"	CONVERT(VARCHAR,SecondApprovalDate, 101) AS Second_Approval_Date, " & vbCrLf & _
	"	CONVERT(VARCHAR,DenialDate, 101) AS Denial_Date, " & vbCrLf & _ 
	"	CONVERT(VARCHAR,AdministrativelyClosedDate, 101) AS Administratively_Closed_Date, " & vbCrLf & _
	"	ISNULL(ChangesApplied,0) AS Changes_Applied " & vbCrLf & _
	"FROM Grantees G " & vbCrLf & _
	"JOIN [Grants].Main AS H ON H.GranteeID=G.GranteeID " & vbCrLf & _
	"JOIN [Grants].Adjustments AS A ON A.GrantID=H.GrantID " & vbCrLf & _
	"LEFT JOIN (SELECT AdjustmentID, MIN(SubmitTimestamp) AS FirstSubmit FROM [Grants].Adjustments_Log GROUP BY AdjustmentID) AS S ON S.AdjustmentID=A.AdjustmentID " & vbCrLf & _
	"LEFT JOIN (SELECT AdjustmentID, SUM(ABS(TotalExpendituresChange)) AS SumAbsChanges FROM [Grants].AdjustmentDetails GROUP BY AdjustmentID) AS T ON T.AdjustmentID=A.AdjustmentID " & vbCrLf & _
	"WHERE FiscalYear=" & FiscalYear & " " & vbCrLf & _
	"ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grant Adjustment Request Status Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<form name="Selection" id="Selection" method="post" >
<div class="sectiontitle">Grant Adjustment Request Status Report</div>
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to (Year(Date())+1)
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select><br />
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>
</form>
</div>

<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<head>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom;"">" & vbCrLf)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLf)
	Response.Write("<head>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr>" & vbCrLf)
		For i = 0 To (rs.Fields.Count-1)
			If IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "GranteeID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "GrantID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grants\Grant.asp?GrantID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "AdjustmentID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grants\Adjustment.asp?GrantID=" & rs.Fields("GrantID") & "&AdjustmentID=" & rs.Fields(i) & """ target=""_blank"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.WRite("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>

<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->