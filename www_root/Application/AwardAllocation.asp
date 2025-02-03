<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, FiscalYear, Columns, OrderBy, OrderByDescription, OrderByField, _
	ShowQuestionScores, RemoveExcluded, ShowExcel, _
	TotalMVCPAFundsRequested, TotalCashMatch, TotalMVCPAFundsAdjusted
debug = False
Columns = 17

OrderByDescription = Array("App ID", "Grantee Name", "Program Name", "Total Points", "Grant Type", "Adjusted Grant Request")
OrderByField = Array("[App_ID]", "[GranteeSort], [App_ID]", "[Program_Name], [App_ID]", "[Total] DESC, [GranteeSort]", "[GrantTypeID], [GranteeSort]", "[AdjustedGrantRequest]")

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

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	Response.Write("No Fiscal Year specified")
	Response.End
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
End If

If Request.Form("ShowQuestionScores")="1" Then
	ShowQuestionScores = True
ElseIf Request.QueryString("ShowQuestionScores")="1" Then
	ShowQuestionScores = True
Else
	ShowQuestionScores = False
End If

If Request.Form.Count = 0 Then
	RemoveExcluded = True
ElseIf Request.Form("RemoveExcluded")="1" Then
	RemoveExcluded = True
ElseIf Request.QueryString("RemoveExcluded")="1" Then
	RemoveExcluded = True
Else
	RemoveExcluded = False
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

If Debug = True Then
	Response.Write("<pre>ShowQuestionScores=" & ShowQuestionScores & "; RemoveExcluded=" & RemoveExcluded & "</pre>")
End If
If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=Scores" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Application Score Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<div class="sectiontitle">MVCPA <%=FiscalYear%> Award Allocation Worksheet</div>
<table style="margin: auto; "><tr><td><form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to (Year(Date())+1)
		Response.Write("<option value=""" & i & """ " & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%></select>&nbsp;&nbsp;&nbsp;
<label for="OrderBy">Order By:</label> <select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """ " & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%></select>&nbsp;&nbsp;&nbsp;<input type="checkbox" name="ShowQuestionScores" value="1" <%=checked(ShowQuestionScores, true) %> onclick="document.Selection.submit();"/>Show Question Scores
&nbsp;&nbsp;&nbsp;<input type="checkbox" name="RemoveExcluded" value="1" <%=checked(RemoveExcluded, true) %> onclick="document.Selection.submit();"/>Remove Applications marked excluded
&nbsp;&nbsp;&nbsp;<a href="AwardAllocation.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&OrderBy=<%=OrderBy %>&ShowQuestionScores=<%=ShowQuestionScores %>&RemoveExcluded=<%=RemoveExcluded %>" target="_blank">Excel</a>
</form></td></tr></table>
<table style="margin: auto; ">
<%
End If

TotalMVCPAFundsRequested = 0.0
TotalMVCPAFundsAdjusted = 0.0
TotalCashMatch = 0.0

sql = "SELECT A.*, MVCPAFunds, CashMatch, AdjustedGrantRequest, C.ExcludeFromConsideration, C.ConsiderationNotes " & vbCrLf & _
	"FROM Scoring.vwScoringAverages AS A" & vbCrLf & _
	"LEFT JOIN ( " & vbCrLf & _
	"	SELECT AppID, SUM(MVCPAFunds) AS MVCPAFunds, SUM(CashMatch) AS CashMatch, SUM(InKIndMatch) AS InKindMatch,  " & vbCrLf & _
	"		SUM(CASE WHEN ISNULL(UnallowedItem,0)=1 THEN 0 WHEN AllowedAmount IS NOT NULL THEN AllowedAmount ELSE MVCPAFunds END) AS AdjustedGrantRequest " & vbCrLF & _
	"	FROM Application.BudgetDetails " & vbCrLf & _
	"GROUP BY AppID) AS B ON B.AppID=A.App_ID " & vbCrLf & _
	"LEFT JOIN Application.Admin AS C ON C.AppID=A.App_ID " & vbCrLf & _
	"WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"	AND GrantTypeID IN (1,3) " & vbCrLf
If RemoveExcluded=True Then
	sql = sql & " AND ISNULL(ExcludeFromConsideration,0)=0 " & vbCrLf
End If
sql = sql & "ORDER BY " & OrderByField(OrderBy) & " "

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

If rs.EOF = False Then
	TotalMVCPAFundsRequested = TotalMVCPAFundsRequested + rs.Fields("MVCPAFunds")
	TotalMVCPAFundsAdjusted = TotalMVCPAFundsAdjusted + rs.Fields("AdjustedGrantRequest")
	TotalCashMatch = TotalCashMatch + rs.Fields("CashMatch")
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
	Response.Write(vbTab & "<th>Grantee Name</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Program Name</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Grant_Type</th>" & vbCrLf)
	If ShowQuestionScores = True Then
		Response.Write(vbTab & "<th>Q1</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q2</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q3</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q4</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q5</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q6</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q7</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q8</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q9</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q10</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q11</th>" & vbCrLf)
		Response.Write(vbTab & "<th>Q1-Q11</th>" & vbCrLf)
		Response.Write(vbTab & "<th>EC1</th>" & vbCrLf)
		Response.Write(vbTab & "<th>EC2</th>" & vbCrLf)
		Response.Write(vbTab & "<th>EC Total</th>" & vbCrLf)
	End If
	Response.Write(vbTab & "<th>Total</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Meets Needs Requirement</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Meets Other Sections Requirement</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Reason or Issues</th>" & vbCrLf)
	Response.Write(vbTab & "<th>MVCPA Funds Requsted</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Cash Match</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Adjusted Grant Request</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("Grantee_Name").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("Program_Name").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & Replace(rs.Fields("Grant_Type").Value, " Grant","") & "</td>" & vbCrLf)
		If ShowQuestionScores = True Then
			If IsNull(rs.Fields("Score_1").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_1")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_1") & "; text-align: right; "">" & rs.Fields("Score_1").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_1").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_2").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_2")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_2") & "; text-align: right; "">" & rs.Fields("Score_2").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_2").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_3").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_3")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_3") & "; text-align: right; "">" & rs.Fields("Score_3").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_3").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_4").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_4")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_4") & "; text-align: right; "">" & rs.Fields("Score_4").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_4").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_5").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_5")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_5") & "; text-align: right; "">" & rs.Fields("Score_5").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_5").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_6").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_6")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_6") & "; text-align: right; "">" & rs.Fields("Score_6").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_6").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_7").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_7")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_7") & "; text-align: right; "">" & rs.Fields("Score_7").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_7").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_8").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_8")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_8") & "; text-align: right; "">" & rs.Fields("Score_8").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_8").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_9").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_9")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_9") & "; text-align: right; "">" & rs.Fields("Score_9").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_9").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_10").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_10")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_10") & "; text-align: right; "">" & rs.Fields("Score_10").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_10").Value & "</td>" & vbCrLf)
			End If
			If IsNull(rs.Fields("Score_11").Value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Color_11")) = False Then
				Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_11") & "; text-align: right; "">" & rs.Fields("Score_11").Value & "</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_11").Value & "</td>" & vbCrLf)
			End If
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Question 1-11 Total").Value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_12").Value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_13").Value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("EC Total").Value & "</td>" & vbCrLf)
		End If
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Total").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("Meets Needs Requirement").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("Meets Other Sections Requirement").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: left; "">" & rs.Fields("ConsiderationNotes").Value & "</td>" & vbCrLf)
		If IsNull(rs.Fields("MVCPAFunds").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("MVCPAFunds").Value) & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("CashMatch").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("CashMatch").Value) & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("AdjustedGrantRequest").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & prepCurrencyWeb(rs.Fields("AdjustedGrantRequest").Value) & "</td>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext()
	Wend
	Response.Write("</tbody>" & vbCrLf)

	Response.Write("<tfoot>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	Response.Write("<td colspan=3></td>" & vbCrLf)
	If ShowQuestionScores = True Then
		Response.Write("<td colspan=""15""></td>" & vbCrLf)
	End If
	Response.Write("<td colspan=""3""></td>" & vbCrLf)
	Response.Write("<th style=""text-align: center"">Total</th>" & vbCrLf)
	Response.Write("<td style=""text-align: right"">" & prepCurrencyWeb(TotalMVCPAFundsRequested) & "</td>" & vbCrLf)
	Response.Write("<td style=""text-align: right"">" & prepCurrencyWeb(TotalCashMatch) & "</td>" & vbCrLf)
	Response.Write("<td style=""text-align: right"">" & prepCurrencyWeb(TotalMVCPAFundsAdjusted) & "</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</tfoot>" & vbCrLf)
Else
	Response.Write("<tr><td>There are no scores to report</td></tr>")
End If

%>
</table>
<%	If ShowExcel = False Then %>
</body>
</html>
<%	End If 

function Selected(vVariable, vValue)
	If vVariable = vValue Then
		Selected = " selected"
	Else
		Selected = ""
	End If
end function

function checked(vVariable, vValue)
	If vVariable = vValue Then
		checked = " checked"
	Else
		checked = ""
	End If
end function
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->