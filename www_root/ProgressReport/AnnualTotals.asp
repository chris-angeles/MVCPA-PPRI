<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, ShowExcel, FiscalYear, OrderBy, Columns, _
	ShowQuestions, ShowQuestionsDescription, ShowQuestionsClause, _
	ShowGrantees, LastQuestion, LastGoal, LastStrategy, _
	Grouping, GroupingDescription, LastDistrict, TotalOnly
Columns = 7
ShowQuestionsDescription = Array("All", "Mandatory", "Border", "Mandatory and Border", "Exclude No Targets")
ShowQuestionsClause = Array("", " AND Mandatory=1 ", "AND BorderReport=1 "," AND (Mandatory=1 Or BorderReport=1) ", " AND (ISNULL(Mandatory,0)=0 AND ISNULL(NoTarget,0)=0) ")
GroupingDescription = Array("None", "State House District", "State Senate District")
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

LastQuestion = 0
LastGoal = 0
If Request.QueryString("ShowExcel") = "1" Then 
	ShowExcel = True
Else
	ShowExcel = False
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
If Len(Request.Form("ShowQuestions")) > 0 Then
	ShowQuestions = CInt(Request.Form("ShowQuestions"))
ElseIf Len(Request.QueryString("ShowQuestions")) > 0 Then
	ShowQuestions = CInt(Request.QueryString("ShowQuestions"))
Else
	ShowQuestions = 0
End If
If Len(Request.Form("ShowGrantees")) > 0 Then
	ShowGrantees = CInt(Request.Form("ShowGrantees"))
ElseIf Len(Request.QueryString("ShowGrantees")) > 0 Then
	ShowGrantees = CInt(Request.QueryString("ShowGrantees"))
Else
	ShowGrantees = 0
End If
If Len(Request.Form("Grouping")) > 0 Then
	Grouping = CInt(Request.Form("Grouping"))
ElseIf Len(Request.QueryString("Grouping")) > 0 Then
	Grouping = CInt(Request.QueryString("Grouping"))
Else
	Grouping = 0
End If
If Request.Form("TotalOnly") = "1" Then 
	TotalOnly = True
ElseIf Request.QueryString("TotalOnly") = "1" Then 
	TotalOnly = True
Else
	TotalOnly = False
End If
If ShowExcel = True and Debug = False Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ApplicationStatus" & FiscalYear & ".xls"
ElseIf Debug = False Then ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Progress Report Annual Totals for Numeric Questions</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="ShowQuestions">Questions to Show: </label> <select name="ShowQuestions" id="ShowQuestions" onchange="Selection.submit();">
<%
For i = 0 to UBound(ShowQuestionsDescription)
	Response.Write("<option value=""" & i & """" & Selected(ShowQuestions, i) & ">" & ShowQuestionsDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<label for="ShowGrantees">Grantees to Show: </label> <select name="ShowGrantees" id="ShowGrantees" onchange="Selection.submit();">
<%
Response.Write("<option value=""0""" & Selected(ShowGrantees, 0) & ">All</option>" & vbCrLf)
Response.Write("<option value=""-5""" & Selected(ShowGrantees, -5) & ">Border</option>" & vbCrLf)
Response.Write("<option value=""-4""" & Selected(ShowGrantees, -4) & ">Port</option>" & vbCrLf)
Response.Write("<option value=""-3""" & Selected(ShowGrantees, -3) & ">Port 2</option>" & vbCrLf)
Response.Write("<option value=""-2""" & Selected(ShowGrantees, -2) & ">Border and Port</option>" & vbCrLf)
Response.Write("<option value=""-1""" & Selected(ShowGrantees, -1) & ">Border, Port, and Port 2</option>" & vbCrLf)

sql = "SELECT B.GranteeID, REPLACE(B.GranteeName, 'City of ','') AS Grantee " & vbCrLf & _
	"FROM [Grants].Main AS A " & vbCrLf & _
	"JOIN Grantees AS B ON A.GranteeID=B.GranteeID " & vbCrLf & _
	"WHERE FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"ORDER BY 2"
Set rs=Con.Execute(sql)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("GranteeID") & """" & Selected(ShowGrantees, rs.Fields("GranteeID")) & ">" & rs.Fields("grantee") & "</option>" & vbCrLf)
	rs.MoveNext
Wend
%>
</select>&nbsp;&nbsp;
<label for="Grouping">Grouping: </label> <select name="Grouping" id="Grouping" onchange="Selection.submit();">
<%
For i = 0 to UBound(GroupingDescription)
	Response.Write("<option value=""" & i & """" & Selected(Grouping, i) & ">" & GroupingDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<input type="checkbox" name="TotalOnly" value="1" <%=Checked(TotalOnly, True) %> onchange="Selection.submit();" /> Total Only
&nbsp;&nbsp;<a href="AnnualTotals.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&ShowQuestions=<%=ShowQuestions %>&ShowGrantees=<%=ShowGrantees %>&Grouping=<%=Grouping %>&TotalOnly=<%=TotalOnly%> %>" target="_blank">Excel</a>
</form>

<br />
<%	End If 

sql = "SELECT A.* " & vbCrLf
If Grouping > 0 Then
	sql = sql & ", D.District" & vbCrLf
Else
	sql = sql & ", 0 As District" & vbCrLf
End If
sql = sql & "FROM vwProgressReportTotals AS A " & vbCrLf
If Grouping = 1 Then
	sql = sql & "LEFT JOIN Lookup.CountyHouseDistrict AS D ON D.CountyID=A.CountyID " & vbCrLf
ElseIf Grouping = 2 Then
	sql = sql & "LEFT JOIN Lookup.CountySenateDistrict AS D ON D.CountyID=A.CountyID " & vbCrLf
End If

sql = sql & "WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " " & ShowQuestionsClause(ShowQuestions)
If ShowGrantees = -5 Then
	sql = sql & "AND BorderCounty=1 " & vbCrLf
ElseIf ShowGrantees = -4 Then
	sql = sql & "AND PortCounty=1 " & vbCrLf
ElseIf ShowGrantees = -3 Then
	sql = sql & "AND Port2County=1 " & vbCrLf
ElseIf ShowGrantees = -2 Then
	sql = sql & "AND (BorderCounty=1 OR PortCounty=1) " & vbCrLf
ElseIf ShowGrantees = -1 Then
	sql = sql & "AND (BorderCounty=1 OR PortCounty=1 OR Port2County=1) " & vbCrLf
ElseIf ShowGrantees>0 Then
	sql = sql & " AND A.GranteeID=" & ShowGrantees & vbCrLf
End If

If Grouping > 0 Then
	sql = sql & " AND District>0 " & vbCrLf & "ORDER BY Fiscal_Year, District, Mandatory DESC, GoalID, StrategyID, ActivityID, MeasureID, CASE WHEN GrantID IS NULL THEN 1 ELSE 0 END, Grantee "
Else
	sql = sql & "ORDER BY Fiscal_Year, Mandatory DESC, GoalID, StrategyID, ActivityID, MeasureID, CASE WHEN GrantID IS NULL THEN 1 ELSE 0 END, Grantee "
End If

Set rs=Con.Execute(sql)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
%>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	Response.Write(vbTab & "<th>Grant ID</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Grantee</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Program</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Total</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Target</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Responses</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Percent of Target</th>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False

		If Grouping > 0 And rs.Fields("District") <> LastDistrict Then
			Response.Flush
			LastDistrict = rs.Fields("District")
			If Grouping = 1 Then
				Response.Write("<tr style=""background-color: LightSalmon; ""><th colspan=""" & Columns & """>State House District: " & LastDistrict & "</th></tr>" & vbCrLf)
			ElseIf Grouping = 2 Then
				Response.Write("<tr style=""background-color: LightSalmon; ""><th colspan=""" & Columns & """>State Senate District: " & LastDistrict & "</th></tr>" & vbCrLf)
			End If
		End If

		If LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = True Then
			LastGoal = rs.Fields("GoalID")
			If rs.Fields("GoalID") = 1 Then
				Response.Write("<tr><th colspan=""" & (Columns) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Mandatory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLf)
			ElseIf rs.Fields("GoalID")=2 Then
				Response.Write("<tr><th colspan=""" & (Columns) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLf)
			ElseIf rs.Fields("GoalID")=8 Then
					Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Fraud-Related Motor Vehicle Crime Measures must be reported during the grant term if awarded."">Mandatory Fraud-Related Motor Vehicle Crime Measures Required for all Grantees</th></tr>" & vbCrLf)
				End If
		End If
		If LastGoal <> rs.Fields("GoalID") Then
			LastGoal = rs.Fields("GoalID")
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th colspan=""" & Columns & """>Goal " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If LastStrategy <> rs.Fields("StrategyID") Then
			LastStrategy = rs.Fields("StrategyID")
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th colspan=""" & Columns & """>Strategy " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If LastQuestion <> rs.Fields("QuestionID") Then
			LastQuestion = rs.Fields("QuestionID")
			Response.Write("<tr style=""background-color: PowderBlue; vertical-align: top; "">" & vbCrLf)
			Response.Write(vbTab & "<th>" & rs.Fields("GoalID") & "." & rs.Fields("StrategyID") & "." & rs.Fields("ActivityID") & "</th>" & vbCrLf)
			Response.Write(vbTab & "<td colspan=""" & (columns - 1) & """ style=""font-weight: bold; text-align: left; "">" & rs.Fields("Activity") & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
			Response.Write("<tr style=""background-color: PowderBlue; "">" & vbCrLf)
			Response.Write(vbTab & "<th></th>" & vbCrLf)
			Response.Write(vbTab & "<td colspan=""" & (columns - 1) & """ style=""font-weight: bold; text-align: left; "">" & rs.Fields("Measure") & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		If TotalOnly = True And rs.Fields("Grantee") <> " Total" Then
			' Skip
		Else
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: right"">" & rs.Fields("GrantID").value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: left; white-space: nowrap;"">" & rs.Fields("Grantee").value & "</td>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: left; white-space: nowrap;"">" & rs.Fields("ProgramName").value & "</td>" & vbCrLf)
			If rs.Fields("ResponseTypeID") = 1 Then
				If IsNull(rs.Fields("Integer_Total").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("Integer_Total").value,0, True, True, True) & "</td>" & vbCrLf)
				End If
				If IsNull(rs.Fields("IntegerTarget").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("IntegerTarget").value,0, True, True, True) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields("ResponseTypeID") = 2 Then
				If IsNull(rs.Fields("Decimal_Total").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("Decimal_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
				End If
				If IsNull(rs.Fields("DecimalTarget").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("DecimalTarget").value, 2, True, True, True) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields("ResponseTypeID") = 3 Then
				If IsNull(rs.Fields("Dollar_Total").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">$" & formatnumber(rs.Fields("Dollar_Total").value, 2, True, True, True) & "</td>" & vbCrLf)
				End If
				If IsNull(rs.Fields("DecimalTarget").value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				Else
					Response.Write(vbTab & "<td style=""text-align: right; "">$" & formatnumber(rs.Fields("DollarTarget").value, 2, True, True, True) & "</td>" & vbCrLf)
				End If
			End If
			If IsNull(rs.Fields("Responses").value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("Responses").value,0, True, True, True) & "</td>" & vbCrLf)
			End If
			If rs.Fields("Grantee") = " Total" Then ' No percent for total row.
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf IsNull(rs.Fields("Percent_Off_Target").value) = True Then
				Response.Write(vbTab & "<td></td>" & vbCrLf)
			ElseIf CDbl(rs.Fields("Percent_Off_Target").value) <= -25 Then
				Response.Write(vbTab & "<td style=""text-align: right; color: red; "">" & formatnumber(rs.Fields("Percent_Off_Target").value,2, True, False, True) & "%</td>" & vbCrLf)
			'ElseIf CDbl(rs.Fields("Percent_Off_Target").value) < 0 Then
			'	Response.Write(vbTab & "<td style=""text-align: right; color: yellow; font-weight: bold; "">" & formatnumber(rs.Fields("Percent_Off_Target").value,2, True, False, True) & "%</td>" & vbCrLf)
			ElseIf CDbl(rs.Fields("Percent_Off_Target").value) >= 25 Then
				Response.Write(vbTab & "<td style=""text-align: right; color: blue; "">" & formatnumber(rs.Fields("Percent_Off_Target").value,2, True, False, True) & "%</td>" & vbCrLf)
			Else
				Response.Write(vbTab & "<td style=""text-align: right; "">" & formatnumber(rs.Fields("Percent_Off_Target").value,2, True, False, True) & "%</td>" & vbCrLf)
			End If
			Response.Write("</tr>" & vbCrLf)
		End If
		rs.MoveNext
	Wend
	Response.Write("<tbody>" & vbCrLf)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%	If ShowExcel = False Then %>
<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>
</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->