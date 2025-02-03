<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, PermitEdit, Columns, _
	FiscalYear, GranteeID, GranteeName, GrantID, PRogramName, IntegerTarget, DecimalTarget, _
	LastGoal, LastMandatory, LastStrategy
debug = False
Columns = 4
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

If MVCPARights = False Then
	Response.Write("Error: Only MVCPA Users are authorized to use this page.")
	SendMessage "Error: Only MVCPA Users are authorized to use this page."
	Response.End
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear")>0) Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
ElseIf Len(Session("FiscalYear"))>0 Then
	FiscalYear = CInt(Session("FiscalYear"))
Else 
	FiscalYear = 0
End If

If Len(Request.Form("GrantID"))>0 Then
	GrantID = CInt(Request.Form("GrantID"))
ElseIf Len(Request.QueryString("GrantID"))>0 Then
	GrantID = CInt(Request.QueryString("GrantID"))
Else
	GrantID = 0
End If

If IsNumeric(GrantID) = False Then
	Response.Write("Error: Non-numeric GrantID.")
	SendMessage "Error: Non-numeric GrantID."
	Response.End
End If
If Len(GrantID)>0 Then
	GrantID = CInt(GrantID)
Else
	GrantID=0
End If

If GrantID > 0 Then
	sql = "SELECT H.GrantID, H.FiscalYear, G.GranteeID, G.GranteeName, H.ProgramName " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"LEFT JOIN [Grants].Main AS H ON G.GranteeID=H.GranteeID " & vbCrLf & _
		"WHERE H.GrantID=" & prepIntegerSQL(GrantID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF Then
		Response.Write("Error: Grant not found.")
		Response.End
	Else
		GrantID = rs.Fields("GrantID")
		FiscalYear = rs.Fields("FiscalYear")
		GranteeID = rs.Fields("GranteeID")
		GranteeName = rs.Fields("GranteeName")
		ProgramName = rs.Fields("ProgramName")
	End If
End If

IF MVCPARights = True Then
	PermitEdit = True
Else
	PermitEdit = False
End If

If Debug = True Then
	Response.Write("<pre>PermitEdit=" & PermitEdit & ", CheckPermissions=" & _
	CheckPermissions(UserSystemID, GranteeID, True) & "</pre>")
	Response.Flush
End If	

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Progress Report for <%=GranteeName %> <%=ProgramName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style>
	tr, td, th {padding: 5px;}
</style>
<script type="text/javascript">
	function validateForm()
	{
		return true;
	}

	function submitForm()
	{
		validateForm();
		document.PR.submit();
	}

	function changedCurrencyField(field)
	{
		if (checkCurrency(field) == false)
			return false;
		return true;
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body style="width: 100%">
<div style="text-align: center;"><form name="Selection" method="post" action="Targets.asp">
<select name="FiscalYear" onchange="document.Selection.submit();">
<%
For i = 2018 to Application("CurrentFiscalYear")+1
	Response.Write(SelectOption( i, i, FiscalYear))
Next
%>
	</select>&nbsp;&nbsp;&nbsp;
<%
	sql = "SELECT GrantID, GranteeName + ' ' + ProgramName AS Description" & vbCrLf & _
		"FROM [Grants].Main AS A" & vbCrLf & _
		"JOIN Grantees AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
		"WHERE FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"ORDER BY REPLACE(GranteeName, 'City Of ','') "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

%>	<select name="GrantID" onchange="document.Selection.submit();">
		<option value="0">Select Grant</option>
<%
	If FiscalYear>0 Then
		Set rs = Con.Execute(sql)
		While rs.EOF = False
			Response.Write(SelectOption(rs.Fields("GrantID"), rs.Fields("Description"), GrantID) & vbCrLf)
			rs.MoveNext
		Wend
	End If
%>
	</select></form></div>
<h1><%=GranteeName %> MVCPA Progress Report Targets for Fiscal Year <%=FiscalYear %></h1>
<%
sql = "SELECT A.QuestionID, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.NoTarget, A.ResponseTypeID, " & vbCrLf & _
	"	IntegerTarget, DecimalTarget " & vbCrLf & _
	"FROM PR.Goals AS G " & vbCrLf & _
	"LEFT JOIN PR.Strategies AS S ON S.GoalID=G.GoalID AND G.Version=S.Version " & vbCrLf & _
	"LEFT JOIN PR.Activities AS A ON A.GoalID=S.GoalID AND S.StrategyID=A.StrategyID AND A.Version=S.Version " & vbCrLf & _
	"LEFT JOIN PR.GrantQuestions AS Q ON Q.GrantID=" & prepIntegerSQL(GrantID) & " AND Q.QuestionID=A.QuestionID " & vbCrLf & _
	"LEFT JOIN PR.Responses AS R ON R.GrantID=" & prepIntegerSQL(GrantID) & " AND R.QuestionID=A.QuestionID " & vbCrLf & _
	"WHERE Q.GrantID=" & prepIntegerSQL(GrantID) & " AND Mandatory=0 AND NoTarget=0 AND ResponseTypeID IN (1,2,3) " & vbCrLF & _
	"ORDER BY A.Mandatory DESC, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
%>

<form name="PR" method="post" action="TargetSubmit.asp">
<%=HiddenField("GrantID", GrantID) %><%=HiddenField("FiscalYear", FiscalYear) %>
<table style="margin: auto; width: 80%; ">
<thead>
	<tr>
		<th>ID</th>
		<th>Activity</th>
		<th>Measure</th>
		<th>Target</th>
	</tr>
</thead>
<%
LastGoal=0
LastStrategy=0
Set rs=Con.Execute(sql)
While rs.EOF = False
	'If LastMandatory <> rs.Fields("Mandatory") Then
	'	LastMandatory = rs.Fields("Mandatory")
	'	If LastMandatory = False Then
	'		Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: YellowGreen; "">Measures for Grantees. Add Target values for those that you will measure.</th></tr>" & vbCrLf)
	'	End If
	'End If
	If LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = False Then
		Response.Write("<tr>" & vbCrLf)
		LastGoal = rs.Fields("GoalID")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("GoalID") & "</td>" & vbCrLf)
		If rs.Fields("GoalID") < 4 Then
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PowderBlue;"">Goal " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
		Else
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PowderBlue;"">Section " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
	ElseIf LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = True Then
		LastGoal = rs.Fields("GoalID")
		If rs.Fields("GoalID") = 1 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for an MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Mandatory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=2 Then
			Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for an MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLf)
		ElseIf rs.Fields("GoalID")=8 Then
				Response.Write("<tr><td></td><th colspan=""" & (Columns - 1) & """ style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Fraud-Related Motor Vehicle Crime Measures must be reported during the grant term if awarded."">Mandatory Fraud-Related Motor Vehicle Crime Measures Required for all Grantees</th></tr>" & vbCrLf)
			End If
	End If

	' Strategy row
	If LastStrategy <> rs.Fields("StrategyID") And rs.Fields("Mandatory") = False  Then
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		LastStrategy = rs.Fields("StrategyID")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("GoalID") & "." & rs.Fields("StrategyID") & "</td>" & vbCrLf)
		If rs.Fields("GoalID") = 6 Then
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PeachPuff; "">Subsection " & rs.Fields("StrategyID") & "</th>" & vbCrLf)
		ElseIf rs.Fields("GoalID") = 4 Or rs.Fields("GoalID") = 5 Then
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PeachPuff; "">Subsection " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
		Else
			Response.Write("<th colspan=""" & (Columns - 1) & """ style=""background-color: PeachPuff; "">Strategy " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
	End If

	' Question row
	If rs.Fields("QuestionID") = 112 or rs.Fields("QuestionID") = 113 or rs.Fields("QuestionID") = 114 Then
		Response.Write("<tr style=""vertical-align: top; "" title=""Specific LBB Border Security Requirement"">" & vbCrLf)
	Else
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
	End If
	Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("MeasureNumber") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("Activity") & "</td>" & vbCrLf)
	' Measure cell adds the description of the reporting area for the required reporting.
	Response.Write(vbTab & "<td>" & rs.Fields("Measure")) 
	Response.Write("</td>" & vbCrLf)

	' Target cell
	If rs.Fields("Mandatory") = True Then
		Response.Write("<td>Required</td>")
	ElseIf rs.Fields("NoTarget") = True Then
		Response.Write("<td style=""background-color: #e6e6e6; text-align: center; "">No Target</td>" & vbCrLf)
	ElseIf rs.Fields("ResponseTypeID")=1 Then ' Integer
		Response.Write("<td>" & IntegerField("IntegerTarget_" & rs.Fields("QuestionID"), rs.Fields("IntegerTarget"), 8, 10, PermitEdit, "") & "</td>" & vbCrLf)
	ElseIf rs.Fields("ResponseTypeID")=2 Then ' Decimal
		Response.Write("<td>" & NumberField("DecimalTarget_" & rs.Fields("QuestionID"), formatnumber(rs.Fields("DecimalTarget"),2), 8, 10, PermitEdit, "") & "</td>" & vbCrLf)
	ElseIf rs.Fields("ResponseTypeID")=3 Then ' Money
		Response.Write("<td>" & NumberField("DecimalTarget_" & rs.Fields("QuestionID"), formatcurrency(rs.Fields("DecimalTarget"),2, True, False, True), 8, 10, PermitEdit, "") & "</td>" & vbCrLf)
	Else
		Response.Write(vbTab & "<td></td>" & vbCrLf)
	End If
	Response.Write("</tr>" & vbCrLf)
	rs.MoveNext()
Wend
Response.Write("</table>" & vbCrLf)

%>

<div style="text-align: center; margin: auto; ">
<%	If PermitEdit = True Or MVCPARights = True Then %>
	<input type="submit" name="Save" value="Save" title="Save changes" onclick="return submitForm();" />
<%	End If %>
	<input type="button" name="Close" value="Close" title="Ignore any pending changes and close window." 
		onclick="window.close();"/>
</div>
<!--<div style="text-align: right"><a href="Targets.asp?FiscalYear=<%=FiscalYear %>GrantID=<%=GrantID %>&ShowExcel=1" target="_blank">Excel</a></div>-->
</form>
</body>
</html>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->