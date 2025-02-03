<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, AppID, FiscalYear, ProgramName, GranteeName, ApplicationSchema, GrantClassID, GrantClass, Version,  AppURL
Debug = False
ApplicationSchema = "Negotiation"

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

AppID = Request.QueryString("AppID")
GrantClassID = CInt(Request.QueryString("GrantClassID"))

If GrantClassID = 4 Then
	Version = 1001
ElseIf FiscalYear>= 2022 Then
	Version = 5
ElseIf FiscalYear>= 2021 Then
	Version = 4
ElseIf FiscalYear>= 2020 Then
	Version = 3 ' Had been 2 until 5/6/2021
ElseIf FiscalYear>= 2018 Then
	Version = 2
Else
	Version = 1
End If

sql = "SELECT I.AppID, I.FiscalYear, G.GranteeID, G.GranteeName, I.GrantClassID, D.GrantClass," & vbCrLf & _
	"	A.ProgramName, A.SubmitID, U.Name AS SubmitByName, A.SubmitTimestamp " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".Main AS A ON A.AppID=I.AppID" & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=A.SubmitID " & vbCrLf & _
	"LEFT JOIN Lookup.GrantClass AS D ON D.GrantClassID=I.GrantClassID " & vbCrLf & _
	"WHERE I.AppID=" & prepIntegerSQL(AppID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs=Con.Execute(sql)
If rs.EOF = False Then
	FiscalYear = rs.Fields("FiscalYear")
	ProgramName = rs.Fields("ProgramName")
	GranteeName = rs.Fields("GranteeName")
End If

If FiscalYear<2022 Then
	AppURL = "Application.asp"
Else
	AppURL = getHomeNegotiationReferenceByGrantClass(GrantClassID, AppID)
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Goals, Strategy, and Activity Targets</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style type="text/css">	th {
		text-align: center;
	}
</style>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag"><%=GranteeName %> <%=GrantClass %>&nbsp;<%=ApplicationSchema %> for Fiscal Year <%=FiscalYear %>: Goals, Strategies, and Activities Targets</div>

<div class="widecontent">
<br />
<%
outputGSA()
%>
</body>
</html>
<%

Function outputGSA()
	Response.Write("<table style=""margin: auto""><thead><tr><th>ID</th><th>Activity</th><th>Measure</th><th>Target</th></tr></thead>" & vbCrLf)
	Dim vrs, vsql, LastMandatory, LastGoal, LastStrategy
	vsql = "SELECT G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
		"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
		"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
		"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
		"	T.IntegerResponse, T.DecimalResponse " & vbCrLf & _
		"FROM Lookup.Goals AS G " & vbCrLf & _
		"LEFT JOIN Lookup.Strategies AS S ON S.Version=G.Version AND S.GoalID=G.GoalID " & vbCrLf & _
		"LEFT JOIN Lookup.Activities AS A ON A.Version=G.Version AND A.GoalID=S.GoalID AND A.StrategyID=S.StrategyID " & vbCrLf & _
		"LEFT JOIN " & ApplicationSchema & ".GSATargets AS T ON T.AppID=" & prepIntegerSQL(AppID) & " AND T.Version=G.Version AND T.GoalID=G.GoalID AND T.StrategyID=S.StrategyID AND T.ActivityID=A.ActivityID AND T.MeasureID=A.MeasureID " & vbCrLF & _
		"WHERE G.Version=" & prepIntegerSQL(Version) & " AND G.GoalID NOT IN (4,5, 6, 7) AND (Mandatory=1 OR NoTarget=0) " & vbCrLf & _
		"ORDER BY A.Mandatory DESC, G.Version, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	LastMandatory = True
	LastGoal=0
	LastStrategy=0
	Set vrs=Con.Execute(vsql)
	While vrs.EOF = False
		If LastMandatory <> vrs.Fields("Mandatory") Then
			LastMandatory = vrs.Fields("Mandatory")
			If LastMandatory = False Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: YellowGreen; "">Measures for Grantees. Add Target values for those that you will measure.</th></tr>" & vbCrLF)
			End If
		End If
		If LastGoal <> vrs.Fields("GoalID") And vrs.Fields("Mandatory") = False Then
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			LastGoal = vrs.Fields("GoalID")
			Response.Write("<td style=""text-align: right; "">" & vrs.Fields("GoalID") & "</td>" & vbCrLf)
			Response.Write("<th colspan=""3"" style=""background-color: PowderBlue;"">Goal " & vrs.Fields("GoalID") & ": " & vrs.Fields("Goal") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		ElseIf LastGoal <> vrs.Fields("GoalID") And vrs.Fields("Mandatory") = True Then
			LastGoal = vrs.Fields("GoalID")
			If vrs.Fields("GoalID") = 1 Then
				Response.Write("<tr style=""vertical-align: top; ""><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Statutory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLF)
			ElseIf vrs.Fields("GoalID")=2 Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Statutory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLF)
			ElseIf vrs.Fields("GoalID")=8 Then
				Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Fraud must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Statutory Fraud-Related Motor Vehicle Crime Measures Required for all Grantees</th></tr>" & vbCrLF)
			End If
		End If
		If LastStrategy <> vrs.Fields("StrategyID") And vrs.Fields("Mandatory") = False  Then
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			LastStrategy = vrs.Fields("StrategyID")
			Response.Write("<td style=""text-align: right; "">" & vrs.Fields("GoalID") & "." & vrs.Fields("StrategyID") & "</td>" & vbCrLf)
			Response.Write("<th colspan=""3"" style=""background-color: PeachPuff; "">Strategy " & vrs.Fields("StrategyID") & ": " & vrs.Fields("Strategy") & "</th>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & vrs.Fields("MeasureNumber") & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td>" & vrs.Fields("Activity") & "</td>" & vbCrLF)
		Response.Write(vbTab & "<td>" & vrs.Fields("Measure") & "</td>" & vbCrLF)
		If vrs.Fields("Mandatory") Then
			Response.Write(vbTab & "<td class=""usertext""></td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=1 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & vrs.Fields("IntegerResponse") & "</td>" & vbCrLF)
		ElseIf vrs.Fields("ResponseTypeID")=2 Then
			Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLf)
		ElseIf vrs.Fields("ResponseTypeID")=3 Then
				Response.Write(vbTab & "<td style=""text-align: right"" class=""usertext"">" & formatnumber(vrs.Fields("DecimalResponse")) & "</td>" & vbCrLF)
		End If
		Response.Write("</tr>" & vbCrLf)
		vrs.MoveNext()
	Wend
	Response.Write("</table>" & vbCrLf)
	Response.Write("<br />")
End Function

function textarea2html(vText)
	if IsNull(vText) = true Then
		textarea2html = null
	ElseIf Len(vText)=0 Then
		textarea2html = ""
	Else
		textarea2html = Replace(vText, vbCrLf&vbCrLf, "<br /><br />")
	End If
end function
%><!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/HomeRef.asp"-->
