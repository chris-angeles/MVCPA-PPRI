<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, k, PermitEdit, Version, AppID, GrantClassID, LastGoal, LastStrategy, LastMandatory,  _
	FiscalYear, GranteeID, GranteeName, SubmitID, SubmitByName, SubmitTimestamp, AppURL, _
	ApplicationSchema, NegotiationLocked

debug = False
ApplicationSchema = "Negotiation"

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
	Response.Write("</pre>" & vbCrLf)
	Response.Flush
End If

If Request.Form.Count>0 Then
	AppID = Request.Form("AppID")
	GrantClassID = Request.Form("GrantClassID")
Else
	AppID = Request.QueryString("AppID")
	GrantClassID = Request.QueryString("GrantClassID")
End If
IF Len(AppID)>0 Then
	AppID = CInt(AppID)
Else
	AppID=0
End If

sql = "SELECT I.AppID, I.GrantClassID, I.FiscalYear, G.GranteeID, G.GranteeName, I.GrantClassID, D.GrantClass," & vbCrLf & _
	"	TF1.SubmitID, U.Name AS SubmitByName, " & vbCrLf & _
	"	CASE WHEN I.GrantClassID=1 THEN TF1.SubmitTimestamp WHEN I.GrantClassID=4 THEN CC1.SubmitTimestamp ELSE NULL END AS SubmitTimestamp, " & vbCrLf & _
	"	CAST(CASE WHEN I.GrantClassID=1 THEN TF2.NegotiationLocked WHEN I.GrantClassID=4 THEN CC2.NegotiationLocked ELSE NULL END AS BIT) AS NegotiationLocked " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".Main AS TF1 ON TF1.AppID=I.AppID" & vbCrLf & _
	"LEFT JOIN Application.Admin AS TF2 ON TF2.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN CC." & ApplicationSchema & " AS CC1 ON CC1.AppID=I.AppID" & vbCrLf & _
	"LEFT JOIN CC.Admin AS CC2 ON CC2.AppID=I.AppID " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=CASE WHEN I.GrantClassID=1 THEN TF1.SubmitID WHEN I.GrantClassID=4 THEN CC1.SubmitID ELSE NULL END " & vbCrLf & _
	"LEFT JOIN Lookup.GrantClass AS D ON D.GrantClassID=I.GrantClassID " & vbCrLf & _
	"WHERE I.AppID=" & prepIntegerSQL(AppID)

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF Then
	Response.Write("Error: Application not found.")
	Response.End
Else
	AppID = rs.Fields("AppID")
	GrantClassID = rs.Fields("GrantClassID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	SubmitID = rs.Fields("SubmitID")
	SubmitByName = rs.Fields("SubmitByName")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	NegotiationLocked = rs.Fields("NegotiationLocked")
End If

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

If FiscalYear<2022 Then
	AppURL = "Application.asp"
Else
	AppURL = getHomeNegotiationReferenceByGrantClass(GrantClassID, AppID)
End If

If GranteeID>0 Then
	If NegotiationLocked = True Then
		PermitEdit = False
	ElseIf IsNull(SubmitID) = True Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, False)
	ElseIf IsNull(SubmitID) = False Then
		PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, True)
	Else
		PermitEdit = False
	End If
Else
		PermitEdit = False
End If

sql = "SELECT G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID AS MeasureID, " & vbCrLf & _
	"	CAST(G.GoalID AS VARCHAR) + '.' + CAST(S.StrategyID AS VARCHAR) + '.' + CAST(A.ActivityID AS VARCHAR) + " & vbCrLf & _
	"		CASE WHEN A.MeasureID=0 THEN '' ELSE '.' + CAST(A.MeasureID AS VARCHAR) END AS MeasureNumber, " & vbCrLf & _
	"	G.Goal, S.Strategy, A.Activity, A.Measure, A.Mandatory, A.ResponseTypeID, " & vbCrLf & _
	"	T.IntegerResponse, T.DecimalResponse " & vbCrLf & _
	"FROM Lookup.Goals AS G " & vbCrLf & _
	"LEFT JOIN Lookup.Strategies AS S ON S.Version=G.Version AND S.GoalID=G.GoalID " & vbCrLf & _
	"LEFT JOIN Lookup.Activities AS A ON A.Version=G.Version AND A.GoalID=S.GoalID AND A.StrategyID=S.StrategyID " & vbCrLf & _
	"LEFT JOIN " & ApplicationSchema & ".GSATargets AS T ON T.AppID=" & prepIntegerSQL(AppID) & " AND T.Version=G.Version AND T.GoalID=G.GoalID AND T.StrategyID=S.StrategyID AND T.ActivityID=A.ActivityID AND T.MeasureID=A.MeasureID " & vbCrLF & _
	"WHERE G.Version=" & prepIntegerSQL(Version) & " AND G.GoalID NOT IN (4,5,6,7) AND (Mandatory=1 OR NoTarget=0) " & vbCrLf & _
	"ORDER BY A.Mandatory DESC, G.Version, G.GoalID, S.StrategyID, A.ActivityID, A.MeasureID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Grant <%=ApplicationSchema %> for <%=GranteeName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<style>
	tr, td, th {padding: 5px;}
</style>
<script type="text/javascript">
	function validateForm()
	{
		//alert("validate!");
<% If GrantClassID = 1 Then %>
		if (!document.GSA.Target_1_1_17.selectedIndex > 0) {
			alert("Select the jurisdiction for the reported numbers");
			document.GSA.Target_1_1_17.focus();
			return false;
		}
		if (!document.GSA.Target_1_1_18.selectedIndex > 0) {
			alert("Select the jurisdiction for the reported numbers");
			document.GSA.Target_1_1_18.focus();
			return false;
		}
		if (!document.GSA.Target_1_1_19.selectedIndex > 0) {
			alert("Select the jurisdiction for the reported numbers");
			document.GSA.Target_1_1_19.focus();
			return false;
		}
		if (!document.GSA.Target_2_1_12.selectedIndex > 0) {
			alert("Select the jurisdiction for the reported numbers");
			document.GSA.Target_2_1_12.focus();
			return false;
		}
		if (!document.GSA.Target_2_1_13.selectedIndex > 0) {
			alert("Select the jurisdiction for the reported numbers");
			document.GSA.Target_2_1_13.focus();
			return false;
		}
<% End If %>		return true;
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
<h1><%=GranteeName %> MVCPA Grant <%=ApplicationSchema %> for Fiscal Year <%=FiscalYear %></h1>
<h2>Goals, Strategies, and Activities</h2>
<%	If SubmitID>0 Then %>
<p style="text-align: center; font-weight: bold; ">The Application was submitted by <%=SubmitByName%> at <%=SubmitTimestamp %> and is now locked.</p>
<%	End If %>
<form name="GSA" method="post" action="GSASubmit.asp" onsubmit="return validateForm()">
<%=HiddenField("AppID", AppID) %><%=HiddenField("Version", Version) %><%=HiddenField("FiscalYear", FiscalYear) %><%=HiddenField("GrantClassID", GrantClassID) %>
<table style="margin: auto">
<thead>
	<tr>
		<th>ID</th>
		<th>Activity</th>
		<th>Measure</th>
		<th>Target</th>
	</tr>
</thead>
<%
LastMandatory = True
LastGoal=0
LastStrategy=0
Set rs=Con.Execute(sql)
While rs.EOF = False
	If LastMandatory <> rs.Fields("Mandatory") Then
		LastMandatory = rs.Fields("Mandatory")
		If LastMandatory = False Then
			Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: YellowGreen; "">Measures for Grantees. Add Target values for those that you will measure.</th></tr>" & vbCrLF)
		End If
	End If
	If LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = False Then
		Response.Write("<tr>" & vbCrLf)
		LastGoal = rs.Fields("GoalID")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("GoalID") & "</td>" & vbCrLf)
		Response.Write("<th colspan=""3"" style=""background-color: PowderBlue;"">Goal " & rs.Fields("GoalID") & ": " & rs.Fields("Goal") & "</th>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
	ElseIf LastGoal <> rs.Fields("GoalID") And rs.Fields("Mandatory") = True Then
		LastGoal = rs.Fields("GoalID")
		If rs.Fields("GoalID") = 1 Then
			Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Theft must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data"">Mandatory Motor Vehicle Theft Measures Required for all Grantees.</th></tr>" & vbCrLF)
		ElseIf rs.Fields("GoalID")=2 Then
			Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Burglary of Motor Vehicle and Theft from a Motor Vehicle - Parts must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Burglary of a Motor Vehicle Measures Required for all Grantees</th></tr>" & vbCrLF)
		ElseIf rs.Fields("GoalID")=8 Then
			Response.Write("<tr><td></td><th colspan=""3"" style=""background-color: PaleGreen; "" title=""For law enforcement teams that apply for a MVCPA grant the following Motor Vehicle Fraud must be measured and reported during the grant term if awarded. Select the method by which the agency will collect and report the data."">Mandatory Motor Vehicle Fraud Measures Required for all Grantees</th></tr>" & vbCrLF)
		End If
	End If
	If LastStrategy <> rs.Fields("StrategyID") And rs.Fields("Mandatory") = False  Then
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		LastStrategy = rs.Fields("StrategyID")
		Response.Write("<td style=""text-align: right; "">" & rs.Fields("GoalID") & "." & rs.Fields("StrategyID") & "</td>" & vbCrLf)
		Response.Write("<th colspan=""3"" style=""background-color: PeachPuff; "">Strategy " & rs.Fields("StrategyID") & ": " & rs.Fields("Strategy") & "</th>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf)
	End If
	Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("MeasureNumber") & "</td>" & vbCrLF)
	Response.Write(vbTab & "<td>" & rs.Fields("Activity") & "</td>" & vbCrLF)
	Response.Write(vbTab & "<td>" & rs.Fields("Measure") & "</td>" & vbCrLF)
	If rs.Fields("Mandatory") and FiscalYear<2022 Then
		Response.Write(vbTab & "<td>Mandatory. Reporting for <br /><select name=""Target_" & Replace(rs.Fields("MeasureNumber"),".","_") & """>")
		Response.Write(vbTab & vbTab & vbTab & SelectOption(0, "Select Jurisdiction", rs.Fields("IntegerResponse")))
		Response.Write(vbTab & vbTab & vbTab & SelectOption(1, "Taskforce Only", rs.Fields("IntegerResponse")))
		Response.Write(vbTab & vbTab & vbTab & SelectOption(2, "Area of Jurisdiction", rs.Fields("IntegerResponse")))
		Response.Write(vbTab & vbTab & vbTab & SelectOption(3, "Combination of TF and Jurisdiction", rs.Fields("IntegerResponse")))
		Response.Write("</select></td>" & vbCrLf)
	ElseIf rs.Fields("Mandatory") Then
		Response.Write("<td></td>" & vbCrLf)
	ElseIf rs.Fields("ResponseTypeID")=1 Then
		Response.Write(vbTab & "<td>" & TextField("Target_" & Replace(rs.Fields("MeasureNumber"),".","_"),rs.Fields("IntegerResponse"),10, 8, PermitEdit, "return checkInteger(this);") & "</td>" & vbCrLF)
	ElseIf rs.Fields("ResponseTypeID")=2 Then
		If IsNull(rs.Fields("IntegerResponse")) = True Then
			Response.Write(vbTab & "<td>" & TextField("Target_" & Replace(rs.Fields("MeasureNumber"),".","_"),null ,10 , 8, PermitEdit, "return checkDecimal(this);") & "</td>" & vbCrLF)
		Else
			Response.Write(vbTab & "<td>" & TextField("Target_" & Replace(rs.Fields("MeasureNumber"),".","_"),formatnumber(rs.Fields("DecimalResponse"),2),10, 8, PermitEdit, "return checkDecimal(this);") & "</td>" & vbCrLF)
		End If
	ElseIf rs.Fields("ResponseTypeID")=3 Then
		If IsNull(rs.Fields("DecimalResponse")) Then
			Response.Write(vbTab & "<td>" & TextField("Target_" & Replace(rs.Fields("MeasureNumber"),".","_"),"",10, 8, PermitEdit, "return changedCurrencyField(this);") & "</td>" & vbCrLF)
		Else
			Response.Write(vbTab & "<td>" & TextField("Target_" & Replace(rs.Fields("MeasureNumber"),".","_"),"$" & formatnumber(rs.Fields("DecimalResponse"),2),10, 8, PermitEdit, "return changedCurrencyField(this);") & "</td>" & vbCrLF)
		End If
	End If
	Response.Write("</tr>" & vbCrLf)
	rs.MoveNext()
Wend

%>
</table>

<div style="text-align: center">
<%	If PErmitEdit = True Then %>
	<input type="submit" value="Save" title="Save changes and return to Application" />
	<input type="button" value="Cancel" title="Ignore changes and return to Application." 
		onclick="location.href = '<%=AppURL%>';"/>
<%	Else %>
	<input type="button" value="Return" title="Return to main page of Application." 
		onclick="location.href = '<%=AppURL%>';"/>
<%	End If %>
</div>
</form>
</body>
</html>
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/HomeRef.asp"-->
