<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 

Dim debug, i, j, PermitEdit, CanSubmit, AllowUpload, Submitted, _
	RRSID, ProgramName, GranteeID, FiscalYear, GranteeName, ORI, Agency, AuthorizedOfficialID, StatePayeeIDNo, _
	ProposedStartDate, ProposedEndDate, Duration, SituationOverview, ProjectDescription

debug = False

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
				response.write("Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j)) & vbCrLf
			next
		else
			Response.Write("Cookies(""" & i & """)=" & Request.Cookies(i) & "<br>")
		end if
	next
	Response.Write("Now=" & Now() & vbCrLf)
	Response.Write("</pre>" & vbCrLf)
End If

If Len(Request.Form("RRSID")) > 0 Then
	RRSID = CLng(Request.Form("RRSID"))
ElseIf Len(Request.QueryString("RRSID")) > 0 Then
	RRSID = CLng(Request.QueryString("RRSID"))
Else
	RRSID = 0
End If

If Len(Request.Form("GranteeID")) > 0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID")) > 0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = 0
End If

If Len(Request.Form("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear = 0
End If

If (GranteeID=0 And FiscalYear=0) And RRSID=0 Then
	Response.Write("Error: No GranteeID and fiscal yer or RRSID Specified")
	SendMessage "Error: No GranteeID and fiscal yer or RRSID Specified"
	Response.End
End If

sql = "SELECT ISNULL(B.RRSID,0) AS RRSID, A.GranteeID, A.GranteeName, B.FiscalYear, B.ProgramName, A.ORI, C.Agency, A.StatePayeeIDNo, " & vbCrLF & _
	"	A.AuthorizedOfficialID, B.ProposedStartDate, B.ProposedEndDate, B.Duration, B.SituationOverview, " & vbCrLf & _
	"	B.ProjectDescription " & vbCrLf & _
	"FROM Grantees AS A" & vbCrLf & _
	"LEFT JOIN RRS.Main AS B ON A.GranteeID=B.GranteeID " & vbCrLf & _
	"LEFT JOIN Lookup.ORI AS C ON C.ORI=A.ORI " & vbCrLf
If RRSID > 0 Then
	sql = sql & "WHERE B.RRSID=" & prepIntegerSQL(RRSID)
Else
	sql = sql & "WHERE A.GranteeID=" & prepIntegerSQL(GranteeID) & " AND FiscalYear=" & prepIntegerSQL(FiscalYear)
End If
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: Grantee record for GranteeID=" & GranteeID & " not retrieved.")
	SendMessage "Error: Grantee record for GranteeID=" & GranteeID & " not retrieved."
	Response.End
Else
	RRSID = rs.Fields("RRSID")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	FiscalYear = rs.Fields("FiscalYear")
	ProgramName = rs.Fields("ProgramName")
	ORI = rs.Fields("ORI")
	StatePayeeIDNo = rs.Fields("StatePayeeIDNo")
	Agency = rs.Fields("Agency")
	AuthorizedOfficialID = rs.Fields("AuthorizedOfficialID")
	ProposedStartDate = rs.Fields("ProposedStartDate")
	ProposedEndDate = rs.Fields("ProposedEndDate")
	Duration = rs.Fields("Duration")
	SituationOverview = rs.Fields("SituationOverview")
End If

If Debug = True Then
	Response.Write("<pre>Captured database variables</pre>" & vbCrLf)
	Response.Flush
End If

PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, Submitted)
AllowUpload = CheckPermissions(UserSystemID, GranteeID, True) ' Allow Upload after submission.

If PermitEdit = False Then
	CanSubmit = False
ElseIf Submitted = True Then
	CanSubmit = False
ElseIf AuthorizedOfficialID = UserSystemID Then
	CanSubmit = True
Else
	CanSubmit = False
End If

If Debug = True Then
	Response.Write("<pre>PermitEdit=" & PermitEdit & ": AllowUpload=" & AllowUpload & ": CanSubmit=" & CanSubmit & ";</pre>" & vbCrLf)
	Response.Flush
End If

If Debug = True Then
	Response.Write("<pre>Start HTML</pre>" & vbCrLf)
	Response.Flush
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Auxiliary Grant Application</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" />
<style>
	table {
		border: 3px solid black;
		border-collapse: collapse;
	}
td, th {
	border: 1px solid black;
	border-collapse: collapse;
}
</style>
<script type="text/javascript">
	function submitForm(action) {
		return true;
	}

	function calculateDuration() {
		var start, end, daydiff, weekdiff, monthdiff;
		start = new Date(document.RRSApplication.ProposedStartDate.value);
		end = new Date(document.RRSApplication.ProposedEndDate.value);
		daydiff = Math.round((end - start) / (1000 * 60 * 60 * 24));
		weekdiff = Math.round((end - start) / (1000 * 60 * 60 * 24 * 7));
		monthdiff = Math.round((end - start) / (1000 * 60 * 60 * 24 * 30));
		//alert("daydiff=" + daydiff.toString());
		if (monthdiff > 5)
			document.RRSApplication.Duration.value = monthdiff.toString() + " months";
		else if (weekdiff > 5)
			document.RRSApplication.Duration.value = weekdiff.toString() + " weeks";
		else
			document.RRSApplication.Duration.value = daydiff.toString() + " days";
	}
</script>
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<br />
<div class="pagetag" style="padding-top: 4px; padding-bottom: 4px; "><%=GranteeName %> Rapid Response Stikeforce Grant Application</div>

<div class="widecontent">

<form name="RRSApplication" id="RRSApplication" method="post" action="RRSApplicationSubmit.asp">
<table style="margin: auto; width: 100%; ">

<tr>
	<td colspan="2"><span style="font-weight: bold;">Program Name:</span> 
		<input type="text" name="ProgramName" value="<%=ProgramName %>" size="50" maxlength="100" /></td>
	<td><span style="font-weight: bold;">Fiscal Year:</span>
		<input type="text" name="FiscalYear" value="<%=FiscalYear %>" size="4" maxlength="4" /></td>
</tr>

<tr style="border-top: 3px solid black; ">
	<th>Grantee/Administrative Agency Name:</th>
	<th>ORI #</th>
	<th>Vendor Number</th>
</tr>

<tr>
	<td><%=GranteeName %></td>
	<td><%=ORI %>&nbsp;<%=Agency %></td>
	<td><%=StatePayeeIDNo %></td>
</tr>

<tr>
	<th>Participating Agency(ies) Name:</th>
	<td></td>
	<td></td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td></td>
	<td></td>
</tr>

<tr style="border-top: 3px solid black; ">
	<th colspan="3">Proposed Term of Grant</th>
</tr>

<tr>
	<th>Proposed Start Date</th>
	<th>Proposed End Date</th>
	<th>Proposed Duration</th>
</tr>

<tr>
	<td style="text-align: center; "><%=DateFieldChange("ProposedStartDate", ProposedStartDate, "calculateDuration();", PermitEdit)%></td>
	<td style="text-align: center; "><%=DateFieldChange("ProposedEndDate", ProposedEndDate, "calculateDuration();", PermitEdit)%></td>
	<td style="text-align: center; "><input type="text" name="Duration" value="<%=Duration %>" readonly="readonly" style="text-align: center; border: none;"/></td>
</tr>

<tr style="border-top: 3px solid black; ">
	<th colspan="3">Describe Emergency or Exigent Situation and Overview of Proposed RRS Operation</th>
</tr>

<tr>
	<td colspan="3" style="text-align: center; "><%=TextArea2("SituationOverview", SituationOverview, 10, 960, 1996, PermitEdit, "") %></td>
</tr>

<tr style="border-top: 3px solid black; ">
	<th colspan="3">Additional Area of RRS Operation (counties/cities) <span style="font-size: smaller; ">(only complete if not covered above):</span></th>
</tr>

<tr>
	<th>Counties</th>
	<th>Cities</th>
	<th rowspan="4"></th>
</tr>

<tr>
	<td>&nbsp;</td>
	<td></td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td></td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td></td>
</tr>

<tr style="border-top: 3px solid black; border-bottom: none;">
	<th colspan="3" style="border-bottom: none;">Summary of RRS Resources Requested / Provided</th>
</tr>

<tr style="border-top: none;">
	<th style="border-top: none;">Type of resource requested</th>
	<th style="border-top: none;">Grant Resource Needed</th>
	<th style="border-top: none;">Match Resource Needed</th>
</tr>

<tr>
	<td># of personnel</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<td>Overtime Units (estimate hours)</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<td rowspan="2">List type of equipment requested for purchase (surveillance, LPR, bait, etc.}</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<td></td>
	<td></td>
</tr>
<tr>
	<td>Travel Costs</td>
	<td></td>
	<td></td>
</tr>

<tr style="border-top: 3px solid black; ">
	<th colspan="3">Rapid Response Strikeforce Grant Budget Summary</th>
</tr>

<tr>
	<th>&nbsp;</th>
	<th>Amount RSS Funds Requested</th>
	<th>20% Match Provided (Required)</th>
</tr>

<tr>
	<td>Personnel</td>
	<td style="text-align: center; "><span style="font-size: smaller; ">Not Allowed in RRS Reimbursement</span></td>
	<td></td>
</tr>

<tr>
	<td>Fringe</td>
	<td style="text-align: center; "><span style="font-size: smaller; ">Not Allowed in RRS Reimbursement</span></td>
	<td></td>
</tr>

<tr>
	<td>Overtime</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<td>Professional and Contract Services</td>
	<td style="text-align: center; "><span style="font-size: smaller; ">Not Allowed in RRS Reimbursement</span></td>
	<td></td>
</tr>

<tr>
	<td>Travel</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<td>Equipment Costs</td>
	<td></td>
	<td></td>
</tr>

<tr>
	<th>Total Amount of funds Requested/Provided:</th>
	<td></td>
	<td></td>
</tr>

<tr>
	<th colspan="3" style="border-top: 3px solid black; ">Describe the activity/response/equip,ent requested. Include description of the match resource(s) proposed:
		<br /><span style="font-size: smaller; ">(Taskforce program income cannot be used)</span></th>
</tr>

<tr>
	<td colspan="3" style="text-align: center; "><%=TextArea2("ProjectDescription", ProjectDescription, 20, 960, 1996, PermitEdit, "") %></td>
</tr>
</table>
</form>
</body>
</html>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/InputValidation.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->