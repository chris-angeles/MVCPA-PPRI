<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, FiscalYear, ShowRadioButtons, ShowText
Debug = False

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
	Response.Write("Now=" & Now() & vbCrLf)
	Response.Write("</pre>" & vbCrLf)
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))="1" Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear = Session("FiscalYear")
End If

If Request.Form("ShowRadioButtons")="1" Then
	ShowRadioButtons = True
ElseIf Request.QueryString("ShowRadioButtons")="1" Then
	ShowRadioButtons = True
Else
	ShowRadioButtons = False
End If

If Request.Form("ShowText")="1" Then
	ShowText = True
ElseIf Request.QueryString("ShowText")="1" Then
	ShowText = True
Else
	ShowText = False
End If

sql = "SELECT * " & vbCrLf & _
	"FROM [Grants].vwOperationalDataPlanData" & vbCrLf & _
	"WHERE FiscalYear=" & prepIntegerSQL(FiscalYear)

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Operational Plan Data retrieved for " & FiscalYear)
	SendMessage "Error: No Operational Plan Data retrieved for " & FiscalYear
	Response.End
Else

End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
	<title>MVCPA Taskforce Multi-Agency Grant Operational Plan</title>
	<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
	<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="widecontent">

<div style="margin: auto; text-align: center; "><form method="post" action="OperationalPlanDataReport.asp">
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2022 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<%=CheckBoxField("ShowRadioButtons", ShowRadioButtons) %> Show Radio Buttons&nbsp;&nbsp;
<%=CheckBoxField("ShowText", ShowText) %> Show Text Responses&nbsp;&nbsp;
<input type="submit" name="Submit" value="Submit" style="width: 60px; "/>
</form></div>
<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Operational Plan New Data Summary for FY<%=FiscalYear %></div>
<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 1 Co-location</div>

<p><b>Are members of the taskforce co-located?</b><br />
	<b><%=rs.Fields("Colocation_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CoLocation", "", 1)) %> All of the time &nbsp;&nbsp;
	<b><%=rs.Fields("Colocation_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CoLocation", "", 2)) %> Occasionally &nbsp;&nbsp;
	<b><%=rs.Fields("Colocation_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CoLocation", "", 3)) %> Never &nbsp;&nbsp;
	[No Response: <%=rs.Fields("Colocation_NR") %>]
</p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 2 Grantee and Subgrantee Meetings</div>

<p><b>By what primary method do scheduled meetings occur and how often are they held for those in 
	the GRANTEE agency only?</b><br />
	<b>Method:</b>
	 <b><%=rs.Fields("MeetingsGranteeMethod_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeMethod", "", 1)) %> In-Person &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeMethod", "", 2)) %> Virtual &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeMethod", "", 3)) %> EMail &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeMethod", "", 4)) %> Phone &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeMethod", "", 5)) %> Other &nbsp;&nbsp;
	[No Response: <%=rs.Fields("MeetingsGranteeMethod_NR") %>]
	<br />
	<b>Frequency:</b>
	<b><%=rs.Fields("MeetingsGranteeFrequency_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeFrequency", "", 1)) %> Daily &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeFrequency_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeFrequency", "", 2)) %> Weekly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeFrequency_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeFrequency", "", 3)) %> Every two weeks &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeFrequency_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeFrequency", "", 4)) %> Monthly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeFrequency_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeFrequency", "", 5)) %> Quarterly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeFrequency_6") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsGranteeFrequency", "", 6)) %> Yearly &nbsp;&nbsp;
	[No Response: <%=rs.Fields("MeetingsGranteeFrequency_NR") %>]
</p>

<p><b>By what primary method do scheduled meetings occur and how often are they held that 
	include the GRANTEE agency and INDIVIDUAL SUBGRANTEE agencies?</b><br />
	<b>Method:</b>
	<b><%=rs.Fields("MeetingsGranteeMethod_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeMethod", "", 1)) %> In-Person &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeMethod", "", 2)) %> Virtual &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeMethod", "", 3)) %> EMail &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeMethod", "", 4)) %> Phone &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsGranteeMethod_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeMethod", "", 5)) %> Other &nbsp;&nbsp;
	[No Response: <%=rs.Fields("MeetingsSubGranteeMethod_NR") %>]
	<br />
	<b>Frequency:</b>
	<b><%=rs.Fields("MeetingsSubGranteeFrequency_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeFrequency", "", 1)) %> Daily &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsSubGranteeFrequency_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeFrequency", "", 2)) %> Weekly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsSubGranteeFrequency_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeFrequency", "", 3)) %> Every two weeks &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsSubGranteeFrequency_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeFrequency", "", 4)) %> Monthly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsSubGranteeFrequency_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeFrequency", "", 5)) %> Quarterly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsSubGranteeFrequency_6") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsSubGranteeFrequency", "", 6)) %> Yearly &nbsp;&nbsp;
	[No Response: <%=rs.Fields("MeetingsSubGranteeMethod_NR") %>]
</p>

<p><b>By what primary method do scheduled meetings occur and how often are they held that 
	include the ENTIRE TASKFORCE?</b><br />
	<b>Method:</b>
	<b><%=rs.Fields("MeetingsAllTFMethod_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFMethod", "", 1)) %> In-Person &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFMethod_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFMethod", "", 2)) %> Virtual &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFMethod_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFMethod", "", 3)) %> EMail &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFMethod_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFMethod", "", 4)) %> Phone &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFMethod_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFMethod", "", 5)) %> Other &nbsp;&nbsp;
	[No Response: <%=rs.Fields("MeetingsAllTFMethod_NR") %>]
	<br />
	<b>Frequency:</b>
	<b><%=rs.Fields("MeetingsAllTFFrequency_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFFrequency", "", 1)) %> Daily &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFFrequency_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFFrequency", "", 2)) %> Weekly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFFrequency_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFFrequency", "", 3)) %> Every two weeks &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFFrequency_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFFrequency", "", 4)) %> Monthly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFFrequency_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFFrequency", "", 5)) %> Quarterly &nbsp;&nbsp;
	<b><%=rs.Fields("MeetingsAllTFFrequency_6") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("MeetingsAllTFFrequency", "", 6)) %> Yearly &nbsp;&nbsp;
	[No Response: <%=rs.Fields("MeetingsAllTFFrequency_NR") %>]
</p>

<p><b>Describe the taskforce meetings with grantee and subgrantee agencies. Include meeting 
	organization, attendees, information, operational issues and progress report and performance 
	data collection issue.</b><br />
	<%=rs.Fields("MeetingsDescription_Count") %> Responses
<%
	If ShowText Then ShowTextResponses("MeetingsDescription")
%></p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 3 Grantee and Subgrantee Contacts and Communication</div>

<p><b>By what primary method and how often does communication occur among those 
		in the GRANTEE agency only?</b><br />
	<b>Method:</b>
	<b><%=rs.Fields("CommunicationGranteeMethod_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeMethod", "", 1)) %> In-Person &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeMethod_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeMethod", "", 2)) %> Virtual &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeMethod_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeMethod", "", 3)) %> EMail &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeMethod_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeMethod", "", 4)) %> Phone &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeMethod_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeMethod", "", 5)) %> Other &nbsp;&nbsp;
	[No Response: <%=rs.Fields("CommunicationGranteeMethod_NR") %>]
	<br />
	<b>Frequency:</b>
	<b><%=rs.Fields("CommunicationGranteeFrequency_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeFrequency", "", 1)) %> Daily &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeFrequency_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeFrequency", "", 2)) %> Weekly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeFrequency_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeFrequency", "", 3)) %> Every two weeks &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeFrequency_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeFrequency", "", 4)) %> Monthly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeFrequency_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeFrequency", "", 5)) %> Quarterly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationGranteeFrequency_6") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationGranteeFrequency", "", 6)) %> Yearly &nbsp;&nbsp;
	[No Response: <%=rs.Fields("CommunicationGranteeFrequency_NR") %>]
</p>

<p><b>By what primary method and how often does communication occur that include 
	the GRANTEE agency and INDIVIDUAL SUBGRANTEES?</b><br />
	<b>Method:</b>
	<b><%=rs.Fields("CommunicationSubGranteeMethod_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeMethod", "", 1)) %> In-Person &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeMethod_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeMethod", "", 2)) %> Virtual &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeMethod_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeMethod", "", 3)) %> EMail &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeMethod_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeMethod", "", 4)) %> Phone &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeMethod_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeMethod", "", 5)) %> Other &nbsp;&nbsp;
	[No Response: <%=rs.Fields("CommunicationSubGranteeMethod_NR") %>]
	<br />
	<b>Frequency:</b>
	<b><%=rs.Fields("CommunicationSubGranteeFrequency_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeFrequency", "", 1)) %> Daily &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeFrequency_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeFrequency", "", 2)) %> Weekly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeFrequency_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeFrequency", "", 3)) %> Every two weeks &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeFrequency_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeFrequency", "", 4)) %> Monthly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeFrequency_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeFrequency", "", 5)) %> Quarterly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationSubGranteeFrequency_6") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationSubGranteeFrequency", "", 6)) %> Yearly &nbsp;&nbsp;
	[No Response: <%=rs.Fields("CommunicationSubGranteeFrequency_NR") %>]
</p>

<p><b>By what primary method and how often does communication occur that include 
	the ENTIRE TASKFORCE?</b><br />
	<b>Method:</b>
	<b><%=rs.Fields("CommunicationAllTFMethod_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFMethod", "", 1)) %> In-Person &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFMethod_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFMethod", "", 2)) %> Virtual &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFMethod_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFMethod", "", 3)) %> EMail &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFMethod_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFMethod", "", 4)) %> Phone &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFMethod_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFMethod", "", 5)) %> Other &nbsp;&nbsp;
	[No Response: <%=rs.Fields("CommunicationAllTFMethod_NR") %>]
	<br />
	<b>Frequency:</b>
	<b><%=rs.Fields("CommunicationAllTFFrequency_1") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFFrequency", "", 1)) %> Daily &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFFrequency_2") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFFrequency", "", 2)) %> Weekly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFFrequency_3") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFFrequency", "", 3)) %> Every two weeks &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFFrequency_4") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFFrequency", "", 4)) %> Monthly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFFrequency_5") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFFrequency", "", 5)) %> Quarterly &nbsp;&nbsp;
	<b><%=rs.Fields("CommunicationAllTFFrequency_6") %></b><%If ShowRadioButtons = True Then Response.Write(RadioInputField("CommunicationAllTFFrequency", "", 6)) %> Yearly &nbsp;&nbsp;
	[No Response: <%=rs.Fields("CommunicationAllTFFrequency_NR") %>]
</p>

<p><b>Describe the taskforce communication with grantee and subgrantee agencies. 
	Include regular, occasional and ad hoc communication about cases, reporting, and trends.</b><br />
	<%=rs.Fields("CommunicationDescription_Count") %> Responses
<%
	If ShowText Then ShowTextResponses("CommunicationDescription")
%></p>
<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 4 Coverage Agency Meetings</div>
<br />

<p><b>Describe meetings that grantee and subgrantee agencies perform with or for coverage agencies. 
	Include purpose, method and frequency of meetings.</b><br />
	<%=rs.Fields("CoverageAgencyMeetings_Count") %> Responses
<%
	If ShowText Then ShowTextResponses("CoverageAgencyMeetings")
%></p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 5 Coverage Agency Contacts</div>

<p><b>Describe contact that grantee and subgrantee have with coverage agencies. 
Include purpose, method and frequency of contact. </b><br />
	<%=rs.Fields("CoverageAgencyContacts_Count") %> Responses
<%
	If ShowText Then ShowTextResponses("CoverageAgencyContacts")
%></p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 6 Intelligence Sharing</div>

<br />

<p><b>Describe a plan to develop, collect, process, disseminate, and receive feedback, 
	intelligence information.  Describe who (sub grantee, coverage agencies, and or other) and 
	how the intelligence is disseminated. Is the information posted to the Virtual Command Center?</b><br />
	<%=rs.Fields("IntelligenceSharing_Count") %> Responses
<%
	If ShowText Then ShowTextResponses("IntelligenceSharing")
%></p>

<br />

<div style="width: 976px; text-align: center; font-weight: bold; ">Section 7 Operational and Investigative Coordination</div>

<br />

<p><b>Describe how cases are assigned to taskforce personnel.  Include if subgrantees are assigned cases from 
	taskforce commander, the sub grantee agency, or both.</b><br />
	<%=rs.Fields("OperationalCoordination_Count") %> Responses
<%
	If ShowText Then ShowTextResponses("OperationalCoordination")
%></p>

<br />


<div style="width: 976px; text-align: center; font-weight: bold; ">Section 8 Direct Operations</div>

<br />

<p><b>Describe how taskforce personnel conduct operations/activities as a group.  
	Include how and what types of operations occur in participating agency jurisdictions.  
	Include how border/bridge  and port operations are coordinated and occur (planned and unplanned) if applicable.</b><br />
	<%=rs.Fields("DirectOperatations_Count") %> Responses
<%
	If ShowText Then ShowTextResponses("DirectOperatations")
%></p>

<br />


</div>
<br />

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
<script type="text/javascript">

	function validateForm()
	{
		return true;
	}

	function checkTypes()
	{
		// Add validation for things that are required to save and avoid an error.
		document.Application.ProgramName.value = replaceWordChars(document.Application.ProgramName.value);
		document.Application.OtherCoverageText.value = replaceWordChars(document.Application.OtherCoverageText.value);
		return true;
	}
</script>
<script src="../includes/formchanges.js"></script>
<script type="text/javascript">
	var saving = false;
	var form = document.getElementById("Application");

	// form being updated
	form.onsubmit = function () { saving = true; };

	// form not saved warning
	/*window.onunload = function() {
		if (!saving) {
			var f = FormChanges(form);
			if (f.length > 0) 
			{
				if (window.confirm("Your form updates have not be saved. Do you wish to continue without saving?"))
					return true;
				else
					return false;
			}
		}
	};*/

	// show changed messages
	function DetectChanges()
	{
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += "\n" + f[e].id;
		alert((msg ? "Elements changed:" : "No changes made.") + msg);
	}

	// Save changes
	function SaveChanges()
	{
		var f = FormChanges(form), msg = "";
		for (var e = 0, el = f.length; e < el; e++) msg += f[e].id + "\n";
		document.Application.Changes.value = msg;
	}

</script>
</body>
</html>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<%
Sub ShowTextResponses(vQuestionName)
	Dim vsql, vrs
	vsql = "SELECT REPLACE(G.GranteeName, 'City of ','') AS Grantee, " & vQuestionName & " AS Response " & vbCrLf & _
		"FROM Application.IDs AS I " & vbCrLf & _
		"LEFT JOIN Application.Main AS A ON A.AppID=I.AppID " & vbCrLf & _
		"JOIN Grantees AS G ON G.GranteeID=I.GranteeID " & vbCrLf & _
		"JOIN [Grants].OperationalPlan AS P ON P.AppID=A.AppID " & vbCrLf & _
		"WHERE I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"ORDER BY REPLACE(G.GranteeName, 'City of ','')"
	If Debug = True Then
		Response.Write("<!-- " & vsql & " -->" & vbCrLf)
		Response.Flush
	End If
	Response.Write("<br >" & vbCrLf)
	Set vrs = Con.Execute(vsql)
	While vrs.EOF = False
		Response.Write("<b>" & vrs.Fields("Grantee") & "</b>: <i>" & vrs.Fields("Response") & "</i><br />" & vbCrLf)
		vrs.MoveNext
	Wend
End Sub
%>
