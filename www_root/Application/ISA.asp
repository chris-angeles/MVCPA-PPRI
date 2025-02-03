<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, ISAID, FiscalYear, GranteeID, GranteeName, CoverageArea, _
	HistoricalDataYear, MVT1, MVT2, MVT3, MVTLoss1, MVTLoss2, MVTLoss3, _
	BMV1, BMV2, BMV3, BMVLoss1, BMVLoss2, BMVLoss3, _
	ReceiveAuthorization, RegionalTaskForce, CurrentProgram, CurrentProgramDescription, _
	PreviouslyApplied, PreviouslyAwarded, TerminationExplanation, MeetCashMatchRequirement, _
	DedicateResources, NoSupplantation, ProgramCategoryID, BriefNarrative, GrantRequestRangeID, _
	SubmitID, SubmitTimestamp, SubmitByName, DateReviewed, DateResponded, MethodRespondedID, Notes, _
	UpdateID, UpdateTimestamp
debug = False
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
	Response.Write("</pre>" & vbCrLF)
End If

ISAID = Request.QueryString("ISAID")
GranteeID = Request.QueryString("GranteeID")
FiscalYear = Request.QueryString("FiscalYear")

If ISAID="" Then
	ISAID=0
	If GranteeID="" Then
		GranteeID = Session("GranteeID")
	End If
	If GranteeID="" Or GranteeID=0 Then
		Response.Write("Error: No ISAID or GranteeID Specified")
		Response.End
	End If
Else
	ISAID=Cint(ISAID)
End If
If FiscalYear="" Then
	FiscalYear = Year(Now) + 1
End If

If ISAID = 0 Then
	sql = "SELECT G.GranteeID, G.GranteeName, I.ISAID, I.FiscalYear, I.CoverageArea, I.HistoricalDataYear, " & vbCrLf & _
		"	I.MVT1, I.MVT2, I.MVT3, I.MVTLoss1, I.MVTLoss2, I.MVTLoss3, I.BMV1, I.BMV2, I.BMV3, I.BMVLoss1, I.BMVLoss2, I.BMVLoss3, " & vbCrLf & _
		"	I.ReceiveAuthorization, I.RegionalTaskForce, I.CurrentProgram, I.CurrentProgramDescription, I.PreviouslyApplied, " & vbCrLF & _
		"	I.PreviouslyAwarded, I.TerminationExplanation, I.MeetCashMatchRequirement, I.DedicateResources, I.NoSupplantation, " & vbCrLF & _
		"	I.ProgramCategoryID, I.BriefNarrative, I.GrantRequestRangeID, I.SubmitID, I.SubmitTimestamp, I.DateReviewed, " & vbCrLF & _
		"	I.DateResponded, I.MethodRespondedID, I.Notes, I.UpdateID, I.UpdateTimestamp, null AS SubmitByName " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"LEFT JOIN ISA AS I ON I.GranteeID=G.GranteeID AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"WHERE G.GranteeID=" & prepIntegerSQL(GranteeID)
Else
	sql = "SELECT G.GranteeID, G.GranteeName, I.ISAID, I.FiscalYear, I.CoverageArea, I.HistoricalDataYear, " & vbCrLf & _
		"	I.MVT1, I.MVT2, I.MVT3, I.MVTLoss1, I.MVTLoss2, I.MVTLoss3, I.BMV1, I.BMV2, I.BMV3, I.BMVLoss1, I.BMVLoss2, I.BMVLoss3, " & vbCrLf & _
		"	I.ReceiveAuthorization, I.RegionalTaskForce, I.CurrentProgram, I.CurrentProgramDescription, I.PreviouslyApplied, " & vbCrLF & _
		"	I.PreviouslyAwarded, I.TerminationExplanation, I.MeetCashMatchRequirement, I.DedicateResources, I.NoSupplantation, " & vbCrLF & _
		"	I.ProgramCategoryID, I.BriefNarrative, I.GrantRequestRangeID, I.SubmitID, I.SubmitTimestamp, I.DateReviewed, " & vbCrLF & _
		"	I.DateResponded, I.MethodRespondedID, I.Notes, I.UpdateID, I.UpdateTimestamp, SB.Name AS SubmitByName " & vbCrLf & _
		"FROM ISA AS I " & vbCrLf & _
		"LEFT JOIN Grantees AS G ON I.GranteeID=G.GranteeID " & vbCrLf & _
		"LEFT JOIN System.Users AS SB ON SB.SystemID=I.SubmitID " & vbCrLF & _
		"WHERE I.ISAID=" & prepIntegerSQL(ISAID)
End If
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)
If rs.EOF = False Then
	'ISAID = rs.Fields("ISAID")
	If IsNull(rs.Fields("FiscalYear"))=False Then
		FiscalYear = rs.Fields("FiscalYear")
	End If
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	CoverageArea = rs.Fields("CoverageArea")
	HistoricalDataYear = rs.Fields("HistoricalDataYear")
	If IsNull(HistoricalDataYear) Then
		HistoricalDataYear = FiscalYear - 2
	End If
	MVT1 = rs.Fields("MVT1")
	MVT2 = rs.Fields("MVT2")
	MVT3 = rs.Fields("MVT3")
	MVTLoss1 = rs.Fields("MVTLoss1")
	MVTLoss2 = rs.Fields("MVTLoss2")
	MVTLoss3 = rs.Fields("MVTLoss3")
	BMV1 = rs.Fields("BMV1")
	BMV2 = rs.Fields("BMV2")
	BMV3 = rs.Fields("BMV3")
	BMVLoss1 = rs.Fields("BMVLoss1")
	BMVLoss2 = rs.Fields("BMVLoss2")
	BMVLoss3 = rs.Fields("BMVLoss3")
	ReceiveAuthorization = rs.Fields("ReceiveAuthorization")
	RegionalTaskForce = rs.Fields("RegionalTaskForce")
	CurrentProgram = rs.Fields("CurrentProgram")
	CurrentProgramDescription = rs.Fields("CurrentProgramDescription")
	PreviouslyApplied = rs.Fields("PreviouslyApplied")
	PreviouslyAwarded = rs.Fields("PreviouslyAwarded")
	TerminationExplanation = rs.Fields("TerminationExplanation")
	MeetCashMatchRequirement = rs.Fields("MeetCashMatchRequirement")
	DedicateResources = rs.Fields("DedicateResources")
	NoSupplantation = rs.Fields("NoSupplantation")
	ProgramCategoryID = rs.Fields("ProgramCategoryID")
	BriefNarrative = rs.Fields("BriefNarrative")
	GrantRequestRangeID = rs.Fields("GrantRequestRangeID")
	SubmitID = rs.Fields("SubmitID")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	SubmitByName = rs.Fields("SubmitByName")
	DateReviewed = rs.Fields("DateReviewed")
	DateResponded = rs.Fields("DateResponded")
	MethodRespondedID = rs.Fields("MethodRespondedID")
	Notes = rs.Fields("Notes")
	UpdateID = rs.Fields("UpdateID")
	UpdateTimestamp = rs.Fields("UpdateTimestamp")
Else
	FiscalYear = 2018
	CoverageArea = ""
	HistoricalDataYear = FiscalYear - 2
	MVT1 = 0
	MVT2 = 0
	MVT3 = 0
	MVTLoss1 = 0
	MVTLoss2 = 0
	MVTLoss3 = 0
	BMV1 = 0
	BMV2 = 0
	BMV3 = 0
	BMVLoss1 = 0
	BMVLoss2 = 0
	BMVLoss3 = 0
	ReceiveAuthorization = 0
	RegionalTaskForce = 0
	CurrentProgram = 0
	CurrentProgramDescription = ""
	PreviouslyApplied = 0
	PreviouslyAwarded = 0
	TerminationExplanation = 0
	MeetCashMatchRequirement = 0
	DedicateResources = 0
	NoSupplantation = 0
	ProgramCategoryID = 0
	BriefNarrative = ""
	GrantRequestRangeID = 0
	SubmitID = null
	SubmitTimestamp = null
	SubmitByName = null
	DateReviewed = null
	DateResponded = null
	MethodRespondedID = null
	Notes = ""
	UpdateID = null
	UpdateTimestamp = null
End If

If GranteeID>0 Then
	If MVCPARights = True Then
		PermitEdit = True
	ElseIf IsNull(SubmitID) Then
		PermitEdit = CheckPermissions(UserSystemID, GranteeID, False)
	Else
		PermitEdit = False
	End If
Else
		PermitEdit = False
End If
If Debug = True Then
	Response.Write("<pre>PermitEdit='" & PermitEdit & "'</pre>")
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" />
<style type="text/css">
	.heading {text-align: left; padding-right: 6px; }
	td {vertical-align: top; padding-bottom: 8px }
</style>
<script type="text/javascript">
	function submitForm(action)
	{
		document.ISA.Button.value=action
		if (checkTypes() == false)
			return false;
		if (action == "submit") {
			if (validateForm() == false)
				return false;
		}
		document.ISA.submit();
	}

	function checkTypes()
	{
		if (checkInteger(document.ISA.MVT1) == false) {
			return false;
		}
		if (checkInteger(document.ISA.MVT2) == false) {
			return false;
		}
		if (checkInteger(document.ISA.MVT3) == false) {
			return false;
		}
		if (checkInteger(document.ISA.BMV1) == false) {
			return false;
		}
		if (checkInteger(document.ISA.BMV2) == false) {
			return false;
		}
		if (checkInteger(document.ISA.BMV3) == false) {
			return false;
		}
		if (document.ISA.BriefNarrative.value.length > 2850) {
			alert("The brief narrative should be less than 2850 characters. Please reduce the length of the text");
			document.ISA.BriefNarrative.focus();
			return false;
		}
	}

	function validateForm()
	{
		if (document.ISA.CoverageArea.value.length == 0) {
			alert("You must enter a proposed coverage area.");
			document.ISA.CoverageArea.focus();
			return false;
		}
		if (document.ISA.MVT1.value.length == 0) {
			alert("You must enter the number of motor vehicle thefts for <%=(HistoricalDataYear-2)%>.");
			document.ISA.MVT1.focus();
			return false;
		}
		if (document.ISA.MVT2.value.length == 0) {
			alert("You must enter the number of motor vehicle thefts for <%=(HistoricalDataYear-1)%>.");
			document.ISA.MVT2.focus();
			return false;
		}
		if (document.ISA.MVT3.value.length == 0) {
			alert("You must enter the number of motor vehicle thefts for <%=HistoricalDataYear%>.");
			document.ISA.MVT3.focus();
			return false;
		}
		if (document.ISA.MVTLoss1.value.length == 0) {
			alert("You must enter the value of motor vehicle theft losses for <%=(HistoricalDataYear-2)%>.");
			document.ISA.MVTLoss1.focus();
			return false;
		}
		if (document.ISA.MVTLoss2.value.length == 0) {
			alert("You must enter the value of motor vehicle theft losses for <%=(HistoricalDataYear-1)%>.");
			document.ISA.MVTLoss2.focus();
			return false;
		}
		if (document.ISA.MVTLoss3.value.length == 0) {
			alert("You must enter the value of motor vehicle theft losses for <%=HistoricalDataYear%>.");
			document.ISA.MVTLoss3.focus();
			return false;
		}
		if (document.ISA.BMV1.value.length == 0) {
			alert("You must enter the number of motor vehicle burglaries for <%=(HistoricalDataYear-2)%>.");
			document.ISA.BMV1.focus();
			return false;
		}
		if (document.ISA.BMV2.value.length == 0) {
			alert("You must enter the number of motor vehicle burglaries for <%=(HistoricalDataYear-1)%>.");
			document.ISA.BMV2.focus();
			return false;
		}
		if (document.ISA.BMV3.value.length == 0) {
			alert("You must enter the number of motor vehicle burglaries for <%=HistoricalDataYear%>.");
			document.ISA.BMV3.focus();
			return false;
		}
		if (document.ISA.BMVLoss1.value.length == 0) {
			alert("You must enter the value of motor vehicle burglary losses for <%=(HistoricalDataYear-2)%>.");
			document.ISA.BMVLoss1.focus();
			return false;
		}
		if (document.ISA.BMVLoss2.value.length == 0) {
			alert("You must enter the value of motor vehicle burglary losses for <%=(HistoricalDataYear-1)%>.");
			document.ISA.BMVLoss2.focus();
			return false;
		}
		if (document.ISA.BMVLoss3.value.length == 0) {
			alert("You must enter the value of motor vehicle burglary losses for <%=HistoricalDataYear%>.");
			document.ISA.BMVLoss3.focus();
			return false;
		}

		var selectedbutton = false;
		for (i = 0; i < document.ISA.ReceiveAuthorization.length; i++) {
			if (document.ISA.ReceiveAuthorization[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate whether you be able to receive authorization from your governing body to submit a grant application.");
			return false;
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.RegionalTaskForce.length; i++) {
			if (document.ISA.RegionalTaskForce[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate whether you be able to coordinate with other jurisdictions to establish a regional Taskforce.");
			return false;
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.CurrentProgram.length; i++) {
			if (document.ISA.CurrentProgram[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate whether your organization currently has an automobile burglary and/or theft interdiction and/or prevention program.");
			return false;
		}
		if (document.ISA.CurrentProgram[0].checked) {
			if (document.ISA.CurrentProgramDescription.value.length == 0) {
				alert("You must provide provide a brief description of the current project.");
				document.ISA.CurrentProgramDescription.focus();
				return false;
			}
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.PreviouslyApplied.length; i++) {
			if (document.ISA.PreviouslyApplied[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate whether your organization applied for a MVCPA Grant in the past.");
			return false;
		}

		if (document.ISA.PreviouslyApplied[0].checked) {
			selectedbutton = false;
			for (i = 0; i < document.ISA.PreviouslyAwarded.length; i++) {
				if (document.ISA.PreviouslyAwarded[i].checked)
					selectedbutton = true;
			}
			if (selectedbutton == false) {
				alert("You must indicate whether your organization was awarded a grant?.");
				return false;
			}
		}
		if (document.ISA.PreviouslyAwarded[0].checked)
		{
			if (document.ISA.TerminationExplanation.value.length == 0) {
				alert("You must provide an explanation regarding the termination of the grant.");
				document.ISA.TerminationExplanation.focus();
				return false;
			}
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.MeetCashMatchRequirement.length; i++) {
			if (document.ISA.MeetCashMatchRequirement[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate whether your organization will be able to meet the 20% minimum cash match requirement if awarded a grant.");
			return false;
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.DedicateResources.length; i++) {
			if (document.ISA.DedicateResources[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate whether your  entity be able to dedicate personnel and other resources required for the successful implementation of a grant if awarded.");
			return false;
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.NoSupplantation.length; i++) {
			if (document.ISA.NoSupplantation[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate whether you can certify that the funds requested will not be used to supplant existing local funds currently used for motor vehicle burglary and theft or prevention.");
			return false;
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.ProgramCategoryID.length; i++) {
			if (document.ISA.ProgramCategoryID[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate the MVCPA grant category which identifies the main focus of the proposed project.");
			return false;
		}

		if (document.ISA.BriefNarrative.value.length == 0) {
			alert("You must provide a brief narrative on the function of the proposed project which includes the projected goals and activities.");
			document.ISA.BriefNarrative.focus();
			return false;
		}

		selectedbutton = false;
		for (i = 0; i < document.ISA.GrantRequestRangeID.length; i++) {
			if (document.ISA.GrantRequestRangeID[i].checked)
				selectedbutton = true;
		}
		if (selectedbutton == false) {
			alert("You must indicate the estimated grant amount that will be requested.");
			return false;
		}

		return true;
	}

 </script>
 <!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Intent To Apply For a New MVCPA Grant
</div>

<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">
<p>The purpose of the "Intent to Apply" process is to conduct an initial assessment for 
eligibility of potential new grant applicants and provide written recommendations for 
the potential new grant application scope. The "Intent to Apply" process also provides 
information to potential grant applicants regarding grant requirements prior to officially 
applying for MVCPA funding. MVCPA staff will respond to all submissions within five (5) 
business days. For more information contact us at <a href="mailto:grantsMVCPA@txdmv.gov?Subject=ISA">grantsMVCPA@txdmv.gov</a></p>

<p>Applicable Authority and Rules: Motor Vehicle Crime Prevention Authority 
grant programs are governed by one or more of the following statutes, rules, standards 
and guidelines.
<ul>
	<li><a href="http://www.statutes.legis.state.tx.us/Docs/CV/htm/CV.70.9.htm#4413(37)" target="_blank">Texas Revised Civil Statutes Article 4413(37)</a></li>
	<li><a href="http://texreg.sos.state.tx.us/public/readtac$ext.ViewTAC?tac_view=3&ti=43&pt=3" target="_blank">Texas Administrative Code: Title 43; Part 3; Chapter 57</a></li>
	<li><a href="https://comptroller.texas.gov/purchasing/docs/ugms.pdf" target="blank">Uniform Grant Management Standards (UGMS) as promulgated by the Texas Comptroller 
	of Public Accounts</a></li>
	<li><a href="http://www.txdmv.gov/reports-and-data/doc_download/1066-grant-administrative-manual" target="_blank">The current Motor Vehicle Crime Prevention Authority Grant Administrative Guide 
	and subsequent adopted grantee instruction manuals</a></li>
</ul></p>

<form name="ISA" method="post" action="ISASubmit.asp" onsubmit="return validateForm();">
<%=HiddenField("ISAID", ISAID) %>
<%=HiddenField("FiscalYear", FiscalYear) %>
<%=HiddenField("GranteeID", GranteeID) %>
<%=HiddenField("Button", "save") %>
<%=HiddenField("HistoricalDataYear", HistoricalDataYear) %>
<table style="padding-left: 6px; ">
<%	If SubmitID>0 Then %>
<tr><td colspan="2" style="text-align: center; font-weight: bold; ">The ISA was submitted by <%=SubmitByName%> at <%=SubmitTimestamp %></td></tr>
<%	End If %>
<tr>
	<td>Grantee Name</td>
	<td><%=GranteeName %></td>
</tr>
<tr>
	<td>Fiscal Year</td>
	<td><%=FiscalYear %></td>
</tr>

<tr>
	<td class="heading">Proposed coverage area (Cities and/or Counties</td>
	<td><%=TextArea("CoverageArea", CoverageArea, 3, 50, 1024, PermitEdit, "") %></td>
</tr>
<tr>
	<td class="heading">Provide the Historical Data for Automobile Theft and Burglary for the proposed 
	coverage area from the past three years (Please Use Actual UCR#)</td>
	<td><table align="center">
	
		<tr>
			<th></th>
			<th><%=(HistoricalDataYear-2) %></th>
			<th><%=(HistoricalDataYear-1) %></th>
			<th><%=(HistoricalDataYear) %></th>
		</tr>
		<tr>
			<td title="Theft of a Motor Vehicle Number Count">MVT #</td>
			<td><%=IntegerField("MVT1", MVT1, 10, 10, PermitEdit, "") %></td>
			<td><%=IntegerField("MVT2", MVT2, 10, 10, PermitEdit, "") %></td>
			<td><%=IntegerField("MVT3", MVT3, 10, 10, PermitEdit, "") %></td>
		</tr>
		<tr>
			<td title="Loss from Theft of a Motor Vehicle" style="white-space: nowrap;">MVT Loss Amount</td>
			<td><%=CurrencyField("MVTLoss1", MVTLoss1, 10, 12, PermitEdit, "") %></td>
			<td><%=CurrencyField("MVTLoss2", MVTLoss2, 10, 12, PermitEdit, "") %></td>
			<td><%=CurrencyField("MVTLoss3", MVTLoss3, 10, 12, PermitEdit, "") %></td>
		</tr>
		<tr>
			<td title="Motor Vehicle Burglary Number Count">BMV #</td>
			<td><%=IntegerField("BMV1", BMV1, 10, 10, PermitEdit, "") %></td>
			<td><%=IntegerField("BMV2", BMV2, 10, 10, PermitEdit, "") %></td>
			<td><%=IntegerField("BMV3", BMV3, 10, 10, PermitEdit, "") %></td>
		</tr>
		<tr>
			<td title="Loss from Burglary of a Motor Vehicle" style="white-space: nowrap;">BMV Loss Amount</td>
			<td><%=CurrencyField("BMVLoss1", BMVLoss1, 10, 12, PermitEdit, "") %></td>
			<td><%=CurrencyField("BMVLoss2", BMVLoss2, 10, 12, PermitEdit, "") %></td>
			<td><%=CurrencyField("BMVLoss3", BMVLoss3, 10, 12, PermitEdit, "") %></td>
		</tr>
	    </table></td>
</tr>

<tr>
	<td class="heading">Will you be able to receive authorization from your governing 
		body to submit a grant application?</td>
	<td><%=RadioInputField("ReceiveAuthorization", ReceiveAuthorization, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("ReceiveAuthorization", ReceiveAuthorization, 0) %>No
	</td>
</tr>

<tr>
	<td class="heading">Will you be able to coordinate with other jurisdictions to establish a 
	regional Taskforce?</td>
	<td><%=RadioInputField("RegionalTaskForce", RegionalTaskForce, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("RegionalTaskForce", RegionalTaskForce, 0) %>No
	</td>
</tr>

<tr>
	<td class="heading">Does your organization currently have an automobile burglary and/or 
		theft interdiction and/or prevention program?</td>
	<td><%=RadioInputField("CurrentProgram", CurrentProgram, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("CurrentProgram", CurrentProgram, 0) %>No<br />
		If Yes, please provide a brief description of the current project.
		<%=TextArea("CurrentProgramDescription", CurrentProgramDescription, 3, 50, 1024, PermitEdit, "") %>
	</td>
</tr>

<tr>
	<td class="heading">Has your organization applied for a MVCPA Grant in the past?</td>
	<td><%=RadioInputField("PreviouslyApplied", PreviouslyApplied, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("PreviouslyApplied", PreviouslyApplied, 0) %>No<br /><br />
		If Yes, were you awarded a grant?<br />
		<%=RadioInputField("PreviouslyAwarded", PreviouslyAwarded, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("PreviouslyAwarded", PreviouslyAwarded, 0) %>No<br />
		If yes, please provide an explanation regarding the termination of the grant<br />
		<%=TextArea("TerminationExplanation", TerminationExplanation, 3, 50, 1024, PermitEdit, "") %>
	</td>
</tr>


<tr>
	<td class="heading">Applicants are required to provide a minimum of 20% cash match 
	contribution of the project total budget. Will you be able to meet the 20% minimum 
	cash match requirement if awarded a grant?</td>
	<td><%=RadioInputField("MeetCashMatchRequirement", MeetCashMatchRequirement, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("MeetCashMatchRequirement", MeetCashMatchRequirement, 0) %>No
	</td>
</tr>

<tr>
	<td class="heading">Will your entity be able to dedicate personnel and other resources 
	required for the successful implementation of a grant if awarded? (i.e. 
	Investigators, financial officers etc.)</td>
	<td><%=RadioInputField("DedicateResources", DedicateResources, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("DedicateResources", DedicateResources, 0) %>No
	</td>
</tr>

<tr>
	<td class="heading">Can you certify that the funds requested will not be used to supplant 
	existing local funds currently used for motor vehicle burglary and theft or prevention?</td>
	<td><%=RadioInputField("NoSupplantation", NoSupplantation, 1) %>Yes&nbsp;&nbsp;&nbsp;
		<%=RadioInputField("NoSupplantation", NoSupplantation, 0) %>No
	</td>
</tr>

<tr>
	<td class="heading">Which of the following MVCPA grant categories identifies the main focus of the proposed project? (Select one.)</td>
	<td>
<%
	sql = "SELECT ProgramCategoryID, ProgramCategory FROM Lookup.ProgramCategory WHERE Version=1 ORDER BY ProgramCategoryID "
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(RadioInputField("ProgramCategoryID", ProgramCategoryID, rs.Fields("ProgramCategoryID")) & rs.Fields("ProgramCategory") & "<br />")
		rs.MoveNext
	Wend
%></td>
</tr>

<tr>
	<td class="heading">Provide a brief narrative on the function of the proposed project which includes the projected goals and activities.</td>
	<td><%=TextArea("BriefNarrative", BriefNarrative, 15, 54, 2850, PermitEdit, "") %></td>
</tr>

<tr>
	<td class="heading">Estimated amount that will be requested</td>
	<td>
<%
	sql = "SELECT GrantRangeID, GrantRangeDescription FROM Lookup.GrantRanges ORDER BY GrantRangeID "
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(RadioInputField("GrantRequestRangeID", GrantRequestRangeID, rs.Fields("GrantRangeID")) & rs.Fields("GrantRangeDescription") & "<br />")
		rs.MoveNext
	Wend
%></td>
</tr>


<tr>
	<th colspan="2">MVCPA Staff Only</th>
</tr>
<tr>
	<td>Date Reviewed</td>
	<td><%=DateField("DateReviewed", DateReviewed, MVCPARights) %></td>
</tr>
<tr>
	<td>Date Responded</td>
	<td><%=DateField("DateResponded", DateResponded, MVCPARights) %></td>
</tr>

<tr>
	<td class="heading">Communication method for response</td>
	<td>
<%
	sql = "SELECT CommunicationMethodID, CommunicationMethod FROM Lookup.CommunicationMethods ORDER BY 1 "
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(RadioInputField("MethodRespondedID", MethodRespondedID, rs.Fields("CommunicationMethodID")) & rs.Fields("CommunicationMethod") & "<br />")
		rs.MoveNext
	Wend
%></td>
</tr>
<tr>
	<td class="heading">Notes</td>
	<td><%=TextArea("Notes", Notes, 3, 50, 2048, MVCPARights, "") %></td>
</tr>
<%	If IsNull(SubmitID) = False And MVCPARights = True Then %>
<tr>
	<td>Unsubmit Application</td>
	<td><input type="checkbox" name="UnSubmit" value="1" /> This will clear the submission allowing grantee to edit ISA.</td>
</tr>
<%	End If %>

</table>

<div style="text-align: center;">
<%	If PermitEdit = True Then %>
	<input type="button" value="Save" onclick="submitForm('save');" />
	<input type="button" value="Submit" onclick="submitForm('submit');" />
	<input type="button" value="Cancel" onclick="location.href='../Home/Default.asp?GranteeID=<%=GranteeID%>';" />
<%	Else %>
	<input type="button" value="Home" onclick="location.href='../Home/Default.asp?GranteeID=<%=GranteeID%>';" />
<%	End If %></div>
</div>
</form>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->