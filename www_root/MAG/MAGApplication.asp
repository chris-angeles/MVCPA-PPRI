<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 

Dim debug, i, j, LastCategory, PermitEdit, CanSubmit, DocumentFolder, _
	GranteeID, FiscalYear, GranteeName, ORI, Agency, MAGID, _
	AuthorizedOfficialID, AuthorizedOfficialName, AuthorizedOfficialTitle, _
	ProgramDirectorName, ProgramDirectorTitle, _
	FinancialOfficerName, FinancialOfficerTitle, RequiredOfficials, _
	StolenVehicles, StolenVehicleValue, OptionID, Certification, TFGrant, SubmitID, SubmitTimestamp, SubmitName, _
	Submitted, AllowUpload, NoButtons

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
				response.write("Cookies(" & i & ":" & j & ")=" & Request.Cookies(i)(j))
			next
		else
			Response.Write("Cookies(""" & i & """)=" & Request.Cookies(i) & "<br>")
		end if
	next
	Response.Write("Now=" & Now() & vbCrLf)
	Response.Write("</pre>" & vbCrLf)
End If

If Len(Request.Form("MAGID")) > 0 Then
	MAGID = CInt(Request.Form("MAGID"))
ElseIf Len(Request.QueryString("MAGID")) > 0 Then
	MAGID = CInt(Request.QueryString("MAGID"))
Else
	MAGID = 0
End If

If Len(Request.Form("GranteeID")) > 0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID")) > 0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = 0
End If

If GranteeID=0 And MAGID=0 Then
	Response.Write("Error: No GranteeID or MAGID Specified")
	SendMessage "Error: No GranteeID or MAGID Specified"
	Response.End
End If

If Len(Request.Form("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear = 2022
End If

If Request.QueryString("NoButtons") = "1" Then
	NoButtons = True
Else
	NoButtons = False
End If

sql = "SELECT G.GranteeID, G.GranteeName, ISNULL(G.ORI, 'None') AS ORI, O.Agency, " & vbCrLf & _
	"	ISNULL(M.MAGID,0) AS MAGID, M.OptionID, " & vbCrLf & _
	"	M.Certification, S.Name AS SubmitName, M.SubmitID, M.SubmitTimestamp, " & vbCrLf & _
	"	G.AuthorizedOfficialID, AO.Name AS AuthorizedOfficialName, AO.Title AS AuthorizedOfficialTitle, " & vbCrLf & _
	"	PD.Name AS ProgramDirectorName, PD.Title AS ProgramDirectorTitle, " & vbCrLf & _
	"	FO.Name AS FinancialOfficerName, FO.Title AS FinancialOfficerTitle, " & vbCrLf & _
	"	CAST(CASE WHEN G.AuthorizedOfficialID>0 AND G.ProgramDirectorID>0 AND G.FinancialOfficerID>0 THEN 1 ELSE 0 END AS BIT) AS RequiredOfficials, " & vbCrLf & _
	"	StolenVehicles, StolenVehicleValue, " & vbCrLf & _
	"	TFGrant = ISNULL((SELECT ProgramName FROM [Grants].Main AS GM JOIN [Grants].ParticipatingAgencies AS PA ON PA.GrantID=GM.GrantID WHERE PA.ORI=G.ORI AND GM.FiscalYear=" & prepIntegerSQL(FiscalYear) & "), 'None') " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN MAG.Main AS M ON M.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN Lookup.ORI AS O ON O.ORI=G.ORI " & vbCrLf & _
	"LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
	"LEFT JOIN System.Users AS PD ON PD.SystemID=G.ProgramDirectorID " & vbCrLf & _
	"LEFT JOIN System.Users AS FO ON FO.SystemID=G.FinancialOfficerID " & vbCrLf & _
	"LEFT JOIN System.Users AS S ON S.SystemID=M.SubmitID " & vbCrLf
If MAGID>0 Then
	sql = sql & "WHERE M.MAGID=" & prepIntegerSQL(MAGID)
Else
	sql = sql & "WHERE G.GranteeID=" & prepIntegerSQL(GranteeID) & " AND M.FiscalYear=" & PrepIntegerSQL(FiscalYear) 
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
	MAGID = rs.Fields("MagID")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.FIelds("GranteeName")
	ORI = rs.Fields("ORI")
	Agency = rs.Fields("Agency")
	AuthorizedOfficialID = rs.Fields("AuthorizedOfficialID")
	AuthorizedOfficialName = rs.Fields("AuthorizedOfficialName")
	AuthorizedOfficialTitle = rs.Fields("AuthorizedOfficialTitle")
	ProgramDirectorName = rs.Fields("ProgramDirectorName")
	ProgramDirectorTitle = rs.Fields("ProgramDirectorTitle")
	FinancialOfficerName = rs.Fields("FinancialOfficerName")
	FinancialOfficerTitle = rs.Fields("FinancialOfficerTitle")
	RequiredOfficials = rs.Fields("RequiredOfficials")
	TFGrant = rs.Fields("TFGrant")
	OptionID = rs.Fields("OptionID")
	StolenVehicles = rs.Fields("StolenVehicles")
	StolenVehicleValue = rs.Fields("StolenVehicleValue")
	Certification = rs.Fields("Certification")
	SubmitID = rs.Fields("SubmitID")
	SubmitTimestamp = rs.Fields("SubmitTimestamp")
	SubmitName = rs.Fields("SubmitName")
	If IsNull(SubmitID) = True Then
		Submitted = False
	Else
		Submitted = True
	End If
End If

PermitEdit = CheckPermissionsWithLock(UserSystemID, GranteeID, Submitted)
AllowUpload = CheckPermissions(UserSystemID, GranteeID, True) ' Allow Upload after submission.

If Debug = True Then
	Response.Write("<pre>Submitted=" & Submitted & " (after original CheckPermissionsWithLock)</pre>")
	Response.Write("<pre>PermitEdit=" & PermitEdit & " (after original CheckPermissionsWithLock)</pre>")
End If

If Now()>CDate("06/03/2022 5:00:00 PM") Then
	PermitEdit = False
ElseIf Submitted = True Then
	PermitEdit = False
End If

If Debug = True Then
	Response.Write("<pre>Submitted=" & Submitted & " (after original CheckPermissionsWithLock)</pre>")
	Response.Write("<pre>PermitEdit=" & PermitEdit & " (after original CheckPermissionsWithLock)</pre>")
End If

If PermitEdit = False Then
	CanSubmit = False
ElseIf Submitted = True Then
	CanSubmit = False
ElseIf AuthorizedOfficialID = UserSystemID Then
	CanSubmit = True
Else
	CanSubmit = False
End If

' If they have an existing Taskforce grant, they are ineligble for MAG grant and that overrides other criteria.
If TFGrant <> "None" Then
	CanSubmit = False
End If

If Debug = True Then
	Response.Write("<pre>UserSystemID=" & UserSystemID & "</pre>")
	Response.Write("<pre>AuthorizedOfficialID=" & AuthorizedOfficialID & "</pre>")
	Response.Write("<pre>PermitEdit=" & PermitEdit & "</pre>")
	Response.Write("<pre>CanSubmit=" & CanSubmit & "</pre>")
	Response.Write("<pre>Submitted=" & Submitted & "</pre>")
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Auxiliary Grant Application</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function submitForm(action)
	{
		document.MAGApplication.Button.value = action;
		if (validateForm() == false)
			return false;
		if (action == "submit") {
			if (document.MAGApplication.StolenVehicles.value.length == 0) {
				alert("You must enter a value for \"Reported incidents of motor vehicle theft as published in the Texas Department of Public Safety Crime in Texas Report\" to submit this application.");
				document.MAGApplication.StolenVehicles.focus();
				return false;
			}
			if (document.MAGApplication.StolenVehicleValue.value.length == 0) {
				alert("You must enter a value for \"The value of reported incidents of motor vehicle theft as published in the Texas Department of Public Safety Crime in Texas Report\" to submit this application.");
				document.MAGApplication.StolenVehicleValue.focus();
				return false;
			}
			if (getNumericValue(document.MAGApplication.StolenVehicleValue.value) < 60000.00) {
				alert("The value of reported incidents of motor vehicle theft must be at least three times the amount of the grant request.");
				document.MAGApplication.StolenVehicleValue.focus();
				return false;
			}
			if (document.MAGApplication.Certification.checked == false) {
				alert("You must certify reviewing the TxGMS Standard Assurances by Local Governments before submitting the application!")
				document.MAGApplication.Certification.focus();
				return false;
			}
			if (confirm("By submitting this application I certify that I have been designated by my jurisdiction as the authorized official to accept the terms and conditions of the grant. The statements herein are true, complete, and accurate to the best of my knowledge. I am aware that any false, fictitious, or fraudulent statements or claims may subject me to criminal, civil, or administrative penalties.") == false) {
				return false;
			}
			if (confirm("By submitting this application I certify that my jurisdiction agrees to comply with all terms and conditions if the grant is awarded and accepted. I further certify that my jurisdiction will comply with all applicable state and federal laws, rules and regulations in the application, acceptance, administration and operation of this grant.") == false) {
				return false;
			}
		}
		document.MAGApplication.submit();
	}

	function validateForm()
	{
		if (document.MAGApplication.OptionID[0].checked || document.MAGApplication.OptionID[1].checked) {
			return true;
		}
		else {
			alert("You must select from one of the two funding options to continue.");
			return false;
		}
		return true;
	}
</script>
</head>
<body>
<%	If NoButtons = False Then %>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>
<%	End If %>

<div class="pagetag"><%=GranteeName %> Auxiliary Grant Application for Fiscal Year <%=FiscalYear %></div>

<div class="widecontent">
<br />
<%
If Now()>CDate("06/03/2022 5:00:00 PM") And Submitted = False Then
	Response.Write("<div style=""width: 100%; margin: auto; color: red; font-weight: bold; text-align: center; "">The Application Period is Closed</div></br>" & vbCrLf)
End If
%>
<form name="MAGApplication" id="MAGApplication" method="post" action="MAGApplicationSubmit.asp">
<%
Response.Write(HiddenField("GranteeID", GranteeID))
Response.Write(HiddenField("FiscalYear", FiscalYear))
Response.Write(HiddenField("MAGID", MAGID))
Response.Write(HiddenField("Button","save"))
If Submitted = True Then
	Response.Write("<div style=""width: 100%; text-align: center; color: red; "">The application was submitted by " & SubmitName & " at " & SubmitTimestamp & ".</div>")
	Response.Write("<br />")
End If
%>

<table>
<tr style="white-space: nowrap; ">
	<td>Grantee Legal Name:</td>
	<td><i><%=GranteeName %></i></td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr>
	<td>Organization ORI:</td>
	<td><i><%=ORI %>: <%=Agency %></i>
<%	If ORI = "None" Then 
		Response.Write("<colspan style=""color: red; font-weight: bold; "">Only Law Enforcement Agencies with an ORI are eligible to apply for this grant.</colspan> ")
	End If
%>
	</td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr>
	<td>Fiscal Year:</td>
	<td><i><%=FiscalYear %></i></td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr><td colspan="2" style="height: 5px;">Grantee Officials: These are the current officials recorded
	for the grantee. If officals are missing or need to be added, please make the changes on the 
	grantee edit page. 
<%	If RequiredOfficials=True Then 
		Response.Write("The Authorized Official, Program Director and Financial Officer are required positions. ")
	Else
		Response.Write("<span style=""font-weight: bold; color: red; "">The Authorized Official, Program Director and Financial Officer are required positions.</span> ")
	End If
%>
    </td>
</tr>

<tr>
	<td>Authorized Official:</td>
	<td><i><%=AuthorizedOfficialName %>, <%=AuthorizedOfficialTitle %></i></td>
</tr>

<tr>
	<td>Program Director:</td>
	<td><i><%=ProgramDirectorName %>, <%=ProgramDirectorTitle %></i></td>
</tr>

<tr>
	<td>Financial Officer:</td>
	<td><i><%=FinancialOfficerName %>, <%=FinancialOfficerTitle %></i></td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr>
	<td style="white-space: nowrap; ">Existing Taskforce Grant:</td>
	<td><i><%=TFGrant %></i></td>
</tr>
<%	If TFGrant <> "None" Then %>
<tr>
	<td></td>
	<td style="color: red; font-weight: bold; font-style: italic; ">Existing taskforce grant makes grantee ineligible for auxiliary grant.</td>
</tr>
<%	End If %>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr style="vertical-align: top; ">
	<td>Purpose</td>
	<td>The purpose of the MAG grant is to provide law enforcement agencies with one-time funding 
	for specific interdiction equipment used to combat motor vehicle theft, theft from motor vehicles, 
	and fraud-related motor vehicle crime. This grant opportunity is only for Automatic License Plate 
	Readers (ALPR). Eligible applicants may request funds to buy or lease one or more ALPR and report 
	the results. Funding is subject to availability of state funds and the application being consistent 
	with the information in this RFA, including the requirements and conditions stated in the RFA. 
	The RFA is posted as required by law for at least thirty (30) days prior to the due date for 
	Applications.</td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr style="vertical-align: top; ">
	<td>Eligibility</td>
	<td>Eligible applicants must meet both of the following conditions: 
	1) be a Texas municipal police department or a county Sheriff's office; and 
	2) not currently receiving funds as a grantee or subgrantee through other MVCPA programs.</td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr style="vertical-align: top; ">
	<td>Application Category</td>
	<td>Applicants meeting the eligibility requirements may submit a request to fund specific items 
	for Law Enforcement, Detection and Apprehension. This provides financial support to law enforcement 
	agencies for the purchase of equipment to combat motor vehicle theft and fraud-related motor 
	vehicle crime through the enforcement of law.  This may include equipment designed to increase 
	recovery of vehicles, clearance of criminal cases, arrest of law violators, and disruption of 
	organized motor vehicle crime.</td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr>
	<td colspan="2">Choose from the eligible funding options:</td>
</tr>

<tr style="vertical-align: top; ">
	<td style="text-align: right; "><%=RadioInputField("OptionID", rs.Fields("OptionID"), 1) %>
	<td><b>Purchase</b> of one (1) Mobile or Stationary Automatic License Plate Reader (ALPR): 
	Reimbursement up to $20,000 with a required 20% cash match..</td>
</tr>
<tr style="vertical-align: top; ">
	<td style="text-align: right; "><%=RadioInputField("OptionID", rs.Fields("OptionID"), 2) %></td>
	<td><b>Lease</b> a multiple unit stationary Automatic License Plate Reader (ALPR) system. 
	This system must be high resolution stationary still image license plate reader system with 
	multiple unit connectivity and access to network that includes machine learning integration: 
	Reimbursement up to $20,000 with a required 20% cash match for the first year of the lease 
	with the condition that grantee will commit to funding not less than one year following the 
	end of first year.</td>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr>
	<td style="vertical-align: top; ">Crime Statistics</td>
	<td>Reported incidents of motor vehicle theft as published in the Texas Department of 
		Public Safety Crime in Texas Report: 
		<%=IntegerField("StolenVehicles", prepNumberWeb(StolenVehicles,0), 7, 11, PermitEdit, "return checkInteger(this);") %><br />
		The value of reported incidents of motor vehicle theft as published in the Texas Department of 
		Public Safety Crime in Texas Report: 
		<%=CurrencyField("StolenVehicleValue", prepCurrencyWeb(StolenVehicleValue), 11, 15, PermitEdit, "return checkCurrency(this);") %>
</tr>

<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr style="vertical-align: top;">
	<td>Resolution:</td>
	<td><p style="margin-top: 0px; "> Resolution (Order or Ordinance) by the applicant governing body is required to make application for these funds. The resolution shall provide that the governing body applies for the funds for the purpose provided in statute (Texas Transportation Code, Chapter 1006) to return the grant funds in the event of loss or misuse and designate the officials that the governing body chooses as its agents to make uniform assurances and administer the grant if awarded.</p>
		<p>In the event a governing body has previously delegated the application authority to a 
		city manager, chief of police, sheriff or other official then the applicant must 
		submit on-line a copy of the delegation order (documentation) along with the Resolution 
		signed by the designated official. A sample Resolution that meets the three elements required 
		is displayed at the link below:</p>
		<p>Link to <a href="Resolution.asp?GranteeID=<%=GranteeID %>&FiscalYear=<%=FiscalYear %>" target="_blank">resolution</a></p></td>
</tr>

<%
If MAGID > 0 Then
	Dim fso, folder, files, file
	If MAGID>0 And AllowUpload = True Then
		Response.Write("<tr><td colspan=""2"" style=""text-align: center; ""><a href=""../Upload/Upload.asp?FID=14&MAGID=" & MAGID & """ target=""_blank"">File Upload</a></td></tr>" & vbCrLf)
	End If

	set fso = Server.CreateObject("Scripting.FileSystemOBject")

	DocumentFolder = Application("DocumentRoot") & "\MAG\" & MAGID & "\"
	If fso.FolderExists(DocumentFolder) = False Then
		fso.CreateFolder(DocumentFolder)
	End If

	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		Response.Write("<tr><th colspan=""2"">Current Documents in folder</th></tr>" & vbCrLf)
		If files.count>0 Then 
			Response.Write("<tr><th colspan=""2"" style=""text-align: center; "">")
			For Each file in files
					Response.Write("<a href=""../Documents/MAG/" & MAGID & "/" & file.Name & _
						""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
			Next
			Response.Write("</td></tr>" & vbCrLf)
		Else
			Response.Write("<tr><td colspan=""2"" style=""text-align: center; "">There are no documents in the folder.</td></tr>")
		End If
	End If
End If
%>
<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr>
	<td colspan="2">TxGMS Standard Assurances by Local Governments</td>
</tr>
<tr>
	<td colspan="2"><%=CheckBoxField("Certification", Certification) %> We acknowledge reviewing the 
	<a href="../RFA/UniformAssurances.pdf" target="_blank" class="plainlink">TxGMS Standard Assurances by Local Governments</a> as 
	promulgated by the Texas Comptroller of Public Accounts and agree to abide by the terms stated therein.</td>
</tr>
<%	If NoButtons = False Then %>
<tr><td colspan="2" style="height: 5px;"></td></tr>

<tr><td colspan="2" style="text-align: center; ">
<%	If CanSubmit = True Then %>
		<input name="Save" id="Save1" type="button" value="Save" onclick="submitForm('save');" 
			title="Only the authorized official may submit the application. Others may save."/>
		<input name="Submit" id="Submit" type="button" value="Submit" onclick="submitForm('submit');" 
			title="Only the authorized official may submit the application. Click button to submit form."/>
<%	ElseIf PermitEdit = True and Submitted = False Then %>
		<input name="Save" id="Save2" type="button" value="Save" onclick="submitForm('save');" 
			title="Only the authorized official may submit the application. Others may save."/>
		<input type="button" name="Submit2" id="Submit2" value="Submit" onclick="alert('Only the authorized official for the entity may submit the application. Other users with grantee permissions in the system may edit the form, but the authorized official will need to logon to submit the completed application.');" 
			title="Only the authorized official may submit the application. "/>
<%	End If %>
		<input name="Cancel" id="Cancel" type="button" value="Cancel" onclick="window.close();" 
			title="Cancel any current edits and close window. Be sure you hit save first if you want the data saved."/></td>
</tr>
<%	End If %>
</table>
</form>
</div>
</body>
</html>
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
<!--#include file="../includes/InputValidation.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->

