<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, PermitEdit, ViewDocuments, GranteeID, GranteeName, OrganizationTypeID, ORI, StatePayeeIDNo, _
	OrganizationalUnit, GeneralPhone, Address1, Address2, City, State, zip, _
	VendorOrganizationalUnit, VendorAddress1, VendorAddress2, VendorCity, VendorState, VendorZIP, _
	AuthorizedOfficialID, ProgramDirectorID, ProgramManagerID, FinancialOfficerID, _
	TaskForceCommanderID, ProgramAdministrativeContactID, FinancialAdministrativeContactID, PIOID, _
	BorderCounty, PortCounty, Port2County, TaskforceGrant, AuxiliaryGrant, RapidResponseStrikeforceGrant, _
	CatalyticConverterGrant, GranteeNameSort, ReducedOfficials, Inactive, _
	UpdateID, UpdateTimestamp, UpdateName, MonitorID, MonitorDocumentCounter
debug = False
If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLf)
	For each i in Request.Form
		Response.Write("Request.Form(""" & i & """)='" & Request.Form(i) & "'" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("Request.QueryString(""" & i & """)='" & Request.Form(i) & "'" & vbCrLf)
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

If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	GranteeID = 0
	'Response.Write("No GranteeID Provided to page.")
	'Response.End	
End If

UserGranteeID = GranteeID

sql = "SELECT G.GranteeID, G.GranteeName, G.OrganizationTypeID, G.ORI, G.StatePayeeIDNo, " & vbCrLf & _
	"	G.Address1, G.Address2, G.City, G.State, G.zip, G.OrganizationalUnit, GeneralPhone, " & vbCrLf & _
	"	VendorOrganizationalUnit, VendorAddress1, VendorAddress2, VendorCity, VendorState, VendorZIP, " & vbCrLf & _
	"	ISNULL(AuthorizedOfficialID,0) AS AuthorizedOfficialID, " & vbCrLf & _
	"	ISNULL(ProgramDirectorID,0) AS ProgramDirectorID, " & vbCrLf & _
	"	ISNULL(FinancialOfficerID,0) AS FinancialOfficerID, " & vbCrLf & _
	"	ISNULL(ProgramManagerID,0) AS ProgramManagerID, " & vbCrLf & _
	"	ISNULL(TaskForceCommanderID,0) AS TaskForceCommanderID, " & vbCrLf & _
	"	ISNULL(ProgramAdministrativeContactID,0) AS ProgramAdministrativeContactID, " & vbCrLf & _
	"	ISNULL(FinancialAdministrativeContactID,0) AS FinancialAdministrativeContactID, " & vbCrLf & _
	"	ISNULL(PIOID,0) AS PIOID, " & _
	"	BorderCounty, PortCounty, Port2County, G.TaskforceGrant, G.AuxiliaryGrant, G.RapidResponseStrikeforceGrant, " & vbCrLf & _
	"	G.CatalyticConverterGrant, G.Inactive, G.GranteeNameSort, G.UpdateID, G.UpdateTimestamp, U.Name AS UpdateName, " & vbCrLf & _
	"	CAST(CASE WHEN AuxiliaryGrant=1 AND ISNULL(TaskforceGrant,0)=0 THEN 1 ELSE 0 END AS BIT) AS ReducedOfficials " & vbCrLf & _
	"FROM Grantees AS G" & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=G.UpdateID " & vbCrLf & _
	"WHERE GranteeID=" & GranteeID
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = False Then
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	OrganizationTypeID = rs.Fields("OrganizationTypeID")
	ORI = rs.Fields("ORI")
	OrganizationalUnit = rs.Fields("OrganizationalUnit")
	GeneralPhone = rs.Fields("GeneralPhone")
	Address1 = rs.Fields("Address1")
	Address2 = rs.Fields("Address2")
	City = rs.Fields("City")
	State = rs.Fields("State")
	zip = rs.Fields("zip")
	StatePayeeIDNo = rs.Fields("StatePayeeIDNo")
	VendorOrganizationalUnit = rs.Fields("VendorOrganizationalUnit")
	VendorAddress1 = rs.Fields("VendorAddress1")
	VendorAddress2 = rs.Fields("VendorAddress2")
	VendorCity = rs.Fields("VendorCity")
	VendorState = rs.Fields("VendorState")
	VendorZIP = rs.Fields("VendorZIP")
	AuthorizedOfficialID = rs.Fields("AuthorizedOfficialID")
	ProgramDirectorID = rs.Fields("ProgramDirectorID")
	ProgramManagerID = rs.Fields("ProgramManagerID")
	FinancialOfficerID = rs.Fields("FinancialOfficerID")
	TaskForceCommanderID = rs.Fields("TaskForceCommanderID")
	ProgramAdministrativeContactID = rs.Fields("ProgramAdministrativeContactID")
	FinancialAdministrativeContactID = rs.Fields("FinancialAdministrativeContactID")
	PIOID = rs.Fields("PIOID")
	BorderCounty = rs.Fields("BorderCounty")
	PortCounty = rs.Fields("PortCounty")
	Port2County = rs.Fields("Port2County")
	TaskforceGrant = rs.Fields("TaskforceGrant")
	AuxiliaryGrant = rs.Fields("AuxiliaryGrant")
	RapidResponseStrikeforceGrant = rs.Fields("RapidResponseStrikeforceGrant")
	CatalyticConverterGrant = rs.Fields("CatalyticConverterGrant")
	Inactive = rs.Fields("Inactive")
	GranteeNameSort = rs.Fields("GranteeNameSort")
	UpdateID = rs.Fields("UpdateID")
	UpdateTimestamp = rs.Fields("UpdateTimeStamp")
	UpdateName = rs.Fields("UpdateName")
	ReducedOfficials = rs.Fields("ReducedOfficials")
Else
	GranteeName = ""
	OrganizationTypeID = 0
	ORI = 0
	OrganizationalUnit = ""
	GeneralPhone = ""
	Address1 = ""
	Address2 = ""
	City = ""
	State = ""
	zip = ""
	StatePayeeIDNo = ""
	VendorOrganizationalUnit = ""
	VendorAddress1 = ""
	VendorAddress2 = ""
	VendorCity = ""
	VendorState = ""
	VendorZIP = ""
	AuthorizedOfficialID = 0
	ProgramDirectorID = 0
	ProgramManagerID = 0
	FinancialOfficerID = 0
	TaskForceCommanderID = 0
	ProgramAdministrativeContactID = 0
	FinancialAdministrativeContactID = 0
	PIOID = 0
	BorderCounty = False
	PortCounty = False
	Port2County = False
	TaskforceGrant = False
	AuxiliaryGrant = False
	RapidResponseStrikeforceGrant = False
	CatalyticConverterGrant = False
	GranteeNameSort = ""
	UpdateID = 0
	UpdateTimestamp = ""
	UpdateName = ""
	ReducedOfficials = False
End If

PermitEdit = CheckPermissions(UserSystemID, GranteeID, True)

If MVCPARights = True Then
	ViewDocuments = True 
Else
	ViewDocuments = False
End If

If GranteeID=0 Then 
	' creating a new grantee. Allow edit.
	PermitEdit = True
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grantee Edit</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" /> 
<script type="text/javascript">
	function validateForm()
	{
		if (document.Grantee.GranteeName.value.length == 0) {
			alert("You must enter a legal name for the primary agency of the grant.");
			Grantee.GranteeName.focus();
			return false;
		}
		if (document.Grantee.ORI.selectedIndex < 1) {
			alert("You must select an ORI for a law enforcemnt agency or choose 'Not associated with any law enforcement entity'.");
			Grantee.ORI.focus();
			return false;
		}
		return true;
	}

	function changeOfficial(position, currentid)
	{
		document.ChangeOfficial.Position.value = position;
		document.ChangeOfficial.CurrentID.value = currentid;
		document.ChangeOfficial.submit();
	}
</script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">This screen is used to edit Grantees. A Grantee must be created within 
the system before you can begin an application for a grant. While there might be multiple
agencies involved in a grant, this page is used to describe the primary agency that
will administer the grant.
</div>

<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">
<form name="Grantee" method="post" action="GranteeSubmit.asp" onsubmit="return validateForm();">
<input type="hidden" name="GranteeID" value="<%=GranteeID %>" />
	<fieldset style="width: 748px;">
		<legend>Grantee Information</legend>
		<label for="GranteeName">Primary Agency / Grantee Legal Name:</label>
		<input type="text" name="GranteeName" id="GranteeName" value="<%=GranteeName %>" 
			size="50" maxlength="255" /><br />

		<label for="OrganizationTypeID">Organization Type:</label>
		<select name="OrganizationTypeID" id="OrganizationTypeID">
			<option value="0">Select an Organization Type</option>
<%
	sql = "SELECT * FROM Lookup.OrganizationType WITH (NOLOCK) ORDER BY OrganizationTypeID"
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("OrganizationTypeID") & """" & Selected(OrganizationTypeID, rs.Fields("OrganizationTypeID")) & ">" & _
			rs.Fields("OrganizationType") & "</option>" & vbCrLf)
		rs.MoveNext
	Wend
%>
		</select>

		<label for="ORI">ORI (if applicable):</label>
		<select name="ORI" id="ORI">
			<option value="">Select an ORI Option</option>
<%
	Response.Write("<option value=""None"" " & Selected(ORI, "None") & ">Not associated with any law enforcement entity</option>" & vbCrLf)
	sql = "SELECT A.ORI, A.Agency, B.County, A.CountyID " & vbCrLf & _
		"FROM Lookup.ORI AS A WITH (NOLOCK) " & vbCrLf & _
		"LEFT JOIN Lookup.Counties AS B WITH (NOLOCK) ON A.CountyID=B.CountyID " & vbCrLf & _
		"ORDER BY A.CountyID, A.ORI"
	Set rs = Con.Execute(sql)
	i = 1
	Response.Write("<optgroup label=""" & rs.Fields("County") & """>" & vbCrLf)
	While rs.EOF = False
		If i<>rs.Fields("CountyID") Then
			i = rs.Fields("CountyID")
			Response.Write("</optgroup>" & vbCrLf)
			Response.Write("<optgroup label=""" & rs.Fields("County") & """>" & vbCrLf)
		End If
		Response.Write("<option value=""" & rs.Fields("ORI") & """" & Selected(ORI, rs.Fields("ORI")) & ">" & rs.Fields("Agency") & _
			" [" & rs.Fields("ORI") & "]</option>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</optgroup>" & vbCrLf)
%>
		</select>
	</fieldset>

	<fieldset style="width: 748px;">
		<legend>Official Grantee Mailing Address</legend>
		<label for="OrganizationalUnit">Organizational Unit:</label>
		<input type="text" name="OrganizationalUnit" id="OrganizationalUnit" value="<%=OrganizationalUnit %>" 
			size="50" maxlength="255" /><br />
		<div class="detailnote" style="text-align: center">(Division or Department within organization to administer grant)</div>

		<label for="GeneralPhone">General Phone:</label>
		<input type="text" name="GeneralPhone" id="GeneralPhone" value="<%=GeneralPhone %>" 
			size="12" maxlength="20" /><br />

		<label for="Address1">Address (line 1):</label>
		<input type="text" name="Address1" id="Address1" value="<%=Address1 %>" 
			size="50" maxlength="255" /><br />

		<label for="Address2">Address (line 2):</label>
		<input type="text" name="Address2" id="Address2" value="<%=Address2 %>" 
			size="50" maxlength="255" /><br />

		<label for="City">City:</label>
		<input type="text" name="City" id="City" value="<%=City %>" 
			size="16" maxlength="20" /><br />

		<label for="State">State:</label>
		<input type="text" name="State" id="State" value="<%=State %>" 
			size="2" maxlength="2" /><br />

		<label for="ZIP">ZIP Code:</label>
		<input type="text" name="ZIP" id="ZIP" value="<%=Zip %>" 
			size="10" maxlength="10" /><br />
	</fieldset>


	<fieldset style="width: 748px;">
		<legend>Vendor Information for Grant Payment</legend>

		<label for="StatePayeeIDNo">State Payee ID Number:</label>
		<input type="text" name="StatePayeeIDNo" id="StatePayeeIDNo" value="<%=StatePayeeIDNo %>" size="20" maxlength="20" /><br />
		<div class="detailnote" style="text-align: center">(Issued by Comptroller of Public Accounts and must include designated mail code.)</div>

		<label for="VendorOrganizationalUnit">Organizational Unit:</label>
		<input type="text" name="VendorOrganizationalUnit" id="VendorOrganizationalUnit" value="<%=VendorOrganizationalUnit %>" 
			size="50" maxlength="255" /><br />
		<div class="detailnote" style="text-align: center">(Division or Department within organization receiving payment)</div>

		<label for="VendorAddress1">Address (line 1):</label>
		<input type="text" name="VendorAddress1" id="VendorAddress1" value="<%=VendorAddress1 %>" 
			size="50" maxlength="255" /><br />

		<label for="VendorAddress2">Address (line 2):</label>
		<input type="text" name="VendorAddress2" id="VendorAddress2" value="<%=VendorAddress2 %>" 
			size="50" maxlength="255" /><br />

		<label for="VendorCity">City:</label>
		<input type="text" name="VendorCity" id="VendorCity" value="<%=VendorCity %>" 
			size="16" maxlength="20" /><br />

		<label for="VendorState">State:</label>
		<input type="text" name="VendorState" id="VendorState" value="<%=VendorState %>" 
			size="2" maxlength="2" /><br />

		<label for="VendorZIP">ZIP Code:</label>
		<input type="text" name="VendorZIP" id="VendorZIP" value="<%=VendorZIP %>" 
			size="10" maxlength="10" /><br />
	</fieldset>

	<fieldset style="width: 748px;">
	<legend>Border / Port Designations</legend>

		<label for="BorderCounty">Border County Designation:</label>
		<input type="checkbox" name="BorderCounty" id="BorderCounty" value="1" <%=Checked(BorderCounty, True) %>  /><br />

		<label for="PortCounty">Port County Designation:</label>
		<input type="checkbox" name="PortCounty" id="PortCounty" value="1" <%=Checked(PortCounty, True) %>  /><br />

		<label for="Port2County">Port 2 (non-IW) County Designation (Harris, Houston, Pasadena, Victoria):</label>
		<input type="checkbox" name="Port2County" id="Port2County" value="1" <%=Checked(Port2County, True) %>  /><br />
	</fieldset>

<%	If GranteeID>0 Then %>
	<table style="width: 748px;">
	<tr><th colspan="3">Officials</th></tr>
<%
	ShowPosition "Authorized Official", AuthorizedOfficialID, PermitEdit
	ShowPosition "Program Director", ProgramDirectorID, PermitEdit
	ShowPosition "Financial Officer", FinancialOfficerID, PermitEdit

	If ReducedOfficials = True Then
		Response.Write("<tr><th colspan=""3"">Optional Positions</th></tr>" & vbCrLF)
	End If

	ShowPosition "Program Manager", ProgramManagerID, PermitEdit
	ShowPosition "Program Administrative Contact", ProgramAdministrativeContactID, PermitEdit
	ShowPosition "Financial Administrative Contact", FinancialAdministrativeContactID, PermitEdit

	If ReducedOfficials = False or TaskForceCommanderID > 0 Then
		ShowPosition "Taskforce Commander", TaskForceCommanderID, PermitEdit
	End If
	If ReducedOfficials = False or PIOID > 0 Then
		ShowPosition "PIO / Media Contact", PIOID, PermitEdit
	End If

	If IsNull(FinancialOfficerID) = False Then
		If IsNull(AuthorizedOfficialID) = False Then
			If FinancialOfficerID = AuthorizedOfficialID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Authorized Official cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If ProgramDirectorID > 0 Then
			If FinancialOfficerID = ProgramDirectorID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Program Director cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If ProgramManagerID > 0 Then
			If FinancialOfficerID = ProgramManagerID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Program Manager cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If ProgramAdministrativeContactID > 0 Then
			If FinancialOfficerID = ProgramAdministrativeContactID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Program Administrative Contact cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If IsNull(FinancialAdministrativeContactID) = False Then
			If FinancialOfficerID = FinancialAdministrativeContactID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Financial Administrative Contact cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If IsNull(TaskForceCommanderID) = False Then
			If FinancialOfficerID = TaskForceCommanderID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Task Force Commander cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If IsNull(PIOID) = False Then
			If FinancialOfficerID = PIOID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Public Information Officer cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

End If

If GranteeID>0 And ViewDocuments = True Then
	Dim Folder, file, files, DocumentFolder, fso, sql2, rs2
	MonitorDocumentCounter = 0
	Response.Write("<tr style=""vertical-align: top; ""><td colspan=""3"">&nbsp;</td></tr>")
	Response.Write("<tr style=""vertical-align: top; ""><td colspan=""3"" style=""text-align: center"">Current Documents in Monitoring folder(s): ")
	sql2 = "SELECT MonitorID, FiscalYear FROM Monitor.Main WHERE GranteeID=" & GranteeID
	If Debug = True Then
		Response.Write("<pre>" &sql2 & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs2 = Con.Execute(sql2)
	while rs2.EOF = False
		DocumentFolder = Application("DocumentRoot") & "\Monitor\" & rs2.Fields("MonitorID") & "\"
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		Response.Write("</td>" & vbCrLf)
		If fso.FolderExists(DocumentFolder) Then
			Set folder = fso.GetFolder(DocumentFolder)
			Set files = folder.Files
			If files.count>0 Then 
				Response.Write("<tr style=""vertical-align: top; ""><td colspan=""3"" style=""text-align: center; "">FY" & rs2.Fields("FiscalYear") & "</td></tr>")
				Response.Write("<tr><td colspan=""3"">")
				For Each file in files
					Response.Write("<a href=""../Documents/Monitor/" & rs2.Fields("MonitorID") & "/" & file.Name & _
						""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
					MonitorDocumentCounter = MonitorDocumentCounter + 1
				Next
				Response.Write("</td></tr>" & vbCrLf)
			End If
		End If
		rs2.MoveNext
	Wend
	If MonitorDocumentCounter = 0 Then
		Response.Write("<tr style=""vertical-align: top; ""><td colspan=""3"" style=""text-align: center"">No Documents in folder</td></tr>")
	End If
End If

%>
	</table><br />
<%	If MVCPARights = True Then %>
	<fieldset style="width: 748px;">
		<legend>For MVCPA Use Only</legend>
		<label for="GranteeName">Grantee Name for Sort:</label>
		<input type="text" name="GranteeNameSort" id="GranteeNameSort" value="<%=GranteeNameSort %>" 
			size="50" maxlength="255" /><br />
		<div class="detailnote" style="text-align: center">(Typically just remove "City of ", if present, at start of grantee name.)</div>
	<br />
		<div style="text-align: center; ">Enable grantee for the following types of grants.</div>
		<label for="TaskforceGrant">Taskforce Grant:</label>
		<input type="checkbox" name="TaskforceGrant" id="TaskforceGrant" value="1" <%=Checked(TaskforceGrant, True) %>  /><br />

		<label for="AuxiliaryGrant">Auxiliary Grant:</label>
		<input type="checkbox" name="AuxiliaryGrant" id="AuxiliaryGrant" value="1" <%=Checked(AuxiliaryGrant, True) %>  /><br />

		<label for="RapidResponseStrikeforceGrant">Rapid Response Strikeforce Grant:</label>
		<input type="checkbox" name="RapidResponseStrikeforceGrant" id="RapidResponseStrikeforceGrant" value="1" <%=Checked(RapidResponseStrikeforceGrant, True) %>  /><br />

		<label for="CatalyticConverterGrant">Catalytic Converter Grant:</label>
		<input type="checkbox" name="CatalyticConverterGrant" id="CatalyticConverterGrant" value="1" <%=Checked(CatalyticConverterGrant, True) %>  /><br />
		<br />
		<label for="Inactive">Grantee is Inactive (No grants or Apps):</label>
		<input type="checkbox" name="Inactive" id="Inactive" value="1" <%=Checked(Inactive, True) %>  /><br />
	</fieldset>
	<br />
<%	End If %>

	<div style="width: 748px; text-align: center;">
<%	If PermitEdit = True Then %>
	<input type="button" value="Cancel" title="Return to homepage" style="text-align: center" onclick="location.href = '../Home/Default.asp?GranteeID=<%=GranteeID%>';"/>
	<input type="submit" value="Save Changes" title="Save Grantee Information" style="text-align: center" />
	<input type="reset" value="Reset" title="reset form" style="text-align: center" />
<%	Else %>
	<input type="button" value="Home" onclick="location.href = '../Home/Default.asp?GranteeID=<%=GranteeID%>';" />
<%	End If %>
	</div>
</form>

<form name="ChangeOfficial" method="post" action="../Grantees/ChangeOfficial1.asp">
<input type="hidden" name="ReturnPage" value="../Grantees/Grantee.asp" />
<input type="hidden" name="GranteeID" value="<%=GranteeID %>" />
<input type="hidden" name="Position" value="" />
<input type="hidden" name="CurrentID" value="" />
</form>
</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/ShowPosition.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->