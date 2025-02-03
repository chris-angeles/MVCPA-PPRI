<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, UpdateSystemID, UserID, FirstName, MiddleName, LastName, Name, Suffix, Title, _
	email, Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, _
	LicensedPeaceOfficer, TCOLEPID, DefaultGrantee,  _
	DeveloperRole, MVCPAAdministratorRole, MVCPAAuditorRole, _
	MVCPAGrantCoordinatorRole, MVCPAAdministrativeAssistantRole, _
	MVCPAScorerRole, MVCPAViewerRole, MVCPAStaffRole, AccountDisabled, _
	LastPasswordChange, LastLogin, AccountCreated, Comments
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

If Len(Request.Form("UpdateSystemID"))>0 Then
	UpdateSystemID = Request.Form("UpdateSystemID")
ElseIf Len(Request.QueryString("SystemID"))>0 Then
	UpdateSystemID = Request.QueryString("SystemID")
End If
If Len(Request.Form("LastName"))>0 Then
	LastName = Request.Form("LastName")
ElseIf Len(Request.QueryString("LastName"))>0 Then
	LastName = Request.QueryString("LastName")
End If


sql = "SELECT UserID, FirstName, MiddleName, LastName, Name, Suffix, Title, email, " & vbCrLf & _
	"	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, " & vbCrLF & _
	"	LicensedPeaceOfficer, TCOLEPID, ISNULL(DefaultGrantee,0) AS DefaultGrantee, " & vbCrLf & _
	"	Developer, MVCPAAdministrator, MVCPAAuditor, MVCPAGrantCoordinator, MVCPAAdministrativeAssistant, " & vbCrLf & _
	"	MVCPAScorer, MVCPAViewer, MVCPAStaff, AccountDisabled, LastPasswordChange, Comments, " & vbCrLF & _
	"	LastLogin = ISNULL(CAST((SELECT MAX(LoginTime) AS LastLogin FROM [System].LoginLog WHERE SystemID=System.Users.SystemID) AS VARCHAR),'Never'), " & vbCrLf & _
	"	AccountCreated = CAST((SELECT MIN(UpdateTimestamp) AS AccountCreated FROM [System].Users_Log WHERE SystemID=System.Users.SystemID) AS VARCHAR) " & vbCrLf & _
	"FROM [System].Users " & vbCrLF & _
	"WHERE SystemID=" & prepNumberSQL(UpdateSystemID)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
If rs.EOF = False Then
	UserID = rs.Fields("UserID")
	FirstName = rs.Fields("FirstName")
	MiddleName = rs.Fields("MiddleName")
	LastName = rs.Fields("LastName")
	Name = rs.Fields("Name")
	Suffix = rs.Fields("Suffix")
	Title = rs.Fields("Title")
	email = rs.Fields("email")
	Address1 = rs.Fields("Address1")
	Address2 = rs.Fields("Address2")
	City = rs.Fields("City")
	State = rs.Fields("State")
	ZIP = rs.Fields("ZIP")
	Phone = rs.Fields("Phone")
	Mobile = rs.Fields("Mobile")
	Fax = rs.Fields("Fax")
	LicensedPeaceOfficer = rs.Fields("LicensedPeaceOfficer")
	TCOLEPID = rs.Fields("TCOLEPID")
	DefaultGrantee = rs.Fields("DefaultGrantee")
	DeveloperRole = rs.Fields("Developer")
	MVCPAAdministratorRole = rs.Fields("MVCPAAdministrator")
	MVCPAAuditorRole = rs.Fields("MVCPAAuditor")
	MVCPAGrantCoordinatorRole = rs.Fields("MVCPAGrantCoordinator")
	MVCPAAdministrativeAssistantRole = rs.Fields("MVCPAAdministrativeAssistant")
	MVCPAViewerRole = rs.Fields("MVCPAViewer")
	MVCPAScorerRole = rs.Fields("MVCPAScorer")
	MVCPAStaffRole = rs.Fields("MVCPAStaff")
	AccountDisabled = rs.Fields("AccountDisabled")
	LastPasswordChange = rs.Fields("LastPasswordChange")
	LastLogin = rs.Fields("LastLogin")
	AccountCreated = rs.Fields("AccountCreated")
	Comments = rs.Fields("Comments")
Else
	UserID = ""
	FirstName = ""
	MiddleName = ""
	Name = ""
	Suffix = ""
	Title = ""
	email = ""
	Address1 = ""
	Address2 = ""
	City = ""
	State = "TX"
	ZIP = ""
	Phone = ""
	Fax = ""
	Mobile = ""
	LicensedPeaceOfficer = False
	TCOLEPID = ""
	DefaultGrantee=0
	DeveloperRole = False
	MVCPAAdministratorRole = False
	MVCPAAuditorRole = False
	MVCPAGrantCoordinatorRole = False
	MVCPAAdministrativeAssistantRole = False
	MVCPAScorerRole = False
	MVCPAViewerRole = False
	MVCPAStaffRole = False
	AccountDisabled = False
	LastPasswordChange = null
	LastLogin = null
	AccountCreated = null
	Comments = ""
End If

If Debug = True Then
	Response.Write("<pre>DeveloperRole=" & DeveloperRole & "; MVCPAAuditorRole=" & MVCPAAuditorRole &  "; MVCPAGrantCoordinatorRole=" & MVCPAGrantCoordinatorRole & "; </pre>")
	Response.Flush
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Add or Update A User</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" />
<script type="text/javascript">
	function validateForm()
	{
		if (document.UpdateUser.LastName.value.length == 0) {
			alert("You must enter a last name for the user.");
			document.UpdateUser.LastName.focus();
			return false;
		}
		if (document.UpdateUser.FirstName.value.length == 0) {
			alert("You must enter a first name for the user.");
			document.UpdateUser.FirstName.focus();
			return false;
		}
		if (document.UpdateUser.email.value.length == 0) {
			alert("You must enter an email address for the user.");
			document.UpdateUser.email.focus();
			return false;
		}
		for (i = 0; i < UpdateUser.GranteePermissions.length; i++) {
			UpdateUser.GranteePermissions.options[i].selected = true;
		}
	}

	function addGrantee()
	{
		if (UpdateUser.Grantee.selectedIndex > 0) {
			UpdateUser.GranteePermissions.options[UpdateUser.GranteePermissions.length] = 
				new Option(UpdateUser.Grantee.options[UpdateUser.Grantee.selectedIndex].text, UpdateUser.Grantee[UpdateUser.Grantee.selectedIndex].value);
			UpdateUser.Grantee.remove(UpdateUser.Grantee.selectedIndex);
		}
	}

	function removeGrantee()
	{
		for (i = 0; i < UpdateUser.GranteePermissions.length;i++)
		{
			if (UpdateUser.GranteePermissions.options[i].selected) {
				UpdateUser.Grantee.options[UpdateUser.Grantee.length] =
					new Option(UpdateUser.GranteePermissions.options[i].text, UpdateUser.GranteePermissions.options[i].value);
				UpdateUser.GranteePermissions.remove(i);
				i--;
			}
		}
		UpdateUser.Grantee.selectedIndex = 0;
	}
</script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Add or Update A User</div>

<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">

<form name="UpdateUser" id="UpdateUser" method="post" action="UpdateUser4.asp" onSubmit="return validateForm();">
	<input type="hidden" name="UpdateSystemID" value="<%=UpdateSystemID %>" />
	<fieldset style="width: 760px;">
		<legend>Name and Title</legend>
		<label for="FirstName">First Name:</label>
		<input type="text" name="FirstName" id="FirstName" value="<%=FirstName %>" size="24" maxlength="24" /><br />
		<label for="MiddleName">Middle Name:</label>
		<input type="text" name="MiddleName" id="MiddleName" value="<%=MiddleName %>" size="24" maxlength="24" /><br />
		<label for="LastName">LastName:</label>
		<input type="text" name="LastName" id="LastName" value="<%=LastName %>" size="24" maxlength="24" /><br />
		<label for="Suffix">Suffix:</label>
		<input type="text" name="Suffix" id="Suffix" value="<%=Suffix %>" size="20" maxlength="20" /><br />
		<label for="Title">Position Title:</label>
		<input type="text" name="Title" id="Title" value="<%=Title %>" size="50" maxlength="100" /><br />
	</fieldset>

	<fieldset style="width: 760px;">
		<legend>Contact Information</legend>
		<label for="email">E-Mail Address:</label>
		<input type="text" name="email" id="email" value="<%=email %>" size="50" maxlength="255" /><br />
		<label for="Address1">Address 1:</label>
		<input type="text" name="Address1" id="Address1" value="<%=Address1 %>" size="50" maxlength="50" /><br />
		<label for="Address2">Address 2:</label>
		<input type="text" name="Address2" id="Address2" value="<%=Address2 %>" size="50" maxlength="50" /><br />
		<label for="City">City:</label>
		<input type="text" name="City" id="City" value="<%=City %>" size="20" maxlength="20" /><br />
		<label for="State">State:</label>
		<input type="text" name="State" id="State" value="<%=State %>" size="2" maxlength="2" /><br />
		<label for="ZIP">ZIP:</label>
		<input type="text" name="ZIP" id="ZIP" value="<%=ZIP %>" size="10" maxlength="10" /><br />
		<label for="Phone">Phone:</label>
		<input type="text" name="Phone" id="Phone" value="<%=Phone %>" size="20" maxlength="20" /><br />
		<label for="Fax">Fax:</label>
		<input type="text" name="Fax" id="Fax" value="<%=Fax %>" size="20" maxlength="20" /><br />
		<label for="Mobile">Mobile:</label>
		<input type="text" name="Mobile" id="Mobile" value="<%=Mobile %>" size="20" maxlength="20" /><br />
	</fieldset>

	<fieldset style="width: 540px;">
		<legend>Other</legend>
		<label style="width: 540px; text-align: center;"><input type="checkbox" 
		name="LicensedPeaceOfficer" id="LicensedPeaceOfficer" value="1" <%=Checked(LicensedPeaceOfficer,true) %>/> 
		Check if this person is a Licensed Peace Officer</label>
		<label for="TCOLEPID">TCOLE PID:</label>
		<input type="text" name="TCOLEPID" id="TCOLEPID" value="<%=TCOLEPID %>" size="7" maxlength="7" /><br />
	</fieldset>

	<fieldset style="width: 540px;">
		<legend>MVCPA User roles</legend>
<%	If Developer = True Then %>
		<label style="width: 540px; text-align: left;"><input type="checkbox" 
		name="Developer" id="Developer" value="1" <%=Checked(DeveloperRole,True) %>/> 
		Developer</label>
<%	Else %>
		<%=CheckBoxField2("Developer", DeveloperRole, False) %> Developer
			<input type="hidden" name="Developer" value="<%=DeveloperRole %>" /><br />
<%	End If 
	If MVCPAAdministrator = True Or Developer = True Then %>
		<label style="width: 540px; text-align: left;"><%=CheckBoxField2("MVCPAAdministrator", MVCPAAdministratorRole, True) %>
		MVCPA Director</label>
<%	Else %>
		<%=CheckBoxField2("MVCPAAdministrator", MVCPAAdministratorRole, False) %> MVCPA Director
		<input type="hidden" name="MVCPAAdministrator" value="<%=MVCPAAdministratorRole %>" /><br />
<%	End If 
	If MVCPAAdministrator = True Or Developer = True Or MVCPAAuditor = True Then %>
		<label style="width: 540px; text-align: left;"><%=CheckBoxField2("MVCPAAuditor", MVCPAAuditorRole, True) %> 
		MVCPA Auditor</label>
<%	Else %>
		<%=CheckBoxField2("MVCPAAuditor", MVCPAAuditorRole, False) %> MVCPA Auditor
<%	End If %>
		<label style="width: 540px; text-align: left;"><input type="checkbox" 
		name="MVCPAGrantCoordinator" id="MVCPAGrantCoordinator" value="1" <%=Checked(MVCPAGrantCoordinatorRole,true) %>/> 
		MVCPA Grant Coordinator</label>

		<label style="width: 540px; text-align: left;"><input type="checkbox" 
		name="MVCPAAdministrativeAssistant" id="MVCPAAdministrativeAssistant" value="1" <%=Checked(MVCPAAdministrativeAssistantRole,true) %>/> 
		MVCPA Administrative Assistant</label>

		<label style="width: 540px; text-align: left;"><input type="checkbox" 
		name="MVCPAScorer" id="MVCPAScorer" value="1" <%=Checked(MVCPAScorerRole,true) %>/> 
		MVCPA Scorer</label>

		<label style="width: 540px; text-align: left;"><input type="checkbox" 
		name="MVCPAViewer" id="MVCPAViewer" value="1" <%=Checked(MVCPAViewerRole,true) %>/> 
		MVCPA Viewer</label>

		<label style="width: 540px; text-align: left;"><input type="checkbox" 
		name="MVCPAStaff" id="MVCPAStaff" value="1" <%=Checked(MVCPAStaffRole,true) %>/> 
		MVCPA Staff (used to include user on current staff lists and dropdowns)</label>
	</fieldset>

	<fieldset style="width: 540px;">
		<legend>Account</legend>
		<label style="width: 540px; text-align: left;"><input type="checkbox" 
		name="AccountDisabled" id="AccountDisabled" value="1" <%=Checked(AccountDisabled,true) %>/> 
		Account Disabled</label>
<%	If UpdateSystemID>0 Then %>
		UserID = <%=UserID%>
<%	If UserSystemID = 1 Then
		Response.Write(" <a href=""../Admin/Impersonate.asp?UserID=" & UserID & """>Impersonate</a>")
	End If 
%><br />
		Last Password Change: <%=LastPasswordChange %> <a href="" onclick="document.ResetPassword.submit(); return false;">reset password</a><br />
		Last Login: <%=LastLogin %><br />
		Account Created: <%=AccountCreated %><br />
		SystemID: <%=UpdateSystemID %>
<%	End If %>
	</fieldset>

	<fieldset style="width: 732px;">
		<legend>Access Permissions to Grantees</legend>
	Do not add permssions for administrators. They automatically have all grantee access.<br />
	<select name="Grantee" id="Grantee">
		<option value="0">Select Grantee to add</option>
<%
	If UpdateSystemID>0 Then
		sql = "SELECT A.GranteeID, A.GranteeName " & vbCrLF & _
			"FROM Grantees AS A " & vbCrLF & _
			"LEFT JOIN System.GranteePermissions AS B ON B.GranteeID=A.GranteeID AND SystemID=" & prepIntegerSQL(UpdateSystemID) & " " & vbCrLf & _
			"WHERE B.GranteeID IS NULL " & vbCrLF & _
			"ORDER BY A.GranteeName "
	Else
		sql = "SELECT A.GranteeID, A.GranteeName " & vbCrLF & _
			"FROM Grantees AS A " & vbCrLF & _
			"ORDER BY A.GranteeName "		
	End If
	Set rs = Con.Execute(sql)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("GranteeID") & """>" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
		rs.MoveNext
	Wend
%>
	</select><input type="button" name="AddGrantee" id="AddGrantee" value="Add Grantee" 
		title="Pick a grantee from the dropdown menu and then click on this button to add them to the selected list."
		style="display: inline; width: 100px;" onclick="addGrantee();" />
<br />
	<select name="GranteePermissions" multiple size="4" style="width: 500px">
<%
	sql = "SELECT A.SystemID, A.GranteeID, B.GranteeName " & vbCrLF & _
		"FROM System.GranteePermissions AS A" & vbCrLF & _
		"LEFT JOIN Grantees AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
		"WHERE SystemID = " & prepIntegerSQL(UpdateSystemID) & vbCrLF & _
		"ORDER BY GranteeName "
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<option value=""" & rs.Fields("GranteeID") & """>" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
		rs.MoveNext
	Wend
%>
	</select>
	<input type="button" name="RemoveGrantee" value="Remove" title="Select a grantee in the list and click on this button to remove them from the selected list."
		style="display: inline; width: 100px; vertical-align: top; " onclick="removeGrantee();" /><br />
	</fieldset>

	<fieldset style="width: 732px;">
		<legend>Default Grantee</legend>
	<label for="DefaultGrantee">Default Grantee to be shown at login:</label>
		<select name="DefaultGrantee" id="DefaultGrantee">
			<option value="0">Select Grantee</option>
<%
	If MVCPARights Then
		sql = "SELECT G.GranteeID, G.GranteeName " & vbCrLf & _
			"FROM Grantees AS G " & vbCrLf & _
			"ORDER BY G.GranteeName "
	Else
		sql = "SELECT G.GranteeID, G.GranteeName " & vbCrLf & _
			"FROM Grantees AS G " & vbCrLf & _
			"JOIN System.GranteePermissions AS P ON P.SystemID=" & prepIntegerSQL(UpdateSystemID) & " AND P.GranteeID=G.GranteeID " & vbCrLf & _
			"ORDER BY G.GranteeName "
	End If
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	set rs=Con.Execute(sql)
	While rs.EOF = False
		If rs.Fields("GranteeID") = DefaultGrantee Then
			Response.Write("<option value=""" & rs.Fields("GranteeID") & """ selected >" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
		Else
			Response.Write("<option value=""" & rs.Fields("GranteeID") & """>" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
		End If
		rs.MoveNext()
	Wend
%>
		</select>
	</fieldset>

	<br />
	Comments: 
	<textarea name="Comments" rows="3" cols="80" maxlength="1000" style="vertical-align: top; "><%=Comments %></textarea><br />

	<div style="text-align: center; width: 732px">
		<input type="submit" value="Submit" />
		<input type="button" value="Cancel" onclick="location.href = '../Home/default.asp';" />
	</div>
</form>

<form name="ResetPassword" method="post" action="ResetPassword.asp" target="_blank">
<input type="hidden" name="UserID" value="<%=UserID %>" />
<input type="hidden" name="Source" value="UpdateUser3.asp" />
</form>
</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->