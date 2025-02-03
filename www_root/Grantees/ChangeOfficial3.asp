<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, UpdateSystemID, Position, GranteeID, ReturnPage, fieldname, SearchLastName, GranteeName, CurrentID, CurrentName, _
	FirstName, MiddleName, LastName, Name, Suffix, Title, email, _
	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile
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

If Request.Form.Count>0 Then
	UpdateSystemID = Request.Form("UpdateSystemID")
	Position = Request.Form("Position")
	ReturnPage = Request.Form("ReturnPage")
	GranteeID = Request.Form("GranteeID")
	SearchLastName = Request.Form("SearchLastName")
	CurrentID = Request.Form("CurrentID")
Else
	Response.Write("Error: No values for change official")
	Response.End
End If

If Position = "Authorized Official" Then
	fieldname = "AuthorizedOfficialID"
ElseIf Position = "Program Director" Then
	fieldname = "ProgramDirectorID" 
ElseIf Position = "Program Manager" Then
	fieldname = "ProgramManagerID" 
ElseIf Position = "Financial Officer" Then
	fieldname = "FinancialOfficerID"
ElseIf Position = "Program Administrative Contact" Then
	fieldname = "ProgramAdministrativeContactID"
ElseIf Position = "Financial Administrative Contact" Then
	fieldname = "FinancialAdministrativeContactID"
ElseIf Position = "Taskforce Commander" Then
	fieldname = "TaskForceCommanderID"
ElseIf Position = "PIO / Media Contact" Then
	fieldname = "PIOID"
Else
	Response.Write("Error: Invalid position title")
	Response.End
End If

sql = "SELECT U.Name AS CurrentName, G." & fieldname & " AS CurrentID, G.GranteeName " & vbCrLF & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=G." & fieldname & " " & vbCrLf & _
	"WHERE G.GranteeID=" & prepIntegerSQL(GranteeID)
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: Retrieving Grantee")
	Response.End
Else
	CurrentName = rs.Fields("CurrentName")
	GranteeName = rs.Fields("GranteeName")
	CurrentID = rs.Fields("CurrentID")
End If

sql = "SELECT FirstName, MiddleName, LastName, Name, Suffix, Title, email, " & vbCrLf & _
	"	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile " & vbCrLF & _
	"FROM System.Users " & vbCrLF & _
	"WHERE SystemID=" & prepNumberSQL(UpdateSystemID)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
If rs.EOF = False Then
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
	Fax = rs.Fields("Fax")
	Mobile = rs.Fields("Mobile")
Else
	FirstName = ""
	MiddleName = ""
	LastName = ""
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
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Change <%=Position %> Position for <%=GranteeName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" />
<script type="text/javascript">
	function validateForm()
	{
		if (document.ChangeOfficial.LastName.value.length == 0) {
			alert("You must enter a last name for the user.");
			document.ChangeOfficial.LastName.focus();
			return false;
		}
		if (document.ChangeOfficial.FirstName.value.length == 0) {
			alert("You must enter a first name for the user.");
			document.ChangeOfficial.FirstName.focus();
			return false;
		}
		if (document.ChangeOfficial.email.value.length == 0) {
			alert("You must enter an email address for the user.");
			document.ChangeOfficial.email.focus();
			return false;
		}
		for (i = 0; i < ChangeOfficial.GranteePermissions.length; i++) {
			ChangeOfficial.GranteePermissions.options[i].selected = true;
		}
	}

	function addGrantee()
	{
		if (ChangeOfficial.Grantee.selectedIndex > 0) {
			ChangeOfficial.GranteePermissions.options[ChangeOfficial.GranteePermissions.length] =
				new Option(ChangeOfficial.Grantee.options[ChangeOfficial.Grantee.selectedIndex].text, ChangeOfficial.Grantee[ChangeOfficial.Grantee.selectedIndex].value);
			ChangeOfficial.Grantee.remove(ChangeOfficial.Grantee.selectedIndex);
		}
	}

	function removeGrantee()
	{
		for (i = 0; i < ChangeOfficial.GranteePermissions.length; i++) {
			if (ChangeOfficial.GranteePermissions.options[i].selected) {
				ChangeOfficial.Grantee.options[ChangeOfficial.Grantee.length] =
					new Option(ChangeOfficial.GranteePermissions.options[i].text, ChangeOfficial.GranteePermissions.options[i].value);
				ChangeOfficial.GranteePermissions.remove(i);
				i--;
			}
		}
		ChangeOfficial.Grantee.selectedIndex = 0;
	}
</script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>
<div class="pagetag">Change <%=Position %> Position for <%=GranteeName %>
<%
If IsNull(CurrentName) = True Then
	Response.Write("<p>This position is currently empty.</p>")
Else
	Response.Write("<p>This position is currently filled by " & CurrentName & ".</p>")
End If
%>
</div>
<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>
<div class="content">

<form name="ChangeOfficial" id="ChangeOfficial" method="post" action="ChangeOfficial4.asp" onSubmit="return validateForm();">
	<input type="hidden" name="UpdateSystemID" value="<%=UpdateSystemID %>" />
	<input type="hidden" name="SearchLastName" value="<%=SearchLastName %>" />
	<input type="hidden" name="ReturnPage" value="<%=ReturnPage %>" />
	<input type="hidden" name="GranteeID" value="<%=GranteeID %>" />
	<input type="hidden" name="Position" value="<%=Position %>" />
	<input type="hidden" name="CurrentID" value="<%=CurrentID %>" />
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
		<label for="EMail">E-Mail Address:</label>
<%	If UpdateSystemID=0 Or MVCPARights = True Then %>
		<input type="text" name="email" id="email" value="<%=email %>" size="50" maxlength="255" /><br />
<%	Else %>
		<input type="text" name="email" id="Text1" value="<%=email %>" size="50" maxlength="255" readonly="readonly"
			title="An email address may only be changed by the actual user or by an MVCPA Administrator." /><br />
<%	End If %>
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
		<label for="Fax">Mobile:</label>
		<input type="text" name="Mobile" id="Mobile" value="<%=Mobile %>" size="20" maxlength="20" /><br />
	</fieldset>

	<div style="text-align: center;width: 720px"><input type="submit" value="Submit" /></div>
</form>

</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../menu/DBMenu.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/prepWeb.asp"-->