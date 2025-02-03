<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FirstName, MiddleName, LastName, Name, Suffix, Title, _
	LicensedPeaceOfficer, TCOLEPID, DefaultGrantee, Email, _
	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile
debug = False
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
End If

If Request.Form.Count>0 Then
	FirstName = Request.Form("FirstName")
	MiddleName = Request.Form("MiddleName")
	LastName = Request.Form("LastName")
	Suffix = Request.Form("Suffix")
	Title = Request.Form("Title")
	email = Request.Form("email")
	Address1 = Request.Form("Address1")
	Address2 = Request.Form("Address2")
	City = Request.Form("City")
	State = Request.Form("State")
	ZIP = Request.Form("ZIP")
	Phone = Request.Form("Phone")
	Fax = Request.Form("Fax")
	Mobile = Request.Form("Mobile")
	LicensedPeaceOfficer = Request.Form("LicensedPeaceOfficer")
	TCOLEPID = Request.Form("TCOLEPID")
	DefaultGrantee = Request.Form("DefaultGrantee")

	sql = "UPDATE System.Users SET UserID=" & prepStringSQL(email) & _
		", FirstName=" & prepStringSQL(FirstName) & _
		", MiddleName=" & prepStringSQL(MiddleName) & _
		", LastName=" & prepStringSQL(LastName) & _
		", Suffix=" & prepStringSQL(Suffix) & _
		", Title=" & prepStringSQL(Title) & _
		", email=" & prepStringSQL(email) & _
		", Address1=" & prepStringSQL(Address1) & _
		", Address2=" & prepStringSQL(Address2) & _
		", City=" & prepStringSQL(City) & _
		", State=" & prepStringSQL(State) & _
		", ZIP=" & prepStringSQL(ZIP) & _
		", Phone=" & prepStringSQL(Phone) & _
		", Fax=" & prepStringSQL(Fax) & _
		", Mobile=" & prepStringSQL(Mobile) & _
		", LicensedPeaceOfficer=" & prepBitRequiredSQL(LicensedPeaceOfficer) & _
		", TCOLEPID=" & prepStringSQL(TCOLEPID) & _
		", DefaultGrantee=" & prepIntegerSQL(DefaultGrantee) & _
		", UpdateID=" & UserSystemID & _
		", UpdateTimestamp=getdate() " & vbCrLF & _
		"WHERE SystemID=" & UserSystemID
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)

	If Debug = True Then
		Response.Write("<a href=""/Home/Default.asp"">Go To Home Page</a>") & vbCrLF
	Else
		Response.Redirect("/Home/Default.asp")
	End If

End If

sql = "SELECT FirstName, MiddleName, LastName, Name, Suffix, Title, email, " & vbCrLf & _
	"	Address1, Address2, City, State, ZIP, Phone, Fax, Mobile, " & vbCrLf & _
	"	LicensedPeaceOfficer, TCOLEPID, ISNULL(DefaultGrantee,0) AS DefaultGrantee " & vbCrLf & _
	"FROM System.Users " & vbCrLF & _
	"WHERE SystemID=" & prepNumberSQL(UserSystemID)
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
	LicensedPeaceOfficer = rs.Fields("LicensedPeaceOfficer")
	TCOLEPID = rs.Fields("TCOLEPID")
	DefaultGrantee = rs.Fields("DefaultGrantee")
Else
	FirstName = ""
	MiddleName = ""
	LastName = ""
	Name = UserName
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
	DefaultGrantee = 0
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Update User Profile</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" />
<script type="text/javascript">
	function validateForm()
	{
		if (document.UpdateProfile.email.value.length == 0) {
			alert("You must provide an email address with your profile!");
			document.UpdateProfile.email.focus();
			return false;
		}
		return true;
	}
</script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>
<div class="pagetag">Update Profile For <%=Name %></div>
<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>
<div class="content">

<form name="UpdateProfile" id="UpdateProfile" method="post" action="UpdateProfile.asp" onSubmit="return validateForm();">
	<fieldset style="width: 732px;">
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

	<fieldset style="width: 732px;">
		<legend>Contact Information</legend>
		<label for="EMail">E-Mail Address:</label>
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
		<label for="Fax">Mobile:</label>
		<input type="text" name="Mobile" id="Mobile" value="<%=Mobile %>" size="20" maxlength="20" /><br />
	</fieldset>

	<fieldset style="width: 732px;">
		<legend>For Licensed Peace Officers</legend>
	<label for="LicensedPeaceOfficer" style="width: 650px; text-align: left; display: inline-block;">
		<input type="checkbox" name="LicensedPeaceOfficer" id="LicensedPeaceOfficer" value="1" <%
	If LicensedPeaceOfficer = True Then	Response.Write(" Checked ") %> /> I am a Licensed Peace Officer in the State of Texas.</label><br />
		<label for="TCOLEPID" style="width: 625px; text-align: left; display: inline-block;">Texas Commission on Law Enforcement (TCOLE) Personal Identification Number (PID): </label>
		<input type="text" name="TCOLEPID" id="TCOLEPID" value="<%=TCOLEPID %>" size="7" 
			maxlength="7" style="text-align: left; display: inline-block;"/><br />
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
			"JOIN System.GranteePermissions AS P ON P.SystemID=" & prepIntegerSQL(USerSystemID) & " AND P.GranteeID=G.GranteeID " & vbCrLf & _
			"ORDER BY G.GranteeName "
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

	<div style="text-align: center">
		<input type="submit" value="Submit" />
		<input type="button" value="Cancel" onclick="location.href = '../Home/Default.asp';" />
	</div>
</form>

</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepDB.asp"-->