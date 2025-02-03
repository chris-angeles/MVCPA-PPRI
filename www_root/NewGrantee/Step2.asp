<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"--><% 
Dim debug, i, SystemID, UserID, Name, FirstName, MiddleName, LastName, Suffix, Title, email, Address1, Address2, City, State, ZIP, Phone, Fax
SystemID = 0
debug = True
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
	Name = Request.Form("Name")
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

	sql = "SELECT UserID, Email, Name FROM System.Users WHERE UserID=" & prepStringSQL(email) & _
		" OR email=" & prepStringSQL(email)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then 
		Response.Write("A user with this email address already exists in system. " & _
			"Please login to create application.")
		Response.End
	End If

	sql = "SELECT UserID, Email, Name FROM System.Users WHERE FirstName=" & _
		prepStringSQL(firstname)
	If Len(middlename)=0 Then
		sql = sql & " AND MiddleName IS NULL "
	Else
		sql = sql & " AND MiddleName=" & prepStringSQL(middlename)
	End If
	sql = sql & " AND lastname=" & prepStringSQL(lastname)
	If Len(suffix) = 0 Then
		sql = sql & " AND Suffix IS NULL "
	Else
		sql = sql & " AND suffix=" & prepStringSQL(suffix)
	End If
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = False Then 
		Response.Write("A user with this name address already exists in system. Please login to create application if you are that person. OTherwise, use Back and add a middle initial or other change to make the name unique within the system.")
		Response.End
	End If

	sql = "INSERT INTO System.Users (UserID, FirstName, MiddleName, LastName, Suffix, Title, email, Address1, Address2, City, State, ZIP, Phone, Fax, UpdateID, UpdateTimestamp) " & vbCrLF & _
		"VALUES (" & prepStringSQL(Email) & _
		", " & prepStringSQL(FirstName) & _
		", " & prepStringSQL(MiddleName) & _
		", " & prepStringSQL(LastName) & _
		", " & prepStringSQL(Suffix) & _
		", " & prepStringSQL(Title) & _
		", " & prepStringSQL(email) & _
		", " & prepStringSQL(Address1) & _
		", " & prepStringSQL(Address2) & _
		", " & prepStringSQL(City) & _
		", " & prepStringSQL(State) & _
		", " & prepStringSQL(ZIP) & _
		", " & prepStringSQL(Phone) & _
		", " & prepStringSQL(Fax) & _
		", -1" & _
		", getdate()) "

	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
Else
	Response.Write("Error: Step 2 with no data submitted")
	Response.End

End If

sql = "SELECT SystemID, UserID, Name, Email " & vbCrLf & _
	"FROM System.Users " & vbCrLF & _
	"WHERE email=" & prepStringSQL(Email)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
If rs.EOF = False Then
	SystemID = rs.Fields("SystemID")
	UserID = rs.Fields("UserID")
	Name = rs.Fields("Name")
	If Debug = True Then
		Response.Write("<pre>SystemID=" & SystemID & ", UserID=" & UserID & ", Name=" & Name & "</pre>" & vbCrLF)
		Response.Flush
	End If
Else
	Response.Write("Error: Unable to retrive new user.")
	Response.End
End If
resetPassword(UserID)

%><!DOCTYPE html>
<html>
<head>
<title>Step 2: Select or create your organizaton within the system</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function validateForm()
	{
		return true;
	}
</script>
</head>
<body>
<div class="header"></div>

<div class="pagetag">Step 2: Select or create your organizaton within the system.</div>


<div class="content">

<form name="Step2" id="Step2" method="post" action="Step3.asp" onSubmit="return validateForm();">
	<fieldset style="width: 540px;">
		<legend>User Login Information</legend>
		<p>Your username is "<%=UserID%>".</p>
		<p>An email has been sent with your password to <%=email%>.</p>
	</fieldset>

	<fieldset style="width: 540px;">
		<legend>Organization / Agency Information</legend>
		<label>Organization Type: <select name="OrganizationTypeID" style="display: inline; ">
			<option value="0">Select Organization Type</option>

<%
	sql = "SELECT OrganizationTypeID, OrganizationType FROM Lookup.OrganizationType ORDER BY 1 "

	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & "<option value=""" & rs.Fields("OrganizationTypeID") & """>" & rs.Fields("OrganizationType") & "</option>" & vbCrLf)
		rs.MoveNext
	Wend
%> </select></label>
	</fieldset>

	<fieldset style="width: 540px;">
		<legend>Organization / Agency Contact Information</legend>
		<label for="EMail">E-Mail Address:</label>
		<input type="text" name="email" id="email" value="<%=email %>" size="50" maxlength="50" /><br />
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
		<input type="submit" value="Submit" />
	</fieldset>
</form>

</div>
<div class="clearfix"></div>
<div class="footer">TxABTPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/ResetPasswordInclude.asp"-->
<!--#include file="../includes/Mail.asp"-->