<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, LastName
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

If Len(Request.Form("LastName"))>0 Then
	LastName = Request.Form("LastName")
ElseIf Len(Request.QueryString("LastName"))>0 Then
	LastName = Request.QueryString("LastName")
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Update User</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" />
<script type="text/javascript">
	function validateForm()
	{
		var buttons = document.UpdateUser.UpdateSystemID.length;
		if(document.UpdateUser.UpdateSystemID[buttons-1].checked)
		{
			if (confirm("Have you verified that this person is not currently in the system?\n\nOK to Continue. Cancel to check list."))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		for (var i = 0; i < buttons - 1; i++)
			if (document.UpdateUser.UpdateSystemID[i].checked)
				return true;
		alert("You must select one of the options to continue.")
		return false;
	}
</script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>
<div class="pagetag">Add or Update A User: Search for "<%=LastName %>"</div>
<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>
<div class="content">

<form name="UpdateUser" id="UpdateUser" method="post" action="UpdateUser3.asp" onsubmit="return validateForm();">
<input type="hidden" name="LastName" value="<%=LastName %>" />
<fieldset style="width: 600px;">
		<legend>Step 2: Select user to update</legend><%
sql = "SELECT SystemID, Name, AccountDisabled FROM System.Users WHERE Name LIKE '%" & replace(LastName,"'","''") & "%' ORDER BY LastName, FirstName"
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
i=0
While rs.EOF = False
	'Response.Write("<input type=""radio"" name=""UpdateSystemID"" id=""UpdateSystemID" & i & _
	'""" value=""" & rs.Fields("SystemID") & """><label for=""UpdateSystemID" & i & """ class=""radio"">" & rs.Fields("Name") & " (" & rs.Fields("SystemID")& ")</label>")
	If rs.Fields("AccountDisabled") = True Then
		Response.Write("<label style=""width: 480px; text-align: left;""><input type=""radio"" name=""UpdateSystemID""" & _
		" value=""" & rs.Fields("SystemID") & """>" & rs.Fields("Name") & " (" & rs.Fields("SystemID")& ") (Account Disabled)</label><br />" & vbCrLf)
	Else
		Response.Write("<label style=""width: 480px; text-align: left;""><input type=""radio"" name=""UpdateSystemID""" & _
		" value=""" & rs.Fields("SystemID") & """>" & rs.Fields("Name") & " (" & rs.Fields("SystemID")& ")</label><br />" & vbCrLf)
	End If
	i = i + 1
	rs.MoveNext
Wend

%><label style="width: 600px; text-align: left;"><input type="radio" name="UpdateSystemID" value="0" /> 
User is not listed. This creates a new user.</label><br /> 
</fieldset>
<br />
<p style="text-align: left;">This is a very important step to avoid creating duplicate users. Be sure that the user you 
wish to create is not already in the system and listed above.</p>

<p style="text-align: left;">Also, do not attempt to change the identity of an existing user. 
If the agency or organization has a new person filling an existing position, a new user should 
be created for that person rather than changing the name of an existing user.</p>

<div style="width: 738px; text-align: center;">
<input type="submit" value="Submit" title="Submit last name for search" />
<input type="button" value="Back" title="Go back to previous page and change search." onclick="location.href = 'UpdateUSer1.asp?LastName=<%=Lastname%>';"" />
<input type="button" value="Cancel" title="Return to Home" onclick="location.href = '../Home/Default.asp';" />

</div>

</form>
</div>
<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepDB.asp"-->