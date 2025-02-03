<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, SearchLastName, fieldname, Position, CurrentID, CurrentName, GranteeID, GranteeName, ReturnPage
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

If Len(Request.Form("GranteeID"))>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else 
	GranteeID = 0
End If
If Request.Form.Count>0 Then
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

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Change <%=Position %> for <%=GranteeName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	function validateForm()
	{
		var buttons = document.ChangeOffical2.UpdateSystemID.length;
		if (document.ChangeOffical2.UpdateSystemID[buttons - 1].checked) {
			if (confirm("Have you verified that this person is not currently in the system?\n\nOK to Continue. Cancel to check list.")) {
				return true;
			}
			else {
				return false;
			}
		}
		for (var i = 0; i < buttons - 1; i++)
			if (document.ChangeOffical2.UpdateSystemID[i].checked)
				return true;
		alert("You must select one of the options to continue.")
		return false;
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

<form name="ChangeOffical2" id="ChangeOfficial2" method="post" action="ChangeOfficial3.asp" onsubmit="return validateForm();">
<input type="hidden" name="SearchLastName" value="<%=SearchLastName %>" />
<input type="hidden" name="ReturnPage" value="<%=ReturnPage %>" />
<input type="hidden" name="GranteeID" value="<%=GranteeID %>" />
<input type="hidden" name="Position" value="<%=Position %>" />
<input type="hidden" name="CurrentID" value="<%=CurrentID %>" />
<fieldset style="width: 600px;">
		<legend>Step 2: Select user to fill <%=Position %> for <%=GranteeName %></legend>
<%

sql = "SELECT SystemID, Name FROM System.Users WHERE Name LIKE '%" & Replace(SearchLastName,"'","''") & "%' ORDER BY LastName, FirstName"
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
i=0
While rs.EOF = False
	'Response.Write("<input type=""radio"" name=""UpdateSystemID"" id=""UpdateSystemID" & i & _
	'""" value=""" & rs.Fields("SystemID") & """><label for=""UpdateSystemID" & i & """ class=""radio"">" & rs.Fields("Name") & " (" & rs.Fields("SystemID")& ")</label>")
	Response.Write("<label style=""width: 480px; text-align: left;""><input type=""radio"" name=""UpdateSystemID""" & _
	" value=""" & rs.Fields("SystemID") & """>" & rs.Fields("Name") & " (" & rs.Fields("SystemID")& ")</label><br />" & vbCrLf)
	i = i + 1
	rs.MoveNext
Wend

%><label style="width: 580px; text-align: left;"><input type="radio" name="UpdateSystemID" value="0" /> 
User is not listed. This creates a new user.</label><br /> 
</fieldset>

<p style="text-align: left;">This is a very important step to avoid creating duplicate users. Be sure that the user you 
wish to create is not already in the system and listed above.</p>

<p style="text-align: left;">Also, do not attempt to change the identity of an existing user. 
If the agency or organization has a new person filling an existing position, a new user should 
be created for that person rather than changing the name of an existing user.</p>

<div style="width: 600px; text-align: center;"><input type="submit" value="Submit" title="Submit last name for search" style="text-align: center" /></div>

</form>
</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->