<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, Position, ReturnPage, GranteeID, fieldname, CurrentID, CurrentName, GranteeName, SearchLastName
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
	Position = Request.Form("Position")
	ReturnPage = Request.Form("ReturnPage")
	GranteeID = Request.Form("GranteeID")
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

sql = "SELECT ISNULL(U.Name,'No person assigned') AS CurrentName, G.GranteeName " & vbCrLF & _
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
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Change <%=Position %> for <%=GranteeName %></title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<link rel="stylesheet" href="/styles/fieldset.css" type="text/css" /> 
<script type="text/javascript">
	function validateForm()
	{
		if (document.ChangeOfficial.SearchLastName.value.length>0)
		{
			return true;
		}
		else
		{
			alert("You must enter some text to use for searching existing users.");
			document.ChangeOfficial.SearchLastName.focus();
			return false;
		}
	}
</script>
</head>
<body onload="document.ChangeOfficial.SearchLastName.focus();">
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Step 1: Determine if new person for the position currently 
	exists in the system. Begin by searching on last name.

	<p>The position is currently assigned to <%=CurrentName %>.</p>
</div>

<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">
<form name="ChangeOfficial" method="post" action="ChangeOfficial2.asp" onsubmit="return validateForm();">
<input type="hidden" name="ReturnPage" value="<%=ReturnPage %>" />
<input type="hidden" name="GranteeID" value="<%=GranteeID %>" />
<input type="hidden" name="Position" value="<%=Position %>" />
<input type="hidden" name="CurrentID" value="<%=CurrentID %>" />
<div class="sectiontitle">Change <%=Position %> for <%=GranteeName %></div>

<p>Step 1: Determine if new person for the position currently 
	exists in the system. Begin by searching on last name.</p>
	<fieldset style="width: 640px;">
		<legend>Step 1: Search for User</legend>
		<label for="SearchLastName">Last Name:</label>
		<input type="text" name="SearchLastName" id="SearchLastName" value="<%=SearchLastName %>" size="24" maxlength="24" /><br />
	</fieldset>
	<div style="text-align: center; width: 640px;"><input type="submit" value="Submit" title="Submit last name for search" /></div>
</form>

</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->