<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j
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
	Response.Write("</pre>" & vbCrLf)
	Response.Flush
End If

Dim ContactPhoneCallID, CallDateTime, CallLength, PhoneNumber, ContactID, ContactName, ContactTitleID, _
	GranteeID, Organization, Questions, MVCPAContactID, Answer, EMail, ContactIssueID, ContactTypeID, _
	Positive, Negative, DateComplete, Title, Address1, Address2, City, State, ZIP, _
	PermitEdit, Reload, ReadOnlyString, ReadOnlyButtonString, LongText


If Len(UserSystemID)=0 Then
	Response.Write("You must be logged in to use this page.<br>")
	Response.Write("<a href=""../Default.asp"">Login page</a>")
	Response.End
End If

PermitEdit = True
If PermitEdit=True Then
	ReadOnlyString = ""
	ReadOnlyButtonString = ""
Else
	ReadOnlyString = " readOnly=true tabIndex=-1 "
	ReadOnlyButtonString = " Disabled=True tabIndex=-1 "
End If

If Debug = True Then
	Response.Write("<pre>")
	Response.Write("QueryString Count=" & Request.QueryString.Count & vbCrLf)
	Response.Write("Form Count=" & Request.Form.Count & vbCrLf)
	Response.Write("ContactID IsEmpty=" & IsEmpty(Request.QueryString("ContactID")) & vbCrLf)
	Response.Write("<pre>")
End If

' Determine if this is a request to load an existing record, reload with a new contactID or simply a new record.
If Request.QueryString.Count > 0 And Request.Form.Count > 0 Then 
	If Request.QueryString("Reload") = "Y" Then
		Reload = "Y"
		ContactPhoneCallID = CInt(Request.Form("ContactPhoneCallID"))
		CallDateTime = Request.Form("CallDateTime")
		CallLength = Request.Form("CallLength")
		PhoneNumber = Request.Form("PhoneNumber")
		ContactID = Request.Form("ContactID")
		If ContactID="" Then
			ContactID=0
		End If
		ContactName = Request.Form("ContactName")
		ContactTitleID = Request.Form("ContactTitleID")
		GranteeID = Request.Form("GranteeID")
		Organization = Request.Form("Organization")
		Questions = Request.Form("Questions")
		MVCPAContactID = Request.Form("MVCPAContactID")
		Answer = Request.Form("Answer")
		EMail = Request.Form("EMail")
		ContactIssueID = Request.Form("ContactIssueID")
		ContactTypeID = Request.Form("ContactTypeID")
		If Request.Form("Positive") = 1 then
			Positive = True
		Else
			Positive = False
		End If
		If Request.Form("Negative") = 1 then
			Negative = True
		Else
			Negative = False
		End If
		DateComplete = Request.Form("DateComplete")
		LongText = Request.Form("LongText")
		Title = Request.Form("Title")
		Address1 = Request.Form("Address1")
		Address2 = Request.Form("Address2")
		City = Request.Form("City")
		State = Request.Form("State")
		ZIP = Request.Form("ZIP_Code")
		If ContactID > 0 Then
			sql = "SELECT U.Name As ContactName, U.Phone As PhoneNumber, U.Title, U.EMail, " & vbCrLf & _
				"	U.Address1, U.Address2, U.City, U.State, U.ZIP, " & vbCrLf & _
				"	CASE WHEN AO.GranteeID>0 THEN 3 " & vbCrLf & _
				"		WHEN FO.GranteeID>0 THEN 4 " & vbCrLf & _
				"		WHEN PD.GranteeID>0 THEN 5 " & vbCrLf & _
				"		WHEN PM.GranteeID>0 THEN 6 " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN 7 " & vbCrLf & _
				"		WHEN PA.GranteeID>0 THEN 8 " & vbCrLf & _
				"	ELSE NULL END AS ContactTitleID, " & vbCrLf & _
				"	CASE WHEN AO.GranteeID>0 THEN AO.GranteeID " & vbCrLf & _
				"		WHEN FO.GranteeID>0 THEN FO.GranteeID " & vbCrLf & _
				"		WHEN PD.GranteeID>0 THEN PD.GranteeID " & vbCrLf & _
				"		WHEN PM.GranteeID>0 THEN PM.GranteeID " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN FA.GranteeID " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN PA.GranteeID " & vbCrLf & _
				"	ELSE 0 END AS GranteeID " & vbCrLf & _
				"FROM System.Users AS U " & vbCrLf & _
				"LEFT JOIN Grantees AS AO ON AO.AuthorizedOfficialID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS FO ON FO.FinancialOfficerID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS PD ON PD.ProgramDirectorID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS PM ON PM.ProgramManagerID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS FA ON FA.FinancialAdministrativeContactID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS PA ON PA.ProgramAdministrativeContactID=U.SystemID " & vbCrLf & _
				"WHERE SystemID=" & ContactID
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = Con.Execute(sql)
			If rs.EOF = True Then
				Response.Write("Error Retreiving System_User Record For ContactID=" & ContactID)
			Else
				ContactName = prepStringWeb(rs.Fields("ContactName"))
				PhoneNumber = prepStringWeb(rs.Fields("PhoneNumber"))
				EMail = prepStringWeb(rs.Fields("Email"))
				Title = prepStringWeb(rs.Fields("Title"))
				Address1 = prepStringWeb(rs.Fields("Address1"))
				Address2 = prepStringWeb(rs.Fields("Address2"))
				City = prepStringWeb(rs.Fields("City"))
				State = prepStringWeb(rs.Fields("State"))
				ZIP = prepStringWeb(rs.Fields("ZIP"))
				ContactTitleID = rs.Fields("ContactTitleID")
				GranteeID = rs.Fields("GranteeID")
			End If
		End If
	ElseIf IsEmpty(Request.QueryString("ContactID")) = False Then
		ContactID = Request.QueryString("ContactID")
		ContactPhoneCallID = 0
		CallDateTime = date()
		CallLength = ""
		PhoneNumber = ""
		ContactName = ""
		ContactTitleID = 0
		Organization = ""
		Questions = ""
		MVCPAContactID = UserSystemID
		Answer = ""
		EMail = ""
		ContactIssueID = 0
		ContactTypeID = 1
		Positive = False
		Negative = False
		DateComplete = ""
		LongText = ""
		Title = ""
		Address1 = ""
		Address2 = ""
		City = ""
		State = ""
		ZIP = ""
		If ContactID > 0 Then
			sql = "SELECT U.Name As ContactName, U.Phone As PhoneNumber, U.Title, U.EMail, " & vbCrLf & _
				"	U.Address1, U.Address2, U.City, U.State, U.ZIP, " & vbCrLf & _
				"	CASE WHEN AO.GranteeID>0 THEN 3 " & vbCrLf & _
				"		WHEN FO.GranteeID>0 THEN 4 " & vbCrLf & _
				"		WHEN PD.GranteeID>0 THEN 5 " & vbCrLf & _
				"		WHEN PM.GranteeID>0 THEN 6 " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN 7 " & vbCrLf & _
				"		WHEN PA.GranteeID>0 THEN 8 " & vbCrLf & _
				"	ELSE NULL END AS ContactTitleID, " & vbCrLf & _
				"	CASE WHEN AO.GranteeID>0 THEN AO.GranteeID " & vbCrLf & _
				"		WHEN FO.GranteeID>0 THEN FO.GranteeID " & vbCrLf & _
				"		WHEN PD.GranteeID>0 THEN PD.GranteeID " & vbCrLf & _
				"		WHEN PM.GranteeID>0 THEN PM.GranteeID " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN FA.GranteeID " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN PA.GranteeID " & vbCrLf & _
				"	ELSE 0 END AS GranteeID, " & vbCrLf & _
				"	CASE WHEN AO.GranteeID>0 THEN AO.GranteeName " & vbCrLf & _
				"		WHEN FO.GranteeID>0 THEN FO.GranteeName " & vbCrLf & _
				"		WHEN PD.GranteeID>0 THEN PD.GranteeName " & vbCrLf & _
				"		WHEN PM.GranteeID>0 THEN PM.GranteeName " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN FA.GranteeName " & vbCrLf & _
				"		WHEN FA.GranteeID>0 THEN PA.GranteeName " & vbCrLf & _
				"	ELSE '' END AS GranteeName " & vbCrLf & _
				"FROM System.Users AS U " & vbCrLf & _
				"LEFT JOIN Grantees AS AO ON AO.AuthorizedOfficialID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS FO ON FO.FinancialOfficerID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS PD ON PD.ProgramDirectorID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS PM ON PM.ProgramManagerID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS FA ON FA.FinancialAdministrativeContactID=U.SystemID " & vbCrLf & _
				"LEFT JOIN Grantees AS PA ON PA.ProgramAdministrativeContactID=U.SystemID " & vbCrLf & _
				"WHERE SystemID=" & ContactID
			If Debug = True Then
				Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = Con.Execute(sql)
			If rs.EOF = True Then
				Response.Write("Error Retreiving System_User Record For ContactID=" & ContactID)
			Else
				ContactName = prepStringWeb(rs.Fields("ContactName"))
				PhoneNumber = prepStringWeb(rs.Fields("PhoneNumber"))
				EMail = prepStringWeb(rs.Fields("Email"))
				Title = prepStringWeb(rs.Fields("Title"))
				Address1 = prepStringWeb(rs.Fields("Address1"))
				Address2 = prepStringWeb(rs.Fields("Address2"))
				City = prepStringWeb(rs.Fields("City"))
				State = prepStringWeb(rs.Fields("State"))
				ZIP = prepStringWeb(rs.Fields("ZIP"))
				ContactTitleID = rs.Fields("ContactTitleID")
				GranteeID = rs.Fields("GranteeID")
				Organization = rs.Fields("GranteeName")
			End If
		End If
		Reload = "Y" ' Set reload = "Y"
	End If
ElseIf Len(Request.QueryString("ContactPhoneCallID")) > 0 Then
	If Debug = True Then
		Response.Write("<pre>Not a reload</pre>" & vbCrLf)
	End If
	Reload = "N"
	ContactPhoneCallID = Request.QueryString("ContactPhoneCallID")
	sql = "SELECT PC.*, SU.Title, SU.Address1, SU.Address2, SU.City, SU.State, SU.ZIP " & _
			"FROM Contact.PhoneCalls AS PC " & _
			"LEFT JOIN System.Users AS SU ON PC.ContactID = SU.SystemID " & _
			"WHERE ContactPhoneCallID=" & ContactPhoneCallID
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error Retreiving Phone Call Record For ContactPhoneCallID=" & ContactPhoneCallID)
	Else
		CallDateTime = rs.Fields("CallDateTime")
		CallLength = rs.Fields("CallLength")
		PhoneNumber = rs.Fields("PhoneNumber")
		ContactID = rs.Fields("ContactID")
		If IsNull(ContactID)="" Then
			ContactID=0
		End If
		ContactName = rs.Fields("ContactName")
		ContactTitleID = rs.Fields("ContactTitleID")
		GranteeID = rs.Fields("GranteeID")
		Organization = rs.Fields("Organization")
		Questions = prepStringWeb(rs.Fields("Questions"))
		MVCPAContactID = rs.Fields("MVCPAContactID")
		Answer = prepStringWeb(rs.Fields("Answer"))
		EMail = prepStringWeb(rs.Fields("EMail"))
		ContactIssueID = rs.Fields("ContactIssueID")
		ContactTypeID = rs.Fields("ContactTypeID")
		Positive = rs.Fields("Positive")
		Negative = rs.Fields("Negative")
		DateComplete = prepStringWeb(rs.Fields("DateComplete"))
		LongText = rs.Fields("LongText")
		Title = prepStringWeb(rs.Fields("Title"))
		Address1 = prepStringWeb(rs.Fields("Address1"))
		Address2 = prepStringWeb(rs.Fields("Address2"))
		City = prepStringWeb(rs.Fields("City"))
		State = prepStringWeb(rs.Fields("State"))
		ZIP = prepStringWeb(rs.Fields("ZIP"))
	End If
Else
	ContactPhoneCallID = 0
	CallDateTime = Date()
	CallLength = ""
	PhoneNumber = ""
	ContactID = 0
	ContactName = ""
	ContactTitleID = 0
	GranteeID = 0
	Organization = ""
	Questions = ""
	MVCPAContactID = UserSystemID
	Answer = ""
	EMail = ""
	ContactIssueID = 0
	ContactTypeID = 1
	DateComplete = ""
End If
If IsNull(ContactID) = True Then ContactID = 0
If IsNull(MVCPAContactID) = True Then MVCPAContactID = 0
If IsNull(ContactIssueID) = True Then ContactIssueID = 0
If ISNULL(ContactTypeID) = True Then ContactTypeID = 0
If IsNull(ContactTitleID) = True Then ContactTitleID = 0

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Contact Datase: Add a new contact item</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
	<script language="JavaScript">
		function submitPage()
		{
			if (document.ContactItems.MVCPAContactID.selectedIndex == 0) {
				alert("You must select a MVCPA Contact before submitting");
				return false;
			}
			document.ContactItems.submit();
			return true;
		}

		function ValidateDate(field)
		{
			if (validDate(field) == false) {
				return false;
			}
			return true;
		}

		function ContactIDChange()
		{
			document.ContactItems.action = "ContactItems.asp?Reload=Y";
			submitPage();
		}

		function CheckTextarea(field)
		{
			if (field.value.length > 2000) {
				alert("The maximum length that can be saved for this field is 2000 characters. There are currently " +
					field.value.length + " characters. Please reduce the size to 2000 characters. " +
					"If you have something very long, paste it into the Documents field which allows for much more text.");
				return false;
			}
			return true;
		}
	</script>
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Contact Database: Add a new contact item</div>


<div class="widecontent">

<form name="ContactItems" method="post" action="ContactItemsSubmit.asp" onsubmit="return submitPage();">
<input type="hidden" name="ContactPhoneCallID" value="<%=ContactPhoneCallID%>">
<input type="hidden" name="SystemID" value="<%=UserSystemID%>">
<table style="margin: auto;">

<tr>
	<th colspan="2">Add Contact List Item</th>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>

<%
If Cint(ContactPhoneCallID) > 0 Then
	Response.Write("<tr><td>ID</td><td>" & ContactPhoneCallID & "</td></tr" & vbCrLf)
End If
%>
<tr>
	<td>Date of Call:</td>
	<td><input type="text" name="CallDateTime" size="8" maxLength="10" value="<%=CallDateTime%>" style="text-align: right" onChange="return ValidateDate(this);">
</tr>

<tr>
	<td>Length of Call (in minutes):</td>
	<td><input type="text" name="CallLength" size="4" maxLength="4" value="<%=CallLength%>" style="text-align: right">
</tr>

<tr>
	<td>Contact (if already in system):</td>
	<td><select name="ContactID" onChange="ContactIDChange()">
			<option value="0">Select the contact's name</option>
<%
	sql = "SELECT '<option value=""' + CAST(SystemID AS VARCHAR) + '""' + CASE WHEN SystemID=" & ContactID & " THEN ' selected' ELSE '' END + '>' + Name + '</option>' As OptionItem " & vbCrLf & _
		"FROM System.Users " & vbCrLf & _
		"ORDER BY LastName, FirstName "
	If Debug = True Then 
		Response.Write("<!--" & sql & "-->")
	End If
	Set rs = Con.Execute(sql) 
	While rs.EOF = False
		Response.Write(vbTab & vbTab & vbTab & rs.Fields("OptionItem") & vbCrLf)
		rs.MoveNext
	Wend
%>
		</select>
</tr>

<%
If Len(Title) > 0 Then
	Response.Write("<tr><td>Title from System User record</td><td>" & Title & "</td></tr>" & vbCrLf)
End If

If IsNumeric(ContactID) = True Then
	If CInt(ContactID) > 0 Then
		Response.Write("<tr><td>Address from System User record</td><td>")
		If Len(Address1) > 0 Then Response.Write(Address1 & "; ")
		If Len(Address2) > 0 Then Response.Write(Address2 & "; ")
		If Len(City) > 0 Then Response.Write(City & ", " & State & " " & Zip)
		Response.Write("</td></tr>" & vbCrLf)
	End If
End If
%>
<tr>
	<td>Name of Contact (if not in system):</td>
	<td><input type="text" name="ContactName" size="30" maxLength="50" value="<%=ContactName%>" style="text-align: left">
</tr>

<tr>
	<td>Phone Number of Caller:</td>
	<td><input type="text" name="PhoneNumber" size="18" maxLength="40" value="<%=PhoneNumber%>" style="text-align: left">
</tr>

<tr>
	<td>E-mail address of Caller:</td>
	<td><input type="text" name="EMail" size="50" maxLength="64" value="<%=EMail%>" style="text-align: left">
</tr>

<tr>
	<td>Contact's Title or Role:</td>
	<td><select name="ContactTitleID">
			<option value="0">Select Contact's Title</option>
<%
	sql = "SELECT '<option value=""' + CAST(TitleID AS VARCHAR) + '""' + CASE WHEN TitleID=" & ContactTitleID & " THEN ' selected' ELSE '' END + '>' + Title + '</option>' As OptionItem " & vbCrLf & _
		"FROM Lookup.Titles " & vbCrLf & _
		"ORDER BY TitleSort "
	If Debug = True Then 
		Response.Write("<!--" & sql & "-->")
	End If
	Set rs = Con.Execute(sql) 
	While rs.EOF = False
		Response.Write(vbTab & vbTab & vbTab & rs.Fields("OptionItem") & vbCrLf)
		rs.MoveNext
	Wend
%>
		</select>
</tr>

<tr>
	<td>Grantee (for existing grantees)</td>
	<td><select name="GranteeID">
		<option value="0">Select Grantee</option>
<%
	sql = "SELECT '<option value=""' + CAST(GranteeID AS VARCHAR) + '""' + CASE WHEN GranteeID=" & GranteeID & " THEN ' selected' ELSE '' END + '>' + GranteeName + '</option>' As OptionItem " & vbCrLf & _
	"FROM Grantees ORDER BY GranteeName "
	If Debug = True Then 
		Response.Write("<!--" & sql & "-->")
	End If
	Set rs = Con.Execute(sql) 
	While rs.EOF = False
		Response.Write(vbTab & vbTab & vbTab & rs.Fields("OptionItem") & vbCrLf)
		rs.MoveNext
	Wend
%>
	</select></td>
</tr>

<tr>
	<td>Organization (Grantee, Company or agency)<br>
		 of caller for grantees not in system:</td>
	<td><input type="text" name="Organization" size="50" maxLength="64" value="<%=Organization%>" style="text-align: left">
</tr>

<tr>
	<td colspan="2">Narrative of Call / Questions raised (limited to 2000 characters):</td>
</tr>
<tr>
	<td colspan="2"><textarea name="Questions" Rows="5" Cols="100" onChange="return CheckTextarea(this)"><%=Questions%></textarea>
</tr>

<tr>
	<td>MVCPA Contact</td>
	<td><select name="MVCPAContactID">
			<option value="0">Select MVCPA Contact</option>
<%
	sql = "SELECT '<option value=""' + CAST(SystemID AS VARCHAR) + '""' + CASE WHEN SystemID=" & MVCPAContactID & " THEN ' selected' ELSE '' END + '>' + Name + '</option>' As OptionItem " & vbCrLf & _
	"FROM System.Users " & vbCrLf & _
	"WHERE MVCPAStaff=1 OR SystemID=" & prepIntegerSQL(MVCPAContactID) & " " & vbCrLf
	If UserSystemID=1 Then 
		sql = sql & " Or SystemID=1 " & vbCrLf
	End If
	sql = sql & "ORDER BY LastName, FirstName "
	If Debug = True Then 
		Response.Write("<!--" & sql & "-->")
	End If
	Set rs = Con.Execute(sql) 
	While rs.EOF = False
		Response.Write(vbTab & vbTab & vbTab & rs.Fields("OptionItem") & vbCrLf)
		rs.MoveNext
	Wend
%>
		</select>
</tr>

<tr>
	<td colspan="2">Answers / Response To Call (limited to 2000 characters):</td>
</tr>
<tr>
	<td colspan="2"><textarea name="Answer" Rows="5" Cols="100"><%=Answer%></textarea>
</tr>

<tr>
	<td>Issue:</td>
	<td><select name="ContactIssueID">
		<option value="0">Select Issue</option>
<%
	sql = "SELECT '<option value=""' + CAST(IssueID AS VARCHAR) + '""' + CASE WHEN IssueID=" & ContactIssueID & " THEN ' selected' ELSE '' END + '>' + IssueDescription + '</option>' As OptionItem " & vbCrLf & _
	"FROM Contact.Issues " & vbCrLf & _
	"ORDER BY IssueSort "
	If Debug = True Then 
		Response.Write("<!--" & sql & "-->")
	End If
	Set rs = Con.Execute(sql) 
	While rs.EOF = False
		Response.Write(vbTab & vbTab & vbTab & rs.Fields("OptionItem") & vbCrLf)
		rs.MoveNext
	Wend
%>
		</select></td>
</tr>

<tr>
	<td>Contact Type</td>
	<td><select name="ContactTypeID">
		<option value="0">Select Issue</option>
<%
	sql = "SELECT '<option value=""' + CAST(ContactTypeID AS VARCHAR) + '""' + CASE WHEN ContactTypeID=" & ContactTypeID & " THEN ' selected' ELSE '' END + '>' + ContactType + '</option>' As OptionItem " & vbCrLf & _
	"FROM Contact.Types " & vbCrLf & _
	"ORDER BY ContactTypeSort "
	If Debug = True Then 
		Response.Write("<!--" & sql & "-->")
	End If
	Set rs = Con.Execute(sql) 
	While rs.EOF = False
		Response.Write(vbTab & vbTab & vbTab & rs.Fields("OptionItem") & vbCrLf)
		rs.MoveNext
	Wend
%>
		</select>
</tr>

<tr>
	<td colspan=2 align=center>
		<input type="checkbox" name="Positive" value=1 <%=Checked(Positive,True)%> onClick="document.ContactItems.DocumentUpdated.value='Y';">Positive / Compliment
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="checkbox" name="Negative" value=1 <%=Checked(Negative,True)%> onClick="document.ContactItems.DocumentUpdated.value='Y';">Negative / Complaint
</tr>

<tr>
	<td>Date Complete / Issue Resolved:</td>
	<td><img src="../images/Date.gif" Title="Insert Today's Date" alt="Calendar icon" onClick="document.ContactItems.DateComplete.value='<%=Date()%>';" 
		alt="Calendar icon" title="Click to enter today's date for date complete" width="24" height="24">
		<input type="text" name="DateComplete" size="8" maxLength="10" value="<%=DateComplete%>" onChange="return ValidateDate(this);"></td>
</tr>

<tr>
	<td colspan="2">Longer items such as the text of email or other supporting documents may be posted here:</td>
</tr>
<tr>
	<td colspan="2"><textarea name="LongText" Rows="5" Cols="100" 
	onChange="document.ContactItems.DocumentUpdated.value='Y';"><%=LongText%></textarea></td>
</tr>

<tr>
	<td colspan="2" align="center"><a href="" ><input type="button" name="Submit"  value="Save"
		alt="Save Button" onclick="submitPage(); return false;"></a>&nbsp;&nbsp;
		<%	If Reload="Y" and ContactPhoneCallID > 0 Then %>
		<input type="button" value="Reset" onclick="location.href='ContactItems.asp?ContactPhoneCallID=<%=ContactPhoneCallID%>'" 
			alt="Reset Button" title="This will return fields to thier initial values">&nbsp;&nbsp;
		<%	ElseIf Reload="Y" Then %>
		<input type="button" value="Reset" onclick="location.href='ContactItems.asp'" alt="Reset Button"
			title="This will return fields to thier initial values">&nbsp;&nbsp;
		<%	Else %>
		<input type="button" value="Reset" onClick="document.ContactItems.reset()"  alt="Reset Button" 
			title="This will return fields to thier initial values">&nbsp;&nbsp;
		<%	End If %>
		<input type="button" value="Cancel" onclick="window.close();" alt="Cancel Button"
			title="This will cancel any updates and Close this window">&nbsp;&nbsp;
		<input type="button" value="Search" onclick="location.href='ContactItemsSearch.asp';" alt="Search Button"
			title="This will cancel any updates and take you to the Contact Search Page"></td>
</tr>


</table>
</form>

</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<%
Function Selected(vVariable,vValue)
	If vVariable = vValue Then
		Selected = " SELECTED "
	Else
		Selected = ""
	End If
End Function

Function Checked(vVariable,vValue)
	If vVariable = vValue Then
		Checked = " CHECKED "
	Else
		Checked = ""
	End If
End Function
%>
