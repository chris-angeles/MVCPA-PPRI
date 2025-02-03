<%@ language=VBScript %>
<% Option Explicit %>
<!--#include file="../includes/EnsureLogin.asp"--> 
<!--#include file="../includes/adovbs.asp"--> 
<!--#include file="../includes/OpenConnection.asp"--> 
<%
Dim Debug, i, MVCPAContactID, GranteeID, ContactIssueID, FirstSearchItem, ShowColumns, _
	ShowContactInfo, ShowQandA, ShowDaysToComplete, StatusToShow, StartDate, EndDate, SortBy, _
	RecordCounter, CompleteCount, IncompleteCount, CompleteTotal
Debug = False

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

ShowColumns = 9
RecordCounter = 0
CompleteCount = 0
IncompleteCount = 0
FirstSearchItem = True

If Request.Form.count > 0 Then
	MVCPAContactID = Request.Form("MVCPAContactID")
	If Len(MVCPAContactID) > 0 Then 
		MVCPAContactID = CInt(MVCPAContactID)
	End If
	GranteeID = Request.Form("GranteeID")
	If Request.Form("ShowContactInfo") = "Y" Then
		ShowContactInfo = "Y"
		ShowColumns = ShowColumns + 1
	Else
		ShowContactInfo = "N"
	End If
	If Request.Form("ShowQandA") = "Y" Then
		ShowQandA = "Y"
		ShowColumns = ShowColumns + 2
	Else
		ShowQandA = "N"
	End If
	If Request.Form("ShowDaysToComplete") = "Y" Then
		ShowDaysToComplete = "Y"
		ShowColumns = ShowColumns + 1
	Else
		ShowDaysToComplete = "N"
	End If
	StatusToShow = Request.Form("StatusToShow")
	StartDate = Request.Form("StartDate")
	EndDate = Request.Form("EndDate")
	ContactIssueID = Request.Form("ContactIssueID")
	SortBy = Request.Form("SortBy")
Else
	MVCPAContactID = UserSystemID
	ShowContactInfo = "Y"
	ShowDaysToComplete = "N"
	StatusToShow = "Outstanding"
	StartDate= StartOfFiscalYear()
	EndDate = Date()
	ContactIssueID = -1
	SortBy = "2"
End If
If IsNull(GranteeID) = True Then GranteeID = 0
If GranteeID = "" Then GranteeID = 0
If IsNull(MVCPAContactID)=True Then
	MVCPAContactID = 0
ElseIf MVCPAContactID="" Then
	MVCPAContactID = 0
End If
If MVCPAContactID > 0 Then ShowColumns = ShowColumns - 1

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Contact Datase: Search contact items</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script language="JavaScript">
	function ValidateDate(field)
	{
		if (validDate(field)==false)
		{
			return false;
		}
		document.ContactItemsSearch.submit();
		return true;
	}

	function ShowAllItems()
	{
		document.ContactItemsSearch.MVCPAContactID.selectedIndex=0;
		document.ContactItemsSearch.GranteeID.selectedIndex=0;
		document.ContactItemsSearch.StartDate.value="9/1/2016";
		document.ContactItemsSearch.EndDate.value="<%=Date()%>";
		document.ContactItemsSearch.ContactIssueID.selectedIndex=0;
		document.ContactItemsSearch.StatusToShow.selectedIndex=0;
		document.ContactItemsSearch.SortBy.selectedIndex=1;
		document.ContactItemsSearch.submit();
	}
		
	function ShowDefault()
	{
		var i;
		for (i=0;i<document.ContactItemsSearch.MVCPAContactID.length;i++)
		{
			if (document.ContactItemsSearch.MVCPAContactID.options(i).value==<%=UserSystemID%>)
			{
				document.ContactItemsSearch.MVCPAContactID.selectedIndex = i;
			}
		}
		document.ContactItemsSearch.StartDate.value="9/1/2009";
		document.ContactItemsSearch.EndDate.value="<%=Date()%>";
		document.ContactItemsSearch.ContactIssueID.selectedIndex=0;
		document.ContactItemsSearch.StatusToShow.selectedIndex=2;
		document.ContactItemsSearch.SortBy.selectedIndex=1;
		document.ContactItemsSearch.submit();
	}
		
	function CompleteItem(field,vid)
	{
		if (field.checked==true)
		{
			window.open("CompleteItem.asp?ContactPhoneCallID="+vid,'hidden','height=100;width=100')
			window.focus();
			field.checked=true;
		}
		else
		{
			field.checked = true
		}
	}
</script>
<!--#include file="../includes/validDate.asp"--> 
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Contact Datase: Search contact items</div>

<div class="widecontent">

<form name="ContactItemsSearch" method="Post" Action="ContactItemsSearch.asp">
<table style="border: none; margin: auto;">

<tr>
	<th colspan="3">Contact Database Contact List Report</th>
</tr>

<tr><td colspan="3">&nbsp;</td></tr>

<tr>
	<td colspan="2"><input type="CheckBox" name="ShowContactInfo" value="Y" onClick="document.ContactItemsSearch.submit();" <%=CheckedValue(ShowContactInfo,"Y")%>>
		Show Contact Information</td>
	<td rowspan="8" valign="top">
		<input type=button value="Show All" onclick="ShowAllItems();" alt="Show All Button" 
			title="Show All Button"><br><br>
		<input type=button value="Default" onclick="ShowDefault();" alt="Default Button" 
			title="Show your outstanding items for the current county in reverse chronological order."><br><br>
		<input type=button value="Close" onclick="window.close()" alt="Close Button"
			title="Close Window"><br><br>
		<input type=button value="Add" onclick="location.href='../Contact/ContactItems.asp'" alt="Add Button" 
			title="Add a new contact item."></td>
</tr>

<tr>
	<td colspan="2"><input type="CheckBox" name="ShowQandA" value="Y" onClick="document.ContactItemsSearch.submit();" <%=CheckedValue(ShowQandA,"Y")%>>
		Show Question and Answer</td>
</tr>

<tr>
	<td colspan="2"><input type="CheckBox" name="ShowDaysToComplete" value="Y" onClick="document.ContactItemsSearch.submit();" <%=CheckedValue(ShowDaysToComplete,"Y")%>>
		Show number of days to complete</td>
</tr>

<tr>
	<td>MVCPA Contact</td>
	<td><%
	sql = "SELECT '<option value=""' + CAST(SystemID AS VARCHAR) + '""' + CASE WHEN SystemID=" & MVCPAContactID & " THEN ' selected' ELSE '' END + '>' + Name + '</option>' As OptionItem " & vbCrLF & _
	"FROM System.Users " & vbCrLF & _
	"WHERE MVCPAStaff=1 OR SystemID IN (SELECT DISTINCT MVCPAContactID FROM Contact.PhoneCalls) " & vbCrLf
	If UserSystemID=1 Then 
		sql = sql & " Or SystemID=1 " & vbCrLf
	End If
	sql = sql & "ORDER BY LastName, FirstName "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
 %><select name="MVCPAContactID" onChange="document.ContactItemsSearch.submit();">
		<option value="-1">All MVCPA Contacts</option>
<%
	Set rs = Con.Execute(sql) 
	While rs.EOF = False
		Response.Write(vbTab & vbTab & vbTab & rs.Fields("OptionItem") & vbCrLf)
		rs.MoveNext
	Wend
%>
		</select></td>
</tr>

<tr>
	<td>Grantee</td>
	<td><select name="GranteeID" onChange="document.ContactItemsSearch.submit();">
		<option value="-1">All Grantees</option>
<%
	sql = "SELECT '<option value=""' + CAST(GranteeID AS VARCHAR) + '""' + CASE WHEN GranteeID=" & GranteeID & " THEN ' selected' ELSE '' END + '>' + GranteeName + '</option>' As OptionItem " & vbCrLF & _
	"FROM Grantees ORDER BY GranteeName "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
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
	<td>Date Range (mm/dd/yyyy)</td>
	<td><input type="text" name="StartDate" value="<%=StartDate%>" size="10" maxLength="10" onChange="return ValidateDate(this)"> 
		through 
		<input type="text" name="EndDate" value="<%=EndDate%>" size="10" maxLength="10" onChange="return ValidateDate(this)">
	</td>
</tr>

<tr>
	<td>Issue to Show:</td>
	<td><select name="ContactIssueID" onChange="document.ContactItemsSearch.submit();">
		<option value="-1" <%=SelectedValue(ContactIssueID,"-1")%>>All Issues</option>
<%
	sql = "SELECT '<option value=""' + CAST(IssueID AS VARCHAR) + '""' + CASE WHEN IssueID=" & ContactIssueID & " THEN ' selected' ELSE '' END + '>' + IssueDescription + '</option>' As OptionItem " & vbCrLF & _
	"FROM Contact.Issues " & vbCrLF & _
	"ORDER BY IssueSort "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
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
	<td>Status to Show</td>
	<td><select name="StatusToShow" onChange="document.ContactItemsSearch.submit();">
			<option value="All" <%=SelectedValue(StatusToShow,"All")%>>All</option>
			<option value="Complete" <%=SelectedValue(StatusToShow,"Complete")%>>Complete</option>
			<option value="Outstanding" <%=SelectedValue(StatusToShow,"Outstanding")%>>Outstanding</option>
		</select></td>
</tr>

<tr>
	<td>Sort By</td>
	<td><select name="SortBy" onChange="document.ContactItemsSearch.submit();">
			<option value="1" <%=SelectedValue(SortBy,"1")%>>Call Date Ascending</option>
			<option value="2" <%=SelectedValue(SortBy,"2")%>>Call Date Descending</option>
		</select></td>
</tr>
</table>
</form>

<table border="0">
<%
sql = "SELECT A.*, B.IssueDescription AS Issue, C.ContactType, D.GranteeName, " & vbCrLf & _
	"	E.Name AS MVCPAContact, F.Title AS ContactTitle, " & vbCrLf & _
	"	ISNULL(SUBSTRING(E.FirstName,1,1),'') + ISNULL(SUBSTRING(E.MiddleName,1,1),'') + ISNULL(SUBSTRING(E.LastName,1,1),'') AS Initials, " & vbCrLf & _
	"	DateDiff(day, CallDateTime, DateComplete) AS DaysToComplete, " & vbCrLf & _
	"	CAST(CASE WHEN DateComplete IS NOT NULL THEN 1 ELSE 0 END AS BIT) AS Complete " & vbCrLf & _
	"FROM Contact.PhoneCalls AS A " & vbCrLf & _
	"LEFT JOIN Contact.Issues AS B ON B.IssueID=A.ContactIssueID " & vbCrLf & _
	"LEFT JOIN Contact.Types AS C ON C.ContactTypeID=A.ContactTypeID " & vbCrLf & _
	"LEFT JOIN Grantees AS D ON D.GranteeID=A.GranteeID " & vbCrLf & _
	"LEFT JOIN System.Users AS E ON E.SystemID=A.MVCPAContactID " & vbCrLf & _
	"LEFT JOIN Lookup.Titles AS F ON F.TitleID=A.ContactTitleID "

If MVCPAContactID > 0 Then
	If FirstSearchItem = True Then
		sql = sql + "WHERE "
	Else
		sql = SQL + " AND "
	End If
	sql = sql + "MVCPAContactID=" & MVCPAContactID
	FirstSearchItem = False
End If

If GranteeID > 0 Then
	If FirstSearchItem = True Then
		sql = sql + "WHERE "
	Else
		sql = SQL + " AND "
	End If
	If GranteeID > 0 Then
		sql = sql + "A.GranteeID=" & prepIntegerSQL(GranteeID)
	Else
		'
	End If
	FirstSearchItem = False
End If

If ContactIssueID > 0 Then
	If FirstSearchItem = True Then
		sql = sql + "WHERE "
	Else
		sql = SQL + " AND "
	End If
	sql = sql + " ContactIssueID=" & ContactIssueID
	FirstSearchItem = False
ElseIf ContactIssueID=0 Then
	If FirstSearchItem = True Then
		sql = sql + "WHERE "
	Else
		sql = SQL + " AND "
	End If
	sql = sql + "ContactIssueID IS NULL"
	FirstSearchItem = False
End If

If Len(StartDate) > 0 Then
	If FirstSearchItem = True Then
		sql = sql + "WHERE "
	Else
		sql = SQL + " AND "
	End If
	sql = sql + "CallDateTime>='" & StartDate & "'"
	FirstSearchItem = False
End If

If Len(EndDate) > 0 Then
	If FirstSearchItem = True Then
		sql = sql + "WHERE "
	Else
		sql = SQL + " AND "
	End If
	sql = sql + "CallDateTime < '" & DateAdd("d", 1, EndDate) & "'"
	FirstSearchItem = False
End If

If StatusToShow = "Complete" Then
	If FirstSearchItem = True Then
		sql = sql + " WHERE "
	Else
		sql = SQL + " AND "
	End If
	sql = sql + "DateComplete IS NOT NULL"
	FirstSearchItem = False
ElseIf StatusToShow = "Outstanding" Then
	If FirstSearchItem = True Then
		sql = sql + " WHERE "
	Else
		sql = SQL + " AND "
	End If
	sql = sql + "DateComplete IS NULL"
	FirstSearchItem = False
End If

If SortBy = "1" Then
	sql = sql & vbCrLF & "ORDER BY CallDateTime ASC, ContactPhoneCallID ASC"
ElseIf SortBy = "2" Then
	sql = sql & vbCrLf & "ORDER BY CallDateTime DESC, ContactPhoneCallID DESC"
End If
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs = Con.Execute(sql)

Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
Response.Write(vbTab & "<th>ID</th>" & vbCrLf)
Response.Write(vbTab & "<th>Date</th>" & vbCrLf)
'Response.Write(vbTab & "<th>Length (min.)</th>" & vbCrLf)
Response.Write(vbTab & "<th NOWRAP>Contact Name</th>" & vbCrLf)
Response.Write(vbTab & "<th>Contact Title</th>" & vbCrLf)
If ShowContactInfo = "Y" Then
	Response.Write(vbTab & "<th>Phone Number<br>Email</th>" & vbCrLf)
End If
Response.Write(vbTab & "<th>Grantee or Organization</th>" & vbCrLf)
Response.Write(vbTab & "<th>Issue</th>" & vbCrLf)
If ShowQandA = "Y" Then
	Response.Write(vbTab & "<th>Questions</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Answer</th>" & vbCrLf)
End If
If MVCPAContactID = 0 Then
	Response.Write(vbTab & "<th>TIDC Contact</th>" & vbCrLf)
End If
Response.Write(vbTab & "<th>Contact Type</th>" & vbCrLf)
Response.Write(vbTab & "<th>Who</th>" & vbCrLf)
Response.Write(vbTab & "<th>+<br>-</th>" & vbCrLf)
Response.Write(vbTab & "<th>Com-<br>plete</th>" & vbCrLf)
If ShowDaysToComplete = "Y" Then
	Response.Write(vbTab & "<th>Days To Complete</th>" & vbCrLf)
End If
Response.Write("</tr>" & vbCrLf)


While rs.EOF = False
	RecordCounter = RecordCounter + 1
	If rs.Fields("Complete") = True Then 
		CompleteCount = CompleteCount + 1
		CompleteTotal = CompleteTotal + rs.Fields("DaysToComplete")
	Else
		IncompleteCount = IncompleteCount + 1
	End If
	Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
	Response.Write(vbTab & "<td><a href=""ContactItems.asp?ContactPhoneCallID=" & _
		rs.Fields("ContactPhoneCallID") & """>" & rs.Fields("ContactPhoneCallID") & "</a></td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("CallDateTime") & "</td>" & vbCrLf)
	'Response.Write(vbTab & "<td>" & rs.Fields("CallLength") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td NOWRAP>" & rs.Fields("ContactName") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("ContactTitle") & "</td>" & vbCrLf)
	If ShowContactInfo = "Y" Then
		Response.Write(vbTab & "<td NOWRAP>" & rs.Fields("PhoneNumber"))
		If IsNull(rs.Fields("Email")) = False Then
			Response.Write(vbTab & "<br><a href=""mailto:" & rs.Fields("Email") & """>" & rs.Fields("Email") & "</a>")
		End If
		Response.Write("</td>" & vbCrLf)
	End If
	Response.Write(vbTab & "<td>")
	If IsNull(rs.Fields("Organization")) = False Then
		Response.Write(rs.Fields("Organization") & " ")
	End If
	If IsNull("GranteeName")=False Then
		Response.Write(rs.Fields("GranteeName"))
	End If
	Response.Write("</td>" & vbCrLf)
	Response.Write(vbTab & "<td>" & rs.Fields("Issue") & "</td>" & vbCrLf)
	If ShowQandA = "Y" Then
		Response.Write(vbTab & "<td>" & rs.Fields("Questions") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td>" & rs.Fields("Answer") & "</td>" & vbCrLf)
	End If
	If MVCPAContactID = 0 Then
		Response.Write(vbTab & "<td NOWRAP>" & rs.Fields("MVCPAContact") & "</td>" & vbCrLf)
	End If
	Response.Write(vbTab & "<td>" & rs.Fields("ContactType") & "</td>" & vbCrLf)
	Response.Write(vbTab & "<td align=center>" & rs.Fields("Initials") & "</td>" & vbCrLf)
	If rs.Fields("Positive") = True and rs.Fields("Negative") = True Then
		Response.Write("<td>+ / -</td>" & vbCrLF)
	ElseIf rs.Fields("Positive") = True Then
		Response.Write("<td>+</td>" & vbCrLF)
	ElseIf rs.Fields("Negative") = True Then
		Response.Write("<td>-</td>" & vbCrLF)
	Else
		Response.Write("<td>&nbsp;</td>" & vbCrLF)
	End If
	If IsNull(rs.Fields("DateComplete")) = True Then
		Response.Write(vbTab & "<td align=center>" & CheckBoxComplete(rs.Fields("ContactPhoneCallID")) & "</td>" & vbCrLf)
	Else
		Response.Write(vbTab & "<td align=center>" & CheckBox(True) & "</td>" & vbCrLf)
	End If
	If ShowDaysToComplete = "Y" Then
		Response.Write(vbTab & "<td>" & rs.Fields("DaysToComplete") & "</td>" & vbCrLf)
	End If
	Response.Write("</tr>" & vbCrLf)
	rs.MoveNext
WEnd

Response.Write("<tr valign=top>" & vbCrLf)
Response.Write("<td colspan=1>" & RecordCounter & "</td>" & vbCrLf)
Response.Write("<td colspan=2>Total Contacts Listed</td>" & vbCrLF)
Response.Write("<td>&nbsp;</td>" & vbCrLf)
If ShowContactInfo = "Y" Then
	Response.Write("<td>&nbsp;</td>" & vbCrLf)
	Response.Write("<td>&nbsp;</td>" & vbCrLf)
End If
Response.Write("<td>&nbsp;</td>" & vbCrLf)
Response.Write("<td>&nbsp;</td>" & vbCrLf)
If ShowQandA = "Y" Then
	Response.Write("<td>&nbsp;</td>" & vbCrLf)
	Response.Write("<td>&nbsp;</td>" & vbCrLf)
End If
Response.Write("<td>&nbsp;</td>" & vbCrLf)
Response.Write("<td>&nbsp;</td>" & vbCrLf)
Response.Write("<td align=center>C=" & CompleteCount & "<br>I=" & IncompleteCount & "</td>" & vbCrLf)
If ShowDaysToComplete = "Y" Then
	If CompleteCount > 0 Then
		Response.Write("<td align=right>" & formatnumber(CompleteTotal/CompleteCount,1,True,True,True) & "</td>" & vbCrLf)
	Else
		Response.Write("<td>&nbsp;</td>" & vbCrLf)
	End If
End If
Response.Write("</tr>" & vbCrLf)
%>

<tr>
	<td colspan="<%=ShowColumns%>" align="center"><input type="button" value="Close" 
		onClick="window.close()" alt="Close Button"> &nbsp;&nbsp;&nbsp;
		<input type="button" value="Add" onClick="location.href='../Contact/ContactItems.asp'" alt="Add Button"
			title="Add a new contact item."></td>
</tr>
</table>

</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<%
Function SelectedValue(vVariable,vValue)
	If vVariable = vValue Then
		SelectedValue = " SELECTED "
	Else
		SelectedValue = ""
	End If
End Function

Function CheckedValue(vVariable,vValue)
	If vVariable = vValue Then
		CheckedValue = " CHECKED "
	Else
		CheckedValue = ""
	End If
End Function

Function CheckBox(vBit)
	If vBit = True Then
		CheckBox = "<input type=checkbox CHECKED readOnly=True tabIndex=-1 onClick=""this.checked=true;"" id=checkbox1 name=checkbox1>"
	Else
		CheckBox = "<input type=checkbox readOnly=True tabIndex=-1 onClick=""this.checked=false;"" id=checkbox1 name=checkbox1>"
	End If
End Function

Function CheckBoxComplete(vID)
	CheckBoxComplete = "<input type=checkbox name=Complete" & vID & " readOnly=True tabIndex=-1 onClick=""CompleteItem(this," & vid & ")"" id=""checkbox" & vid & """>"
End Function

Function StartOfFiscalYear()
	dim currentdate
	currentdate = date()
	If month(currentdate)>8 Then
		StartOfFiscalYear = "9/1/" & year(currentdate)
	Else
		StartOfFiscalYear = "9/1/" & (year(currentdate)-1)
	End If
End Function
%>