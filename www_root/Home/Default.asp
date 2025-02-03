<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, GranteeID, GranteeCount, FiscalYear, _
	GranteeName, OrganizationType, ORI, ORIAgency, StatePayeeIDNo, _
	GeneralPhone, Address1, Address2, City, State, zip, _
	AuthorizedOfficialID, ProgramDirectorID, FinancialOfficerID, ProgramManagerID, TaskForceCommanderID, _
	ProgramAdministrativeContactID, FinancialAdministrativeContactID, PIOID, ReducedOfficials, _
	TaskforceGrant, AuxiliaryGrant, RapidResponseStrikeforceGrant, CatalyticConverterGrant

debug = False
If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLf)
	For each i in Request.Form
		Response.Write("Request.Form(""" & i & """)='" & Request.Form(i) & "'" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'" & vbCrLf)
	Next
	For each i in Application.Contents
		Response.Write("Application(""" & i & """)='" & Application(i) & "'" & vbCrLf)
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
	Session("GranteeID") = GranteeID
	Response.Cookies("GranteeID") = GranteeID
	UserGranteeID = GranteeID
ElseIf Len(Request.QueryString("GranteeID"))>0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
	Session("GranteeID") = GranteeID
	Response.Cookies("GranteeID") = GranteeID
	UserGranteeID = GranteeID
ElseIf Len(Session("GranteeID"))>0 Then
	GranteeID = Session("GranteeID")
	Response.Cookies("GranteeID") = GranteeID
ElseIf Len(Response.Cookies("GranteeID"))>0 Then
	GranteeID = CInt(Response.Cookies("GranteeID"))
Else 
	GranteeID = 0
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
	Session("FiscalYear") = FiscalYear
	Response.Cookies("FiscalYear") = FiscalYear
	UserFiscalYear = FiscalYear
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
	Session("FiscalYear") = GranteeID
	Response.Cookies("FiscalYear") = FiscalYear
	UserFiscalYear = FiscalYear
ElseIf Len(Session("FiscalYear"))>0 Then
	FiscalYear = CInt(Session("FiscalYear"))
	UserFiscalYear = FiscalYear
	Response.Cookies("FiscalYear") = FiscalYear
ElseIf Len(Response.Cookies("FiscalYear"))>0 Then
	FiscalYear = CInt(Response.Cookies("FiscalYear"))
	UserFiscalYear = FiscalYear
Else 
	FiscalYear = Application("DefaultFiscalYear")
	Response.Cookies("FiscalYear") = FiscalYear
	UserFiscalYear = FiscalYear
End If

PermitEdit = CheckPermissions(UserSystemID, GranteeID, True)

UserGranteeID = GranteeID
'UserFiscalYear= Session("FiscalYear")
'Response.Cookies("GranteeID") = GranteeID

If Debug = True Then
	Response.Write("<pre>")
	Response.Write("FiscalYear=" & FiscalYear & vbCrLf)
	Response.Write("Type Of Fiscal Year=" & TypeName(FiscalYear) & vbCrLf)
	Response.Write("GranteeID=" & GranteeID) 
	Response.Write("<pre>")
End If

' The year in the next line should allow for any year in which applications are still being accepted so that all active grantees can apply.
If FiscalYear>2024 Then
	sql = "SELECT G.GranteeID, G.GranteeNameSort AS GranteeName, G.ORI, " & vbCrLf & _
		"	CAST(CASE WHEN ISNULL(CatalyticConverterGrant,0)=1 AND ISNULL(TaskforceGrant,0)=0 THEN 1 ELSE 0 END AS BIT) AS SB224 " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"WHERE ISNULL(G.Inactive,0)=0 Or GranteeID=" & prepIntegerSQL(GranteeID) & " OR " & vbCrLf & _
		"	EXISTS (SELECT GranteeID FROM Grants.Main AS A WHERE A.GranteeID=G.GranteeID AND FiscalYear=" & prepIntegerSQL(FiscalYear) & ") OR " & vbCrLf & _
		"	EXISTS (SELECT GranteeID FROM MAG.Main AS B WHERE B.GranteeID=G.GranteeID AND FiscalYear=" & prepIntegerSQL(FiscalYear) & ") OR " & vbCrLF & _
		"	EXISTS (SELECT GranteeID FROM Application.IDs AS C WHERE C.GranteeID=G.GranteeID AND FiscalYear=" & prepIntegerSQL(FiscalYear) & ") " & vbCrLf & _
		"ORDER BY GranteeNameSort "
Else
	sql = "SELECT G.GranteeID, G.GranteeNameSort AS GranteeName, G.ORI, " & vbCrLf & _
		"	CAST(CASE WHEN ISNULL(CatalyticConverterGrant,0)=1 AND ISNULL(TaskforceGrant,0)=0 THEN 1 ELSE 0 END AS BIT) AS SB224 " & vbCrLf & _
		"FROM Grantees AS G " & vbCrLf & _
		"WHERE EXISTS (SELECT GranteeID FROM Grants.Main AS A WHERE A.GranteeID=G.GranteeID AND FiscalYear=" & prepIntegerSQL(FiscalYear) & ") OR " & vbCrLf & _
		"	EXISTS (SELECT GranteeID FROM MAG.Main AS B WHERE B.GranteeID=G.GranteeID AND FiscalYear=" & prepIntegerSQL(FiscalYear) & ") OR " & vbCrLF & _
		"	EXISTS (SELECT GranteeID FROM Application.IDs AS C WHERE C.GranteeID=G.GranteeID AND FiscalYear=" & prepIntegerSQL(FiscalYear) & ") OR " & vbCrLf & _
		"	GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf & _
		"ORDER BY G.GranteeNameSort;"
End If

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Texas MVCPA Home Page</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<script type="text/javascript">
	window.name = "Main";

	function changeOfficial(position, currentid)
	{
		document.ChangeOfficial.Position.value = position;
		document.ChangeOfficial.CurrentID.value = currentid;
		document.ChangeOfficial.submit();
	}
</script>
</head>
<!-- UPdated by Jim 11/20/2023 11:50 am. -->
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">Home Page for <%=UserName %></div>

<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">
<!--FiscalYear=<%=UserFiscalYear %>-->
<form name="SelectGrantee" method="post" action="Default.asp">
Grantee to Display: <select name="GranteeID" onchange="document.SelectGrantee.submit();">
	<option value="0">Select a Grantee</option>
<%
Set rs = Con.Execute(sql)
While rs.EOF = False
	If IsNull(rs.Fields("ORI")) Then
		Response.Write("<option value=""" & rs.Fields("GranteeID") & """ " & selected(GranteeID, rs.fields("GranteeID")) & ">" & rs.Fields("GranteeName") & "</option>" & vbCrLf)
	Else
		If rs.Fields("SB224") = True Then
			Response.Write("<option value=""" & rs.Fields("GranteeID") & """ " & selected(GranteeID, rs.fields("GranteeID")) & ">" & rs.Fields("GranteeName") & " [" & rs.Fields("ORI") & "] (SB224)</option>" & vbCrLf)
		Else
			Response.Write("<option value=""" & rs.Fields("GranteeID") & """ " & selected(GranteeID, rs.fields("GranteeID")) & ">" & rs.Fields("GranteeName") & " [" & rs.Fields("ORI") & "]</option>" & vbCrLf)
		End If
	End If
	rs.MoveNext()
Wend
%>
</select> Fiscal Year: <select name="FiscalYear" onchange="document.SelectGrantee.submit();">
<%
For i = 2017 to Application("CurrentFiscalYear")+1
	Response.Write("<option value=""" & i & """ " & selected(i, FiscalYear) & ">FY" & (i mod 100) & "</option>" & vbCrLf)
Next
%>
</select>
</form><br />
<%
sql = "SELECT G.GranteeID, G.GranteeName, OT.OrganizationType, G.ORI, O.Agency AS ORIAgency, StatePayeeIDNo, " & vbCrLf & _
	"	GeneralPhone, Address1, Address2, City, State, zip,  " & vbCrLf & _
	"	ISNULL(AuthorizedOfficialID,0) AS AuthorizedOfficialID, " & vbCrLf & _
	"	ISNULL(ProgramDirectorID,0) AS ProgramDirectorID, " & vbCrLf & _
	"	ISNULL(FinancialOfficerID,0) AS FinancialOfficerID, " & vbCrLf & _
	"	ISNULL(ProgramManagerID,0) AS ProgramManagerID, " & vbCrLf & _
	"	ISNULL(TaskForceCommanderID,0) AS TaskForceCommanderID, " & vbCrLf & _
	"	ISNULL(ProgramAdministrativeContactID,0) AS ProgramAdministrativeContactID, " & vbCrLf & _
	"	ISNULL(FinancialAdministrativeContactID,0) AS FinancialAdministrativeContactID, " & vbCrLf & _
	"	ISNULL(PIOID,0) AS PIOID, " & _
	"	CAST(CASE WHEN (AuxiliaryGrant=1 Or RapidResponseStrikeforceGrant=1) AND ISNULL(TaskforceGrant,0)=0 THEN 1 ELSE 0 END AS BIT) AS ReducedOfficials, " & vbCrLf & _
	"	TaskforceGrant, AuxiliaryGrant, RapidResponseStrikeforceGrant, CatalyticConverterGrant " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN Lookup.OrganizationType AS OT ON OT.OrganizationTypeID=G.OrganizationTypeID " & vbCrLf & _
	"LEFT JOIN Lookup.ORI AS O ON O.ORI = G.ORI " & vbCrLf & _
	"WHERE GranteeID=" & prepIntegerSQL(GranteeID)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs = Con.Execute(sql)

If rs.EOF = False Then
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	OrganizationType = rs.Fields("OrganizationType")
	ORI = rs.Fields("ORI")
	ORIAgency = rs.Fields("ORIAgency")
	StatePayeeIDNo = rs.Fields("StatePayeeIDNo")
	GeneralPhone = rs.Fields("GeneralPhone")
	Address1 = rs.Fields("Address1")
	Address2 = rs.Fields("Address2")
	City = rs.Fields("City")
	State = rs.Fields("State")
	zip = rs.Fields("zip")
	AuthorizedOfficialID = rs.Fields("AuthorizedOfficialID")
	ProgramDirectorID = rs.Fields("ProgramDirectorID")
	FinancialOfficerID = rs.Fields("FinancialOfficerID")
	ProgramManagerID = rs.Fields("ProgramManagerID")
	TaskForceCommanderID = rs.Fields("TaskForceCommanderID")
	ProgramAdministrativeContactID = rs.Fields("ProgramAdministrativeContactID")
	FinancialAdministrativeContactID = rs.Fields("FinancialAdministrativeContactID")
	PIOID = rs.Fields("PIOID")
	ReducedOfficials = rs.Fields("ReducedOfficials")
	TaskforceGrant = rs.Fields("TaskforceGrant")
	AuxiliaryGrant = rs.Fields("AuxiliaryGrant")
	RapidResponseStrikeforceGrant = rs.Fields("RapidResponseStrikeforceGrant")
	CatalyticConverterGrant = rs.Fields("CatalyticConverterGrant")

	If Debug = True Then
		Response.Write("<pre>AuthorizedOfficialID=" & AuthorizedOfficialID & "</pre>")
		Response.Write("<pre>ProgramDirectorID=" & ProgramDirectorID & "</pre>")
		Response.Write("<pre>FinancialOfficerID=" & FinancialOfficerID & "</pre>")
		Response.Write("<pre>ProgramManagerID=" & ProgramManagerID & "</pre>")
		Response.Write("<pre>TaskForceCommanderID=" & TaskForceCommanderID & "</pre>")
		Response.Write("<pre>ProgramAdministrativeContactID=" & ProgramAdministrativeContactID & "</pre>")
		Response.Write("<pre>FinancialAdministrativeContactID=" & FinancialAdministrativeContactID & "</pre>")
		Response.Write("<pre>PIOID=" & PIOID & "</pre>")
		Response.Write("<pre>ReducedOfficials=" & ReducedOfficials & "</pre>")
		Response.Write("<pre>RapidResponseStrikeforceGrant=" & RapidResponseStrikeforceGrant & "</pre>")
		Response.Write("<pre>CatalyticConverterGrant=" & CatalyticConverterGrant & "</pre>")
	End If
%><table style="width: 738px">
<tr>
	<td>Primary Agency / Grantee Legal Name:</td>
	<td colspan="2"><%=GranteeName %></td>
</tr>
<tr>
	<td>Organization Type:</td>
	<td colspan="2"><%=OrganizationType %></td>
</tr><tr>
	<td>ORI (if applicable):</td>
	<td colspan="2"><%=ORI %>&nbsp;&nbsp;<%=ORIAgency %></td>
</tr>
<tr>
	<td>State Payee Identification Number:</td>
	<td colspan="2"><%=StatePayeeIDNo %></td>
</tr>
<tr>
	<td>General Phone Number:</td>
	<td colspan="2"><%=GeneralPhone %></td>
</tr>
<tr>
	<td>Official Address:</td>
	<td colspan="2"><%
	Response.Write(Address1 & "<br />")
	If IsNull(Address2) = False Then
		Response.Write(Address2 & "<br />")
	End If
	If IsNull(City) = False And IsNull(State) = False Then
		Response.Write(City & ", " & State & " " & zip)
	End If
%></td></tr>
<tr style="vertical-align: top; "><td>Grant Eligibility Tracks:</td><td><%
If TaskforceGrant Then
	Response.Write("TaskForce Grants")
Else
	Response.Write("<span style=""text-decoration: line-through; "">TaskForce Grants</span>")
End If
If AuxiliaryGrant Then
	Response.Write("; Auxiliary Grants")
Else
	Response.Write("; <span style=""text-decoration: line-through; "">Auxiliary Grants</span>")
End If
If RapidResponseStrikeforceGrant Then
	Response.Write("; Rapid Response Strikeforce Grants")
Else
	Response.Write("; <span style=""text-decoration: line-through; "">Rapid Response Strikeforce Grants</span>")
End If
If CatalyticConverterGrant Then
	Response.Write("; Catalytic Converter Grants")
Else
	Response.Write("; <span style=""text-decoration: line-through; "">Catalytic Converter Grants</span>")
End If

%></td></tr>
<tr><td colspan="3">&nbsp;</td></tr>
<tr><th colspan="3">Grant Officials</th></tr>
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

	If FinancialOfficerID > 0 Then
		If AuthorizedOfficialID > 0 Then
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

	If rs.Fields("FinancialOfficerID") > 0 Then
		If rs.Fields("ProgramManagerID") > 0 Then
			If rs.Fields("FinancialOfficerID") = rs.Fields("ProgramManagerID") Then
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
		If FinancialAdministrativeContactID > 0 Then
			If FinancialOfficerID = FinancialAdministrativeContactID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Financial Administrative Contact cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If TaskForceCommanderID > 0 Then
			If FinancialOfficerID = TaskForceCommanderID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Task Force Commander cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If

	If FinancialOfficerID > 0 Then
		If PIOID > 0 Then
			If FinancialOfficerID = PIOID Then
				Response.Write("<tr><td colspan=""3"" style=""text-align: center; color: red; font-weight: bold;"">The Public Information Officer cannot be the same person as the Financial Officer</td></tr>")
			End If
		End If
	End If
'End If
%>
<tr><td colspan="3">&nbsp;</td></tr>
  </table>

<%
sql = "SELECT U.Name, U.SystemID, U.email, U.Title " & vbCrLf & _
	"FROM Grantees AS G " & vbCrLf & _
	"LEFT JOIN system.GranteePermissions AS GP ON GP.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=GP.SystemID " & vbCrLf & _
	"WHERE G.GranteeID=" & prepIntegerSQL(GranteeID) & " AND ISNULL(AccountDisabled,0)=0 AND NOT(U.MVCPAGrantCoordinator=1 OR U.MVCPAGrantCoordinator=1 OR U.MVCPAAdministrativeAssistant=1 OR U.Developer=1) " & vbCrLf & _
	"	AND NOT(ISNULL(G.AuthorizedOfficialID,0)=U.SystemID OR ISNULL(G.ProgramDirectorID,0)=U.SystemID OR ISNULL(G.ProgramManagerID,0)=U.SystemID OR ISNULL(G.FinancialOfficerID,0)=U.SystemID OR ISNULL(G.ProgramAdministrativeContactID,0)=U.SystemID OR ISNULL(G.FinancialAdministrativeContactID,0)=U.SystemID OR ISNULL(G.TaskForceCommanderID,0)=U.SystemID)"
Set rs = Con. Execute(sql)
If rs.EOF = False Then
	Response.Write("<table align=""center"">")
	Response.Write("<tr><th>Others with write permissions for grantee</th></tr>" & vbCrLf)
	While rs.EOF = False
		Response.Write("<tr><td style=""text-align: center""><a href=""mailto:" & rs.Fields("Email") & "?subject=MVCPA"" class=""plainlink"">" & rs.Fields("Name"))
		If IsNull(rs.Fields("Title")) = false Then
			Response.Write(", " & rs.Fields("Title"))
		End If
		Response.Write("</a></td></tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</table>")
End If
%>

<%
sql = "SELECT FiscalYear, ProgramName, AwardAmount, GrantNumber " & vbCrLf & _
	"FROM [Grants].Main " & vbCrLf & _
	"WHERE GranteeID=" & prepIntegerSQL(GranteeID) & " " & vbCrLf & _
	"ORDER BY FiscalYEar DESC"
Set rs = Con. Execute(sql)
If rs.EOF = False Then
	Response.Write(vbTab & "<div class=""sectiontitle"">Grants Awarded</div>" & vbCrLf)
	Response.Write(vbTab & "<table align=""center"" style=""cell-padding: 4px; "">" & vbCrLf)
	Response.Write(vbTab & "<thead>" & vbCrLf)
	Response.Write(vbTab & "<tr>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Year</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Name</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Award</th>" & vbCrLf)
	Response.Write(vbTab & vbTab & "<th>Grant Number</th>" & vbCrLf)
	Response.Write(vbTab & "</tr>" & vbCrLf)
	Response.Write(vbTab & "</thead>" & vbCrLf)
	Response.Write(vbTab & "<tbody>" & vbCrLf)
	While rs.EOF = False
		Response.Write(vbTab & "<tr>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td>" & rs.Fields("FiscalYear") & "</td>" & vbCrLf)
		Response.Write(vbTab & vbTab & "<td>" & rs.Fields("ProgramName") & "</td>" & vbCrLf)
		If IsNull(rs.Fields("AwardAmount")) = False Then
			Response.Write(vbTab & vbTab & "<td>" & formatcurrency(rs.Fields("AwardAmount"),2,true,true,true) & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & vbTab & "<td></td>" & vbCrLf)
		End If
		Response.Write(vbTab & vbTab & "<td>" & rs.Fields("GrantNumber") & "</td>" & vbCrLf)
		Response.Write(vbTab & "</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write(vbTab & "</tbody>" & vbCrLf)
	Response.Write(vbTab & "</table>" & vbCrLf)
End If
%>
<!--<input type="button" value="Edit" title="Edit Grantee Information" style="text-align: center" onclick="location.href='../Grantees/Grantee.asp?GranteeID=<%=GranteeID%>'" />-->
<!--<br />Create an <a href="../Application/ISA.asp?ISAID=0&FiscalYear=2018&GranteeID=<%=GranteeID %>" class="plainlink">ISA</a>.-->
<form name="ChangeOfficial" method="post" action="../Grantees/ChangeOfficial1.asp">
<input type="hidden" name="ReturnPage" value="../Home/Default.asp" />
<input type="hidden" name="GranteeID" value="<%=GranteeID %>" />
<input type="hidden" name="Position" value="" />
<input type="hidden" name="CurrentID" value="" />
</form>
<%	
End If
If MVCPARights = True Then %>
<br />
<div><a href="../Grantees/Grantee.asp?GranteeID=0" class="plainlink">Create a new Grantee / Organization</a></div>
<%	End If %>
<%	If Session("MVCPARights") = True Then
		If GranteeID = 0 Then
			If FiscalYear > 0 Then
				ShowDashboard FiscalYear
			Else
				ShowDashboard Application("CurrentFiscalYear")
			End If
		End if
	Else %>
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
<%	End If %>
</div>

<div class="clearfix"></div>

<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../Home/Dashboard.asp"-->
<!--#include file="../includes/ShowPosition.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->