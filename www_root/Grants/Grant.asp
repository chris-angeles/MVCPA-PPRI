<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, ViewDocuments, GrantID, FiscalYear, GranteeID, GranteeName, _
	GrantClassID, GrantClass, GrantNumber, DisplayQuarterOffset, CurrentYearAllocation, PriorYearAllocation, _
	ProgramName, MatchAmount, AwardAmount, InLieuOfDPSBudget, InLieuOfNICBBudget, _
	ReimbursementRate, ReportsCompleteDate, ProgramGoalsDate, DeficienciesResolvedDate, _
	FundsReturnedDate, CloseoutID, CloseoutName, CloseoutDate, AppID, _
	AdministrativeComments, UpdateID, UpdateName, UpdateTimestamp
debug = false
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

GrantID = Request.QueryString("GrantID")


If GrantID > 0 Then
sql = "SELECT G.GrantID, G.FiscalYear, G.GrantClassID, GC.GrantClass, G.GranteeID, H.GranteeName, G.GrantNumber, " & vbCrLf & _
	"	ISNULL(G.DisplayQuarterOffset,0) AS DisplayQuarterOffset, G.CurrentYearAllocation, G.PriorYearAllocation, " & vbCrLf & _
	"	G.ProgramName, G.MatchAmount, G.AwardAmount, " & vbCrLf & _
	"	G.InLieuOfDPSBudget, G.InLieuOfNICBBudget, G.ReimbursementRate, " & vbCrLf & _
	"	G.ReportsCompleteDate, G.ProgramGoalsDate, G.DeficienciesResolvedDate, G.FundsReturnedDate, " & vbCrLf & _
	"	ISNULL(G.CloseoutID,0) AS CloseoutID, T.Name AS CloseoutName, G.CloseoutDate, G.AppID, " & vbCrLf & _
	"	G.AdministrativeComments, G.UpdateID, G.UpdateTimestamp, ISNULL(U.Name,'') AS UpdateName " & vbCrLf & _
	"FROM [Grants].Main AS G " & vbCrLF & _
	"LEFT JOIN Grantees AS H ON H.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN System.Users AS T ON T.SystemID=G.CloseoutID " & vbCrLf & _
	"LEFT JOIN System.Users AS U ON U.SystemID=G.UpdateID " & vbCrLf & _
	"LEFT JOIN Lookup.GrantClass AS GC ON GC.GrantClassID=G.GrantClassID " & vbCrLF & _
	"WHERE GrantID=" & prepIntegerSQL(GrantID)
Else
	Response.Write("Error: No grant ID provided.")
	Response.End
End If
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)
If rs.EOF = False Then
	GrantID = rs.Fields("GrantID")
	FiscalYear = rs.Fields("FiscalYear")
	GrantClassID = rs.Fields("GrantClassID")
	GrantClass = rs.Fields("GrantClass")
	GranteeID = rs.Fields("GranteeID")
	GranteeName = rs.Fields("GranteeName")
	GrantNumber = rs.Fields("GrantNumber")
	DisplayQuarterOffset = rs.Fields("DisplayQuarterOffset")
	CurrentYearAllocation = rs.Fields("CurrentYearAllocation")
	PriorYearAllocation = rs.Fields("PriorYearAllocation")
	ProgramName = rs.Fields("ProgramName")
	MatchAmount = rs.Fields("MatchAmount")
	AwardAmount = rs.Fields("AwardAmount")
	InLieuOfDPSBudget = rs.Fields("InLieuOfDPSBudget")
	InLieuOfNICBBudget = rs.Fields("InLieuOfNICBBudget")
	ReimbursementRate = rs.Fields("ReimbursementRate")
	ReportsCompleteDate = rs.Fields("ReportsCompleteDate")
	ProgramGoalsDate = rs.Fields("ProgramGoalsDate")
	DeficienciesResolvedDate = rs.Fields("DeficienciesResolvedDate")
	FundsReturnedDate = rs.Fields("FundsReturnedDate")
	CloseoutID = rs.Fields("CloseoutID")
	CloseoutName = rs.Fields("CloseoutName")
	CloseoutDate = rs.Fields("CloseoutDate")
	AppID = rs.Fields("AppID")
	AdministrativeComments = rs.Fields("AdministrativeComments")
	UpdateID = rs.Fields("UpdateID")
	UpdateName = rs.Fields("UpdateName")
	UpdateTimestamp = rs.Fields("UpdateTimestamp")
Else
	Response.Write("Error retrieving grant record.")
	Response.End
End If

UserGranteeID = GranteeID
UserFiscalYear = FiscalYear

PermitEdit = CheckPermissions(UserSystemID, GranteeID, True)
ViewDocuments = PermitEdit

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grant Page</title>
<meta http-equiv="cache-control" content="no-cache" />
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
<!--#include file="../includes/InputValidation.asp"-->
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">The Grant Page displays information about a grant and allows those 
	with appropriate permissions to edit the grant information.</div>

<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">

<form name="Grant" method="post" action="GrantSubmit.asp">
<table>
	<tr>
		<td>GrantID</td>
		<td><%=IntegerField("GrantID", GrantID, 6, 8, False, "") %></td>
	</tr>
	<tr>
		<td>FiscalYear</td>
		<td><%=IntegerField("FiscalYear", FiscalYear, 4, 4, False, "") %></td>
	</tr>
	<tr>
		<td>Grant Class</td>
		<td><%=GrantClass %>:(<%=IntegerField("GrantClassID", GrantClassID, 1, 1, False, "") %>)</td>
	</tr>
	<tr>
		<td>GranteeID</td>
		<td><%=IntegerField("GranteeID", GranteeID, 6, 8, False, "") %></td>
	</tr>
	<tr>
		<td>Grantee Name</td>
		<td><%=TextField("GranteeName", GranteeName, 80, 255, False, "") %></td>
	</tr>
	<tr>
		<td>Program Name</td>
		<td><%=TextField("ProgramName", ProgramNAme, 76, 255, MVCPAAuditor, "")  %></td>
	</tr>
	<tr>
		<td>GrantNumber</td>
		<td><%=TextField("GrantNumber", GrantNumber, 14, 14, MVCPAAuditor, "") %></td>
	</tr>
	<tr>
		<td>Total Award Amount</td>
		<td><%=CurrencyField("AwardAmount", AwardAmount, 12, 16, MVCPAAuditor, "checkCurrency(this)") %></td>
	</tr>
	<tr>
		<td>&nbsp;&nbsp;From Current Year Appropriation</td>
		<td><%=CurrencyField("CurrentYearAllocation", CurrentYearAllocation, 12, 16, MVCPAAuditor, "checkCurrency(this)") %></td>
	</tr>
	<tr>
		<td>&nbsp;&nbsp;From Prior Year Appropriation</td>
		<td><%=CurrencyField("PriorYearAllocation", PriorYearAllocation, 12, 16, MVCPAAuditor, "checkCurrency(this)") %></td>
	</tr>
	<tr>
		<td>Match Amount</td>
		<td><%=CurrencyField("MatchAmount", MatchAmount, 12, 16, MVCPAAuditor, "checkCurrency(this)") %></td>
	</tr>
	<tr>
		<td style="white-space: nowrap;">In Lieu Of Amount from DPS in Budget</td>
		<td><%=CurrencyField("InLieuOfDPSBudget", InLieuOfDPSBudget, 12, 16, MVCPAAuditor, "checkCurrency(this)") %></td>
	</tr>
	<tr>
		<td style="white-space: nowrap;">In Lieu Of Amount from NICB in Budget</td>
		<td><%=CurrencyField("InLieuOfNICBBudget", InLieuOfNICBBudget, 12, 16, MVCPAAuditor, "checkCurrency(this)") %></td>
	</tr>
	<tr>
		<td style="white-space: nowrap;">Reimbursement Rate</td>
		<td><%=prepNumberWeb(ReimbursementRate,8) %>%</td>
	</tr>
	<tr>
		<td style="white-space: nowrap;">Display Quarter Offset</td>
		<td><%=IntegerField("DisplayQuarterOffset", [DisplayQuarterOffset], 1, 1, (Developer OR MVCPAAdministrator OR MVCPAAuditor) AND FiscalYear>2023, "checkInteger(this)") %>
			Integer to add to one for starting quarter (2024+).
		</td>
	</tr>
	<tr><td colspan="2"><hr /></td></tr>
	<tr><td colspan="2"><table>
	<tr>
		<td style="vertical-align: top; ">All financial, progress, inventory, performance and other reports required as a condition of the grant have been submitted.</td>
		<td><%=DateField("ReportsCompleteDate", ReportsCompleteDate, MVCPARights) %></td>
	</tr>
	<tr style="vertical-align: top; ">
		<td>All program goals were met or explanation provided as to why goals were not achieved.</td>
		<td><%=DateField("ProgramGoalsDate", ProgramGoalsDate, MVCPARights) %></td>
	</tr>
	<tr style="vertical-align: top; ">
		<td>Any deficiencies identified in a monitoring visit have been resolved</td>
		<td><%=DateField("DeficienciesResolvedDate", DeficienciesResolvedDate, MVCPARights) %></td>
	</tr>
	<tr style="vertical-align: top; ">
		<td>Notice sent to grantee about funds returned as a result of credits, rebates, or transaction errors.</td>
		<td><%=DateField("FundsReturnedDate", FundsReturnedDate, MVCPARights) %></td>
	</tr>
	</table></td></tr>
	<tr style="vertical-align: top; ">
		<td style="white-space: nowrap;" title="All reports have been completed, all checks have been issued. The grant process is complete.">Grant Closeout Date</td>
		<td><%=DateField("CloseoutDate", CloseoutDate, MVCPARights) %><%
		If Len(CloseoutName)>0 Then
			Response.Write(" By " & CloseoutName)
		End If
		%></td>
	</tr>
	<tr>
		<td>Last Update</td>
		<td><%
		If GrantID>0 Then
			Response.Write(UpdateName & " (" & UpdateID & "), " & UpdateTimestamp)
		End If 
		%></td>
	</tr>
	<tr><td colspan="2"><hr /></td></tr>
	<tr style="vertical-align: top; ">
		<td>Administrative Comments</td>
		<td><%=TextArea("AdministrativeComments", AdministrativeComments, 3, 80, 1020, MVCPARights, "") %></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td colspan="2">
	<table style="margin: auto;  border: 1px solid #dddddd; ">

	<tr>
		<td style="vertical-align: top; text-align: center"><b>Participating Agencies</b>
		<td style="vertical-align: top; text-align: center "><b>Coverage Agencies</b><br />
	</tr>
	<tr>
		<td style="vertical-align: top">

<%
	sql = "SELECT A.ORI, REPLACE(B.Agency,'&','&amp;') AS Agency" & vbCrLF & _
		"FROM [Grants].ParticipatingAgencies AS A" & vbCrLF & _
		"LEFT JOIN Lookup.ORI AS B ON B.ORI=A.ORI " & vbCrLf & _
		"WHERE A.GrantID = " & prepIntegerSQL(GrantID) & vbCrLF & _
		"ORDER BY A.ORI "
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(rs.Fields("ORI") & " " & rs.Fields("Agency") & "<br />" & vbCrLf)
		rs.MoveNext
	Wend
%></td>
		<td style="vertical-align: top">
<%
	sql = "SELECT A.ORI, REPLACE(B.Agency,'&','&amp;') AS Agency" & vbCrLF & _
		"FROM [Grants].CoverageAgencies AS A" & vbCrLF & _
		"LEFT JOIN Lookup.ORI AS B ON B.ORI=A.ORI " & vbCrLf & _
		"WHERE A.GrantID = " & prepIntegerSQL(GrantID) & vbCrLF & _
		"ORDER BY A.ORI "
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(rs.Fields("ORI") & " " & rs.Fields("Agency") & "<br />" & vbCrLf)
		rs.MoveNext
	Wend
%></td>
	</tr>
	</table></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
<%
If ViewDocuments = True Then
	Dim Folder, file, files, DocumentFolder, fso
	DocumentFolder = Application("DocumentRoot") & "\Grant\" & GrantID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	Response.Write("<tr style=""vertical-align: top; ""><td colspan=""2"" style=""text-align: center"">Current Documents in grant folder: ")
	If PermitEdit = True Then
		Response.Write("<a href=""../Upload/Upload.asp?fid=3&GrantID=" & GrantID & "&AppID=" & AppID & """ class=""plainlink"" target=""_blank"">Upload</a>")
	End If
	Response.Write("</td>" & vbCrLf)
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		Response.Write("<tr><td colspan=""2"">")
		If files.count>0 Then 
			For Each file in files
				Response.Write("<a href=""../Documents/Grant/" & GrantID & "/" & file.Name & _
					""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
			Next
			Response.Write("</td></tr>" & vbCrLf)
		Else
			Response.Write("No files in folder")
		End If
	Else
		Response.Write("<tr style=""vertical-align: top; ""><td colspan=""2"" style=""text-align: center"">No Documents in folder</td></tr>")
	End If

	DocumentFolder = Application("DocumentRoot") & "\Application\" & AppID & "\"
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	Response.Write("<tr style=""vertical-align: top; ""><td colspan=""2"" style=""text-align: center"">Current Documents in application folder: ")
	Response.Write("</td>" & vbCrLf)
	If fso.FolderExists(DocumentFolder) Then
		Set folder = fso.GetFolder(DocumentFolder)
		Set files = folder.Files
		Response.Write("<tr><td colspan=""2"">")
		If files.count>0 Then 
			For Each file in files
				Response.Write("<a href=""../Documents/Application/" & AppID & "/" & file.Name & _
					""" target=""_blank"">" & file.Name & "</a> (" & file.DateLastModified & ")<br />" & vbCrLf)
			Next
			Response.Write("</td></tr>" & vbCrLf)
		Else
			Response.Write("No files in folder")
		End If
	Else
		Response.Write("<tr style=""vertical-align: top; ""><td colspan=""2"" style=""text-align: center"">No Documents in folder</td></tr>")
	End If
End If

If FiscalYear>2017 Then
	Response.Write("<tr><td colspan=""2"">&nbsp;</td></tr>" & vbCrLf)
	sql = "SELECT A.BudgetCategoryID, A.BudgetCategory, TotalExpenditures, MVCPAExpenditures, MatchExpenditures, InKindExpenditures " & vbCrLf & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLf & _
	"LEFT JOIN [Grants].Budget AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND B.GrantID=" & prepIntegerSQL(GrantID) & " " & vbCrLf & _
	"UNION " & vbCrLf & _
	"SELECT 99, 'Total' AS BudgetCategory, SUM(TotalExpenditures) AS TotalExpenditures, SUM(MVCPAExpenditures) AS MVCPAExpenditures, SUM(MatchExpenditures) AS MatchExpenditures, SUM(InKindExpenditures) AS InKindExpenditures " & vbCrLF & _
	"FROM Lookup.BudgetCategories AS A " & vbCrLF & _
	"LEFT JOIN [Grants].Budget AS B ON B.BudgetCategoryID=A.BudgetCategoryID AND B.GrantID=" & prepIntegerSQL(GrantID) & " " & vbCrLF & _
	"ORDER BY A.BudgetCategoryID "
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Response.Write("<tr><td style=""text-align: center"" colspan=""2""><table>" & vbCrLf)
	Response.Write("<tr><th colspan=""5"">Approved Budget</th></tr>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom;"">" & vbCrLf) 
	Response.Write(vbTab & "<td style=""text-align: center"">Budget Category</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center"">Total Expenditures</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center"">MVCPA Expenditures</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center"">Match Expenditures</td>" & vbCrLf)
	Response.Write(vbTab & "<td style=""text-align: center"">In-Kind Expenditures</td>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf) 
	Set rs=Con.Execute(sql)
	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf) 
		Response.Write(vbTab & "<td style=""text-align: left"">" & rs.Fields("BudgetCategory") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWeb(rs.Fields("TotalExpenditures")) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWeb(rs.Fields("MVCPAExpenditures")) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWeb(rs.Fields("MatchExpenditures")) & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right"">" & prepCurrencyWeb(rs.Fields("InKindExpenditures")) & "</td>" & vbCrLf)
		Response.Write("</tr>" & vbCrLf) 
		rs.MoveNext
	Wend
	Response.Write("</table></td></tr>" & vbCrLf)
End If

If MVCPARights = True Then %>
	<tr>
		<td colspan="2" style="text-align: center">
			<input type="submit" name="Submit" value="Submit" title="Submit changes and return." />
			<input type="button" name="Cancel" value="Cancel" title="No changes. Return to home." 
				onclick="location.href = '../home/default.asp?GranteeID=<%=GranteeID%>';"/>
		</td>
	</tr>
<%	
End If 
%>
</table>
</form>

</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/CheckPermissions.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
