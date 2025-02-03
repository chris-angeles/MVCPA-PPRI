<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, vrs, vsql, ShowExcel, NegotiationYear, _
	OrderBy, OrderByDescription, OrderByField1, OrderByField2, GrantTypeID, _
	FilterBy, FilterByDescription, AppToShow, AppToShowDescription, ResolutionStatus, ShowAgreements
Dim DocumentFolder, fso, folder, file, files
OrderByDescription = Array("AppID", "Grantee Name", "Grantee ID", "Program Name", "MVCPA Funds Requested Ascending", "MVCPA Funds Requested Descending", "Cash Match Percentage")
OrderByField1 = Array("[App ID]", "REPLACE([Grantee Name],'City of ','')", "[Grantee ID]", "[Program Name]", "[MVCPA Funds Requested] ASC", "[MVCPA Funds Requested] DESC", "[Cash Match Pct] DESC")
OrderByField2 = Array("[App ID]", "REPLACE([Grantee Name],'City of ','')", "[Grantee ID]", "[Revised Program Name]", "[Revised MVCPA Funds Requested] ASC", "[Revised MVCPA Funds Requested] DESC", "[Revised Cash Match Pct] DESC")
FilterByDescription = Array("Show All", "Submitted", "Certified Complete", "Awarded", "Not Submitted")
AppToShowDescription = Array("Initial Application", "Negotiated Application", "Both")
debug = False
If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.QueryString(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("<pre>Session(""" & i & """)='" & Session(i) & "'</pre>" & vbCrLf)
	Next
End If

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	If Month(Date()) > 9 Then
		FiscalYear = Year(Date)+1
	Else
		FiscalYear = Year(Date)
	End If
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
End If

If Len(Request.Form("FilterBy"))>0 Then
	FilterBy = CInt(Request.Form("FilterBy"))
ElseIf Len(Request.QueryString("FilterBy"))>0 Then
	FilterBy = CInt(Request.QueryString("FilterBy"))
Else
	'If FiscalYear = 2021 Then
		FilterBy = 0
	'Else
	'	FilterBy = 3
	'End If
End If

If Len(Request.Form("GrantTypeID"))>0 Then 
	GrantTypeID = CInt(Request.Form("GrantTypeID"))
ElseIf Len(Request.QueryString("GrantTypeID"))>0 Then 
	GrantTypeID = CInt(Request.QueryString("GrantTypeID"))
Else
	GrantTypeID = 0
End If

If Len(Request.Form("AppToShow"))>0 Then 
	AppToShow = CInt(Request.Form("AppToShow"))
ElseIf Len(Request.QueryString("AppToShow"))>0 Then 
	AppToShow = CInt(Request.QueryString("AppToShow"))
Else
	AppToShow = 0
End If

If Request.Form("ResolutionStatus")="1" Then 
	ResolutionStatus = True
ElseIf Request.QueryString("ResolutionStatus")="1" Then 
	ResolutionStatus = True
Else
	ResolutionStatus = False
End If

If Request.Form("ShowAgreements")="1" Then 
	ShowAgreements = True
ElseIf Request.QueryString("ShowAgreements")="1" Then 
	ShowAgreements = True
Else
	ShowAgreements = False
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

If FiscalYear > 0 Then
	If getCCApplicationSchema(FiscalYear) = "Negotiation" Then 
		NegotiationYear = True
	Else
		NegotiationYear = False
	End If
Else
	NegotiationYear = False
End If
If Debug = True Then
	Response.Write("<pre>Negotiation Year = " & NegotiationYear & "</pre>" & vbCrLf)
End If
sql = "SELECT [App ID], [Fiscal Year], [Grantee ID], [Grantee Name], [Program Name], " & vbCrLf & _
	"	[Grant Type ID], [Grant Type], " & vbCrLf
If AppToShow = 0 Or AppToShow = 2 Then
	sql = sql & vbTab & "[Submitted By], [Submitted], [Application Certified Complete Date], " & vbCrLf & _
	"	[Official Grant Award Letter Date], " & vbCrLf & _
	"	[MVCPA Funds Requested], [Cash Match], [Cash Match Pct], [In-Kind Match], " & vbCrLf & _
	"	[Status], " & vbCrLf
End If
If AppToShow = 1 Or AppToShow = 2 Then
	sql = sql & "[Revised Submitted By], [Revised Submitted], [Revised Accepted Date], " & vbCrLf & _
	"	[Official Grant Award Letter Date], " & vbCrLf & _
	"	[Grant Award Certified Complete], [Revised MVCPA Funds Requested], " & vbCrLf & _ 
	"	[Revised Cash Match], [Revised Cash Match Pct], [Revised In-Kind Match], " & vbCrLf & _
	"	[Revised Status]," & vbCrLf
End If
sql = sql & "	[Grant Award Amount], [Initial Award Transmission Date], [Signed SGA], GrantID AS [Grant ID]"
If ShowAgreements = True Then
	sql = sql & ", " & vbCrLf & _
	"	[Interlocal Agreements Confirmed], [Prosecutor Agreements Confirmed], " & vbCrLf & _
	"	[Operational Plan], [Operational Plan Approved], [Multi-Agency Grant] " & vbCrLf
Else
	sql = sql & vbCrLf
End If
If NegotiationYear = True Then
	sql = sql & ", [Negotiation] AS [Negotiation Records Created] " & vbCrLf
Else
End If
sql = sql & "FROM [CC].[vwApplicationSummary] " & vbCrLf & _
	"WHERE GrantClassID=4 AND [Fiscal Year]=" & prepIntegerSQL(FiscalYear) & " " & vbCrLF
	If AppToShow = 1 Or AppToShow = 2 Then
		If GrantTypeID > 0 Then
			sql = sql & vbTab & "AND [Revised Grant Type ID]=" & prepIntegerSQL(GrantTypeID) & " " & vbCrLf
		End If
		If FilterBy = 1 Then
			sql = sql & vbTab & "AND [Revised Submitted] IS NOT NULL "
		ElseIf FilterBy = 2 Then
			sql = sql & vbTab & "AND [Grant Award Certified Complete] IS NOT NULL " & vbCrLf
		ElseIf FilterBy = 3 Then
			sql = sql & vbTab & "AND [Grant Award Amount] > 0 " & vbCrLf
		ElseIf FilterBy = 4 Then
			sql = sql & vbTab & "AND [Revised Submitted] IS NULL " & vbCrLf
		End If
		sql = sql & "ORDER BY " & OrderByField2(OrderBy)
	Else
		If GrantTypeID > 0 Then
			sql = sql & vbTab & "AND [Grant Type ID]=" & prepIntegerSQL(GrantTypeID) & " " & vbCrLf
		End If
		If FilterBy = 1 Then
			sql = sql & vbTab & "AND [Submitted] IS NOT NULL "
		ElseIf FilterBy = 2 Then
			sql = sql & vbTab & "AND [Application Certified Complete Date] IS NOT NULL " & vbCrLf
		ElseIf FilterBy = 3 Then
			sql = sql & vbTab & "AND [Grant Award Amount] > 0 " & vbCrLf
		ElseIf FilterBy = 4 Then
			sql = sql & vbTab & "AND [Submitted] IS NULL " & vbCrLf
		End If
		sql = sql & "ORDER BY " & OrderByField1(OrderBy)
	End If
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

Set rs=Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=ApplicationStatus" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Application Status Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<div class="sectiontitle" style="white-space: nowrap;"><%=FiscalYear%> Catalytic Converter Grant Applications</div>
<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to Year(Date())+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="OrderBy">Order By:</label> <select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """" & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<label for="GrantTypeID">Show Grant Type:</label> <select name="GrantTypeID" id="GrantTypeID" onchange="Selection.submit();">
<option value="0">All Grant Types</option>
<%
vsql = "SELECT GrantTypeID, GrantType FROM Lookup.GrantType WHERE Version=1 ORDER BY GrantTypeID"
Set vrs = Con.Execute(vsql)
While vrs.EOF = False
	Response.Write("<option value=""" & vrs.Fields("GrantTypeID") & """" & Selected(GrantTypeID, vrs.Fields("GrantTypeID")) & ">" & vrs.Fields("GrantType") & "</option>" & vbCrLf)
	vrs.MoveNext()
Wend
%></select>&nbsp;&nbsp;
<label for="FilterBy">Filter By:</label> <select name="FilterBy" id="FilterBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(FilterByDescription)
	Response.Write("<option value=""" & i & """" & Selected(FilterBy, i) & ">" & FilterByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<label for="AppToShow">Show:</label> <select name="AppToShow" id="AppToShow" onchange="Selection.submit();">
<%
For i = 0 to UBound(AppToShowDescription)
	Response.Write("<option value=""" & i & """" & Selected(AppToShow, i) & ">" & AppToShowDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<label for="Resolution">Resolution:</label> <select name="ResolutionStatus" id="ResolutionStatus" onchange="Selection.submit();">
<%
If ResolutionStatus = True Then
	Response.Write("<option value=""0"">No</option>" & vbCrLf)
	Response.Write("<option value=""1"" selected=""selected"">Yes</option>" & vbCrLf)
Else
	Response.Write("<option value=""0"" selected=""selected"">No</option>" & vbCrLf)
	Response.Write("<option value=""1"">Yes</option>" & vbCrLf)
End If
%>
</select>&nbsp;&nbsp;
<label for="ShowAgreements">Agreements:</label> <select name="ShowAgreements" id="ShowAgreements" onchange="Selection.submit();">
<%
If ShowAgreements = True Then
	Response.Write("<option value=""0"">No</option>" & vbCrLf)
	Response.Write("<option value=""1"" selected=""selected"">Yes</option>" & vbCrLf)
Else
	Response.Write("<option value=""0"" selected=""selected"">No</option>" & vbCrLf)
	Response.Write("<option value=""1"">Yes</option>" & vbCrLf)
End If
%>
</select>
&nbsp;&nbsp;<a href="ApplicationStatus.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&OrderBy=<%=OrderBy %>&GrantTypeID=<%=GrantTypeID %>&FilterBy=<%=FilterBy %>&AppToShow=<%=AppToShow %>&ResolutionStatus=<% If ResolutionStatus=True Then Response.Write("1") Else Response.Write("0") End If %>&ShowAgreements=<% If ShowAgreements=True Then Response.Write("1") Else Response.Write("0") End If %>" target="_blank">Excel</a>
</form>

<br />
<table class="reporttable">
<%
End If ' End of html only code.

If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	For i = 0 To (rs.Fields.Count-1)
		Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
		If MVCPARights = True and rs.Fields(i).Name="App ID" and ShowExcel = False Then
			'Response.Write("<th title=""Edit"">E</th><th title=""Manage"">M</th>")
			Response.Write("<th title=""Manage"">M</th>")
		End If
	Next
	If ResolutionStatus = True Then
		Response.Write("<th title=""Resolution"">Res</th>")
	End If
	Response.Write(vbCrLf & "</tr>" & vbCrLF)
	Response.Write("</thead>" & vbCrLf)

	While rs.EOF = False
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLF)
		For i = 0 To (rs.Fields.Count-1)
			If rs.Fields(i).Name = "Operational Plan" Then
				If IsNull(rs.Fields(i)) = False Then
					Response.Write("<td style=""text-align: center""><a href=""/OperationalPlan/OperationalPlan.asp?AppID=" & rs.Fields(i) & """ target=""_blank"">Form</a></td>")
				Else
					DocumentFolder = Application("DocumentRoot") & "\Application\" & rs.Fields("App ID") & "\"
					If Debug = True Then
						Response.Write("<pre>" & DocumentFolder & "</pre>" & vbCrLf)
					End If
					ShowOpPlanDocument()
				End If
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "Grantee ID" Then
				If MVCPARights = True And ShowExcel = False Then
					Response.Write("<td style=""text-align: right""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name = "App ID" And ShowExcel = False Then
				If MVCPARights = True Then
					If AppToShow = 0 Then
						Response.Write("<td style=""text-align: right"" title=""View/Print Application""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\CatalyticConverter\PrintApplication.asp?AppID=" & rs.Fields(i) & "&FiscalYear=" & FiscalYear & """ target=""_blank"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
					Else
						Response.Write("<td style=""text-align: right"" title=""View/Print Negotiation Application""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\CatalyticConverter\PrintApplication.asp?AppID=" & rs.Fields(i) & "&FiscalYear=" & FiscalYear & """ target=""_blank"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
					End If
					'Response.Write("<td style=""text-align: right"" title=""Edit Application""><a href=""..\Application\Application.asp?AppID=" & rs.Fields(i) & """ target=""Main"" class=""plainlink"">E</a></td>" & vbCrLf)
					Response.Write("<td style=""text-align: right"" title=""Manage Application""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\CatalyticConverter\AppAdmin.asp?AppID=" & rs.Fields(i) & "&FiscalYear=" & FiscalYear & """ target=""_blank"" class=""plainlink"">M</a></td>" & vbCrLf)
				ElseIF ShowExcel = False Then
					If FiscalYear > 2021 Then
						Response.Write("<td style=""text-align: right"" title=""View/Print Application""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\CatalyticConverter\PrintApplication.asp?AppID=" & rs.Fields(i) & "&FiscalYear=" & FiscalYear & """ target=""_blank"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
					Else
						Response.Write("<td style=""text-align: right"" title=""View/Print Application""><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "\CatalyticConverter\PrintApplication.asp?AppID=" & rs.Fields(i) & "&FiscalYear=" & FiscalYear & """ target=""_blank"" class=""plainlink"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
					End If
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="Fiscal Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Name="Cash Match Pct" Or rs.Fields(i).Name="Revised Cash Match Pct"  Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, false, false) & "%</td>")
			ElseIf rs.Fields(i).Name = "Application Certified Complete Date" _
				Or rs.Fields(i).Name = "Revised Accepted Date" _
				Or rs.Fields(i).Name="Grant Award Certified Complete" _
				Or rs.Fields(i).Name = "Initial Award Transmission Date" _
				Or rs.Fields(i).Name = "Official Grant Award Letter Date" _
				Or rs.Fields(i).Name = "Interlocal Agreements Confirmed" _ 
				Or rs.Fields(i).Name = "Prosecutor Agreements Confirmed" _
				Or rs.Fields(i).Name = "Operational Plan Approved" Then
				Response.Write("<td style=""text-align: right"">" & formatdatetime(rs.Fields(i).value, vbGeneralDate) & "</td>")
			ElseIf rs.Fields(i).Type = adBoolean Then
				If rs.Fields(i).Value = True Then
					Response.Write("<td style=""text-align: center"">X</td>" & vbCrLf)
				Else
					Response.Write("<td></td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>" & vbCrLf)
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>" & vbCrLF)
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>" & vbCrLf)
			End If
		Next
		If ResolutionStatus = True Then
			DocumentFolder = Application("DocumentRoot") & "\Application\" & rs.Fields("App ID") & "\"
			set fso = Server.CreateObject("Scripting.FileSystemOBject")
			If fso.FolderExists(DocumentFolder) Then
				If fso.FileExists(DocumentFolder & "Resolution.pdf") Then
					Response.Write("<td><a href=""https://" & Request.ServerVariables("SERVER_NAME")& "/Documents/Application/" & rs.Fields("App ID") & "/Resolution.pdf"" target=""_blank"">Res</a></td>" & vbCrLf)
				Else
					Response.Write("<td>" & "No" & "</td>" & vbCrLf)
				End If
			Else
				Response.Write("<td title=""Directory does not exist"">No</td>" & vbCrLf)
			End If
		End If
		'Response.Write("<td>" & rs.Fields("Application Certified Complete Date").Type & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%	If ShowExcel = False Then %>
<div style="width: 100%; text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%	End If 

'					ShowOpPlanDocument()
'				ShowOpPlanDocument("Multi-Agency Operational Plan Document.pdf")

Sub ShowOpPlanDocument()
	Dim vOpPlanDocument
	set fso = Server.CreateObject("Scripting.FileSystemOBject")
	If fso.FolderExists(DocumentFolder) Then
		Response.Write("<td>")
		vOpPlanDocument = "Operational or Multi-Agency Plan.pdf" 
		If fso.FileExists(DocumentFolder & vOpPlanDocument) Then
			Response.Write("<a href=""https://" & Request.ServerVariables("SERVER_NAME") & "/Documents/Application/" & rs.Fields("App ID") & "/" & vOpPlanDocument & """ target=""_blank"">Doc</a>" & vbCrLf)
		Else
			' Show nothing.
		End If
		vOpPlanDocument = "Multi-Agency Operational Plan Document.pdf" 
		If fso.FileExists(DocumentFolder & vOpPlanDocument) Then
			Response.Write("<a href=""https://" & Request.ServerVariables("SERVER_NAME") & "/Documents/Application/" & rs.Fields("App ID") & "/" & vOpPlanDocument & """ target=""_blank"">Doc</a>" & vbCrLf)
		Else
			' Show nothing.
		End If
		Response.Write("</td>" & vbCrLf)
	Else
		Response.Write("<td title=""Directory does not exist"">No</td>" & vbCrLf)
	End If
End Sub
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/getApplicationSchema.asp"-->