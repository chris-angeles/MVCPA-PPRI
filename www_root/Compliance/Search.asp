<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, MonitorType, OrderByDescription, MonitorTypeDescription, OrderByField, _
	ShowExcel, OpenClosedType, OpenClosedTypeDescription, CurrentDate, _
	fso, folder, file, DocumentRoot, Directory, filename, FileCount
OrderByDescription = Array("Monitor ID", "Grantee Name")
MonitorTypeDescription = Array("All", "Desk Review","Site Visit", "CAFR", "Other Audit")
OpenClosedTypeDescription = Array("All", "Open","Closed")
OrderByField = Array("MonitorID", "REPLACE(GranteeName,'City of ','')")
debug = False
CurrentDate = Date()
DocumentRoot = Application("DocumentRoot")

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
End If
If Len(Request.Form("MonitorType"))>0 Then
	MonitorType = CInt(Request.Form("MonitorType"))
ElseIf Len(Request.QueryString("MonitorType"))>0 Then
	MonitorType = CInt(Request.QueryString("MonitorType"))
End If
If Len(Request.Form("OpenClosedType")) > 0 Then
	OpenClosedType = CInt(Request.Form("OpenClosedType"))
ElseIf Len(Request.QueryString("ShowAmounts"))>0 Then
	OpenClosedType = CInt(Request.QueryString("OpenClosedType"))
Else
	OpenClosedType = 0
End If
If Request.Form("ShowExel") = "1" Then
	ShowExcel = True
ElseIf Request.QueryString("ShowExcel") = "1" Then
	ShowExcel = True
Else
	ShowExcel = False
End If

sql = "SELECT * FROM Monitor.vwMain" & vbCrLf & _
	"WHERE 1=1 " 
If FiscalYear>0 Then
	sql = sql & "	AND FiscalYear=" & FiscalYear & " " & vbCrLf
End If
If MonitorType>0 Then
	If MonitorType = 1 Then
		sql = sql & " AND DeskReview=1"
	ElseIf MonitorType = 2 Then
		sql = sql & " AND SiteVisit=1"
	ElseIf MonitorType = 3 Then
		sql = sql & " AND CAFR=1"
	ElseIf MonitorType = 4 Then
		sql = sql & " AND (ExternalAudit=1 OR OtherStateAgencyAudit=1 OR OtherAudit=1)"
	End If
End If
If OpenClosedType > 0 Then
	If OpenClosedType = 1 Then
		sql = sql & " AND CompletionClosedDate IS NULL "
	Else
		sql = sql & " AND CompletionClosedDate IS NOT NULL "
	End If
End If
sql = sql & vbCrLf & "ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)


If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=Search.xls"
Else
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Search Grant Monitoring</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<div style="margin: auto; text-align: center">
<form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
	<option value="0">All Years</option>
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """" & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="MonitorType">Monitor Type:</label> <select name="MonitorType" id="MonitorType" onchange="Selection.submit();">
<%
	For i = 0 to 4
		Response.Write("<option value=""" & i & """" & selected(MonitorType, i) & ">" & MonitorTypeDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<label for="OpenClosedType">Open/Closed:</label> <select name="OpenClosedType" id="OpenClosedType" onchange="Selection.submit();">
<%
	For i = 0 to 2
		Response.Write("<option value=""" & i & """" & selected(OpenClosedType, i) & ">" & OpenClosedTypeDescription(i) & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;
<a href="Search.asp?ShowExcel=1&FiscalYear=<%=FiscalYear%>&MonitorType=<%=MonitorType %>&OpenClosedType=<%=OpenClosedType%>&OrderBy=<%=OrderBy %>" target="_blank">Excel</a>
</form>
</div>
<br />
<%
End If
%>
<table class="reporttable">
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLF)
	Response.Write(vbCrLf & "<th rowspan=""2"">ID</th>")
	Response.Write(vbCrLf & "<th rowspan=""2"">Grantee</th>")
	Response.Write(vbCrLf & "<th rowspan=""2"">Type</th>")
	Response.Write(vbCrLf & "<th colspan=""2"">Years Reviewed</th>")
	Response.Write(vbCrLf & "<th colspan=""2"">Site Visit Dates</th>")
	Response.Write(vbCrLf & "<th rowspan=""2"" style=""width: 125px; "">Data Collection Complete</th>" & vbCrLF)
	Response.Write(vbCrLf & "<th colspan=""4"">Action Plan</th>")
	Response.Write(vbCrLf & "<th rowspan=""2"">Closed</th>")
	Response.Write(vbCrLf & "</tr>" & vbCrLf)

	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
	Response.Write("<th>Start</th>" & vbCrLf)
	Response.Write("<th>End</th>" & vbCrLf)
	Response.Write("<th>Start</th>" & vbCrLf)
	Response.Write("<th>End</th>" & vbCrLf)
	Response.Write("<th>Required</th>" & vbCrLf)
	Response.Write("<th>Due Date</th>" & vbCrLf)
	Response.Write("<th>Follow-up Date</th>" & vbCrLf)
	Response.Write("<th>Complete Date</th>" & vbCrLf)
	Response.Write("<th>Files</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLF)
	Response.Write("<thead>" & vbCrLf)

	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		FileCount=0
		Directory = Application("DocumentRoot") & "Monitor/" & rs.Fields("MonitorID")
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		If fso.FolderExists(Directory) Then
			Set folder = fso.GetFolder(Directory)
			For Each file in folder.Files
				FileCount = FileCount + 1
			Next
		Else
			FileCount = 0
		End If
		Response.Write("<tr style=""vertical-align: top;"">" & vbCrLf)
		Response.Write("<td style=""text-align: center; ""><a href=""Monitor.asp?MonitorID=" & rs.Fields("MonitorID") & """ target=""_blank"">" & rs.Fields("MonitorID") & "</a></td>" & vbCrLf)
		Response.Write("<td>" & rs.Fields("GranteeName").value & "</td>")
		Response.Write("<td>" & rs.Fields("MonitorType").value & "</td>")
		Response.Write("<td>" & rs.Fields("YearsReviewedStart").value & "</td>")
		Response.Write("<td>" & rs.Fields("YearsReviewedEnd").value & "</td>")
		Response.Write("<td>" & rs.Fields("StartDate").value & "</td>")
		Response.Write("<td>" & rs.Fields("EndDate").value & "</td>")
		Response.Write("<td style=""text-align: center; "">" & rs.Fields("DataCollectionCompleteDate").value & "</td>")
		Response.Write("<td style=""text-align: center; "">" & rs.Fields("ActionPlanRequiredText").value & "</td>")
		Response.Write("<td style=""text-align: center; "">" & rs.Fields("ActionPlanDueDate").value & "</td>")
		Response.Write("<td style=""text-align: center; "">" & rs.Fields("ActionPlanFollowUpDate").value & "</td>")
		Response.Write("<td style=""text-align: center; "">" & rs.Fields("ActionPlanCompleteDate").value & "</td>")
		Response.Write("<td style=""text-align: center; "">" & rs.Fields("CompletionClosedDate").value & "</td>")
		Response.Write("<td style=""text-align: center; "">" & FileCount & "</td>")
		'Response.Write("<td>" & rs.Fields(7).Type & "</td>")
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
	Response.Write("</tbody>" & vbCrLf)
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
%>
</table>
<%
If ShowExcel = False Then
%>
<div style="text-align: center"><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%
End If
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->