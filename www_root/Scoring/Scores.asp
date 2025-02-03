<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, LastAppID, FiscalYear, Columns, IncludeComments, AppToShow, NoOfScorers, _
	GrantTypeID, OrderBy, OrderByDescription, OrderByField, ShowExcel
debug = False
Columns = 17
FiscalYear = 2024

OrderByDescription = Array("App ID", "Grantee Name", "Program Name")
OrderByField = Array("[App_ID]", "[GranteeSort], [App_ID]", "[Program_Name], [App_ID]")

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

If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
End If

If FiscalYear > 2023 Then
	Response.Redirect("Scores3.asp?FiscalYear=" & FiscalYear)
ElseIf FiscalYear = 2023 Then
	Response.Redirect("Scores3.asp?FiscalYear=2024")
End If

If Len(Request.Form("AppToShow"))>0 Then
	AppToShow = CInt(Request.Form("AppToShow"))
ElseIf Len(Request.QueryString("AppToShow"))>0 Then
	AppToShow = CInt(Request.QueryString("AppToShow"))
Else
	AppToShow = 0
End If

If Len(Request.Form("GrantTypeID"))>0 Then
	GrantTypeID = CInt(Request.Form("GrantTypeID"))
ElseIf Len(Request.QueryString("GrantTypeID"))>0 Then
	GrantTypeID = CInt(Request.QueryString("GrantTypeID"))
Else
	GrantTypeID = 4
End If

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
End If

If Request.Form("IncludeComments")="1" Then 
	IncludeComments = True
ElseIf Request.QueryString("IncludeComments")="1" Then 
	IncludeComments = True
Else
	IncludeComments = False
End If

If Request.QueryString("ShowExcel")="1" Then 
	ShowExcel = True
Else
	ShowExcel = False
End If

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=Scores" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	Response.ContentType = "text/html"
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Application Score Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">

<div class="sectiontitle">MVCPA <%=FiscalYear%> Application Score Report</div>
<table style="margin: auto; "><tr><td><form name="Selection" id="Selection" method="post" >
<label for="FiscalYear">Fiscal Year:</label> <select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2018 to 2024 step 2
		Response.Write("<option value=""" & i & """ " & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%></select>&nbsp;&nbsp;&nbsp;
<label for="OrderBy">Order By:</label> <select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """ " & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="GrantTypeID">Show Grant Type:</label> <select name="GrantTypeID" id="GrantTypeID" onchange="Selection.submit();">
<option value="0">All Grant Types</option>
<%
sql = "SELECT GrantTypeID, GrantType FROM Lookup.GrantType WHERE Version=1 UNION SELECT 4,'Continued and New' ORDER BY GrantTypeID"
Set rs = Con.Execute(sql)
While rs.EOF = False
	Response.Write("<option value=""" & rs.Fields("GrantTypeID") & """" & Selected(GrantTypeID, rs.Fields("GrantTypeID")) & ">" & rs.Fields("GrantType") & "</option>" & vbCrLf)
	rs.MoveNext()
Wend
%></select>&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="IncludeComments" id="IncludeComments" value="1" <% If IncludeComments = True Then Response.Write("checked ") %>  onchange="Selection.submit(); " />
<label for="IncludeComments">Include Comments</label>
&nbsp;&nbsp;&nbsp;<a href="Scores.asp?ShowExcel=1&FiscalYear=<%=FiscalYear %>&OrderBy=<%=OrderBy %>&GrantTypeID=<%=GrantTypeID %>" target="_blank">Excel</a><br />
<label for="AppToShow">Application(s) To Show:</label> <select name="AppToShow" id="AppToShow" onchange="Selection.submit();">
<option value="0">All Applications</option>

<%
sql = "SELECT I.AppID, A.ProgramName, C.GranteeName, A.GrantTypeID, REPLACE(B.GrantType, ' Grant','') AS GrantType " & vbCrLF & _
	"FROM Application.IDs AS I " & vbCrLf & _
	"LEFT JOIN Application.Main AS A ON A.AppID=I.AppID AND I.GrantClassID=1 " & vbCrLf & _
	"LEFT JOIN Lookup.GrantType AS B ON B.GrantTypeID=A.GrantTypeID AND B.Version=1 " & vbCrLF & _
	"LEFT JOIN Grantees AS C ON C.GranteeID=I.GranteeID " & vbCrLf & _
	"WHERE I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND SubmitTimestamp IS NOT NULL" & vbCrLf
If GrantTypeID=4 Then
	sql = sql & " AND A.GrantTypeID IN (1,3) " & vbCrLf
ElseIf GrantTypeID>0 Then
	sql = sql & " AND A.GrantTypeID=" & prepIntegerSQL(GrantTypeID) & " " & vbCrLf
End If
sql = sql &	"ORDER BY REPLACE(C.GranteeName,'City of ',''), A.GrantTypeID "
If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
While rs.EOF = False
	If rs.Fields("AppID") = AppToShow Then
		Response.Write(vbTab & "<option value=""" & rs.Fields("AppID") & """ selected>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
	Else
		Response.Write(vbTab & "<option value=""" & rs.Fields("AppID") & """>" & rs.Fields("ProgramName") & ", " & rs.Fields("GranteeName") & ", " & rs.Fields("GrantType") & "</option>" & vbCrLf)
	End If
	rs.MoveNext()
Wend
%></select>
</form></td></tr></table>
<table style="margin: auto; ">
<%
End If
LastAppID=0
sql = "SELECT *, 1 AS AverageToEnd " & vbCrLF & _
	"FROM Scoring.vwScoringAverages WITH (NOLOCK) " & vbCrLF & _
	"WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " " & vbCrLF
	If GrantTypeID = 4 Then
		sql = sql & vbTab & "AND GrantTypeID IN (1,3) " & vbCrLf
	ElseIf GrantTypeID > 0 Then
		sql = sql & vbTab & "AND GrantTypeID=" & prepIntegerSQL(GrantTypeID) & " " & vbCrLf
	End If
	If AppToShow > 0 Then
		sql = sql & " AND App_ID=" & prepIntegerSQL(AppToShow) & " " & vbCrLf
	End If
	If IncludeComments = False Then
		sql = sql & " AND OfficialScorer=1 " & vbCrLf
	End If
sql = sql &	"UNION " & vbCrLF & _
	"SELECT *, 0 AS AverageToEnd " & vbCrLF & _
	"FROM Scoring.vwScores WITH (NOLOCK) " & vbCrLf & _
	"WHERE Fiscal_Year=" & prepIntegerSQL(FiscalYear) & " " & vbCrLF
	If GrantTypeID = 4 Then
		sql = sql & vbTab & "AND GrantTypeID IN (1,3) " & vbCrLf
	ElseIf GrantTypeID > 0 Then
		sql = sql & vbTab & "AND GrantTypeID=" & prepIntegerSQL(GrantTypeID) & " " & vbCrLf
	End If
	If AppToShow > 0 Then
		sql = sql & " AND App_ID=" & prepIntegerSQL(AppToShow) & " " & vbCrLf
	End If
	If IncludeComments = False Then
		sql = sql & vbTab & " AND OfficialScorer=1 " & vbCrLf
	End If
	sql = sql & "ORDER BY " & OrderByField(OrderBy) & ", AverageToEnd, Scorer "

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)

If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
	Response.Write(vbTab & "<th colspan=""" & Columns & """>Grantee / Program / Type" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("<tr style=""vertical-align: bottom; "">" & vbCrLf)
	Response.Write(vbTab & "<th>Scorer</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q1</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q2</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q3</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q4</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q5</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q6</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q7</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q8</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q9</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q10</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q11</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Q1-Q11</th>" & vbCrLf)
	Response.Write(vbTab & "<th>EC1</th>" & vbCrLf)
	Response.Write(vbTab & "<th>EC2</th>" & vbCrLf)
	Response.Write(vbTab & "<th>EC Total</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Total</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Meets Needs Requirement</th>" & vbCrLf)
	Response.Write(vbTab & "<th>Meets Other Sections Requirement</th>" & vbCrLf)
	Response.Write("</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)
	While rs.EOF = False
		If LastAppID<>rs.Fields("App_ID") Then
			LastAppID=rs.Fields("App_ID")
			Response.Write("<tr>" & vbCrLf)
			Response.Write(vbTab & "<td style=""text-align: center; "" colspan=""" & columns & """>" & rs.Fields("Grantee_Name").Value & " / ")
			Response.Write(rs.Fields("Program_Name").Value & " / " & vbCrLf)
			Response.Write(Replace(rs.Fields("Grant_Type").Value, " Grant","") & "</td>" & vbCrLf)
			Response.Write("</tr>" & vbCrLf)
		End If
		Response.Write("<tr>" & vbCrLf)
		'Response.Write(vbTab & "<td>" & rs.Fields("Grantee_Name").Value & "</td>" & vbCrLf)
		'Response.Write(vbTab & "<td>" & rs.Fields("Program_Name").Value & "</td>" & vbCrLf)
		'Response.Write(vbTab & "<td>" & Replace(rs.Fields("Grant_Type").Value, " Grant","") & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""white-space: nowrap; "">" & rs.Fields("Scorer").Value & "</td>" & vbCrLf)
		If IsNull(rs.Fields("Score_1").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_1")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_1") & "; text-align: right; "">" & rs.Fields("Score_1").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_1").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_2").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_2")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_2") & "; text-align: right; "">" & rs.Fields("Score_2").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_2").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_3").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_3")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_3") & "; text-align: right; "">" & rs.Fields("Score_3").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_3").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_4").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_4")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_4") & "; text-align: right; "">" & rs.Fields("Score_4").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_4").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_5").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_5")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_5") & "; text-align: right; "">" & rs.Fields("Score_5").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_5").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_6").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_6")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_6") & "; text-align: right; "">" & rs.Fields("Score_6").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_6").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_7").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_7")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_7") & "; text-align: right; "">" & rs.Fields("Score_7").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_7").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_8").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_8")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_8") & "; text-align: right; "">" & rs.Fields("Score_8").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_8").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_9").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_9")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_9") & "; text-align: right; "">" & rs.Fields("Score_9").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_9").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_10").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_10")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_10") & "; text-align: right; "">" & rs.Fields("Score_10").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_10").Value & "</td>" & vbCrLf)
		End If
		If IsNull(rs.Fields("Score_11").Value) = True Then
			Response.Write(vbTab & "<td></td>" & vbCrLf)
		ElseIf IsNull(rs.Fields("Color_11")) = False Then
			Response.Write(vbTab & "<td style=""background-color: " & rs.Fields("Color_11") & "; text-align: right; "">" & rs.Fields("Score_11").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_11").Value & "</td>" & vbCrLf)
		End If
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Question 1-11 Total").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_12").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Score_13").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("EC Total").Value & "</td>" & vbCrLf)
		Response.Write(vbTab & "<td style=""text-align: right; "">" & rs.Fields("Total").Value & "</td>" & vbCrLf)
		If rs.Fields("Meets Needs Requirement").Value = "No" Then
			Response.Write(vbTab & "<td style=""text-align: center; font-weight: bold; "">" & rs.Fields("Meets Needs Requirement").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("Meets Needs Requirement").Value & "</td>" & vbCrLf)
		End If
		If rs.Fields("Meets Other Sections Requirement").Value = "No" Then
			Response.Write(vbTab & "<td style=""text-align: center; font-weight: bold; "">" & rs.Fields("Meets Other Sections Requirement").Value & "</td>" & vbCrLf)
		Else
			Response.Write(vbTab & "<td style=""text-align: center; "">" & rs.Fields("Meets Other Sections Requirement").Value & "</td>" & vbCrLf)
		End If
		Response.Write("</tr>" & vbCrLf)
		If IncludeComments = True Then
			Response.Write("<tr><td></td><td colspan=""" & (Columns - 1) & """>")
			For j = 1 to 13
				If IsNull(rs.Fields("Comments_" & j)) = False Then
					Response.Write("Q" & j & ": " & rs.Fields("Comments_" & j) & "<br />" & vbCrLf)
				End If
			Next
			Response.Write("</td></tr>")
		End If
		rs.MoveNext()
	Wend
	Response.Write("</tbody>" & vbCrLf)
Else
	Response.Write("<tr><td>There are no scores to report</td></tr>")
End If

%>
</table>
<%	If ShowExcel = False Then %>
</body>
</html>
<%	End If 

function Selected(vVariable, vValue)
	If vVariable = vSelected Then
		selected = " selected"
	Else
		seelcted = ""
	End If
end function
%>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->