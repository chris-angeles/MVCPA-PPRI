<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, columns,  counter, FiscalYear, OrderBy, OrderByDescription, OrderByField, ShowExcel, GranteesToShow, _
	LEOOnly, BorderPort, BorderPortDescription, BorderPortClause
BorderPortDescription = Array ("All", "Border", "Port", "Port 2", "Border and Port", "Border, Port, and Port 2")
BorderPortClause = Array ("1=1", "G.BorderCounty=1", "G.PortCounty=1", "G.Port2County=1", "(G.BorderCounty=1 OR G.PortCounty=1)", "(G.BorderCounty=1 OR G.PortCounty=1 OR G.Port2County=1)")
OrderByDescription = Array("Grantee ID", "Grantee Name", "Name")
OrderByField = Array("G.GranteeID", "REPLACE(G.GranteeName,'City of ','')", "U.LastName, U.FirstName")

debug = False
columns = 9
counter = 0

If Debug = True Then
	For each i in Request.Form
		Response.Write("<pre>Request.Form(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Request.QueryString
		Response.Write("<pre>Request.QueryString(""" & i & """)='" & Request.Form(i) & "'</pre>" & vbCrLf)
	Next
	For each i in Session.Contents
		Response.Write("<pre>Session(""" & i & """)='" & Session(i) & "'</pre>" & vbCrLf)
	Next
	Response.Flush
End If

If Len(Request.Form("GranteesToShow"))>0 Then
	GranteesToShow = CInt(Request.Form("GranteesToShow"))
ElseIf Len(Request.QueryString("GranteesToShow"))>0 Then
	GranteesToShow = CInt(Request.QueryString("GranteesToShow"))
Else
	GranteesToShow = 5
End If
If Len(Request.Form("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear"))>0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	If Month(Date()) >= 9 Then
		FiscalYear = Year(Date)+1
	Else
		FiscalYear = Year(Date)
	End If
End If
If Request.Form("LEOOnly") = "1" Then
	LEOOnly = True
ElseIf Request.QueryString("LEOOnly") = "1" Then
	LEOOnly = True
Else
	LEOOnly = False
End If
If Len(Request.Form("BorderPort")) > 0 Then
	BorderPort = CInt(Request.Form("BorderPort"))
ElseIf Len(Request.QueryString("BorderPort")) > 0 Then
	BorderPort = CInt(Request.QueryString("BorderPort"))
Else
	BorderPort = 0
End If
'Response.Write("<pre>Date=" & Date & "</pre>")
'Response.Write("<pre>Month(Date())=" & Month(Date()) & "</pre>")
'Response.Write("<pre>FiscalYear=" & FiscalYear & "</pre>")

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 2
End If
If Len(Request.Form("ShowExcel"))>0 Then
	If Request.Form("ShowExcel")="1" Then 
		ShowExcel = True
	Else
		ShowExcel = False
	End If
ElseIf Len(Request.QueryString("ShowExcel"))>0 Then
	If Request.QueryString("ShowExcel")="1" Then 
		ShowExcel = True
	Else
		ShowExcel = False
	End If
Else
	ShowExcel = False
End If

sql = "SELECT Email, Name, G.GranteeName AS Grantee_Name, " & vbCrLf & _
	"	REVERSE(SUBSTRING(REVERSE(CASE WHEN G.ProgramDirectorID=U.SystemID THEN 'Program Director, ' ELSE '' END + " & vbCrLf & _
	"	CASE WHEN G.ProgramManagerID=U.SystemID THEN 'Program Manager, ' ELSE '' END + " & vbCrLf & _
	"	CASE WHEN G.FinancialOfficerID=U.SystemID THEN 'Financial Officer, ' ELSE '' END + " & vbCrLf & _
	"	CASE WHEN G.ProgramAdministrativeContactID=U.SystemID THEN 'Program Administrative Contact, ' ELSE '' END + " & vbCrLf & _
	"	CASE WHEN G.FinancialAdministrativeContactID=U.SystemID THEN 'Financial Administrative Contact, ' ELSE '' END + " & vbCrLf & _
	"	CASE WHEN G.TaskForceCommanderID=U.SystemID THEN 'Taskforce Commander, ' ELSE '' END), 3, 1000)) AS Position,  " & vbCrLf & _
	"	CASE WHEN TCOLEPID>0 THEN 'X' WHEN LicensedPeaceOfficer=1 THEN '<i>X</i>' ELSE '' END AS LEO, " & vbCrLf & _
	"	Phone, Mobile, ISNULL(G.BorderCounty,0) AS Border_County, " & vbCrLF & _
	"	ISNULL(G.PortCounty,0) AS Port_County, ISNULL(G.Port2County,0) AS Port_2_County " & vbCrLf & _
	"FROM [System].Users AS U " & vbCrLf & _
	"LEFT JOIN Grantees AS G ON U.SystemID IN (G.ProgramDirectorID, G.ProgramManagerID, G.FinancialOfficerID, G.ProgramAdministrativeContactID, G.FinancialAdministrativeContactID, G.TaskForceCommanderID)" & vbCrLf & _
	"LEFT JOIN [Grants].Main AS GR ON GR.GranteeID=G.GranteeID AND GR.FiscalYear=" & prepIntegerSQL(FiscalYear) & " AND GR.AwardAmount>0 " & vbCrLf & _
	"LEFT JOIN (SELECT DISTINCT GranteeID, FiscalYear FROM [Application].IDs) AS AP ON AP.GranteeID=G.GranteeID AND AP.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"LEFT JOIN MAG.Main AS MAG ON MAG.GranteeID=G.GranteeID AND MAG.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"LEFT JOIN MAG.Admin AS MA ON MA.MAGID=MAG.MAGID " & vbCrLf & _
	"WHERE ISNULL(AccountDisabled,0)=0 "
If GranteesToShow = 1 Then
	sql = sql & " AND G.GranteeID IS NOT NULL " & vbCrLf
ElseIf GranteesToShow = 2 Then
	sql = sql & " AND AP.GranteeID IS NOT NULL " & vbCrLf
ElseIf GranteesToShow = 3 Then
	sql = sql & " AND GR.GranteeID IS NOT NULL " & vbCrLf
ElseIf GranteesToShow = 4 Then
	sql = sql & " AND GR.AwardAmount>0 " & vbCrLf
ElseIf GranteesToShow = 5 Then
	sql = sql & " AND MA.GrantAwardAmount >0 " & vbCrLf
End If
If LEOOnly = True Then
	sql = sql & " AND ISNULL(U.LicensedPeaceOfficer,0)=1 " & vbCrLf
End If
If BorderPort > 0 Then
	sql = sql & " AND " & BorderPortClause(BorderPort)
End If
sql = sql &	"ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If ShowExcel = True Then
	If Debug = False Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "content-disposition", "filename=MailingList.xls"
	End If
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Mailing List</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="GranteesToShow">Grantees to Show:</label>
<select name="GranteesToShow" onchange="Selection.submit();">
<%
Response.Write("<option value=""1"" "& Selected(GranteesToShow,1) & ">All in Grantee Table (every applicant, any year)</option>")
Response.Write("<option value=""2"" "& Selected(GranteesToShow,2) & ">Applicants</option>")
Response.Write("<option value=""3"" "& Selected(GranteesToShow,3) & ">Awarded Grants</option>" )
Response.Write("<option value=""4"" "& Selected(GranteesToShow,4) & ">Awarded TaskforceGrants</option>" )
Response.Write("<option value=""5"" "& Selected(GranteesToShow,5) & ">Awarded Auxiliary Grants</option>" )
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="FiscalYear">Fiscal Year:</label> 
<select name="FiscalYear" id="FiscalYear" onchange="Selection.submit();">
<%
	For i = 2017 to Application("CurrentFiscalYear")+1
		Response.Write("<option value=""" & i & """ " & selected(FiscalYear, i) & ">" & i & "</option>" & vbCrLf)
	Next
%>
</select>&nbsp;&nbsp;&nbsp;
<label for="OrderBy">Order By:</label><select name="OrderBy" id="OrderBy" onchange="Selection.submit();">
<%
For i = 0 to UBound(OrderByDescription)
	Response.Write("<option value=""" & i & """ " & Selected(OrderBy, i) & ">" & OrderByDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;&nbsp;
<input name="LEOOnly" type="checkbox" <%=Checked(LEOOnly, True) %> value="1" onchange="Selection.submit();" /> Show LEO Only&nbsp;&nbsp;&nbsp;
<label for="BorderPort">Border/Port:</label> <select name="BorderPort" id="BorderPort" onchange="Selection.submit();">
<%
For i = 0 to UBound(BorderPortDescription)
	Response.Write("<option value=""" & i & """" & Selected(BorderPort, i) & ">" & BorderPortDescription(i) & "</option>" & vbCrLf)
Next
%>
</select>&nbsp;&nbsp;
<a href="MailingList.asp?GranteesToShow=<%=GranteesToShow %>&OrderBy=<%=OrderBy%>&FiscalYear=<%=FiscalYear %>&LEOOnly=<%If LEOOnly=True Then Response.Write("1") Else Response.Write("0") End If %>&BorderPort=<%=BorderPort %>&ShowExcel=1" target="_blank">Show Excel</a>
</form>

<br />
<table class="reporttable">
<%	End If %>
<%
If rs.EOF = False Then
	Response.Write("<thead>" & vbCrLf)
	Response.Write("<tr>" & vbCrLf)
	For i = 0 To (rs.Fields.Count-1)
		If ShowExcel = True Then
			Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
		ElseIf InStr(rs.Fields(i).Name, "_Email")>0 Then
			' skip field
		Else
			Response.Write("<th>" & Replace(rs.Fields(i).Name,"_"," ") & "</th>")
		End If
	Next
	Response.Write(vbCrLf & "</tr>" & vbCrLf)
	Response.Write("</thead>" & vbCrLf)
	Response.Write("<tbody>" & vbCrLf)

	While rs.EOF = False
		counter = counter + 1
		Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
		For i = 0 To (rs.Fields.Count-1)
			If ShowExcel = True Then
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			ElseIf InStr(rs.Fields(i).Name, "_Email")>0 Then
				' skip field
			ElseIf IsNull(rs.Fields(i).value) = True Then
				Response.Write("<td></td>")
			ElseIf rs.Fields(i).Name = "ID" Then
				If MVCPARights = True Then
					Response.Write("<td style=""text-align: right""><a href=""..\Grantees\Grantee.asp?GranteeID=" & rs.Fields(i) & """ target=""Main"">" & rs.Fields(i) & "</a></td>" & vbCrLf)
				Else
					Response.Write("<td style=""text-align: right"">" & rs.Fields(i) & "</td>" & vbCrLf)
				End If
			ElseIf rs.Fields(i).Name="FiscalYear" Or rs.Fields(i).Name="Fiscal_Year" Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, false, false) & "</td>")
			ElseIf rs.Fields(i).Type = adCurrency Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,2, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Type=adBigInt Or rs.Fields(i).Type=adInteger Or rs.Fields(i).Type=adSmallInt Or rs.Fields(i).Type=adUnsignedTinyInt Then
				Response.Write("<td style=""text-align: right"">" & formatnumber(rs.Fields(i).value,0, true, true, true) & "</td>")
			ElseIf rs.Fields(i).Name="Authorized_Official" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("Authorized_Official_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Name="Program_Director" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("Program_Director_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Name="Program_Manager" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("Program_Manager_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Name="Financial_Officer" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("Financial_Officer_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Name="Program_Administrative_Contact" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("Program_Administrative_Contact_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Name="Financial_Administrative_Contact" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("Financial_Administrative_Contact_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Name="Task_Force_Commander" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("Task_Force_Commander_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Name="PIO / Media Contact" Then
				Response.Write("<td><a href=""mailto:" & rs.Fields("PIO_EMail") & "&subject=MVCPA"" class=""plainlink"">" & rs.Fields(i).value & "</a></td>")
			ElseIf rs.Fields(i).Type = adBoolean Then
				If rs.Fields(i).value = True Then
					Response.Write("<td style=""text-align: center;"">X</td>")
				Else
					Response.Write("<td style=""text-align: center;""></td>")
				End If
			Else
				Response.Write("<td>" & rs.Fields(i).value & "</td>")
			End If
		Next
		Response.Write("</tr>" & vbCrLf)
		rs.MoveNext
	Wend
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
Response.Write("</tbody>" & vbCrLf)
Response.Write("<tfoot><tr><td colspan=""" & columns & """ style=""text-align: center"">" & counter & " records.</td></tr></tfoot>" & vbCrLf)
%>

</table>
<%If ShowExcel = False Then %>
<div style="text-align: center; "><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->