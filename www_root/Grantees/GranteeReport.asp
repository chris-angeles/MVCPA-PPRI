<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, FiscalYear, OrderBy, OrderByDescription, OrderByField, ShowExcel, GranteesToShow, LEOOnly, counter
OrderByDescription = Array("Grantee ID", "Grantee Name", "ORI")
OrderByField = Array("G.GranteeID", "REPLACE(G.GranteeName,'City of ','')", "G.ORI")

debug = False
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
End If

Counter = 0
If Len(Request.Form("GranteesToShow"))>0 Then
	GranteesToShow = CInt(Request.Form("GranteesToShow"))
ElseIf Len(Request.QueryString("GranteesToShow"))>0 Then
	GranteesToShow = CInt(Request.QueryString("GranteesToShow"))
Else
	GranteesToShow = 6
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
'Response.Write("<pre>Date=" & Date & "</pre>")
'Response.Write("<pre>Month(Date())=" & Month(Date()) & "</pre>")
'Response.Write("<pre>FiscalYear=" & FiscalYear & "</pre>")

If Len(Request.Form("OrderBy"))>0 Then
	OrderBy = CInt(Request.Form("OrderBy"))
ElseIf Len(Request.QueryString("OrderBy"))>0 Then
	OrderBy = CInt(Request.QueryString("OrderBy"))
Else
	OrderBy = 1
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

sql = "SELECT G.GranteeID AS ID, G.ORI, G.GranteeName AS Grantee_Name, " & vbCrLf & _
	"	AO.Name AS Authorized_Official, CASE WHEN ISNULL(AO.AccountDisabled,'')=0 THEN AO.EMail ELSE NULL END AS Authorized_Official_Email,  " & vbCrLf & _
	"	CASE WHEN G.ProgramDirectorID IS NOT NULL THEN PD.Name " & vbCrLf & _
	"		WHEN G.AuthorizedOfficialID IS NULL AND G.ProgramDirectorID IS NULL AND " & vbCrLf & _
	"			G.ProgramManagerID IS NULL AND G.FinancialOfficerID IS NULL AND " & vbCrLf & _
	"			G.ProgramAdministrativeContactID IS NULL AND G.FinancialAdministrativeContactID IS NULL AND " & vbCrLf & _
	"			G.TaskForceCommanderID IS NULL AND G.PIOID IS NULL " & vbCrLf & _
	"		THEN D.Name END AS Program_Director, " & vbCrLf & _
	"	CASE WHEN G.ProgramDirectorID IS NOT NULL THEN PD.Email " & vbCrLf & _
	"		WHEN G.AuthorizedOfficialID IS NULL AND G.ProgramDirectorID IS NULL AND " & vbCrLf & _
	"			G.ProgramManagerID IS NULL AND G.FinancialOfficerID IS NULL AND " & vbCrLf & _
	"			G.ProgramAdministrativeContactID IS NULL AND G.FinancialAdministrativeContactID IS NULL AND " & vbCrLf & _
	"			G.TaskForceCommanderID IS NULL AND G.PIOID IS NULL " & vbCrLf & _
	"		THEN D.email END AS Program_Director_Email, " & vbCrLf & _
	"	PM.Name AS Program_Manager, CASE WHEN ISNULL(PM.AccountDisabled,'')=0 THEN PM.EMail ELSE NULL END AS Program_Manager_Email,  " & vbCrLf & _
	"	FO.Name AS Financial_Officer, CASE WHEN ISNULL(FO.AccountDisabled,'')=0 THEN FO.EMail ELSE NULL END AS Financial_Officer_Email,  " & vbCrLf & _
	"	PAC.Name AS Program_Administrative_Contact, CASE WHEN ISNULL(PAC.AccountDisabled,'')=0 THEN PAC.EMail ELSE NULL END AS Program_Administrative_Contact_Email,  " & vbCrLf & _
	"	FAC.Name AS Financial_Administrative_Contact, CASE WHEN ISNULL(FAC.AccountDisabled,'')=0 THEN FAC.EMail ELSE NULL END AS Financial_Administrative_Contact_Email,  " & vbCrLf & _
	"	TFC.Name AS Task_Force_Commander, CASE WHEN ISNULL(TFC.AccountDisabled,'')=0 THEN TFC.EMail ELSE NULL END AS Task_Force_Commander_Email,  " & vbCrLf & _
	"	PIO.Name AS [PIO / Media Contact], CASE WHEN ISNULL(PIO.AccountDisabled,'')=0 THEN PIO.EMail ELSE NULL END AS PIO_Email, " & vbCrLf & _
	"	ISNULL(G.BorderCounty,0) AS Border_County, ISNULL(G.PortCounty,0) AS Port_County, " & vbCrLf & _
	"	ISNULL(G.Port2County,0) AS Port_2_County"
If GranteesToShow = 4 Then
	sql = sql & ", " & vbCrLf & "	ISNULL(TaskforceGrant,0) AS [TF], ISNULL(CatalyticConverterGrant,0) AS [CC], ISNULL(AuxiliaryGrant,0) AS [MAG] " & vbCrLf
Else
	sql = sql & " " & vbCrLf
End If
sql = sql & "FROM Grantees G " & vbCrLf
If LEOOnly = True Then
	sql = sql & "LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID AND ISNULL(AO.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS PD ON PD.SystemID=G.ProgramDirectorID AND ISNULL(PD.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS PM ON PM.SystemID=G.ProgramManagerID AND ISNULL(PM.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS FO ON FO.SystemID=G.FinancialOfficerID AND ISNULL(FO.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS PAC ON PAC.SystemID=G.ProgramAdministrativeContactID AND ISNULL(PAC.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS FAC ON FAC.SystemID=G.FinancialAdministrativeContactID AND ISNULL(FAC.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS TFC ON TFC.SystemID=G.TaskForceCommanderID AND ISNULL(TFC.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN System.Users AS PIO ON PIO.SystemID=G.PIOID AND ISNULL(PIO.LicensedPeaceOfficer,0)=1 " & vbCrLf & _
	"LEFT JOIN (SELECT GranteeID, MIN(C2.SystemID) AS SystemID FROM System.GranteePermissions AS C1 JOIN System.Users AS C2 ON C2.SystemID=C1.SystemID WHERE ISNULL(C2.AccountDisabled,0)=0 AND ISNULL(C2.LicensedPeaceOfficer,0)=1 GROUP BY C1.GranteeID) AS C ON C.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN System.Users AS D ON D.SystemID=C.SystemID " & vbCrLf
Else
	sql = sql & "LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
	"LEFT JOIN System.Users AS PD ON PD.SystemID=G.ProgramDirectorID " & vbCrLf & _
	"LEFT JOIN System.Users AS PM ON PM.SystemID=G.ProgramManagerID " & vbCrLf & _
	"LEFT JOIN System.Users AS FO ON FO.SystemID=G.FinancialOfficerID " & vbCrLf & _
	"LEFT JOIN System.Users AS PAC ON PAC.SystemID=G.ProgramAdministrativeContactID " & vbCrLf & _
	"LEFT JOIN System.Users AS FAC ON FAC.SystemID=G.FinancialAdministrativeContactID " & vbCrLf & _
	"LEFT JOIN System.Users AS TFC ON TFC.SystemID=G.TaskForceCommanderID " & vbCrLf & _
	"LEFT JOIN System.Users AS PIO ON PIO.SystemID=G.PIOID " & vbCrLf & _
	"LEFT JOIN (SELECT GranteeID, MIN(C2.SystemID) AS SystemID FROM System.GranteePermissions AS C1 JOIN System.Users AS C2 ON C2.SystemID=C1.SystemID WHERE ISNULL(C2.AccountDisabled,0)=0 GROUP BY C1.GranteeID) AS C ON C.GranteeID=G.GranteeID " & vbCrLf & _
	"LEFT JOIN System.Users AS D ON D.SystemID=C.SystemID " & vbCrLf
End If
	sql = sql & "LEFT JOIN (select GranteeID, FiscalYear from [Grants].Main group by GranteeID, FiscalYear) AS GR ON GR.GranteeID=G.GranteeID AND GR.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"LEFT JOIN (SELECT DISTINCT GranteeID, FiscalYear FROM [Application].IDs) AS AP ON AP.GranteeID=G.GranteeID AND AP.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"LEFT JOIN MAG.Main AS MAGM ON MAGM.GranteeID=G.GranteeID AND MAGM.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"LEFT JOIN MAG.Admin AS MAGA ON MAGA.MAGID=MAGM.MAGID " & vbCrLf
If GranteesToShow = 1 Then
	sql = sql & "WHERE G.TaskForceGrant=1 " & vbCrLf
ElseIf GranteesToShow = 2 Then
	sql = sql & "WHERE AP.GranteeID IS NOT NULL " & vbCrLf
ElseIf GranteesToShow = 3 Then
	sql = sql & "WHERE GR.GranteeID IS NOT NULL " & vbCrLf
ElseIf GranteesToShow = 4 Then
	' No where clause needed.
ElseIf GranteesToShow = 5 Then
	sql = sql & "WHERE G.AuxiliaryGrant=1 " & vbCrLf
ElseIf GranteesToShow = 6 Then
	sql = sql & "WHERE MAGA.GrantAwardAmount>0 " & vbCrLf
ElseIf GranteesToShow = 7 Then
	sql = sql & "WHERE G.RapidResponseStrikeforceGrant=1 " & vbCrLf
ElseIf GranteesToShow = 8 Then
	sql = sql & "WHERE G.CatalyticConverterGrant=1 " & vbCrLf
End If
sql = sql &	"ORDER BY " & OrderByField(OrderBy)
If Debug = True Then
	Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
	Response.Flush
End If

Set rs=Con.Execute(sql)

If ShowExcel = True Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "filename=GranteeReport" & FiscalYear & ".xls"
	Response.Write("<table>" & vbCrLf)
Else ' Start of Web only code
	If Debug = False Then
		Response.ContentType = "text/html"
	End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>Grantee Report</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width: 100%">


<form name="Selection" id="Selection" method="post" >
<label for="GranteesToShow">Grantees to Show:</label>
<select name="GranteesToShow" onchange="Selection.submit();">
<%
Response.Write("<option value=""1"" "& Selected(GranteesToShow,1) & ">All Taskforce Grantees</option>")
Response.Write("<option value=""2"" "& Selected(GranteesToShow,2) & ">Taskforce Grant Applicants</option>")
Response.Write("<option value=""3"" "& Selected(GranteesToShow,3) & ">Awarded Taskforce Grants</option>" )
Response.Write("<option value=""4"" "& Selected(GranteesToShow,4) & ">All Grantees</option>" )
Response.Write("<option value=""5"" "& Selected(GranteesToShow,5) & ">Auxiliary Grant Grantees</option>" )
Response.Write("<option value=""6"" "& Selected(GranteesToShow,6) & ">Awarded Auxiliary Grants</option>" )
Response.Write("<option value=""7"" "& Selected(GranteesToShow,7) & ">Rapid Response Strikeforce Grantees</option>" )
Response.Write("<option value=""8"" "& Selected(GranteesToShow,8) & ">Catalytic Converter Grantees</option>" )
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
<a href="GranteeReport.asp?OrderBy=<%=OrderBy%>&LEOOnly=<%If LEOOnly=True Then Response.Write("1") Else Response.Write("0") End If %>&FiscalYear=<%=FiscalYear %>&GranteesToShow=<%=GranteesToShow %>&ShowExcel=1" target="_blank">Show Excel</a>
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
	If ShowExcel = False Then
		Response.Write("<tr><td colspan=""14"" style=""text-align: center; "">count=" & counter & "</td></td>" & vbCrLf)
	End If
Else
	Response.Write("<tr><td>Nothing to show</td></tr>" & vbCrLf)
End If
Response.Write("</tbody>" & vbCrLf)
%>

</table>
<%If ShowExcel = False Then %>
<div style="text-align: center; "><input type="button" value="Close" onclick="window.close();" /></div>

</body>
</html>
<%	End If %>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->