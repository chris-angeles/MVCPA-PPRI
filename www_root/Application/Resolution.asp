<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"--><% 
Dim debug, i, j, AppID, GranteeID, FiscalYear, GranteeName, AuthorizedOfficial, AuthorizedOfficialTitle, _
	ProgramDirector, ProgramDirectorTitle, FinancialOfficer, FinancialOfficerTitle, Statute, GrantProgramTitle, para2
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

AppID = Request.QueryString("AppID")
GranteeID = Request.QueryString("GranteeID")
FiscalYear = Request.QueryString("FiscalYear")
para2 = "WHEREAS, this grant program will assist this jurisdiction to combat motor vehicle burglary and theft; and "

If Len(AppID)>0 Then 
	AppID=CInt(AppID)
Else
	AppID = 0
End If

If Len(GranteeID)>0 Then 
	GranteeID=CInt(GranteeID)
Else
	GranteeID = 0
End If

If Len(FiscalYear)=4 Then
	FiscalYear=CInt(FiscalYear)
Else
	FiscalYear = 2021
End If

If FiscalYear > 2020 Then
	Statute = "Texas Transportation Code Chapter 1006"
Else
	Statute = "Texas Revised Civil Statutes Article 4413(37)"
End If

If FiscalYear > 2021 Then
    If FiscalYear = 2025 Then
        GrantProgramTitle = "SB224 Catalytic Converter Grant Program"
		para2="WHEREAS, this grant program will assist this jurisdiction to combat catalytic converter theft; and"
    Else
        GrantProgramTitle = "Motor Vehicle Crime Prevention Authority Grant Program"
    End If
Else
    GrantProgramTitle = "Taskforce Grant Program"
End If


If AppID=0 and GranteeID=0 Then
	FiscalYear=2018
	GranteeName="{City/County/AgencyName}"
	ProgramDirector="{Position-Example- MVCPA Commander, Chief of Police, etc...}"
	AuthorizedOfficial = "{County Judge / Sheriff / City Manager / Police Chief / Executive Director, etc...} of this {county / city / agency}"
	FinancialOfficer = "{Position-Example- County Auditor, City CFO, etc...}"
	AuthorizedOfficialTitle = "County Judge /Mayor/ Executive Director/City Manager"
	ProgramDirectorTitle = "{ProgramDirectorTitle}"
	FinancialOfficerTitle = "{FinancialOfficerTitle}"
Else
	sql = "SELECT G.GranteeID, G.GranteeName, ISNULL(I.FiscalYear," & FiscalYear & ") AS FiscalYear, " & vbCrLf & _
		"	AO.Name AS AuthorizedOfficial, AO.Title AS AuthorizedOfficialTitle, " & vbCrLf & _
		"	PD.Name AS ProgramDirector, PD.Title AS ProgramDirectorTitle, " & vbCrLf & _
		"	FO.Name AS FinancialOfficer, FO.Title AS FinancialOfficerTitle, A.AppID AS AppID " & vbCrLF & _
		"FROM Grantees AS G " & vbCrLF & _
		"LEFT JOIN Application.IDs AS I ON I.GranteeID=G.GranteeID" & " AND I.FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
		"LEFT JOIN Application.Main AS A ON A.AppID=I.AppID " & vbCrLf & _
		"LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
		"LEFT JOIN System.Users AS PD ON PD.SystemID=G.ProgramDirectorID " & vbCrLf & _
		"LEFT JOIN System.Users AS FO ON FO.SystemID=G.FinancialOfficerID " & vbCrLf
	If AppID > 0 Then
		sql = sql & "WHERE I.AppID=" & PrepIntegerSQL(AppID)
	ElseIf GranteeID>0 Then
		sql = sql & "WHERE G.GranteeID=" & prepIntegerSQL(GranteeID)
	End If
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: No Grantee and Application record retrieved")
		Response.End
	Else
		AppID = rs.Fields("AppID")
		FiscalYear = rs.Fields("FiscalYear")
		GranteeName = rs.Fields("GranteeName")
		AuthorizedOfficial = rs.Fields("AuthorizedOfficial")
		ProgramDirector = rs.Fields("ProgramDirector")
		ProgramDirectorTitle = rs.Fields("ProgramDirectorTitle")
		FinancialOfficer = rs.Fields("FinancialOfficer")
		FinancialOfficerTitle = rs.Fields("FinancialOfficerTitle")
		AuthorizedOfficialTitle = rs.Fields("AuthorizedOfficialTitle")
	End If
End If
%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Resolution</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width:80%; margin: auto;">

<h1>Motor Vehicle Crime Prevention Authority Resolution</h1>

<h2><%=FiscalYear %>&nbsp;<%=GranteeName %> Resolution</h2>

<h2><%=GrantProgramTitle %></h2>

 	 	 
<p>WHEREAS, under the provisions of the <%=Statute %> and Texas 
Administrative Code Title 43; Part 3; Chapter 57, entities are eligible to receive grants from 
the Motor Vehicle Crime Prevention Authority to provide financial support to law 
enforcement agencies for economic automobile theft enforcement teams and to combat motor 
vehicle burglary in the jurisdiction; and</p>
 
<p><%=para2%></p>
 
<p>WHEREAS, <%=GranteeName %> has agreed that in the event of loss or misuse of the grant funds,
<%=GranteeName %> assures that the grant funds will be returned in full to the Motor Vehicle Crime Prevention Authority.</p>
 
<p>NOW THEREFORE, BE IT RESOLVED and ordered that <%=AuthorizedOfficial %>, 
<%=AuthorizedOfficialTitle %>, is designated as the Authorized Official to apply for, accept, 
decline, modify, or cancel the grant application for the Motor Vehicle Crime Prevention Authority Grant Program and all other necessary documents to accept said grant; and</p>
 
<p>BE IT FURTHER RESOLVED that <%=ProgramDirector%>, <%=ProgramDirectorTitle %>, is designated 
as the Program Director and <%=FinancialOfficer %>, <%=FinancialOfficerTitle %>, is designated 
as the Financial Officer for this grant.</p>
 
<p>Adopted this ______day of ________________, <%=(FiscalYear) %>.</p>
 <br />	

<p>________________________________________<br />
<%=AuthorizedOfficial %><br />
<%=AuthorizedOfficialTitle %></p>






</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->