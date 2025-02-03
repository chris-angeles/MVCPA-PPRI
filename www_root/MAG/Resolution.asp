<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"--><% 
Dim debug, i, j, MAGID, GranteeID, FiscalYear, GranteeName, AuthorizedOfficial, AuthorizedOfficialTitle, _
	ProgramDirector, ProgramDirectorTitle, FinancialOfficer, FinancialOfficerTitle, GrantProgramTitle
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

If Len(Request.Form("GranteeID")) > 0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
ElseIf Len(Request.QueryString("GranteeID")) > 0 Then
	GranteeID = CInt(Request.QueryString("GranteeID"))
Else
	Response.Write("Error: No GranteeID Specified")
	SendMessage "Error: No GranteeID Specified"
	Response.End
End If

If Len(Request.Form("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.Form("FiscalYear"))
ElseIf Len(Request.QueryString("FiscalYear")) > 0 Then
	FiscalYear = CInt(Request.QueryString("FiscalYear"))
Else
	FiscalYear = 2023
End If

GrantProgramTitle = "Motor Vehicle Crime Prevention Authority Auxiliary Grant Program"

sql = "SELECT G.GranteeID, G.GranteeName, ISNULL(A.FiscalYear," & FiscalYear & ") AS FiscalYear, " & vbCrLf & _
	"	AO.Name AS AuthorizedOfficial, AO.Title AS AuthorizedOfficialTitle, " & vbCrLf & _
	"	PD.Name AS ProgramDirector, PD.Title AS ProgramDirectorTitle, " & vbCrLf & _
	"	FO.Name AS FinancialOfficer, FO.Title AS FinancialOfficerTitle, A.MAGID AS MAGID " & vbCrLF & _
	"FROM Grantees AS G " & vbCrLF & _
	"LEFT JOIN MAG.Main AS A ON A.GranteeID=G.GranteeID " & " AND FiscalYear=" & prepIntegerSQL(FiscalYear) & " " & vbCrLf & _
	"LEFT JOIN System.Users AS AO ON AO.SystemID=G.AuthorizedOfficialID " & vbCrLf & _
	"LEFT JOIN System.Users AS PD ON PD.SystemID=G.ProgramDirectorID " & vbCrLf & _
	"LEFT JOIN System.Users AS FO ON FO.SystemID=G.FinancialOfficerID " & vbCrLf & _
	"WHERE G.GranteeID=" & PrepIntegerSQL(GranteeID) & " "

If Debug = True Then
	Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
	Response.Flush
End If
Set rs = Con.Execute(sql)
If rs.EOF = True Then
	Response.Write("Error: No Grantee and Application record retrieved")
	Response.End
Else
	MAGID = rs.Fields("MAGID")
	FiscalYear = rs.Fields("FiscalYear")
	GranteeName = rs.Fields("GranteeName")
	AuthorizedOfficial = rs.Fields("AuthorizedOfficial")
	ProgramDirector = rs.Fields("ProgramDirector")
	ProgramDirectorTitle = rs.Fields("ProgramDirectorTitle")
	FinancialOfficer = rs.Fields("FinancialOfficer")
	FinancialOfficerTitle = rs.Fields("FinancialOfficerTitle")
	AuthorizedOfficialTitle = rs.Fields("AuthorizedOfficialTitle")
End If

If IsNull(ProgramDirector) = True Then
	ProgramDirector="{Program Director}"
End If
If IsNull(AuthorizedOfficial) = True Then
	AuthorizedOfficial = "{Authorized Official}"
End If
If IsNull(FinancialOfficer) = True Then
	FinancialOfficer = "{Financial Officer}"
End If
If IsNull(AuthorizedOfficialTitle) = True Then
	AuthorizedOfficialTitle = "Authorized Official Title"
End If
If IsNull(ProgramDirectorTitle) = True Then
	ProgramDirectorTitle = "{Program Director Title}"
End If
If IsNull(FinancialOfficerTitle) = True Then
	FinancialOfficerTitle = "{Financial Officer Title}"
End If

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Resolution</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="width:80%; margin: auto;">

<h1>Motor Vehicle Crime Prevention Authority</h1>

<h2><%=FiscalYear %>&nbsp;<%=GranteeName %> Resolution</h2>

<h2><%=GrantProgramTitle %></h2>

 	 	 
<p>WHEREAS, under the provisions of the Texas Transportation Code Chapter 1006 and Texas 
Administrative Code Title 43; Part 3; Chapter 57, entities are eligible to receive grants from 
the Motor Vehicle Crime Prevention Authority to provide financial support to law 
enforcement agencies for economic motor vehicle theft enforcement teams; and</p>
 
<p>WHEREAS, this grant program will assist this jurisdiction to combat motor vehicle theft, 
motor vehicle burglary and fraud-related motor vehicle crime; and </p>
 
<p>WHEREAS, <%=GranteeName %> has agreed that in the event of loss or misuse of the grant funds,
<%=GranteeName %> assures that the grant funds will be returned in full to the Motor Vehicle Crime 
Prevention Authority.</p>
 
<p>NOW THEREFORE, BE IT RESOLVED and ordered that <%=AuthorizedOfficial %>, 
<%=AuthorizedOfficialTitle %>, is designated as the Authorized Official to apply for, accept, 
decline, modify, or cancel the grant application for the Motor Vehicle Crime Prevention 
Authority Grant Program and all other necessary documents to accept said grant; and</p>
 
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