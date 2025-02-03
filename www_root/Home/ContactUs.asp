<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j
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

%><!DOCTYPE html>
<html lang="en-us">
<head>
<title>MVCPA Contact Us</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body>
<div class="header" title="MVCPA logo banner. Outline of a car with eyes below and text Watch Your Car"></div>

<div class="pagetag">The <strong>Motor Vehicle Crime Prevention Authority</strong> (MVCPA) 
	awards financial grants to agencies, organizations, and concerned parties in an effort to 
	raise public awareness of vehicle theft and burglary and implement education and prevention 
	initiatives.</div>

<div class="menu"><%=displayDBMenu(UserSystemID, UserFiscalYear, UserGranteeID) %></div>

<div class="content">

<h1>Contact Us</h1>

<p>For technical problems with this website:<br />

<a href="mailto: mvcpa@tamu.edu?subject=MVCPA">Gourov Singla</a><br />
	Technology Services - Arts & Sciences<br />
	Texas A&M University<br />
	</p>

<p style="font-weight: bold; ">Always send grant management questions or comments to 
	<a href="mailto:grantsMVCPA@txdmv.gov">grantsMVCPA@txdmv.gov</a>.  
	The following is provided to help facilitate communication:</p>

<p>Questions regarding general Motor Vehicle Crime Prevention Authority (MVCPA) Division, 
	reports, publications, or information about MVCPA Board activities and meetings contact:<br />
	<a href="mailto:Yessenia.Benavides@txdmv.gov">Yessenia Benavides</a> - MVCPA Board Data and Reporting Specialist,
	512-465-4011</p>

<p>Questions about law enforcement (LE) programs and assistance with LE agency coordination, 
	TCOLE approved training; taskforce grant operations, Rapid Response Strikeforce, 
	and MVCPA Accessory grants contact:<br /> 
	<a href="mailto: Joe.Canady@txdmv.gov">Joe Canady</a>, 
	512-465-1383</p>

<p>Questions about grant activity and progress reports, inventory, obtaining MVCPA printed 
	literature / promotional material, establishing service mark agreements, and communication 
	of MVCPA messaging contact:<br />
	<a href="mailto:Gresham.Kay@txdmv.gov">Gresham Kay</a>, 
	512-465-1408</p>

<p>Questions regarding grant payments, reasonable and allowable expenses, expenditure report 
	questions and Comprehensive Annual Financial Review (CAFR) reports contact:<br />
	<a href="mailto: Daniel.price@txdmv.gov">Dan Price</a>, MVCPA Grant Auditor<br />
	512-465-1486</p>

<p>To report a criminal issue or civil action involving the local grant program contact:<br /> 
	<a href="mailto:David.Richards@txdmv.gov">David Richards</a>, MVCPA General Counsel. 
	512-465-1423 (send copy to MVCPA Director)</p>

<p>About the MVCPA Board or program operation:<br />
	<a href="mailto: Joe.Canady@txdmv.gov">Joe Canady</a>, MVCPA Interim Director<br />
	512-465-1383</p>

<p>See more information about MVCPA board and reports issued: 
	<a href="https://www.txdmv.gov/about-us/MVCPA" target="_blank">https://www.txdmv.gov/about-us/MVCPA</a></p>

<p>See more information about MVCPA grantees:  
<a href="https://www.txdmv.gov/mvcpa-grantees" target="_blank">https://www.txdmv.gov/mvcpa-grantees</a></p>

<p>See more information about MVCPA crime prevention: 
<a href="www.txwatchyourcar.com" target="_blank">www.txwatchyourcar.com</a></p>



</div>

<div class="clearfix"></div>
<div class="footer">TxDMV - MVCPA, ppri.tamu.edu &copy; 2017</div>
</body>
</html>
<!--#include file="../Menu/DBMenu.asp"-->
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->