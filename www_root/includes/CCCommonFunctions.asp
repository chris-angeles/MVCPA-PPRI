<%
Function getRFAReference(vFiscalYear)
	If vFiscalYear = 2025 Then
		' needs to be updated....
		getRFAReference = "https://www.txdmv.gov/sites/default/files/body-files/FY%202024%20SB%20224%20Catalytic%20Converter%20Grant%20RFA.pdf"
	ElseIf vFiscalYear = 2024 Then
		getRFAReference = "https://www.txdmv.gov/sites/default/files/body-files/FY%202024%20SB%20224%20Catalytic%20Converter%20Grant%20RFA.pdf"
	Else
		getRFAReference = ""
	End If
End Function

Function getRFALink(vFiscalYear)
	Dim vURL
	vURL = getRFAReference(vFiscalYear)
	If vURL = "" Then
		getRFALink = "<a href=""JavaScript: alert('Reference to document is not available');"">Request for Application (RFA)</a>"
	Else
		getRFALink = "<a href=""" & getRFAReference(FiscalYear) & """ target=""_blank"">Request for Application (RFA)</a>"
	End If
End Function

Function getDeadline(vFiscalYear, vApplicationSchema)
	If vApplicationSchema = "Application" Then
		If vFiscalYear = 2024 Then
			getDeadline = CDate("05/11/2024 11:59:59 PM")
		ElseIf vFiscalYear = 2025 Then
			getDeadline = CDate("11/12/2024 11:59:59 PM")
		End If
	ElseIf vApplicationSchema = "Negotiation" Then
		If vFiscalYear = 2024 Then
			getDeadline = CDate("09/30/2024 11:59:59 PM")
		ElseIf vFiscalYear = 2025 Then
			getDeadline = CDate("02/28/2025 11:59:59 PM")
		End If
	End If
End Function

%>