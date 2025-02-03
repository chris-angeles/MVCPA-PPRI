<%
Function getApplicationSchema(vFiscalYear)
	If vFiscalYear = 2018 Then
		getApplicationSchema = "Negotiation"
	ElseIf vFiscalYear = 2019 Then
		getApplicationSchema = "Application"
	ElseIf vFiscalYear = 2020 Then
		getApplicationSchema = "Negotiation"
	ElseIf vFiscalYear = 2021 Then
		getApplicationSchema = "Negotiation"
	ElseIf vFiscalYear = 2022 Then
		getApplicationSchema = "Negotiation"
	ElseIf vFiscalYear = 2025 Then ' 2025 decided to be a negotiation year because of additional funding.
		getApplicationSchema = "Negotiation"
	ElseIf vFiscalYear Mod 2 = 0 And vFiscalYear > 2017 Then
		getApplicationSchema = "Negotiation"
	Else
		getApplicationSchema = "Application"
	End If
End Function

Function getCCApplicationSchema(vFiscalYear)
	If vFiscalYear = 2025 Then
		'getCCApplicationSchema = "Application"
		getCCApplicationSchema = "Negotiation" ' Change after Negotiation process starts.
	ElseIf vFiscalYear = 2024 Then
		'getCCApplicationSchema = "Application"
		getCCApplicationSchema = "Negotiation" ' Change after Negotiation process starts.
	Else
		getCCApplicationSchema = "Application"
	End If
End Function
%>