<%
Function PRVersion(vGrantClassID, vFiscalYear)
	If vGrantClassID=4 AND vFiscalYear >=2024 Then
		PRVersion = 1001
	ElseIf vFiscalYear>2021 Then 
		PRVersion = 5
	ElseIf FiscalYear>2020 Then
		PRVersion = 4
	ElseIf FiscalYear>2019 Then
		PRVersion = 3
	Else
		PRVersion = 2
	End If
End Function
%>