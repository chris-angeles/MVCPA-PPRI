<%
Sub ShowDashboard(vFiscalYear)
	Dim vsql, vrs, vColumns, vi
	Debug = False
	Response.Write("<!-- Start Dashboard.asp include file -->" & vbCrLf)
	vsql = "SELECT /*Grantee_ID AS ID,*/ Grantee_Name, /*Grant_ID,*/ " & vbCRLF & _
	"	Program_Name, Fiscal_Year, Award, Match, Program_Total, " & vbCrLf & _
	"	Cash_Expenditure_Total, Reimbursable, Reimbursed, " & vbCrLf & _
	"	[ER Q1], [ER Q2], [ER Q3], [ER Q4], Adj, " & vbCrLf & _
	"	[PR Q1], [PR Q2], [PR Q3], [PR Q4], [PR YE], IC_Status, GranteeSort  " & vbCrLf & _
	"FROM [Grants].vwGrantStatus " & vbCrLf & _
	"WHERE GrantClassID=1 AND Fiscal_Year=" & FiscalYear & " " & vbCrLf & _
	"ORDER BY Fiscal_Year, GranteeSort" & vbCrLf

	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If

	Response.Write("<br />" & vbCrLf)
	Response.Write("<div style=""width: 100%; text-align: center; font-size: small; font-weight: bold; "">Taskforce Grants</div>" & vbCrLf)

	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		vColumns = vrs.Fields.Count
		Response.Write("<table style=""margin: auto; font-size: x-small; "">" & vbCrLf)
		Response.Write("<thead><tr style=""vertical-align: bottom; "">" & vbCrLf)
		For vi = 0 to vColumns - 2
			Response.Write(vbTab & "<th>" & Replace(vrs.Fields(vi).Name, "_", " ") & "</th>" & vbCrLf)
		Next
		Response.Write("</tr></thead>" & vbCrLf)
		Response.Write("<tbody>" & vbCrLf)
		While vrs.EOF = False
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			For vi = 0 to vColumns - 2
				If IsNull(vrs.Fields(vi).value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				ElseIf vrs.Fields(vi).Type = vbCurrency Then
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatcurrency(vrs.Fields(vi).value, 2, true, true, true) & "</td>" & vbCrLf)
				ElseIf vrs.Fields(vi).Type = vbBoolean Then
					If vrs.Fields(vi).value = true Then
						Response.Write(vbTab & "<td style=""text-align: center; "">X</td>" & vbCrLf)
					Else
						Response.Write(vbTab & "<td></td>" & vbCrLf)
					End If
				Else
					Response.Write(vbTab & "<td>" & vrs.Fields(vi).value & "</td>" & vbCrLf)
				End If
			Next
			Response.Write("</tr>" & vbCrLf)
			vrs.MoveNext()
		Wend
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("<tfoot>" & vbCrLF)
		Response.Write("</table>" & vbCrLF)

		If Debug = True Then
			Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
			Response.Flush
		End If

		If FiscalYear>2023 Then

			Response.Write("<br />" & vbCrLf)
			Response.Write("<div style=""width: 100%; text-align: center; font-size: small; font-weight: bold; "">Catalytic Converter Grants</div>" & vbCrLf)

			vsql = "SELECT /*Grantee_ID AS ID,*/ Grantee_Name, /*Grant_ID,*/ " & vbCRLF & _
				"	Program_Name, Fiscal_Year, Award, Match, Program_Total, " & vbCrLf & _
				"	Cash_Expenditure_Total, Reimbursable, Reimbursed, " & vbCrLf & _
				"	[ER Q1], [ER Q2], [ER Q3], [ER Q4], Adj, " & vbCrLf & _
				"	[PR Q1], [PR Q2], [PR Q3], [PR Q4], [PR YE], IC_Status, GranteeSort  " & vbCrLf & _
				"FROM [Grants].vwGrantStatus " & vbCrLf & _
				"WHERE GrantClassID=4 AND Fiscal_Year=" & FiscalYear & " " & vbCrLf & _
				"ORDER BY Fiscal_Year, GranteeSort" & vbCrLf
			If Debug = True Then
				Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
				Response.Flush
			End If

			Set vrs = Con.Execute(vsql)
			If vrs.EOF = False Then
				vColumns = vrs.Fields.Count
				Response.Write("<table style=""margin: auto; font-size: x-small; "">" & vbCrLf)
				Response.Write("<thead><tr style=""vertical-align: bottom; "">" & vbCrLf)
				For vi = 0 to vColumns - 2
					Response.Write("<th>" & Replace(vrs.Fields(vi).Name, "_", " ") & "</th>" & vbCrLf)
				Next
				Response.Write("</tr></thead>" & vbCrLf)
				Response.Write("<tbody>" & vbCrLf)
				While vrs.EOF = False
					Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
					For vi = 0 to vColumns - 2
						If IsNull(vrs.Fields(vi).value) = True Then
							Response.Write(vbTab & "<td></td>" & vbCrLf)
						ElseIf vrs.Fields(vi).Type = vbCurrency Then
							Response.Write(vbTab & "<td style=""text-align: right; "">" & formatcurrency(vrs.Fields(vi).value, 2, true, true, true) & "</td>" & vbCrLf)
						ElseIf vrs.Fields(vi).Type = vbBoolean Then
							If vrs.Fields(vi).value = true Then
								Response.Write(vbTab & "<td style=""text-align: center; "">X</td>" & vbCrLf)
							Else
								Response.Write(vbTab & "<td></td>" & vbCrLf)
							End If
						Else
							Response.Write(vbTab & "<td>" & vrs.Fields(vi).value & "</td>" & vbCrLf)
						End If
					Next
					Response.Write("</tr>" & vbCrLf)
					vrs.MoveNext()
				Wend
				Response.Write("</tbody>" & vbCrLf)
				Response.WRite("</table>" & vbCrLF)
				Response.WRite("</tfoot>" & vbCrLF)
			End If
		End If
		Response.Write("<table style=""margin: auto; font-size: x-small; "">" & vbCrLf)
		Response.Write(vbTab & "<tr><td colspan=""" & vColumns & """ style=""text-align: center; "">Q1, Q2, Q3, Q4 Expenditure Report Status: <strong>S</strong> Submitted, <strong>A</strong> Approved, <strong>P</strong> Paid</td></tr>" & vbCrLf)
		Response.Write(vbTab & "<tr><td colspan=""" & vColumns & """ style=""text-align: center; "">Adj: The Adjustment ID of first adjustment requiring an approval.</td></tr>" & vbCrLf)
		Response.Write(vbTab & "<tr><td colspan=""" & vColumns & """ style=""text-align: center; "">Q1, Q2, Q3, Q4 Progress Report Status: <strong>S</strong> Submitted, <strong>A</strong> Approved.</td></tr>" & vbCrLf)
		Response.Write(vbTab & "<tr><td colspan=""" & vColumns & """ style=""text-align: center; "">Q1, Q2, Q3, Q4 Inventory Certification Status: <strong>S</strong> Submitted, <strong>A</strong> Approved.</td></tr>" & vbCrLf)
		Response.WRite("</table>" & vbCrLF)
	End If

	vsql = "SELECT A.GrantID AS [Grant ID], A.FiscalYear AS [Fiscal Year], B.GranteeName AS [Grantee], A.ProgramName AS [Program Name], 'No Grant Number' AS Issue, B.GranteeNameSort " & vbCrLf & _
		"FROM Grants.Main AS A " & vbCrLf & _
		"JOIN Grantees AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
		"WHERE GrantNumber IS NULL" & vbCrLf & _
		"UNION " & vbCrLf & _
		"SELECT A.GrantID AS [Grant ID], A.FiscalYear AS [Fiscal Year], B.GranteeName AS [Grantee], A.ProgramName AS [Program Name], 'No Award Amount' AS Issue, B.GranteeNameSort " & vbCrLf & _
		"FROM Grants.Main AS A " & vbCrLf & _
		"JOIN Grantees AS B ON B.GranteeID=A.GranteeID " & vbCrLf & _
		"WHERE ISNULL(A.AwardAmount,0.0)=0.0 " & vbCrLf & _
		"ORDER BY [Fiscal Year], [GranteeNameSort] "
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If

	Set vrs = Con.Execute(vsql)
	If vrs.EOF = False Then
		vColumns = vrs.Fields.Count
		Response.Write("<br />" & vbCrLf)
		Response.Write("<table style=""margin: auto; font-size: x-small; "">" & vbCrLf)
		Response.Write("<caption style=""font-size: small; "">Detected Issues</caption>" & vbCrLf)
		Response.Write("<thead><tr style=""vertical-align: bottom; "">" & vbCrLf)
		For vi = 0 to vColumns - 2
			Response.Write(vbTab & "<th>" & Replace(vrs.Fields(vi).Name, "_", " ") & "</th>" & vbCrLf)
		Next
		Response.Write("</tr></thead>" & vbCrLf)
		Response.Write("<tbody>" & vbCrLf)
		While vrs.EOF = False
			Response.Write("<tr style=""vertical-align: top; "">" & vbCrLf)
			For vi = 0 to vColumns - 2
				If IsNull(vrs.Fields(vi).value) = True Then
					Response.Write(vbTab & "<td></td>" & vbCrLf)
				ElseIf vrs.Fields(vi).Name = "Issue" Then
					Response.Write(vbTab & "<td style=""text-align: left; color: red; "">" & vrs.Fields(vi).value & "</td>" & vbCrLf)
				ElseIf vrs.Fields(vi).Type = vbCurrency Then
					Response.Write(vbTab & "<td style=""text-align: right; "">" & formatcurrency(vrs.Fields(vi).value, 2, true, true, true) & "</td>" & vbCrLf)
				ElseIf vrs.Fields(vi).Type = vbBoolean Then
					If vrs.Fields(vi).value = true Then
						Response.Write(vbTab & "<td style=""text-align: center; "">X</td>" & vbCrLf)
					Else
						Response.Write(vbTab & "<td></td>" & vbCrLf)
					End If
				Else
					Response.Write(vbTab & "<td>" & vrs.Fields(vi).value & "</td>" & vbCrLf)
				End If
			Next
			Response.Write("</tr>" & vbCrLf)
			vrs.MoveNext()
		Wend
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("<tfoot>" & vbCrLF)
		Response.Write("</table>" & vbCrLF)

	End If

	Response.Write("<!-- End Dashboard.asp include file -->" & vbCrLf)
End Sub
%>