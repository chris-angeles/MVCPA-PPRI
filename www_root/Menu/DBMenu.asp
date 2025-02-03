<%
Function displayDBMenu(vSystemID, vFiscalYear, vGranteeID)
	Dim vrs, vsql, vLastCategory
	vLastCategory=0
	Response.Write("<!-- Start of database generated menu. -->" & vbCrLf)
	Response.Write(vbTab & "<div id=""sidenavigation""><ul>" & vbCrLf)
	Response.Write("<li style=""text-align: center; color: black; "">FY" & (vFiscalYear mod 100) & " Selected</li>")
	vsql = "EXEC Menu.spMenu2 @SystemID=" & prepIntegerSQL(vSystemID) & ",  @FiscalYear=" & prepIntegerSQL(vFiscalYear) & ", @GranteeID=" & prepIntegerSQL(vGranteeID)
	If Debug = True Then
		Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	While vrs.EOF = False
		If vrs.Fields("CategoryAndLink") = True Then
			Response.Write("<li style=""border-left: 1px;"">" & Replace(vrs.Fields("anchortag"), "<a ", "<a style=""border-left: 1px;"" ") & "</li>" & vbCrLf)
		Else
			If vrs.Fields("Category") <> vLastCategory And vrs.Fields("CategoryAndLink")=False Then
				vLastCategory = vrs.Fields("Category")
				Response.Write("<li>" & vLastCategory & "</li>" & vbCrLf)
			End If
			Response.Write("<li>" & vrs.Fields("anchortag") & "</li>" & vbCrLf)
		End If
		vrs.MoveNext()
	Wend
	Response.Write(vbTab & "</ul></div>" & vbCrLf)
	Response.Write("<!-- End of database generated menu. -->" & vbCrLf)
End Function
%>