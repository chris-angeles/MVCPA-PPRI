<%
function getHomeApplicationReferenceByAppID(vAppID)
	dim vsql, vrs, vcon, vGrantClassID, vFiscalYear
	vsql = "SELECT AppID, GrantClassID, FiscalYear FROM Application.IDs WHERE AppID=" & PrepIntegerSQL(vAppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = True Then
		getHomeApplicationReferenceByAppID = "/Main/Main.asp"	
	Else
		vGrantClassID = vrs.Fields("GrantClassID")
		vFiscalYear = vrs.Fields("FiscalYear")
		If vGrantClassID=1 And vFiscalYear<2022 Then
			getHomeApplicationReferenceByAppID = "/Application/Application.asp?AppID=" & vAppID
		Else
			getHomeApplicationReferenceByAppID = getHomeApplicationReferenceByGrantClass(vGrantClassID, vAppID)
		End If
	End If
End Function

Function getHomeApplicationReferenceByGrantClass(vGrantClassID, vAppID)
	IF vGrantClassID=1 Then
		getHomeApplicationReferenceByGrantClass = "/Application/TFGApplication.asp?AppID=" & vAppID
	ElseIf vGrantClassID = 4 Then
		getHomeApplicationReferenceByGrantClass = "/CatalyticConverter/Application.asp?AppID=" & vAppID
	End If
End Function

function getHomeNegotiationReferenceByAppID(vAppID)
	dim vsql, vrs, vcon, vGrantClassID, vFiscalYear
	vsql = "SELECT AppID, GrantClassID, FiscalYear FROM Application.IDs WHERE AppID=" & PrepIntegerSQL(vAppID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set vrs = Con.Execute(vsql)
	If vrs.EOF = True Then
		getHomeNegotiationReferenceByAppID = "/Main/Main.asp"	
	Else
		vGrantClassID = vrs.Fields("GrantClassID")
		vFiscalYear = vrs.Fields("FiscalYear")
		If vGrantClassID=1 And vFiscalYear<2022 Then
			getHomeNegotiationReferenceByAppID = "/Negotiation/Application.asp?AppID=" & vAppID
		Else
			getHomeNegotiationReferenceByAppID = getHomeNegotiationReferenceByGrantClass(vGrantClassID, vAppID)
		End If
	End If
End Function

Function getHomeNegotiationReferenceByGrantClass(vGrantClassID, vAppID)
	IF vGrantClassID=1 Then
		getHomeNegotiationReferenceByGrantClass = "/Negotiation/TFGApplication.asp?AppID=" & vAppID
	ElseIf vGrantClassID = 4 Then
		getHomeNegotiationReferenceByGrantClass = "/CatalyticConverter/Negotiation.asp?AppID=" & vAppID
	End If
End Function

%>