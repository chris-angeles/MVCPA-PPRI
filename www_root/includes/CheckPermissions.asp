<%
Function CheckPermissions(uid, gid, includeAPTPA)
	Dim vsql, vrs

	If uid = "" Or gid = "" Then
		CheckPermissions = False
		Exit Function
	End If	
	vsql = "EXEC System.spPermissions @SystemID=" & uid & ", @GranteeID=" & gid
	Set vrs = con.execute(vsql)
	'Response.Write("<!--" & vsql & "-->")
	'Response.Write("<pre>" & vsql & "</pre>")

	If vrs.bof = True Then
		Response.Write("Error obtaining permissions for user")
		Response.End
	End If

	If includeAPTPA = True and vrs.Fields("Developer") = True Then
		CheckPermissions = True
		Exit Function
	ElseIf includeAPTPA = True and vrs.Fields("MVCPAAdministrator") = True Then
		CheckPermissions = True
		Exit Function
	ElseIf includeAPTPA = True and vrs.Fields("MVCPAAuditor") = True Then
		CheckPermissions = True
		Exit Function
	ElseIf includeAPTPA = True and vrs.Fields("MVCPAGrantCoordinator") = True Then
		CheckPermissions = True
		Exit Function
	ElseIf includeAPTPA = True and vrs.Fields("MVCPAAdministrativeAssistant") = True Then
		CheckPermissions = True
		Exit Function
	ElseIf vrs.Fields("GranteePermission") = True Then
		CheckPermissions = True
		Exit Function
	Else
		CheckPermissions = False
		Exit Function
	End If	
End Function

Function CheckPermissionsWithLock(uid, gid, submitted)
	Dim vsql, vrs

	If uid = "" Or gid = "" Then
		CheckPermissionsWithLock = False
		Exit Function
	End If	
	vsql = "EXEC System.spPermissions @SystemID=" & uid & ", @GranteeID=" & gid
	Set vrs = con.execute(vsql)
	'Response.Write("<!--" & vsql & "-->")

	If vrs.bof = True Then
		Response.Write("Error obtaining permissions for user")
		Response.End
	End If

	If submitted = False and vrs.Fields("Developer") = True Then
		CheckPermissionsWithLock = True
		Exit Function
	ElseIf submitted = False and vrs.Fields("MVCPAAdministrator") = True Then
		CheckPermissionsWithLock = True
		Exit Function
	ElseIf submitted = False and vrs.Fields("MVCPAAuditor") = True Then
		CheckPermissionsWithLock = True
		Exit Function
	ElseIf submitted = False and vrs.Fields("MVCPAGrantCoordinator") = True Then
		CheckPermissionsWithLock = True
		Exit Function
	ElseIf submitted = False and vrs.Fields("MVCPAAdministrativeAssistant") = True Then
		CheckPermissionsWithLock = True
		Exit Function
	ElseIf submitted = False and vrs.Fields("GranteePermission") = True Then
		CheckPermissionsWithLock = True
		Exit Function
	Else
		CheckPermissionsWithLock = False
		Exit Function
	End If	
End Function

%>