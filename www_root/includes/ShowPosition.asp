<%
Sub ShowPosition(vPositionTitle, vSystemID, vPermitEdit)
	Dim vsql, vrs

	Response.Write("<tr title=""" & PositionDescription(vPositionTitle) & """>" & vbCrLf)
	Response.Write("<td style=""vertical-align: top; font-weight: bold; "">" & vPositionTitle & "</td>" & vbCrLf)
	If IsNull(vSystemID)=True Or vSystemID=0 Then
		Response.Write("<td>Position not filled</td>")
		vSystemID=0
	Else
		vsql = "SELECT SystemID, Name, Title, Email, Address1, Address2, City, State, ZIP, phone, fax, Mobile " & vbCrLF & _
			"FROM System.Users WITH (NOLOCK) " & vbCrLF & _
			"WHERE SystemID=" & prepIntegerSQL(vSystemID)
		If Debug = True Then
			Response.Write("<pre>" & vsql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		Set vrs=Con.Execute(vsql)
		If vrs.EOF = False Then
			Response.Write("<td style=""padding-bottom: 10px"">" & vrs.Fields("Name"))
			If MVCPARights = True Then
				Response.Write("<a href=""../Contact/ContactItems.asp?ContactID=" & vSystemID & """ target=""_blank""><img src=""../images/document2_add.png"" />")
			End If
			Response.Write("<br>" & vbCrLf)
			If IsNull(vrs.Fields("Title"))= False Then
				Response.Write(vrs.Fields("Title") & "<br />" & vbCrLf)
			End If
			If IsNull(vrs.Fields("Address1"))= False Then
				Response.Write(vrs.Fields("Address1") & "<br />" & vbCrLf)
			End If
			If IsNull(vrs.Fields("City"))= False Then
				Response.Write(vrs.Fields("City") & ", " & vrs.Fields("State") & " " & vrs.Fields("zip") & "<br />" & vbCrLf)
			End If
			If IsNull(vrs.Fields("Phone"))= False Then
				Response.Write("Phone: " & vrs.Fields("Phone") & "<br />" & vbCrLf)
			End If
			If IsNull(vrs.Fields("Mobile"))= False Then
				Response.Write("Mobile: " & vrs.Fields("Mobile") & "<br />" & vbCrLf)
			End If
			If IsNull(vrs.Fields("Email"))= False Then
				Response.Write("<a href=""mailto:" & vrs.Fields("Email") & "?subject=TxMVCPA"" class=""plainlink"">" & vrs.Fields("Email") & "</a>" & vbCrLf)
			End If
		End If
		Response.Write("</td>" & vbCrLf)
	End If
	If vPermitEdit = True Then
		If vSystemID=0 Then
			Response.Write("<td><input type=""button"" value=""Add"" style=""width: 70px; "" onclick=""changeOfficial('" & _
				vPositionTitle & "', " & vSystemID & ");""></td>" & vbCrLf)
		Else
			Response.Write("<td><input type=""button"" value=""Change"" style=""width: 70px; "" onclick=""changeOfficial('" & _
				vPositionTitle & "', " & vSystemID & ");""></td>" & vbCrLf)
		End If
	Else
		Response.Write("<td></td>" & vbCrLF)
	End If
	Response.Write("</tr>" & vbCrLf)
End Sub

Function PositionDescription(vPosition)
	If vPosition = "Authorized Official" Then
		PositionDescription = "The Authorized Official is a County Judge or Mayor or appointed official when clearly authorized to enter into agreements on behalf of the governing body. The Authorized Official must be authorized to apply for, accept, decline, modify, or cancel the grant for the applicant agency."
	ElseIf vPosition = "Program Director" Then
		PositionDescription = "The Program Director must be an officer or employee responsible for the program operation or monitoring and who will serve as the point-of-contact regarding the program's day-to-day operations." 
	ElseIf vPosition = "Program Manager" Then
		PositionDescription = "The Program Manager is designated by governing authority or program director and responsible for: day-to-day operation, monitoring of the grant, act as the Program Director in his/her absence, responsible for record keeping, reviewing and approving financial expenditures, maintaining program files, approving required program reports, evaluating the program, responding to APTPA monitor reports, etc."
	ElseIf vPosition = "Financial Officer" Then
		PositionDescription = "The Financial Officer must be the County Auditor or city or agency chief financial officer. The financial officer may not serve as the program director or the authorized official."
	ElseIf vPosition = "Program Administrative Contact" Then
		PositionDescription = "The Program Administrative Contact is an optional secondary contact for program director and program manager."
	ElseIf vPosition = "Financial Administrative Contact" Then
		PositionDescription = "The Financial Administrative Contact is an optional secondary contact for the financial officer."
	ElseIf vPosition = "Taskforce Commander" Then
		PositionDescription = "Taskforce Commander is the primary licensed law enforcement person in a MVCPA funded program responsible for coordinating with other MVCPA taskforces or law enforcement agencies to respond to requests for information or tactical assistance. This person is usually the Program Director or Program Manager."
	ElseIf vPosition = "PIO / Media Contact" Then
		PositionDescription = "PIO / Media Contact is the person to contact for public information regarding the grantee or Taskforce. The public should be referred to this person for information."
	Else
		Response.Write("Error: Invalid position title")
		Response.End
	End If
End Function
%>