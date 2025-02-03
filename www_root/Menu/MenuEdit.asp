<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, copy, objFSO, RootFolder, objFolder, objSubFolder, _
	found, objFiles, objFile, savechanges, _
	ItemID, Page, Directory, MenuText, MenuDescription, CategoryID, _
	ItemSort, StartFiscalYear, EndFiscalYear, _
	ExternalLink, CategoryAndLink, NewWindow, PermissionLevelID, LinkID, _
	GranteeRequired, GrantRequired,  TaskforceGrantee, AuxiliaryGrantee, CCGrantee, _
	ISARequired, AppRequired, NegotiationRequired, MAGRequired, RRSRequired, CCRequired, _
	GranteeLink, GrantLink, ISALink, AppLink, NegotiationLink, RRSLink, CCLink, _
	SendGranteeID, SendGrantID, SendISAID, SendNegotiationID, _
	SendAppID, SendMAGID, SendFiscalYear, Inactive, Notes
debug = False

If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLf)
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
	Response.Write("</pre>" & vbCrLf)
End If

If Developer = True Then
	PermitEdit = True
Else
	PermitEdit = False
	Response.Write("You do not have permission to access this page")
	Response.End
End If

If Len(Request.Form("ItemID"))>0 Then
	ItemID = CInt(Request.Form("ItemID"))
ElseIf Len(Request.QueryString("ItemID"))>0 Then
	ItemID = CInt(Request.QueryString("ItemID"))
End If

If Request.Form.Count > 0 Then
	copy = false
	Page = Request.Form("Page")
	Directory = Request.Form("Directory")
	MenuText = Request.Form("MenuText")
	MenuDescription = Request.Form("MenuDescription")
	CategoryID = Request.Form("CategoryID")
	ItemSort = Request.Form("ItemSort")
	StartFiscalYear = Request.Form("StartFiscalYear")
	EndFiscalYear = Request.Form("EndFiscalYear")
	If EndFiscalYear = "0" Then
		EndFiscalYear = null
	ElseIf Len(EndFiscalYear)>0 Then
		EndFiscalYEar = CInt(EndFiscalYear)
	End If
	ExternalLink = Request.Form("ExternalLink")
	CategoryAndLink = Request.Form("CategoryAndLink")
	NewWindow = Request.Form("NewWindow")
	PermissionLevelID = Request.Form("PermissionLevelID")
	LinkID = Request.Form("LinkID")
	If Len(LinkID)=0 Then
		LinkID=0
	Else
		LinkID=CInt(LinkID)
	End If
	GranteeRequired = Request.Form("GranteeRequired")
	GrantRequired = Request.Form("GrantRequired")
	MAGRequired = Request.Form("MAGRequired")
	RRSRequired = Request.Form("RRSRequired")
	CCRequired = Request.Form("CCRequired")
	TaskforceGrantee = Request.Form("TaskforceGrantee")
	AuxiliaryGrantee = Request.Form("AuxiliaryGrantee")
	CCGrantee = Request.Form("CCGrantee")
	ISARequired = Request.Form("ISARequired")
	AppRequired = Request.Form("AppRequired")
	NegotiationRequired = Request.Form("NegotiationRequired")
	GranteeLink = Request.Form("GranteeLink")
	GrantLink = Request.Form("GrantLink")
	ISALink = Request.Form("ISALink")
	AppLink = Request.Form("AppLink")
	RRSLink = Request.Form("RRSLink")
	CCLink = Request.Form("CCLink")
	NegotiationLink = Request.Form("NegotiationLink")
	SendGranteeID = Request.Form("SendGranteeID")
	SendGrantID = Request.Form("SendGrantID")
	SendISAID = Request.Form("SendISAID")
	SendAppID = Request.Form("SendAppID")
	SendNegotiationID = Request.Form("SendNegotiationID")
	SendMAGID = Request.Form("SendMAGID")
	SendFiscalYear = Request.Form("SendFiscalYear")
	Inactive = Request.Form("Inactive")
	Notes = Request.Form("Notes")
	savechanges = Request.Form("savechanges")

	If savechanges = "1" Then
		If ItemID=0 Then
			sql = "INSERT INTO Menu.Items (Page, Directory, MenuText, MenuDescription, " & vbCrLf & _
				"	CategoryID, ItemSort, StartFiscalYear, EndFiscalYear, ExternalLink, " & vbCrLf & _
				"	CategoryAndLink, NewWindow, PermissionLevelID, LinkID, GranteeRequired, " & vbCrLf & _
				"	TaskforceGrantee, AuxiliaryGrantee,  CCGrantee, GrantRequired, " & vbCrLf & _
				"	ISARequired, AppRequired, NegotiationRequired, " & vbCrLF & _
				"	MAGRequired, RRSRequired, CCRequired, " & vbCrLf & _
				"	GranteeLink, GrantLink, ISALink, AppLink, NegotiationLink, RRSLink, CCLink, " & vbCrLf & _
				"	SendGranteeID, SendGrantID, SendISAID, SendAppID, SendNegotiationID, " & vbCrLf & _
				"	SendMAGID, SendFiscalYear, Inactive, Notes) " & vbCrLf & _
				"OUTPUT Inserted.ItemID " & vbCrLf & _
				"VALUES (" & prepStringSQL(Page) & ", " & vbCrLf & _
				prepStringSQL(Directory) & ", " & vbCrLf & _
				prepStringSQL(MenuText) & ", " & vbCrLf & _
				prepStringSQL(MenuDescription) & ", " & vbCrLf & _
				prepIntegerSQL(CategoryID) & ", " & vbCrLf & _
				prepIntegerSQL(ItemSort) & ", " & vbCrLf & _
				prepIntegerSQL(StartFiscalYear) & ", " & vbCrLf & _
				prepIntegerSQL(EndFiscalYear) & ", " & vbCrLf & _
				prepStringSQL(ExternalLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(CategoryAndLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(NewWindow) & ", " & vbCrLf & _
				prepIntegerSQL(PermissionLevelID) & ", " & vbCrLf & _
				prepIntegerSQL(LinkID) & ", " & vbCrLf & _
				prepBitRequiredSQL(GranteeRequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(TaskforceGrantee) & ", " & vbCrLf & _
				prepBitRequiredSQL(AuxiliaryGrantee) & ", " & vbCrLf & _
				prepBitRequiredSQL(CCGrantee) & ", " & vbCrLf & _
				prepBitRequiredSQL(GrantRequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(ISARequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(AppRequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(NegotiationRequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(MAGRequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(RRSRequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(CCRequired) & ", " & vbCrLf & _
				prepBitRequiredSQL(GranteeLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(GrantLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(ISALink) & ", " & vbCrLf & _
				prepBitRequiredSQL(AppLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(NegotiationLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(RRSLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(CCLink) & ", " & vbCrLf & _
				prepBitRequiredSQL(SendGranteeID) & ", " & vbCrLf & _
				prepBitRequiredSQL(SendGrantID) & ", " & vbCrLf & _
				prepBitRequiredSQL(SendISAID) & ", " & vbCrLf & _
				prepBitRequiredSQL(SendAppID) & ", " & vbCrLf & _
				prepBitRequiredSQL(SendNegotiationID) & ", " & vbCrLf & _
				prepBitRequiredSQL(SendMAGID) & ", " & vbCrLf & _
				prepBitRequiredSQL(SendFiscalYear) & ", " & vbCrLf & _
				prepBitRequiredSQL(Inactive) & ", " & vbCrLf & _
				prepStringSQL(Notes) & ") " & vbCrLf 
			If Debug = True Then
				Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			Set rs = Con.Execute(sql)
			If rs.EOF = False Then
				ItemID = rs.Fields("ItemID")
			Else
				Response.Write("Error retriving ItemID from Database.")
				Response.End
			End If

		Else
			sql = "UPDATE Menu.Items SET " & vbCrLf & _
				"Page=" & prepStringSQL(Page) & ", " & vbCrLf & _
				"Directory=" & prepStringSQL(Directory) & ", " & vbCrLf & _
				"MenuText=" & prepStringSQL(MenuText) & ", " & vbCrLf & _
				"MenuDescription=" & prepStringSQL(MenuDescription) & ", " & vbCrLf & _
				"CategoryID=" & prepIntegerSQL(CategoryID) & ", " & vbCrLf & _
				"ItemSort=" & prepIntegerSQL(ItemSort) & ", " & vbCrLf & _
				"StartFiscalYear=" & prepIntegerSQL(StartFiscalYear) & ", " & vbCrLf & _
				"EndFiscalYear=" & prepIntegerSQL(EndFiscalYear) & ", " & vbCrLf & _
				"ExternalLink=" & prepStringSQL(ExternalLink) & ", " & vbCrLf & _
				"CategoryAndLink=" & prepBitRequiredSQL(CategoryAndLink) & ", " & vbCrLf & _
				"NewWindow=" & prepBitRequiredSQL(NewWindow) & ", " & vbCrLf & _
				"PermissionLevelID=" & prepIntegerSQL(PermissionLevelID) & ", " & vbCrLf & _
				"LinkID=" & prepIntegerSQL(LinkID) & ", " & vbCrLf & _
				"GranteeRequired=" & prepBitRequiredSQL(GranteeRequired) & ", " & vbCrLf & _
				"TaskforceGrantee=" & prepBitRequiredSQL(TaskforceGrantee) & ", " & vbCrLf & _
				"AuxiliaryGrantee=" & prepBitRequiredSQL(AuxiliaryGrantee) & ", " & vbCrLf & _
				"CCGrantee=" & prepBitRequiredSQL(CCGrantee) & ", " & vbCrLf & _
				"GrantRequired=" & prepBitRequiredSQL(GrantRequired) & ", " & vbCrLf & _
				"ISARequired=" & prepBitRequiredSQL(ISARequired) & ", " & vbCrLf & _
				"AppRequired=" & prepBitRequiredSQL(AppRequired) & ", " & vbCrLf & _
				"NegotiationRequired=" & prepBitRequiredSQL(NegotiationRequired) & ", " & vbCrLf & _
				"MAGRequired=" & prepBitRequiredSQL(MAGRequired) & ", " & vbCrLf & _
				"RRSRequired=" & prepBitRequiredSQL(RRSRequired) & ", " & vbCrLf & _
				"CCRequired=" & prepBitRequiredSQL(CCRequired) & ", " & vbCrLf & _
				"GranteeLink=" & prepBitRequiredSQL(GranteeLink) & ", " & vbCrLf & _
				"GrantLink=" & prepBitRequiredSQL(GrantLink) & ", " & vbCrLf & _
				"ISALink=" & prepBitRequiredSQL(ISALink) & ", " & vbCrLf & _
				"AppLink=" & prepBitRequiredSQL(AppLink) & ", " & vbCrLf & _
				"NegotiationLink=" & prepBitRequiredSQL(NegotiationLink) & ", " & vbCrLf & _
				"RRSLink=" & prepBitRequiredSQL(RRSLink) & ", " & vbCrLf & _
				"CCLink=" & prepBitRequiredSQL(CCLink) & ", " & vbCrLf & _
				"SendGranteeID=" & prepBitRequiredSQL(SendGranteeID) & ", " & vbCrLf & _
				"SendGrantID=" & prepBitRequiredSQL(SendGrantID) & ", " & vbCrLf & _
				"SendISAID=" & prepBitRequiredSQL(SendISAID) & ", " & vbCrLf & _
				"SendAppID=" & prepBitRequiredSQL(SendAppID) & ", " & vbCrLf & _
				"SendNegotiationID=" & prepBitRequiredSQL(SendNegotiationID) & ", " & vbCrLf & _
				"SendMAGID=" & prepBitRequiredSQL(SendMAGID) & ", " & vbCrLf & _
				"SendFiscalYear=" & prepBitRequiredSQL(SendFiscalYear) & ", " & vbCrLf & _
				"Inactive=" & prepBitRequiredSQL(Inactive) & ", " & vbCrLf & _
				"Notes=" & prepStringSQL(Notes) & " " & vbCrLf & _
				"OUTPUT Inserted.ItemID " & vbCrLf & _
				"WHERE ItemID=" & prepIntegerSQL(ItemID)
			If Debug = True Then
				Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
				Response.Flush
			End If
			rs = Con.Execute(sql)
		End If
	End If
End If

If Request.QueryString.Count>0 Or SaveChanges=1 Then
	If Request.QueryString("copyitem") = "1" Then
		copy = true
	Else
		copy = false
	End If
	sql = "SELECT * " & vbCrLf & _
		"FROM Menu.Items AS I" & vbCrLf & _
		"WHERE ItemID=" & prepIntegerSQL(ItemID)
	Set rs=Con.Execute(sql)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

	If rs.EOF = False Then
		ItemID = rs.Fields("ItemID")
		Page = rs.Fields("Page")
		Directory = rs.Fields("Directory")
		MenuText = rs.Fields("MenuText")
		MenuDescription = rs.Fields("MenuDescription")
		CategoryID = rs.Fields("CategoryID")
		ItemSort = rs.Fields("ItemSort")
		StartFiscalYear = rs.Fields("StartFiscalYear")
		EndFiscalYear = rs.Fields("EndFiscalYear")
		ExternalLink = rs.Fields("ExternalLink")
		CategoryAndLink = rs.Fields("CategoryAndLink")
		NewWindow = rs.Fields("NewWindow")
		PermissionLevelID = rs.Fields("PermissionLevelID")
		LinkID = rs.Fields("LinkID")
		If IsNull(LinkID) Then
			LinkID = 0
		End If
		GranteeRequired = rs.Fields("GranteeRequired")
		TaskforceGrantee = rs.Fields("TaskforceGrantee")
		AuxiliaryGrantee = rs.Fields("AuxiliaryGrantee")
		CCGrantee = rs.Fields("CCGrantee")
		GrantRequired = rs.Fields("GrantRequired")
		ISARequired = rs.Fields("ISARequired")
		AppRequired = rs.Fields("AppRequired")
		NegotiationRequired = rs.Fields("NegotiationRequired")
		MAGRequired = rs.Fields("MAGRequired")
		RRSRequired = rs.Fields("RRSRequired")
		CCRequired = rs.Fields("CCRequired")
		GranteeLink = rs.Fields("GranteeLink")
		GrantLink = rs.Fields("GrantLink")
		ISALink = rs.Fields("ISALink")
		AppLink = rs.Fields("AppLink")
		NegotiationLink = rs.Fields("NegotiationLink")
		RRSLink = rs.Fields("RRSLink")
		CCLink = rs.Fields("CCLink")
		SendGranteeID = rs.Fields("SendGranteeID")
		SendGrantID = rs.Fields("SendGrantID")
		SendISAID = rs.Fields("SendISAID")
		SendAppID = rs.Fields("SendAppID")
		SendNegotiationID = rs.Fields("SendNegotiationID")
		SendMAGID = rs.Fields("SendMAGID")
		SendFiscalYear = rs.Fields("SendFiscalYear")
		Inactive = rs.Fields("Inactive")
		Notes = rs.Fields("Notes")
		If IsNull(MenuText) = True Then
			MenuText = ""
		End If
	Else
		ItemID = 0
		Page = ""
		Directory = ""
		MenuText = ""
		MenuDescription = ""
		CategoryID = 0
		ItemSort = 0
		StartFiscalYear = ""
		EndFiscalYear = ""
		ExternalLink = ""
		CategoryAndLink = False
		NewWindow = False
		PermissionLevelID = null
		LinkID = null
		GranteeRequired = False
		MAGRequired = False
		TaskforceGrantee = False
		AuxiliaryGrantee = False
		CCGrantee = False
		GrantRequired = False
		ISARequired = False
		AppRequired = False
		NegotiationRequired = False
		RRSRequired = False
		CCRequired = False
		GranteeLink = False
		GrantLink = False
		ISALink = False
		AppLink = False
		NegotiationLink = False
		RRSLink = False
		CCLink = False
		SendGranteeID = False
		SendGrantID = False
		SendISAID = False
		SendAppID = False
		SendNegotiationID = False
		SendMAGID = False
		SendFiscalYear = False
		Inactive = False
		Notes = ""
	End If
End If

%><!DOCTYPE html>
<html lang="en-us">
<title>Select Menu Item To Edit</title>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" /> 
<link rel="stylesheet" href="/styles/main.css" type="text/css" /> 
</head>
<body style="margin: 0px; width: 100%">

<div class="sectiontitle">Select Menu Item To Edit</div>

<%

%>
<form name="MenuEdit" method="post" action="MenuEdit.asp">
<input type="hidden" name="savechanges" value="1" />
<input type="hidden" name="ItemID" value="<%=ItemID %>" />
<table style="margin: auto">
	<tr>
		<td>ItemID:</td>
		<td><%=ItemID %></td>
	</tr>

	<tr>
		<td>Directory:</td>
		<td><select name="Directory" onchange="MenuEdit.savechanges.value='0';document.MenuEdit.submit();">
			<option value=""></option>
<%
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		RootFolder = Server.MapPath("..")
		Set objFolder = objFSO.GetFolder(RootFolder)
		For Each objSubFolder In objFolder.SubFolders
			REsponse.WRite("<!--" & objSubFolder.Name & "-->")
			Response.Write(vbTab & vbTab & vbTab & SelectOption("/" & objSubFolder.Name & "/", "/" & objSubFolder.Name & "/", Directory))
		Next
		Response.Write(vbTab & vbTab & vbTab & SelectOption("/", "/ (Root)", Directory))
%></select></td>
	</tr>

	<tr title="filename of target">
		<td>Page Name</td>
		<td><select name="Page">
			<option value=""></option>
<%

	If Len(Directory)>0 And Left(Directory,4)<>"http" Then
		Set objFolder = objFSO.GetFolder(Server.MapPath(Directory))
		'Response.Write(objFolder.Name)
		found = false
		Set objFiles = objFolder.Files
		For Each objFile in objFiles
			If Right(objFile.Name,4) = ".asp" or Right(objFile.Name,4) = ".htm" or Right(objFile.Name,5) = ".aspx" Then
				Response.Write(SelectOption(objFile.Name, objFile.Name, Page))
				If objFile.Name = Page Then
					Found = True
				End If
			End If
		Next
		If Found = False Then
			Response.Write(SelectOption(Page, Page & " (not found)", Page))
		End If
		Set objFSO = nothing
	End If
		%></select></td>
	</tr>

	<tr>
		<td>Category</td>
		<td><select name="CategoryID"><%
	If CategoryID>0 Then
		sql = "SELECT '<option value=""' + CAST(CategoryID AS VARCHAR) + '""' + CASE WHEN CategoryID=" & CategoryID & " THEN ' SELECTED' ELSE '' END + '>' + Category + '</option>' AS SelectItem FROM Menu.Categories ORDER BY CategorySort "
	Else
		sql = "SELECT '<option value=""' + CAST(CategoryID AS VARCHAR) + '"">' + Category + '</option>' AS SelectItem FROM Menu.Categories ORDER BY CategorySort "
	End If
	If Debug = True Then
		Response.Write(sql)
		Response.Flush
	End If
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & vbTab & rs.Fields("SelectItem") & vbCrLf)
		rs.MoveNext
	Wend
%></select></td>
	</tr>

	<tr>
		<td>Menu Text</td>
		<td><%=TextField("MenuText", Server.HTMLEncode(MenuText), 50, 127, PermitEdit, "")%></td>
	</tr>

	<tr>
		<td>Menu Description</td>
		<td><%=TextField("MenuDescription", MenuDescription, 80, 255, PermitEdit, "")%></td>
	</tr>

	<tr>
		<td>Item Sort</td>
		<td><%=IntegerField("ItemSort", ItemSort, 3, 3, PermitEdit, "")%></td>
	</tr>

	<tr>
		<td>Start Fiscal Year</td>
		<td><%=IntegerField("StartFiscalYear", StartFiscalYEar, 4, 4, PermitEdit, "")%></td>
	</tr>

	<tr>
		<td>End Fiscal Year</td>
		<td><%=IntegerField("EndFiscalYear", EndFiscalYEar, 4, 4, PermitEdit, "")%></td>
	</tr>

	<tr>
		<td>External Link</td>
		<td><%=TextField("ExternalLink", ExternalLink, 60, 255, PermitEdit, "")%></td>
	</tr>

	<tr>
		<td>Category and Link</td>
		<td><%=	CheckBoxField("CategoryAndLink",CategoryAndLink)%></td>
	</tr>

	<tr>
		<td>New Window</td>
		<td><%=	CheckBoxField("NewWindow",NewWindow)%></td>
	</tr>

	<tr>
		<td>Permission Level</td>
		<td><select name="PermissionLevelID"><%
	If PermissionLevelID>0 Then
		sql = "SELECT '<option value=""' + CAST(PermissionLevelID AS VARCHAR) + '""' + CASE WHEN PermissionLevelID=" & PermissionLevelID & " THEN ' SELECTED' ELSE '' END + '>' + PermissionLevelDescription + '</option>' AS SelectItem " & vbCrLf & _
			"FROM Menu.PermissionLevels " & vbCrLf & _
			"ORDER BY PermissionLevelID "
	Else
		sql = "SELECT '<option value=""' + CAST(PermissionLevelID AS VARCHAR) + '"">' + PermissionLevelDescription + '</option>' AS SelectItem " & vbCrLf & _
		"FROM Menu.PermissionLevels " & vbCrLf & _
		"ORDER BY PermissionLevelID "
	End If
	'Response.Write(sql)
	'Response.Flush
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & vbTab & rs.Fields("SelectItem") & vbCrLf)
		rs.MoveNext
	Wend
%></select></td>
	</tr>

	<tr>
		<td>Link ID reference</td>
		<td><select name="LinkID">
			<option value=""0"">none</option><%
	If PermissionLevelID>0 Then
		sql = "SELECT '<option value=""' + CAST(LinkID AS VARCHAR) + '""' + CASE WHEN LinkID=" & LinkID & " THEN ' SELECTED' ELSE '' END + '>' + LinkDescription + '</option>' AS SelectItem " & vbCrLf & _
			"FROM Menu.Links " & vbCrLf & _
			"ORDER BY LinkID "
	Else
		sql = "SELECT '<option value=""' + CAST(LinkID AS VARCHAR) + '"">' + LinkDescription + '</option>' AS SelectItem " & vbCrLf & _
			"FROM Menu.Links " & vbCrLf & _
			"ORDER BY LinkID "
	End If
	'Response.Write(sql)
	'Response.Flush
	Set rs = Con.Execute(sql)
	While rs.EOF = False
		Response.Write(vbTab & vbTab & rs.Fields("SelectItem") & vbCrLf)
		rs.MoveNext
	Wend
%></select></td>
	</tr>

	<tr>
		<td>Grantee Required</td>
		<td><%=	CheckBoxField("GranteeRequired",GranteeRequired)%></td>
	</tr>

	<tr>
		<td>Taskforce Grantee</td>
		<td><%=	CheckBoxField("TaskforceGrantee",TaskforceGrantee)%></td>
	</tr>

	<tr>
		<td>Auxiliary Grantee</td>
		<td><%=	CheckBoxField("AuxiliaryGrantee",AuxiliaryGrantee)%></td>
	</tr>

	<tr>
		<td>Catalytic Converter Grantee</td>
		<td><%=	CheckBoxField("CCGrantee",CCGrantee)%></td>
	</tr>

	<tr>
		<td>Grant Required</td>
		<td><%=	CheckBoxField("GrantRequired",GrantRequired)%></td>
	</tr>

	<tr>
		<td>ISA Required</td>
		<td><%=	CheckBoxField("ISARequired",ISARequired)%></td>
	</tr>

	<tr>
		<td>Application Required</td>
		<td><%=	CheckBoxField("AppRequired",AppRequired)%></td>
	</tr>

	<tr>
		<td>Negotiation Required</td>
		<td><%=	CheckBoxField("NegotiationRequired",NegotiationRequired)%></td>
	</tr>

	<tr>
		<td>MVCPA Auxilary Grant Required</td>
		<td><%=	CheckBoxField("MAGRequired",MAGRequired)%></td>
	</tr>

	<tr>
		<td>Rapid Response StrikeForce Grant Required</td>
		<td><%=	CheckBoxField("RRSRequired",RRSRequired)%></td>
	</tr>

	<tr>
		<td>Catalytic Converter Grant Required</td>
		<td><%=	CheckBoxField("CCRequired",CCRequired)%></td>
	</tr>

	<tr>
		<td>Grantee Link</td>
		<td><%=	CheckBoxField("GranteeLink",GranteeLink)%></td>
	</tr>

	<tr>
		<td>Grant Link</td>
		<td><%=	CheckBoxField("GrantLink",GrantLink)%></td>
	</tr>

	<tr>
		<td>ISA Link</td>
		<td><%=	CheckBoxField("ISALink",ISALink)%></td>
	</tr>

	<tr>
		<td>Application Link</td>
		<td><%=	CheckBoxField("AppLink",AppLink)%></td>
	</tr>

	<tr>
		<td>Negotiation Link</td>
		<td><%=	CheckBoxField("NegotiationLink",NegotiationLink)%></td>
	</tr>

	<tr>
		<td>Rapid Response Taskforce Link</td>
		<td><%=	CheckBoxField("RRSLink",RRSLink)%></td>
	</tr>

	<tr>
		<td>Catalytic Converter Taskforce Link</td>
		<td><%=	CheckBoxField("CCLink",CCLink)%></td>
	</tr>

	<tr>
		<td>Send Grantee ID</td>
		<td><%=	CheckBoxField("SendGranteeID",SendGranteeID)%></td>
	</tr>

	<tr>
		<td>Send Grant ID</td>
		<td><%=	CheckBoxField("SendGrantID",SendGrantID)%></td>
	</tr>

	<tr>
		<td>Send ISA ID</td>
		<td><%=	CheckBoxField("SendISAID",SendISAID)%></td>
	</tr>

	<tr>
		<td>Send Application ID</td>
		<td><%=	CheckBoxField("SendAppID",SendAppID)%></td>
	</tr>

	<tr>
		<td>Send Negotiation AppID</td>
		<td><%=	CheckBoxField("SendNegotiationID",SendNegotiationID)%></td>
	</tr>

	<tr>
		<td>Send MVCPA Auxiliary Grant MAGID</td>
		<td><%=	CheckBoxField("SendMAGID",SendMAGID)%></td>
	</tr>

	<tr>
		<td>Send Fiscal Year</td>
		<td><%=	CheckBoxField("SendFiscalYear",SendFiscalYear)%></td>
	</tr>

	<tr>
		<td>Inactive</td>
		<td><%=	CheckBoxField("Inactive",Inactive)%></td>
	</tr>

	<tr>
		<td>Notes</td>
		<td><%=TextField("Notes", Notes, 80, 1023, PermitEdit, "")%></td>
	</tr>

</table>
<div style="text-align: center; "><input type="submit" value="Submit" />
<input type="button" value="Return to List" onclick="location.href = 'MenuEditList.asp';" />
</div>
</form>

</body>
</html>
<!--#include file="../includes/prepWeb.asp"-->
<!--#include file="../includes/prepDB.asp"-->
<!--#include file="../includes/InputHelpers.asp"-->
