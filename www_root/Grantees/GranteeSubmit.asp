<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, TimeStamp, UpdateSystemID, GranteeID, GranteeName, OrganizationTypeID, ORI, StatePayeeIDNo, _	
	OrganizationalUnit, GeneralPhone, Address1, Address2, City, State, ZIP, _
	VendorOrganizationalUnit, VendorAddress1, VendorAddress2, VendorCity, VendorState, VendorZIP, _
	BorderCounty, PortCounty, Port2County, Inactive, GranteeNameSort, _
	TaskforceGrant, AuxiliaryGrant, RapidResponseStrikeforceGrant, CatalyticConverterGrant
debug = False
Timestamp = Now()
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

If Request.Form.Count>0 Then
	GranteeID = CInt(Request.Form("GranteeID"))
	GranteeName = Request.Form("GranteeName")
	OrganizationTypeID = Request.Form("OrganizationTypeID")
	ORI = Request.Form("ORI")
	StatePayeeIDNo = Request.Form("StatePayeeIDNo")
	OrganizationalUnit = Request.Form("OrganizationalUnit")
	GeneralPhone = Request.Form("GeneralPhone")
	Address1 = Request.Form("Address1")
	Address2 = Request.Form("Address2")
	City = Request.Form("City")
	State = Request.Form("State")
	ZIP = Request.Form("ZIP")
	VendorOrganizationalUnit = Request.Form("VendorOrganizationalUnit")
	VendorAddress1 = Request.Form("VendorAddress1")
	VendorAddress2 = Request.Form("VendorAddress2")
	VendorCity = Request.Form("VendorCity")
	VendorState = Request.Form("VendorState")
	VendorZIP = Request.Form("VendorZIP")
	BorderCounty = Request.Form("BorderCounty")
	PortCounty = Request.Form("PortCounty")
	Port2County = Request.Form("Port2County")
	Inactive = Request.Form("Inactive")
	GranteeNameSort = Request.Form("GranteeNameSort")
	If Len(GranteeNameSort) = 0 Then
		GranteeNameSort = GranteeName
	End If
	TaskforceGrant = Request.Form("TaskforceGrant")
	AuxiliaryGrant = Request.Form("AuxiliaryGrant")
	RapidResponseStrikeforceGrant = Request.Form("RapidResponseStrikeforceGrant")
	CatalyticConverterGrant = Request.Form("CatalyticConverterGrant")
End If

If GranteeID= 0 Then
	' Do an insert
	sql = "INSERT INTO Grantees (GranteeName, OrganizationTypeID, ORI, StatePayeeIDNo, " & vbCrLf & _
		"	OrganizationalUnit, GeneralPhone, Address1, Address2, City, State, zip, " & vbCrLf & _
		"	VendorOrganizationalUnit, VendorAddress1, VendorAddress2, VendorCity, " & vbCrLf & _
		"	VendorState, VendorZIP, BorderCounty, PortCounty, Port2County, " & vbCrLf & _
		"	GranteeNameSort, TaskforceGrant, AuxiliaryGrant, RapidResponseStrikeforceGrant, " & vbCrLf & _
		"	CatalyticConverterGrant, UpdateID, UpdateTimestamp) " & vbCrLf & _
		"OUTPUT Inserted.GranteeID " & vbCrLf & _
		"VALUES (" & vbCrLf & _
		prepStringSQL(GranteeName) & ", " & vbCrLf & _
		prepStringSQL(OrganizationTypeID) & ", " & vbCrLf & _
		prepStringSQL(ORI) & ", " & vbCrLf & _
		prepStringSQL(StatePayeeIDNo) & ", " & vbCrLf & _
		prepStringSQL(OrganizationalUnit) & ", " & vbCrLf & _
		prepStringSQL(GeneralPhone) & ", " & vbCrLf & _
		prepStringSQL(Address1) & ", " & vbCrLf & _
		prepStringSQL(Address2) & ", " & vbCrLf & _
		prepStringSQL(City) & ", " & vbCrLf & _
		prepStringSQL(State) & ", " & vbCrLf & _
		prepStringSQL(zip) & ", " & vbCrLf & _
		prepStringSQL(VendorOrganizationalUnit) & ", " & vbCrLf & _
		prepStringSQL(VendorAddress1) & ", " & vbCrLf & _
		prepStringSQL(VendorAddress2) & ", " & vbCrLf & _
		prepStringSQL(VendorCity) & ", " & vbCrLf & _
		prepStringSQL(VendorState) & ", " & vbCrLf & _
		prepStringSQL(VendorZIP) & ", " & vbCrLf & _
		prepBitSQL(BorderCounty) & ", " & vbCrLf & _
		prepBitSQL(PortCounty) & ", " & vbCrLf & _
		prepBitSQL(Port2County) & ", " & vbCrLf & _
		prepStringSQL(GranteeNameSort) & ", " & vbCrLf & _
		prepBitSQL(TaskforceGrant) & ", " & vbCrLf & _
		prepBitSQL(AuxiliaryGrant) & ", " & vbCrLf & _
		prepBitSQL(RapidResponseStrikeforceGrant) & ", " & vbCrLf & _
		prepBitSQL(CatalyticConverterGrant) & ", " & vbCrLf & _
		prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _ 
		prepStringSQL(Timestamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

	Set rs=Con.Execute(sql)
	If rs.EOF = True Then
		Response.Write("Error: No GranteeID Created.")
		Response.End
	Else 
		GranteeID = rs.Fields("GranteeID")
		Session("GranteeID") = GranteeID
	End If

	sql = "SELECT SystemID, GranteeID FROM System.GranteePermissions WHERE SystemID=" & _
		prepIntegerSQL(UpdateSystemID) & " AND GranteeID=" & prepIntegerSQL(GranteeID)
	If Debug = True Then
		Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Set rs = con.execute(sql)
	If rs.EOF = True Then
		sql = "INSERT INTO System.GranteePermissions(SystemID, GranteeID, UpdateID, UpdateTimeStamp) VALUES (" & _
			UserSystemID & ", " & prepIntegerSQL(GranteeID) & ", " & prepIntegerSQL(UserSystemID) & _
			", " & prepStringSQL(TimeStamp) & ")"
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)
	End If

	If MVCPARights = False Then
		sql = "UPDATE [System].Users SET DefaultGrantee=" & prepIntegerSQL(GranteeID) & " " & vbCrLF & _
			"WHERE SystemID=" & prepIntegerSQL(UserSystemID) & " AND (DefaultGrantee IS NULL OR DefaultGrantee=0) "
		If Debug = True Then
			Response.Write("<pre>" &sql & "</pre>" & vbCrLf)
			Response.Flush
		End If
		con.execute(sql)
	End If
Else
	' Do an update
	sql = "UPDATE Grantees " & vbCrLf & _
	"SET GranteeName=" & prepStringSQL(GranteeName) & ", " & vbCrLf & _
	"	OrganizationTypeID=" & prepIntegerSQL(OrganizationTypeID) & ", " & vbCrLf & _	
	"	ORI=" & prepStringSQL(ORI) & ", " & vbCrLf & _	
	"	StatePayeeIDNo=" & prepStringSQL(StatePayeeIDNo) & ", " & vbCrLf & _	
	"	OrganizationalUnit=" & prepStringSQL(OrganizationalUnit) & ", " & vbCrLf & _
	"	GeneralPhone=" & prepStringSQL(GeneralPhone) & ", " & vbCrLf & _
	"	Address1=" & prepStringSQL(Address1) & ", " & vbCrLf & _	
	"	Address2=" & prepStringSQL(Address2) & ", " & vbCrLf & _	
	"	City=" & prepStringSQL(City) & ", " & vbCrLf & _	
	"	State=" & prepStringSQL(State) & ", " & vbCrLf & _	
	"	ZIP=" & prepStringSQL(ZIP) & ", " & vbCrLf & _	
	"	VendorOrganizationalUnit=" & prepStringSQL(VendorOrganizationalUnit) & ", " & vbCrLf & _
	"	VendorAddress1=" & prepStringSQL(VendorAddress1) & ", " & vbCrLf & _
	"	VendorAddress2=" & prepStringSQL(VendorAddress2) & ", " & vbCrLf & _
	"	VendorCity=" & prepStringSQL(VendorCity) & ", " & vbCrLf & _
	"	VendorState=" & prepStringSQL(VendorState) & ", " & vbCrLf & _
	"	VendorZIP=" & prepStringSQL(VendorZIP) & ", " & vbCrLf & _
	"	BorderCounty=" & prepBitSQL(BorderCounty) & ", " & vbCrLf & _
	"	PortCounty=" & prepBitSQL(PortCounty) & ", " & vbCrLf & _
	"	Port2County=" & prepBitSQL(Port2County) & ", " & vbCrLf
If MVCPARights = True Then
	sql = sql & _
	"	GranteeNameSort=" & prepStringSQL(GranteeNameSort) & ", " & vbCrLf & _
	"	Inactive=" & prepBitSQL(Inactive) & ", " & vbCrLf & _
	"	TaskforceGrant=" & prepBitSQL(TaskforceGrant) & ", " & vbCrLf & _
	"	AuxiliaryGrant=" & prepBitSQL(AuxiliaryGrant) & ", " & vbCrLf & _
	"	RapidResponseStrikeforceGrant=" & prepBitSQL(RapidResponseStrikeforceGrant) & ", " & vbCrLf & _
	"	CatalyticConverterGrant=" & prepBitSQL(CatalyticConverterGrant) & ", " & vbCrLf
End If
	sql = sql & _
	"	UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & vbCrLf & _	
	"	UpdateTimestamp=" & prepStringSQL(Timestamp) & " " & vbCrLf & _	
	"WHERE GranteeID=" & prepIntegerSQL(GranteeID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If

	Con.Execute(sql)
End If


If Debug = True Then
	Response.Write("<a href=""../Home/Default.asp?GranteeID=" & GranteeID & """>Home</a>" & vbCrLf)
Else
	Response.Redirect("../Home/Default.asp?GranteeID=" & GranteeID)
End If

%><!--#include file="../includes/prepDB.asp"-->