<%@ language=VBScript %><% Option Explicit%><!--#include file="../includes/adovbs.asp"-->
<!--#include file="../includes/OpenConnection.asp"-->
<!--#include file="../includes/EnsureLogin.asp"--><% 
Dim debug, i, j, PermitEdit, GrantID, Quarter, EquipmentID, GranteeID, GranteeName, Submitted, _
	AssetClassID, ItemDescription, ModelYear, MakeManufacturer, Model, UseID, SerialNo, _
	SourceID, OwnerID, AcquisitionDate, Cost, MVCPAPercentOfCost, Location, ConditionID, _
	OtherInformation, UpdateID, UpdateTimestamp, TimeStamp
TimeStamp = Now()

debug = False
If Debug = True Then
	Response.Write("<pre>Dubugging Information: " & vbCrLF)
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
	Response.Flush
End If

GrantID = Request.Form("GrantID")
Quarter = Request.Form("Quarter")
EquipmentID = Request.Form("EquipmentID")
If Len(GrantID) = 0 Then
	Response.Write("Error: No GrantID provided.")
	sendWarning("Error: No GrantID provided.")
	Response.End
Else
	GrantID = CInt(GrantID)
End If
If Len(EquipmentID) = 0 Then
	Response.Write("Error: No EquipmentID provided.")
	sendWarning("Error: No EquipmentID provided.")
	Response.End
Else
	EquipmentID = CInt(EquipmentID)
End If


AssetClassID = Request.Form("AssetClassID")
ItemDescription = Request.Form("ItemDescription")
ModelYear = Request.Form("ModelYear")
MakeManufacturer = Request.Form("MakeManufacturer")
Model = Request.Form("Model")
UseID = Request.Form("UseID")
SerialNo = Request.Form("SerialNo")
SourceID = Request.Form("SourceID")
OwnerID = Request.Form("OwnerID")
AcquisitionDate = Request.Form("AcquisitionDate")
Cost = Request.Form("Cost")
MVCPAPercentOfCost = Request.Form("MVCPAPercentOfCost")
Location = Request.Form("Location")
ConditionID = Request.Form("ConditionID")
OtherInformation = Request.Form("OtherInformation")

If Debug = True Then
	Response.Write("<pre>Delete?" & vbCrLf)
	Response.Write("EquipmentID>0: " & (EquipmentID>0) & vbCrLf)
	Response.Write("AssetClassID=""0"": " & (AssetClassID="0") & vbCrLf)
	Response.Write("ItemDescription="""": " & (ItemDescription="") & vbCrLf)
	Response.Write("ModelYear="""": " & (ModelYear="") & vbCrLf)
	Response.Write("SerialNo="""": " & (SerialNo="") & vbCrLf)
	Response.Write("Cost="""": " & (Cost="") & vbCrLf)
	Response.Write("</pre>")
	Response.Flush
End If

If EquipmentID = 0 Then
	sql = "INSERT INTO ER.EquipmentDetail (GrantID, Quarter, AssetClassID, ItemDescription, " & vbCrLF & _
		"	ModelYear, MakeManufacturer, Model, UseID, SerialNo, SourceID, OwnerID, " & vbCrLF & _
		"	AcquisitionDate, Cost, MVCPAPercentOfCost, Location, ConditionID, " & vbCrLf & _
		"	OtherInformation, UpdateID, UpdateTimestamp) " & vbCrLF & _
	"VALUES (" & prepIntegerSQL(GrantID) & ", " & _
	prepIntegerSQL(Quarter) & ", " & _
	prepStringSQL(AssetClassID) & ", " & _
	prepStringSQL(ItemDescription) & ", " & _
	prepIntegerSQL(ModelYear) & ", " & _
	prepStringSQL(MakeManufacturer) & ", " & _
	prepStringSQL(Model) & ", " & _
	prepIntegerSQL(UseID) & ", " & _
	prepStringSQL(SerialNo) & ", " & _
	prepIntegerSQL(SourceID) & ", " & _
	prepIntegerSQL(OwnerID) & ", " & _
	prepDateSQL(AcquisitionDate) & ", " & _
	prepNumberSQL(Cost) & ", " & _
	prepIntegerSQL(MVCPAPercentOfCost) & ", " & _
	prepStringSQL(Location) & ", " & _
	prepIntegerSQL(ConditionID) & ", " & _
	prepStringSQL(OtherInformation) & ", " & _
	prepIntegerSQL(UserSystemID) & ", " & _
	prepStringSQL(TimeStamp) & ")"
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
ElseIf EquipmentID>0 AND AssetClassID="0" AND ItemDescription="" AND ModelYear="" AND SerialNo="" AND Cost="" Then
	sql = "DELETE FROM ER.EquipmentDetail WHERE EquipmentID=" & prepIntegerSQL(EquipmentID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
	EquipmentID=0
ElseIf EquipmentID>0 Then
	sql = "UPDATE ER.EquipmentDetail " & vbCrLF & _
		"SET GrantID=" & prepIntegerSQL(GrantID) & ", " & _
	"Quarter=" & prepIntegerSQL(Quarter) & ", " & _
	"AssetClassID=" & prepStringSQL(AssetClassID) & ", " & _
	"ItemDescription=" & prepStringSQL(ItemDescription) & ", " & _
	"ModelYear=" & prepIntegerSQL(ModelYear) & ", " & _
	"MakeManufacturer=" & prepStringSQL(MakeManufacturer) & ", " & _
	"Model=" & prepStringSQL(Model) & ", " & _
	"UseID=" & prepIntegerSQL(UseID) & ", " & _
	"SerialNo=" & prepStringSQL(SerialNo) & ", " & _
	"SourceID=" & prepIntegerSQL(SourceID) & ", " & _
	"OwnerID=" & prepIntegerSQL(OwnerID) & ", " & _
	"AcquisitionDate=" & prepDateSQL(AcquisitionDate) & ", " & _
	"Cost=" & prepNumberSQL(Cost) & ", " & _
	"MVCPAPercentOfCost=" & prepIntegerSQL(MVCPAPercentOfCost) & ", " & _
	"Location=" & prepStringSQL(Location) & ", " & _
	"ConditionID=" & prepIntegerSQL(ConditionID) & ", " & _
	"OtherInformation=" & prepStringSQL(OtherInformation) & ", " & _
	"UpdateID=" & prepIntegerSQL(UserSystemID) & ", " & _
	"UpdateTimestamp=" & prepStringSQL(TimeStamp) & " " & vbCrLF & _
	"WHERE EquipmentID=" & prepIntegerSQL(EquipmentID)
	If Debug = True Then
		Response.Write("<pre>" & sql & "</pre>" & vbCrLf)
		Response.Flush
	End If
	Con.Execute(sql)
End If

%><html>
<%	If Debug = True Then %>
<body>
<%	Else %>
<body onload="document.form.submit();">
<%	End If %>
<form name="form" action=Equipment.asp method="post">
<input type="hidden" name="GrantID" value="<%=GrantID %>">
<input type="hidden" name="Quarter" value="<%=Quarter %>">
<input type="hidden" name="EquipmentID" value="<%=EquipmentID %>">
<input type="submit" value="Sumbit" />
</form>
</body>
</html>
<!--#include file="../includes/prepDB.asp"-->
